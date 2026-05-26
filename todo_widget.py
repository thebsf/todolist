# -*- coding: utf-8 -*-
"""Low-presence Windows desktop widget for local and Microsoft To Do tasks."""

from __future__ import annotations

import json
import os
import queue
import sys
import threading
import uuid
import ctypes
import copy
from ctypes import wintypes
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import tkinter as tk
from tkinter import colorchooser, messagebox, ttk

try:
    import msal
except ImportError:
    msal = None

import requests

try:
    from PIL import Image, ImageOps, ImageStat
except ImportError:
    Image = ImageOps = ImageStat = None

try:
    import winreg
except ImportError:
    winreg = None


APP_NAME = "QuietTodoWidget"
APP_TITLE = "静默待办"
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["Tasks.ReadWrite"]
INSTANCE_MUTEX_NAME = "Local\\QuietTodoWidget.Singleton"

DATA_DIR = Path(os.getenv("APPDATA", Path.home())) / APP_NAME
SETTINGS_PATH = DATA_DIR / "settings.json"
TASKS_PATH = DATA_DIR / "local_tasks.json"
REMOTE_TASKS_PATH = DATA_DIR / "remote_tasks.json"
TOKEN_CACHE_PATH = DATA_DIR / "msal_cache.bin"

DEFAULT_SETTINGS = {
    "geometry": "356x460+120+140",
    "opacity": 0.76,
    "background": "#263136",
    "foreground": "#e5e9e8",
    "muted": "#a4aeac",
    "accent": "#9fb8b3",
    "auto_theme": True,
    "theme_preset": "wallpaper",
    "locked": False,
    "task_view": "pending",
    "task_order": [],
    "subtask_order": {},
    "collapsed_nodes": [],
    "remote_nested": {},
    "font_family": "Microsoft YaHei UI",
    "font_size": 10,
    "refresh_minutes": 5,
    "auto_start": False,
    "client_id": "",
    "tenant": "common",
    "todo_list_id": "",
    "todo_list_name": "",
}

THEME_PRESETS = {
    "wallpaper": {
        "label": "自动融入壁纸",
    },
    "dusk": {
        "label": "暮色灰蓝",
        "background": "#283238",
        "foreground": "#e1e6e5",
        "muted": "#98a4a2",
        "accent": "#596d70",
    },
    "sage": {
        "label": "雾松绿",
        "background": "#303a36",
        "foreground": "#e3e7e1",
        "muted": "#a1aaa1",
        "accent": "#68796e",
    },
    "sand": {
        "label": "暖砂岩",
        "background": "#403b35",
        "foreground": "#ebe5dc",
        "muted": "#aaa093",
        "accent": "#746958",
    },
    "ink": {
        "label": "墨夜蓝",
        "background": "#222b33",
        "foreground": "#dde4e9",
        "muted": "#929eaa",
        "accent": "#506171",
    },
    "custom": {
        "label": "自定义颜色",
    },
}
THEME_LABELS = [definition["label"] for definition in THEME_PRESETS.values()]
THEME_KEYS_BY_LABEL = {
    definition["label"]: key for key, definition in THEME_PRESETS.items()
}


class GraphError(RuntimeError):
    pass


class GraphAuthRequired(GraphError):
    pass


class SingleInstance:
    def __init__(self) -> None:
        self.handle = None

    def acquire(self) -> bool:
        if os.name != "nt":
            return True
        self.handle = ctypes.windll.kernel32.CreateMutexW(
            None, False, INSTANCE_MUTEX_NAME
        )
        return bool(self.handle) and ctypes.windll.kernel32.GetLastError() != 183

    def release(self) -> None:
        if self.handle:
            ctypes.windll.kernel32.CloseHandle(self.handle)
            self.handle = None


def ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def load_json(path: Path, fallback):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return fallback


def write_json(path: Path, value) -> None:
    ensure_data_dir()
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(
        json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    temp.replace(path)


class DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def dpapi_transform(data: bytes, protect: bool) -> bytes | None:
    """Encrypt or decrypt a token cache using the current Windows user context."""
    if os.name != "nt" or not data:
        return None
    buffer = ctypes.create_string_buffer(data)
    in_blob = DataBlob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
    out_blob = DataBlob()
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    function = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    if not function(
        ctypes.byref(in_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(out_blob),
    ):
        return None
    try:
        return ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        kernel32.LocalFree(out_blob.pbData)


def normalize_color(value: str, fallback: str) -> str:
    if isinstance(value, str) and len(value) == 7 and value.startswith("#"):
        try:
            int(value[1:], 16)
            return value
        except ValueError:
            pass
    return fallback


def normalize_settings(raw: dict | None) -> dict:
    raw = raw if isinstance(raw, dict) else {}
    settings = DEFAULT_SETTINGS.copy()
    settings.update(raw)
    settings.pop("topmost", None)
    try:
        settings["opacity"] = min(1.0, max(0.35, float(settings["opacity"])))
    except (TypeError, ValueError):
        settings["opacity"] = DEFAULT_SETTINGS["opacity"]
    try:
        settings["font_size"] = min(20, max(8, int(settings["font_size"])))
    except (TypeError, ValueError):
        settings["font_size"] = DEFAULT_SETTINGS["font_size"]
    try:
        settings["refresh_minutes"] = min(
            60, max(1, int(settings["refresh_minutes"]))
        )
    except (TypeError, ValueError):
        settings["refresh_minutes"] = DEFAULT_SETTINGS["refresh_minutes"]
    for key in ("background", "foreground", "muted", "accent"):
        settings[key] = normalize_color(settings[key], DEFAULT_SETTINGS[key])
    for key in ("auto_theme", "locked", "auto_start"):
        settings[key] = bool(settings[key])
    preset = str(settings.get("theme_preset", "")).strip()
    if preset not in THEME_PRESETS:
        preset = "wallpaper" if settings["auto_theme"] else "custom"
    settings["theme_preset"] = preset
    for key in ("geometry", "font_family", "client_id", "tenant", "todo_list_id"):
        settings[key] = str(settings.get(key, DEFAULT_SETTINGS[key])).strip()
    settings["todo_list_name"] = str(settings.get("todo_list_name", "")).strip()
    settings["tenant"] = settings["tenant"] or "common"
    settings["task_view"] = (
        settings.get("task_view") if settings.get("task_view") in ("pending", "completed") else "pending"
    )
    order = settings.get("task_order", [])
    settings["task_order"] = list(dict.fromkeys(str(item) for item in order)) if isinstance(order, list) else []
    collapsed = settings.get("collapsed_nodes", [])
    settings["collapsed_nodes"] = (
        list(dict.fromkeys(str(item) for item in collapsed))
        if isinstance(collapsed, list)
        else []
    )
    for key in ("subtask_order", "remote_nested"):
        settings[key] = settings[key] if isinstance(settings.get(key), dict) else {}
    return settings


def color_rgb(color: str) -> tuple[int, int, int]:
    return tuple(int(color[index : index + 2], 16) for index in (1, 3, 5))


def rgb_color(rgb: tuple[int, int, int]) -> str:
    return "#" + "".join(f"{min(255, max(0, int(value))):02x}" for value in rgb)


def blend(source: tuple[int, int, int], target: tuple[int, int, int], amount: float):
    return tuple(round(a * (1 - amount) + b * amount) for a, b in zip(source, target))


def luminance(rgb: tuple[int, int, int]) -> float:
    return 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2]


def local_task(title: str) -> dict:
    return {
        "id": uuid.uuid4().hex,
        "title": title.strip(),
        "completed": False,
        "pinned": False,
        "source": "local",
        "subtasks": [],
    }


def local_subtask(title: str) -> dict:
    return {
        "id": uuid.uuid4().hex,
        "title": title.strip(),
        "completed": False,
        "source": "local",
        "subtasks": [],
    }


class LocalTaskStore:
    def __init__(self, path: Path, root_source: str = "local"):
        self.path = path
        self.root_source = root_source

    def load(self) -> list[dict]:
        tasks = load_json(self.path, [])
        normalized = []
        if not isinstance(tasks, list):
            return normalized
        for raw in tasks:
            if not isinstance(raw, dict) or not str(raw.get("title", "")).strip():
                continue
            task = local_task(str(raw["title"]))
            task["id"] = str(raw.get("id", task["id"]))
            task["completed"] = bool(raw.get("completed", False))
            task["pinned"] = bool(raw.get("pinned", False))
            task["source"] = str(raw.get("source", self.root_source))
            task["subtasks"] = self.normalize_subtasks(raw.get("subtasks", []), depth=1)
            normalized.append(task)
        return normalized

    def normalize_subtasks(self, raw_items, depth: int) -> list[dict]:
        normalized = []
        if depth > 4 or not isinstance(raw_items, list):
            return normalized
        for raw in raw_items:
            if not isinstance(raw, dict) or not str(raw.get("title", "")).strip():
                continue
            item = local_subtask(str(raw["title"]))
            item["id"] = str(raw.get("id", item["id"]))
            item["completed"] = bool(raw.get("completed", False))
            item["source"] = str(raw.get("source", "local"))
            item["subtasks"] = self.normalize_subtasks(raw.get("subtasks", []), depth + 1)
            normalized.append(item)
        return normalized

    def save(self, tasks: list[dict]) -> None:
        write_json(self.path, tasks)


class GraphTodoClient:
    def __init__(self, client_id: str, tenant: str):
        self.client_id = client_id.strip()
        self.tenant = tenant.strip() or "common"
        self.cache = None
        self.app = None
        if self.client_id and msal is not None:
            ensure_data_dir()
            self.cache = msal.SerializableTokenCache()
            if TOKEN_CACHE_PATH.exists():
                try:
                    decrypted = dpapi_transform(TOKEN_CACHE_PATH.read_bytes(), protect=False)
                    if decrypted:
                        self.cache.deserialize(decrypted.decode("utf-8"))
                except OSError:
                    pass
            authority = f"https://login.microsoftonline.com/{self.tenant}"
            self.app = msal.PublicClientApplication(
                self.client_id, authority=authority, token_cache=self.cache
            )

    @property
    def configured(self) -> bool:
        return bool(self.client_id) and self.app is not None

    @property
    def missing_dependency(self) -> bool:
        return bool(self.client_id) and msal is None

    @property
    def has_account(self) -> bool:
        return bool(self.app and self.app.get_accounts())

    def _save_cache(self) -> None:
        if self.cache and self.cache.has_state_changed:
            protected = dpapi_transform(self.cache.serialize().encode("utf-8"), protect=True)
            if protected:
                TOKEN_CACHE_PATH.write_bytes(protected)

    def login(self) -> str:
        if not self.client_id:
            raise GraphError("请先填写 Microsoft 应用 Client ID")
        if msal is None:
            raise GraphError("当前环境未安装 msal，请重新构建或执行 pip install msal")
        result = self.app.acquire_token_interactive(
            scopes=GRAPH_SCOPES, prompt="select_account"
        )
        self._save_cache()
        if "access_token" not in result:
            raise GraphError(result.get("error_description", "Microsoft 登录失败"))
        claims = result.get("id_token_claims", {})
        return claims.get("name") or claims.get("preferred_username") or "已登录"

    def logout(self) -> None:
        if self.app:
            for account in self.app.get_accounts():
                self.app.remove_account(account)
        if TOKEN_CACHE_PATH.exists():
            TOKEN_CACHE_PATH.unlink(missing_ok=True)

    def _token(self) -> str:
        if not self.configured:
            raise GraphAuthRequired("尚未配置 Microsoft To Do")
        accounts = self.app.get_accounts()
        if not accounts:
            raise GraphAuthRequired("请先在设置中登录 Microsoft To Do")
        result = self.app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
        self._save_cache()
        if not result or "access_token" not in result:
            raise GraphAuthRequired("Microsoft 登录已过期，请重新登录")
        return result["access_token"]

    def request(self, method: str, path: str, payload: dict | None = None):
        headers = {"Authorization": f"Bearer {self._token()}"}
        if payload is not None:
            headers["Content-Type"] = "application/json"
        try:
            response = requests.request(
                method,
                GRAPH_ROOT + path,
                headers=headers,
                json=payload,
                timeout=20,
            )
        except requests.RequestException as exc:
            raise GraphError(f"连接 Microsoft To Do 失败: {exc}") from exc
        if response.status_code == 401:
            raise GraphAuthRequired("Microsoft 登录已过期，请重新登录")
        if response.status_code >= 400:
            try:
                detail = response.json()["error"]["message"]
            except (ValueError, KeyError, TypeError):
                detail = response.text[:180]
            raise GraphError(f"Microsoft To Do 请求失败 ({response.status_code}): {detail}")
        if response.status_code == 204:
            return None
        return response.json()

    def list_task_lists(self) -> list[dict]:
        return self.request("GET", "/me/todo/lists").get("value", [])

    def list_tasks(self, list_id: str) -> list[dict]:
        encoded_list = quote(list_id, safe="")
        data = self.request("GET", f"/me/todo/lists/{encoded_list}/tasks?$top=100")
        tasks = []
        while data:
            for item in data.get("value", []):
                encoded_task = quote(item["id"], safe="")
                child_data = self.request(
                    "GET",
                    f"/me/todo/lists/{encoded_list}/tasks/{encoded_task}/checklistItems",
                )
                tasks.append(
                    {
                        "id": item["id"],
                        "title": item.get("title", "未命名任务"),
                        "completed": item.get("status") == "completed",
                        "pinned": item.get("importance") == "high",
                        "source": "todo",
                        "subtasks": [
                            {
                                "id": child["id"],
                                "title": child.get("displayName", "未命名子任务"),
                                "completed": bool(child.get("isChecked", False)),
                                "source": "todo_subtask",
                                "subtasks": [],
                            }
                            for child in child_data.get("value", [])
                        ],
                    }
                )
            next_url = data.get("@odata.nextLink")
            if not next_url:
                break
            path = next_url.replace(GRAPH_ROOT, "", 1)
            data = self.request("GET", path)
        return tasks

    def create_task(self, list_id: str, title: str) -> dict:
        encoded_list = quote(list_id, safe="")
        return self.request("POST", f"/me/todo/lists/{encoded_list}/tasks", {"title": title})

    def update_task_completed(self, list_id: str, task_id: str, value: bool) -> None:
        self.request(
            "PATCH",
            f"/me/todo/lists/{quote(list_id, safe='')}/tasks/{quote(task_id, safe='')}",
            {"status": "completed" if value else "notStarted"},
        )

    def update_task_pinned(self, list_id: str, task_id: str, value: bool) -> None:
        self.request(
            "PATCH",
            f"/me/todo/lists/{quote(list_id, safe='')}/tasks/{quote(task_id, safe='')}",
            {"importance": "high" if value else "normal"},
        )

    def delete_task(self, list_id: str, task_id: str) -> None:
        self.request(
            "DELETE",
            f"/me/todo/lists/{quote(list_id, safe='')}/tasks/{quote(task_id, safe='')}",
        )

    def create_subtask(self, list_id: str, task_id: str, title: str) -> None:
        self.request(
            "POST",
            (
                f"/me/todo/lists/{quote(list_id, safe='')}/tasks/"
                f"{quote(task_id, safe='')}/checklistItems"
            ),
            {"displayName": title},
        )

    def update_subtask_completed(
        self, list_id: str, task_id: str, item_id: str, value: bool
    ) -> None:
        self.request(
            "PATCH",
            (
                f"/me/todo/lists/{quote(list_id, safe='')}/tasks/"
                f"{quote(task_id, safe='')}/checklistItems/{quote(item_id, safe='')}"
            ),
            {"isChecked": value},
        )

    def delete_subtask(self, list_id: str, task_id: str, item_id: str) -> None:
        self.request(
            "DELETE",
            (
                f"/me/todo/lists/{quote(list_id, safe='')}/tasks/"
                f"{quote(task_id, safe='')}/checklistItems/{quote(item_id, safe='')}"
            ),
        )


class QuietTodoWidget:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.settings = normalize_settings(load_json(SETTINGS_PATH, {}))
        self.local_store = LocalTaskStore(TASKS_PATH)
        self.remote_store = LocalTaskStore(REMOTE_TASKS_PATH, root_source="todo")
        self.local_tasks = self.local_store.load()
        self.remote_tasks: list[dict] = self.remote_store.load()
        self.todo_lists: list[dict] = []
        self.graph = GraphTodoClient(self.settings["client_id"], self.settings["tenant"])
        self.settings_window: tk.Toplevel | None = None
        self.status_text = tk.StringVar(value="本地待办")
        self.input_text = tk.StringVar()
        self.subtask_draft_text = tk.StringVar()
        self.subtask_draft_parent_key: str | None = None
        self.root_draft_entry: tk.Entry | None = None
        self.root_draft_frame: tk.Frame | None = None
        self.root_draft_circle: tk.Canvas | None = None
        self.subtask_draft_entry: tk.Entry | None = None
        self.refresh_job = None
        self.theme_job = None
        self.bottom_job = None
        self.last_wallpaper_signature = None
        self.drag_origin = None
        self.resize_origin = None
        self.node_drag_key = None
        self.node_drag_parent_key = None
        self.node_drag_target = None
        self.node_drag_reparent = False
        self.drag_rows: list[tuple[str, str, tk.Widget]] = []
        self.worker_results: queue.Queue = queue.Queue()
        self.applied_opacity = None
        self.appearance_signature = None

        self._configure_window()
        self._build_window()
        self.apply_appearance(refresh_theme=True)
        self.render_tasks()
        self.root.after(120, self._poll_results)
        self.schedule_sync()
        self.schedule_theme_refresh()
        self.root.after(100, self.keep_at_desktop_bottom)
        if self.graph.has_account and self.settings["todo_list_id"]:
            self.sync_remote()

    def _configure_window(self) -> None:
        self.root.title(APP_TITLE)
        self.root.geometry(self.settings["geometry"])
        self.root.minsize(280, 260)
        self.root.overrideredirect(True)
        try:
            self.root.attributes("-toolwindow", True)
        except tk.TclError:
            pass
        self.root.attributes("-topmost", False)
        self.root.attributes("-alpha", self.settings["opacity"])
        self.root.protocol("WM_DELETE_WINDOW", self.close)

    def _build_window(self) -> None:
        self.panel = tk.Frame(self.root, bd=0, highlightthickness=0)
        self.panel.pack(fill="both", expand=True)

        self.controls = tk.Frame(self.panel, bd=0, highlightthickness=0)
        self.controls.pack(fill="x", padx=9, pady=(4, 2))
        self.pending_button = tk.Button(
            self.controls,
            text="待办",
            command=lambda: self.set_task_view("pending"),
            relief="flat",
            bd=0,
        )
        self.pending_button.pack(side="left", padx=(0, 4))
        self.completed_button = tk.Button(
            self.controls,
            text="已完成",
            command=lambda: self.set_task_view("completed"),
            relief="flat",
            bd=0,
        )
        self.completed_button.pack(side="left")
        self.lock_button = tk.Button(
            self.controls, command=self.toggle_lock, relief="flat", bd=0
        )
        self.lock_button.pack(side="right", padx=(5, 0))
        self.sync_button = tk.Button(
            self.controls, text="↻", command=self.request_sync, relief="flat", bd=0
        )
        self.sync_button.pack(side="right", padx=(5, 0))
        self.collapse_all_button = tk.Button(
            self.controls,
            text="\u25b8\u25b8",
            command=self.toggle_all_collapsed,
            relief="flat",
            bd=0,
        )
        self.collapse_all_button.pack(side="right", padx=(5, 0))
        self.settings_button = tk.Button(
            self.controls, text="···", command=self.open_settings, relief="flat", bd=0
        )
        self.settings_button.pack(side="right")
        self.controls.bind("<ButtonPress-1>", self.start_drag)
        self.controls.bind("<B1-Motion>", self.drag_window)
        self.controls.bind("<ButtonRelease-1>", self.stop_drag)

        self.canvas = tk.Canvas(self.panel, bd=0, highlightthickness=0)
        self.tasks_frame = tk.Frame(self.canvas, bd=0, highlightthickness=0)
        self.tasks_window = self.canvas.create_window(
            (0, 0), window=self.tasks_frame, anchor="nw"
        )
        self.canvas.configure(yscrollcommand=self.update_scroll_thumb)
        self.canvas.pack(side="top", fill="both", expand=True, padx=(10, 4))
        self.scroll_track = tk.Canvas(
            self.panel, width=6, bd=0, highlightthickness=0, cursor="sb_v_double_arrow"
        )
        self.scroll_thumb = self.scroll_track.create_rectangle(2, 0, 5, 0, outline="")
        self.scroll_track.place(relx=1, rely=0.12, relheight=0.68, x=-4, anchor="ne")
        self.scroll_track_drag_y = None
        self.scroll_track.bind("<ButtonPress-1>", self.start_scroll_thumb_drag)
        self.scroll_track.bind("<B1-Motion>", self.drag_scroll_thumb)
        self.scroll_track.bind("<ButtonRelease-1>", self.stop_scroll_thumb_drag)
        self.tasks_frame.bind("<Configure>", self.update_scroll_region)
        self.canvas.bind("<Configure>", self.resize_task_canvas)
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)

        self.grip = tk.Label(self.panel, text="◢", anchor="se", cursor="size_nw_se")
        self.grip.place(relx=1, rely=1, anchor="se", x=-3, y=-2)
        self.grip.bind("<ButtonPress-1>", self.start_resize)
        self.grip.bind("<B1-Motion>", self.resize_window)
        self.grip.bind("<ButtonRelease-1>", lambda _event: self.save_geometry())

    def font(self, extra: int = 0, weight: str = "normal") -> tuple:
        return (
            self.settings["font_family"],
            self.settings["font_size"] + extra,
            weight,
        )

    def button_style(self, button: tk.Button, accent: bool = False) -> None:
        bg = self.settings["accent"] if accent else self.settings["background"]
        fg = self.settings["background"] if accent else self.settings["muted"]
        button.configure(
            bg=bg,
            fg=fg,
            disabledforeground=fg,
            activebackground=self.settings["accent"],
            activeforeground=self.settings["background"],
            font=self.font(),
            highlightthickness=0,
            cursor="hand2",
        )

    def completion_button(
        self, parent: tk.Widget, completed: bool, command, small: bool = False
    ) -> tk.Canvas:
        bg = self.settings["background"]
        size = max(12, self.settings["font_size"] + (2 if small else 3))
        control = tk.Canvas(
            parent,
            width=size,
            height=size,
            bg=bg,
            bd=0,
            highlightthickness=0,
            cursor="hand2" if command else "",
        )
        inset = 1
        outline = self.settings["accent"] if completed else self.settings["muted"]
        fill = self.settings["accent"] if completed else bg
        control.create_oval(
            inset,
            inset,
            size - inset,
            size - inset,
            outline=outline,
            fill=fill,
            width=1,
        )
        if completed:
            control.create_line(
                size * 0.27,
                size * 0.53,
                size * 0.44,
                size * 0.69,
                size * 0.75,
                size * 0.34,
                fill=bg,
                width=1,
                smooth=True,
            )
        if command:
            control.bind("<Button-1>", lambda _event: command())
        control.pack(side="left", padx=(2, 5))
        return control

    def draft_circle(self, parent: tk.Widget, small: bool = False) -> tk.Canvas:
        bg = self.settings["background"]
        size = max(12, self.settings["font_size"] + (2 if small else 3))
        control = tk.Canvas(
            parent,
            width=size,
            height=size,
            bg=bg,
            bd=0,
            highlightthickness=0,
        )
        control.pack(side="left", padx=(2, 5))
        return control

    def paint_draft_circle(self, control: tk.Canvas, visible: bool) -> None:
        control.delete("draft_circle")
        if not visible:
            return
        size = max(1, int(control.cget("width")))
        control.create_oval(
            1,
            1,
            size - 1,
            size - 1,
            outline=self.settings["muted"],
            fill=self.settings["background"],
            width=1,
            tags="draft_circle",
        )

    def show_root_draft_circle(self, _event=None) -> None:
        if self.root_draft_circle and self.root_draft_circle.winfo_exists():
            self.paint_draft_circle(self.root_draft_circle, True)

    def defer_hide_root_draft_circle(self, _event=None) -> None:
        self.root.after_idle(self.hide_root_draft_circle_if_outside)

    def hide_root_draft_circle_if_outside(self) -> None:
        frame = self.root_draft_frame
        circle = self.root_draft_circle
        if not frame or not circle or not frame.winfo_exists() or not circle.winfo_exists():
            return
        pointer_x = frame.winfo_pointerx()
        pointer_y = frame.winfo_pointery()
        inside = (
            frame.winfo_rootx() <= pointer_x < frame.winfo_rootx() + frame.winfo_width()
            and frame.winfo_rooty() <= pointer_y < frame.winfo_rooty() + frame.winfo_height()
        )
        if not inside:
            self.paint_draft_circle(circle, False)

    def disclosure_slot(
        self, parent: tk.Widget, key: str, has_children: bool, small: bool = False
    ) -> tk.Frame:
        bg = self.settings["background"]
        slot = tk.Frame(parent, bg=bg, width=14, height=20)
        slot.pack(side="left", padx=(0, 3))
        slot.pack_propagate(False)
        if has_children:
            collapse = tk.Button(
                slot,
                text="\u25b8" if self.is_collapsed(key) else "\u25be",
                command=lambda k=key: self.toggle_collapsed(k),
                relief="flat",
                bd=0,
                padx=0,
                pady=0,
            )
            collapse.pack(fill="both", expand=True)
            self.button_style(collapse)
            if small:
                collapse.configure(font=self.font(-1))
            if self.settings["locked"]:
                collapse.configure(state="disabled")
        return slot

    def apply_appearance(self, refresh_theme: bool = False) -> None:
        if refresh_theme and self.settings["auto_theme"]:
            palette = self.wallpaper_palette()
            if palette:
                self.settings.update(palette)
        bg = self.settings["background"]
        muted = self.settings["muted"]
        signature = (
            bg,
            self.settings["foreground"],
            muted,
            self.settings["accent"],
            self.settings["opacity"],
            self.settings["font_family"],
            self.settings["font_size"],
            self.settings["task_view"],
            self.settings["locked"],
        )
        if self.appearance_signature == signature:
            return
        self.appearance_signature = signature
        if self.applied_opacity != self.settings["opacity"]:
            self.root.attributes("-alpha", self.settings["opacity"])
            self.applied_opacity = self.settings["opacity"]
        self.root.attributes("-topmost", False)
        self.root.configure(bg=bg)
        for widget in (
            self.panel,
            self.controls,
            self.canvas,
            self.tasks_frame,
            self.scroll_track,
        ):
            widget.configure(bg=bg)
        self.grip.configure(bg=bg, fg=muted, font=self.font(-1))
        self.scroll_track.itemconfigure(self.scroll_thumb, fill=self.settings["muted"])
        self.button_style(self.settings_button)
        self.button_style(self.sync_button)
        self.button_style(self.collapse_all_button)
        self.button_style(self.lock_button)
        self.button_style(self.pending_button, accent=self.settings["task_view"] == "pending")
        self.button_style(
            self.completed_button, accent=self.settings["task_view"] == "completed"
        )
        self.update_locked_layout()
        self.send_to_bottom()
        write_json(SETTINGS_PATH, self.settings)
        if hasattr(self, "tasks_frame"):
            self.render_tasks()

    def update_scroll_region(self, _event=None) -> None:
        content_height = max(1, self.tasks_frame.winfo_reqheight())
        viewport_height = max(1, self.canvas.winfo_height())
        viewport_width = max(1, self.canvas.winfo_width())
        self.canvas.configure(
            scrollregion=(0, 0, viewport_width, max(content_height, viewport_height))
        )
        if content_height <= viewport_height:
            self.canvas.yview_moveto(0)
        self.update_scroll_thumb(*self.canvas.yview())

    def resize_task_canvas(self, event) -> None:
        self.canvas.itemconfigure(self.tasks_window, width=event.width)
        self.update_scroll_region()

    def on_mousewheel(self, event) -> None:
        if not self.root.winfo_exists():
            return
        left = self.canvas.winfo_rootx()
        top = self.canvas.winfo_rooty()
        right = left + self.canvas.winfo_width()
        bottom = top + self.canvas.winfo_height()
        if not (left <= event.x_root <= right and top <= event.y_root <= bottom):
            return
        if self.settings["locked"]:
            return "break"
        content_height = self.tasks_frame.winfo_reqheight()
        viewport_height = self.canvas.winfo_height()
        if content_height <= viewport_height:
            self.canvas.yview_moveto(0)
            return "break"
        direction = int(-event.delta / 120)
        first, last = self.canvas.yview()
        if (direction < 0 and first <= 0.0) or (direction > 0 and last >= 1.0):
            return "break"
        self.canvas.yview_scroll(direction, "units")
        return "break"

    def update_scroll_thumb(self, first: str | float, last: str | float) -> None:
        first, last = float(first), float(last)
        if last - first >= 0.999:
            self.scroll_track.place_forget()
            return
        if not self.scroll_track.winfo_manager():
            self.scroll_track.place(
                relx=1, rely=0.12, relheight=0.68, x=-4, anchor="ne"
            )
        height = max(1, self.scroll_track.winfo_height())
        thumb_top = round(first * height)
        thumb_bottom = max(thumb_top + 14, round(last * height))
        thumb_bottom = min(height, thumb_bottom)
        self.scroll_track.coords(self.scroll_thumb, 2, thumb_top, 5, thumb_bottom)

    def start_scroll_thumb_drag(self, event) -> None:
        if self.settings["locked"]:
            return
        self.scroll_track_drag_y = event.y

    def drag_scroll_thumb(self, event) -> None:
        if self.settings["locked"]:
            return
        if self.scroll_track_drag_y is None:
            return
        delta = event.y - self.scroll_track_drag_y
        height = max(1, self.scroll_track.winfo_height())
        self.canvas.yview_moveto(max(0.0, self.canvas.yview()[0] + delta / height))
        self.scroll_track_drag_y = event.y

    def stop_scroll_thumb_drag(self, _event=None) -> None:
        self.scroll_track_drag_y = None

    def start_drag(self, event) -> None:
        if self.settings["locked"]:
            return
        self.drag_origin = (
            event.x_root,
            event.y_root,
            self.root.winfo_x(),
            self.root.winfo_y(),
        )

    def drag_window(self, event) -> None:
        if not self.drag_origin:
            return
        start_x, start_y, window_x, window_y = self.drag_origin
        self.root.geometry(f"+{window_x + event.x_root - start_x}+{window_y + event.y_root - start_y}")

    def stop_drag(self, _event=None) -> None:
        self.drag_origin = None
        self.save_geometry()
        if self.settings["auto_theme"]:
            self.apply_appearance(refresh_theme=True)

    def start_resize(self, event) -> None:
        if self.settings["locked"]:
            return
        self.resize_origin = (
            event.x_root,
            event.y_root,
            self.root.winfo_width(),
            self.root.winfo_height(),
        )

    def resize_window(self, event) -> None:
        if not self.resize_origin:
            return
        start_x, start_y, width, height = self.resize_origin
        new_width = max(280, width + event.x_root - start_x)
        new_height = max(260, height + event.y_root - start_y)
        self.root.geometry(f"{new_width}x{new_height}")

    def save_geometry(self) -> None:
        self.resize_origin = None
        self.settings["geometry"] = self.root.geometry()
        write_json(SETTINGS_PATH, self.settings)

    def toggle_lock(self) -> None:
        self.settings["locked"] = not self.settings["locked"]
        self.node_drag_key = None
        self.node_drag_parent_key = None
        self.node_drag_target = None
        self.node_drag_reparent = False
        self.scroll_track_drag_y = None
        write_json(SETTINGS_PATH, self.settings)
        self.update_locked_layout()
        self.render_tasks()
        self.send_to_bottom()

    def update_locked_layout(self) -> None:
        self.lock_button.configure(
            text="\ue785" if self.settings["locked"] else "\ue72e",
            font=("Segoe MDL2 Assets", max(7, self.settings["font_size"] - 2)),
        )
        for button in (
            self.pending_button,
            self.completed_button,
            self.settings_button,
            self.sync_button,
            self.collapse_all_button,
        ):
            button.configure(state="normal")

    def set_task_view(self, view: str) -> None:
        if view not in ("pending", "completed"):
            return
        if view == "completed":
            self.input_text.set("")
            self.subtask_draft_text.set("")
            self.subtask_draft_parent_key = None
        self.settings["task_view"] = view
        write_json(SETTINGS_PATH, self.settings)
        self.button_style(self.pending_button, accent=view == "pending")
        self.button_style(self.completed_button, accent=view == "completed")
        self.render_tasks()

    def send_to_bottom(self) -> None:
        if not self.root.winfo_exists():
            return
        self.root.update_idletasks()
        if os.name == "nt":
            hwnd_bottom = 1
            flags = 0x0001 | 0x0002 | 0x0010 | 0x0200
            ctypes.windll.user32.SetWindowPos(
                self.root.winfo_id(), hwnd_bottom, 0, 0, 0, 0, flags
            )
        else:
            self.root.lower()

    def keep_at_desktop_bottom(self) -> None:
        if os.name == "nt" and ctypes.windll.user32.IsIconic(self.root.winfo_id()):
            ctypes.windll.user32.ShowWindow(self.root.winfo_id(), 4)  # SW_SHOWNOACTIVATE
            self.root.update_idletasks()
            self.send_to_bottom()
        if self.bottom_job:
            self.root.after_cancel(self.bottom_job)
        self.bottom_job = self.root.after(1000, self.keep_at_desktop_bottom)

    def display_tasks(self) -> list[dict]:
        tasks = self.local_tasks + self.remote_tasks
        self.ensure_task_order(tasks)
        completed = self.settings["task_view"] == "completed"
        shown = [task for task in tasks if task["completed"] == completed]
        position = {key: index for index, key in enumerate(self.settings["task_order"])}
        return sorted(
            shown,
            key=lambda task: (
                not task.get("pinned", False),
                position.get(self.task_key(task), len(position)),
            ),
        )

    def task_key(self, task: dict) -> str:
        return f"{task['source']}:{task['id']}"

    def child_key(self, parent_key: str, item: dict) -> str:
        return f"{parent_key}/{item.get('source', 'local')}:{item['id']}"

    def find_node(self, wanted_key: str) -> dict | None:
        def find_children(root_task: dict, parent: dict, parent_key: str, depth: int):
            for item in parent.get("subtasks", []):
                key = self.child_key(parent_key, item)
                record = {
                    "key": key,
                    "parent_key": parent_key,
                    "node": item,
                    "parent": parent,
                    "root": root_task,
                    "depth": depth,
                }
                if key == wanted_key:
                    return record
                found = find_children(root_task, item, key, depth + 1)
                if found:
                    return found
            return None

        for task in self.local_tasks + self.remote_tasks:
            key = self.task_key(task)
            if key == wanted_key:
                return {
                    "key": key,
                    "parent_key": "root",
                    "node": task,
                    "parent": None,
                    "root": task,
                    "depth": 0,
                }
            found = find_children(task, task, key, 1)
            if found:
                return found
        return None

    def subtree_height(self, node: dict) -> int:
        children = node.get("subtasks", [])
        if not children:
            return 1
        return 1 + max(self.subtree_height(child) for child in children)

    def node_key_mapping(self, node: dict, old_key: str, new_key: str) -> dict[str, str]:
        mapping = {old_key: new_key}
        for child in node.get("subtasks", []):
            old_child = self.child_key(old_key, child)
            new_child = self.child_key(new_key, child)
            mapping.update(self.node_key_mapping(child, old_child, new_child))
        return mapping

    def remap_saved_node_keys(self, mapping: dict[str, str]) -> None:
        self.settings["collapsed_nodes"] = [
            mapping.get(key, key) for key in self.settings["collapsed_nodes"]
        ]
        updated_orders = {}
        for parent_key, order in self.settings["subtask_order"].items():
            mapped_parent = mapping.get(parent_key, parent_key)
            mapped_order = [mapping.get(key, key) for key in order]
            combined = updated_orders.setdefault(mapped_parent, [])
            for key in mapped_order:
                if key not in combined:
                    combined.append(key)
        self.settings["subtask_order"] = updated_orders

    def is_collapsed(self, key: str) -> bool:
        return key in self.settings["collapsed_nodes"]

    def toggle_collapsed(self, key: str) -> None:
        if self.settings["locked"]:
            return
        collapsed = self.settings["collapsed_nodes"]
        if key in collapsed:
            collapsed.remove(key)
        else:
            collapsed.append(key)
        write_json(SETTINGS_PATH, self.settings)
        self.render_tasks()

    def collapsible_keys(self, tasks: list[dict] | None = None) -> list[str]:
        keys = []

        def collect(parent: dict, parent_key: str) -> None:
            for child in parent.get("subtasks", []):
                child_key = self.child_key(parent_key, child)
                if child.get("subtasks"):
                    keys.append(child_key)
                collect(child, child_key)

        for task in tasks if tasks is not None else self.display_tasks():
            task_key = self.task_key(task)
            if task.get("subtasks"):
                keys.append(task_key)
            collect(task, task_key)
        return keys

    def update_collapse_all_button(self, tasks: list[dict] | None = None) -> None:
        keys = self.collapsible_keys(tasks)
        all_collapsed = bool(keys) and all(
            key in self.settings["collapsed_nodes"] for key in keys
        )
        self.collapse_all_button.configure(
            text="\u25be\u25be" if all_collapsed else "\u25b8\u25b8",
            state="normal" if keys else "disabled",
            disabledforeground=self.settings["muted"],
        )

    def toggle_all_collapsed(self) -> None:
        keys = self.collapsible_keys()
        if not keys:
            return
        collapsed = self.settings["collapsed_nodes"]
        if all(key in collapsed for key in keys):
            self.settings["collapsed_nodes"] = [
                key for key in collapsed if key not in keys
            ]
        else:
            self.settings["collapsed_nodes"] = list(
                dict.fromkeys([*collapsed, *keys])
            )
        write_json(SETTINGS_PATH, self.settings)
        self.render_tasks()

    def ensure_task_order(self, tasks: list[dict]) -> None:
        order = self.settings["task_order"]
        changed = False
        for task in tasks:
            key = self.task_key(task)
            if key not in order:
                order.append(key)
                changed = True
        if changed:
            write_json(SETTINGS_PATH, self.settings)

    def ordered_subtasks(self, parent: dict, parent_key: str) -> list[dict]:
        children = parent.get("subtasks", [])
        order = self.settings["subtask_order"].setdefault(parent_key, [])
        changed = False
        for child in children:
            key = self.child_key(parent_key, child)
            if key not in order:
                order.append(key)
                changed = True
        if changed:
            write_json(SETTINGS_PATH, self.settings)
        position = {key: index for index, key in enumerate(order)}
        return sorted(
            children,
            key=lambda child: position.get(self.child_key(parent_key, child), len(position)),
        )

    def merge_remote_nested(self, tasks: list[dict]) -> list[dict]:
        for task in tasks:
            task_key = self.task_key(task)
            for item in task.get("subtasks", []):
                key = self.child_key(task_key, item)
                item["subtasks"] = copy.deepcopy(self.settings["remote_nested"].get(key, []))
        return tasks

    def save_nested_changes(self, root_task: dict) -> None:
        if root_task["source"] == "local":
            self.local_store.save(self.local_tasks)
            return
        root_key = self.task_key(root_task)
        for child in root_task.get("subtasks", []):
            key = self.child_key(root_key, child)
            self.settings["remote_nested"][key] = child.get("subtasks", [])
        write_json(SETTINGS_PATH, self.settings)
        self.remote_store.save(self.remote_tasks)

    def save_all_task_trees(self) -> None:
        self.local_store.save(self.local_tasks)
        for task in self.remote_tasks:
            root_key = self.task_key(task)
            for child in task.get("subtasks", []):
                if child.get("source") == "todo_subtask":
                    child_key = self.child_key(root_key, child)
                    self.settings["remote_nested"][child_key] = child.get("subtasks", [])
        write_json(SETTINGS_PATH, self.settings)
        self.remote_store.save(self.remote_tasks)

    def put_task_first(self, key: str) -> None:
        self.settings["task_order"] = [
            key,
            *[existing for existing in self.settings["task_order"] if existing != key],
        ]
        write_json(SETTINGS_PATH, self.settings)

    def render_tasks(self) -> None:
        previous_position = self.canvas.yview()[0]
        self.set_redraw(False)
        try:
            for child in self.tasks_frame.winfo_children():
                child.destroy()
            self.drag_rows = []
            self.root_draft_entry = None
            self.root_draft_frame = None
            self.root_draft_circle = None
            self.subtask_draft_entry = None
            tasks = self.display_tasks()
            if not tasks:
                empty = (
                    "没有已完成任务"
                    if self.settings["task_view"] == "completed"
                    else "输入一项任务开始"
                )
                tk.Label(
                    self.tasks_frame,
                    text=empty,
                    bg=self.settings["background"],
                    fg=self.settings["muted"],
                    font=self.font(),
                    anchor="w",
                ).pack(fill="x", padx=4, pady=12)
            else:
                for task in tasks:
                    self.render_task_row(task)
            if self.settings["task_view"] == "pending":
                self.render_root_draft_row()
            self.update_collapse_all_button(tasks)
            self.root.update_idletasks()
            self.update_scroll_region()
            content_height = self.tasks_frame.winfo_reqheight()
            viewport_height = max(1, self.canvas.winfo_height())
            if content_height > viewport_height:
                max_position = max(0.0, 1.0 - viewport_height / content_height)
                self.canvas.yview_moveto(min(previous_position, max_position))
        finally:
            self.set_redraw(True)

    def set_redraw(self, enabled: bool) -> None:
        if os.name != "nt" or not self.root.winfo_exists():
            return
        user32 = ctypes.windll.user32
        user32.SendMessageW(self.root.winfo_id(), 0x000B, int(enabled), 0)  # WM_SETREDRAW
        if enabled:
            user32.RedrawWindow(
                self.root.winfo_id(), None, None, 0x0001 | 0x0080 | 0x0100
            )

    def render_task_row(self, task: dict) -> None:
        bg = self.settings["background"]
        fg = self.settings["muted"] if task["completed"] else self.settings["foreground"]
        group = tk.Frame(
            self.tasks_frame,
            bg=bg,
            bd=0,
            highlightthickness=1,
            highlightbackground=bg,
        )
        group.pack(fill="x", pady=(1, 3))
        key = self.task_key(task)
        self.drag_rows.append((key, "root", group))
        row = tk.Frame(group, bg=bg, bd=0)
        row.pack(fill="x")
        handle = tk.Label(
            row,
            text="\u2261",
            bg=bg,
            fg=self.settings["muted"],
            font=self.font(),
            cursor="sb_v_double_arrow",
        )
        handle.pack(side="left", padx=(0, 3))
        handle.bind(
            "<ButtonPress-1>", lambda event, k=key: self.start_node_drag(k, "root", event)
        )
        handle.bind("<B1-Motion>", self.drag_node)
        handle.bind("<ButtonRelease-1>", self.stop_node_drag)
        self.completion_button(
            row,
            task["completed"],
            None
            if self.settings["locked"]
            else lambda t=task: self.toggle_task(t, not t["completed"]),
        )
        pin = tk.Button(
            row,
            text="\u2605" if task.get("pinned", False) else "\u2606",
            command=lambda t=task: self.toggle_pin(t),
            relief="flat",
            bd=0,
        )
        pin.pack(side="left", padx=(0, 3))
        self.button_style(pin)
        if self.settings["locked"]:
            pin.configure(state="disabled")
        self.disclosure_slot(row, key, bool(task.get("subtasks")))
        title = tk.Label(
            row,
            text=task["title"],
            bg=bg,
            fg=fg,
            font=self.font(),
            justify="left",
            anchor="w",
            wraplength=max(130, self.root.winfo_width() - 136),
        )
        title.pack(side="left", fill="x", expand=True, pady=2)
        delete = tk.Button(
            row,
            text="\u00d7",
            command=lambda t=task: self.delete_task(t),
            relief="flat",
            bd=0,
        )
        delete.pack(side="right")
        add_sub = None
        if self.settings["task_view"] == "pending":
            add_sub = tk.Button(
                row,
                text="\uff0b",
                command=lambda t=task, k=key: self.add_subtask(t, t, k, 0),
                relief="flat",
                bd=0,
            )
            add_sub.pack(side="right", padx=(2, 0))
            self.button_style(add_sub)
        self.button_style(delete)
        if self.settings["locked"]:
            if add_sub:
                add_sub.configure(state="disabled")
            delete.configure(state="disabled")
        if not self.is_collapsed(key):
            for item in self.ordered_subtasks(task, key):
                self.render_subtask_row(group, task, task, item, key, 1)
            if (
                self.settings["task_view"] == "pending"
                and self.subtask_draft_parent_key == key
            ):
                self.render_subtask_draft_row(group, key, 1)

    def render_subtask_row(
        self,
        group: tk.Frame,
        root_task: dict,
        parent: dict,
        item: dict,
        parent_key: str,
        depth: int,
    ) -> None:
        bg = self.settings["background"]
        fg = self.settings["muted"] if item["completed"] else self.settings["foreground"]
        row = tk.Frame(group, bg=bg, bd=0)
        row.pack(fill="x", padx=(19 + (depth * 10), 0), pady=1)
        item_key = self.child_key(parent_key, item)
        self.drag_rows.append((item_key, parent_key, row))
        handle = tk.Label(
            row,
            text="\u2261",
            bg=bg,
            fg=self.settings["muted"],
            font=self.font(-1),
            cursor="sb_v_double_arrow",
        )
        handle.pack(side="left", padx=(0, 3))
        handle.bind(
            "<ButtonPress-1>",
            lambda event, k=item_key, p=parent_key: self.start_node_drag(k, p, event),
        )
        handle.bind("<B1-Motion>", self.drag_node)
        handle.bind("<ButtonRelease-1>", self.stop_node_drag)
        self.completion_button(
            row,
            item["completed"],
            None
            if self.settings["locked"]
            else lambda r=root_task, i=item: self.toggle_subtask(
                r, i, not i["completed"]
            ),
            small=True,
        )
        self.disclosure_slot(row, item_key, bool(item.get("subtasks")), small=True)
        tk.Label(
            row,
            text=item["title"],
            bg=bg,
            fg=fg,
            font=self.font(-1),
            anchor="w",
        ).pack(side="left", fill="x", expand=True)
        delete = tk.Button(
            row,
            text="\u00d7",
            command=lambda r=root_task, p=parent, i=item: self.delete_subtask(r, p, i),
            relief="flat",
            bd=0,
        )
        delete.pack(side="right")
        add_sub = None
        if depth < 4 and self.settings["task_view"] == "pending":
            add_sub = tk.Button(
                row,
                text="\uff0b",
                command=lambda r=root_task, p=item, k=item_key, d=depth: self.add_subtask(
                    r, p, k, d
                ),
                relief="flat",
                bd=0,
            )
            add_sub.pack(side="right", padx=(2, 0))
            self.button_style(add_sub)
        self.button_style(delete)
        if self.settings["locked"]:
            delete.configure(state="disabled")
            if add_sub:
                add_sub.configure(state="disabled")
        if not self.is_collapsed(item_key):
            for child in self.ordered_subtasks(item, item_key):
                self.render_subtask_row(
                    group, root_task, item, child, item_key, depth + 1
                )
            if (
                self.settings["task_view"] == "pending"
                and self.subtask_draft_parent_key == item_key
            ):
                self.render_subtask_draft_row(group, item_key, depth + 1)

    def draft_entry(self, parent: tk.Widget, variable: tk.StringVar) -> tk.Entry:
        locked = self.settings["locked"]
        entry = tk.Entry(
            parent,
            textvariable=variable,
            relief="flat",
            bd=0,
            bg=self.settings["background"],
            fg=self.settings["foreground"],
            insertbackground=self.settings["foreground"],
            font=self.font(-1),
            insertwidth=1,
            highlightthickness=0,
            state="disabled" if locked else "normal",
            disabledbackground=self.settings["background"],
            disabledforeground=self.settings["muted"],
        )
        return entry

    def render_root_draft_row(self) -> None:
        bg = self.settings["background"]
        frame = tk.Frame(self.tasks_frame, bg=bg, bd=0)
        frame.pack(fill="x", pady=(5, 9))
        spacer = tk.Label(frame, text="\u2261", bg=bg, fg=bg, font=self.font())
        spacer.pack(side="left", padx=(0, 3))
        circle = self.draft_circle(frame)
        entry = self.draft_entry(frame, self.input_text)
        entry.configure(font=self.font())
        entry.pack(side="left", fill="x", expand=True, ipady=4, padx=(0, 8))
        entry.bind("<Return>", lambda _event: self.add_task())
        entry.bind("<Escape>", lambda _event: self.input_text.set(""))
        for widget in (frame, spacer, circle, entry):
            widget.bind("<Enter>", self.show_root_draft_circle, add="+")
            widget.bind("<Leave>", self.defer_hide_root_draft_circle, add="+")
        self.root_draft_frame = frame
        self.root_draft_circle = circle
        self.root_draft_entry = entry

    def render_subtask_draft_row(
        self, group: tk.Frame, parent_key: str, depth: int
    ) -> None:
        bg = self.settings["background"]
        frame = tk.Frame(group, bg=bg, bd=0)
        frame.pack(fill="x", padx=(19 + (depth * 10), 6), pady=(2, 4))
        spacer = tk.Label(frame, text="\u2261", bg=bg, fg=bg, font=self.font(-1))
        spacer.pack(side="left", padx=(0, 3))
        self.completion_button(frame, False, None, small=True)
        self.disclosure_slot(frame, parent_key, False, small=True)
        entry = self.draft_entry(frame, self.subtask_draft_text)
        entry.pack(side="left", fill="x", expand=True, ipady=3)
        entry.bind("<Return>", lambda _event: self.commit_subtask_draft())
        entry.bind("<Escape>", lambda _event: self.cancel_subtask_draft())
        entry.bind(
            "<FocusOut>",
            lambda _event, current=entry, key=parent_key: self.defer_cancel_empty_subtask_draft(
                current, key
            ),
        )
        self.subtask_draft_entry = entry
        if not self.settings["locked"]:
            self.root.after_idle(
                lambda item=entry: item.focus_set() if item.winfo_exists() else None
            )

    def toggle_pin(self, task: dict) -> None:
        if self.settings["locked"]:
            return
        pinned = not task.get("pinned", False)
        if task["source"] == "todo":
            task["pinned"] = pinned
            self.remote_store.save(self.remote_tasks)
            self.render_tasks()
            self.run_async(
                lambda: self.graph.update_task_pinned(
                    self.settings["todo_list_id"], task["id"], pinned
                ),
                lambda _result: self.sync_remote(),
                "正在更新置顶状态...",
            )
        else:
            task["pinned"] = pinned
            self.local_store.save(self.local_tasks)
            self.render_tasks()

    def start_node_drag(self, key: str, parent_key: str, _event=None) -> None:
        if self.settings["locked"]:
            return
        self.node_drag_key = key
        self.node_drag_parent_key = parent_key
        self.node_drag_target = None
        self.node_drag_reparent = False

    def drag_node(self, _event=None) -> None:
        if self.settings["locked"]:
            return
        if not self.node_drag_key:
            return
        pointer_y = self.tasks_frame.winfo_pointery()
        self.node_drag_target = None
        siblings = [
            (key, widget)
            for key, parent_key, widget in self.drag_rows
            if parent_key == self.node_drag_parent_key and key != self.node_drag_key
        ]
        for _key, _parent_key, widget in self.drag_rows:
            if hasattr(widget, "configure"):
                widget.configure(highlightbackground=self.settings["background"])
        hovered = []
        for key, parent_key, widget in self.drag_rows:
            if key == self.node_drag_key:
                continue
            top = widget.winfo_rooty()
            if top <= pointer_y <= top + widget.winfo_height():
                hovered.append((key, parent_key, widget))
        if hovered:
            key, hovered_parent_key, widget = hovered[-1]
            self.node_drag_target = key
            moving_record = self.find_node(self.node_drag_key)
            target_record = self.find_node(key)
            shifted_right = (
                self.tasks_frame.winfo_pointerx() > widget.winfo_rootx() + 46
            )
            self.node_drag_reparent = bool(
                moving_record
                and moving_record["depth"] > 0
                and target_record
                and (
                    hovered_parent_key != self.node_drag_parent_key
                    or shifted_right
                )
            )
            widget.configure(highlightthickness=1, highlightbackground=self.settings["accent"])
        if siblings and not self.node_drag_target:
            first_key, first_widget = siblings[0]
            last_key, last_widget = siblings[-1]
            if pointer_y < first_widget.winfo_rooty():
                self.node_drag_target = first_key
                first_widget.configure(
                    highlightthickness=1, highlightbackground=self.settings["accent"]
                )
            elif pointer_y > last_widget.winfo_rooty() + last_widget.winfo_height():
                self.node_drag_target = last_key
                last_widget.configure(
                    highlightthickness=1, highlightbackground=self.settings["accent"]
                )

    def stop_node_drag(self, _event=None) -> None:
        if self.settings["locked"]:
            return
        moving = self.node_drag_key
        parent_key = self.node_drag_parent_key
        target = self.node_drag_target
        reparent = self.node_drag_reparent
        self.node_drag_key = None
        self.node_drag_parent_key = None
        self.node_drag_target = None
        self.node_drag_reparent = False
        for _key, _parent, widget in self.drag_rows:
            widget.configure(highlightbackground=self.settings["background"])
        if not moving or not target or moving == target:
            return
        moving_record = self.find_node(moving)
        target_record = self.find_node(target)
        if not moving_record or not target_record:
            return
        if target_record["parent_key"] != parent_key or reparent:
            self.move_subtask_to_parent(moving_record, target_record)
            return
        storage = (
            self.settings["task_order"]
            if parent_key == "root"
            else self.settings["subtask_order"].setdefault(parent_key, [])
        )
        order = [key for key in storage if key != moving]
        target_index = order.index(target)
        target_widget = next(widget for key, parent, widget in self.drag_rows if key == target and parent == parent_key)
        if self.tasks_frame.winfo_pointery() > target_widget.winfo_rooty() + target_widget.winfo_height() / 2:
            target_index += 1
        order.insert(target_index, moving)
        if parent_key == "root":
            self.settings["task_order"] = order
        else:
            self.settings["subtask_order"][parent_key] = order
        write_json(SETTINGS_PATH, self.settings)
        self.render_tasks()

    def move_subtask_to_parent(self, moving: dict, target: dict) -> None:
        if self.settings["locked"]:
            return
        node = moving["node"]
        if moving["depth"] == 0:
            return
        if node.get("source") == "todo_subtask":
            self.status_text.set("To Do 一级子任务只能在原主任务内排序")
            return
        if target["depth"] + self.subtree_height(node) > 4:
            self.status_text.set("最多支持四级子任务，无法移动到这里")
            return
        if target["node"] is node or self.find_descendant(node, target["node"]):
            self.status_text.set("不能把子任务移动到自身下方")
            return
        if target["depth"] == 0 and target["root"].get("source") == "todo":
            self.status_text.set("To Do 主任务的一级子任务需通过同步创建")
            return
        old_key = moving["key"]
        new_key = self.child_key(target["key"], node)
        mapping = self.node_key_mapping(node, old_key, new_key)
        moving["parent"]["subtasks"].remove(node)
        target["node"].setdefault("subtasks", []).append(node)
        self.remap_saved_node_keys(mapping)
        origin_order = self.settings["subtask_order"].setdefault(moving["parent_key"], [])
        self.settings["subtask_order"][moving["parent_key"]] = [
            key for key in origin_order if key != new_key and key != old_key
        ]
        destination_order = self.settings["subtask_order"].setdefault(target["key"], [])
        if new_key not in destination_order:
            destination_order.append(new_key)
        if target["key"] in self.settings["collapsed_nodes"]:
            self.settings["collapsed_nodes"].remove(target["key"])
        self.save_all_task_trees()
        self.status_text.set("已移动子任务")
        self.render_tasks()

    def find_descendant(self, node: dict, candidate: dict) -> bool:
        for child in node.get("subtasks", []):
            if child is candidate or self.find_descendant(child, candidate):
                return True
        return False

    def using_remote_input(self) -> bool:
        return bool(
            self.graph.has_account
            and self.settings["todo_list_id"]
            and self.graph.configured
        )

    def add_task(self) -> None:
        if self.settings["locked"]:
            return
        title = self.input_text.get().strip()
        if not title:
            return
        self.input_text.set("")
        if self.using_remote_input():
            def added(task):
                if task and task.get("id"):
                    self.put_task_first(f"todo:{task['id']}")
                self.sync_remote()

            self.run_async(
                lambda: self.graph.create_task(self.settings["todo_list_id"], title),
                added,
                "正在添加到 Microsoft To Do...",
            )
        else:
            task = local_task(title)
            self.local_tasks.insert(0, task)
            self.put_task_first(self.task_key(task))
            self.local_store.save(self.local_tasks)
            self.status_text.set("已添加本地待办")
            self.render_tasks()

    def add_subtask(
        self, root_task: dict, parent: dict, parent_key: str, depth: int
    ) -> None:
        if self.settings["locked"]:
            return
        if depth >= 4:
            return
        self.subtask_draft_parent_key = parent_key
        self.subtask_draft_text.set("")
        if parent_key in self.settings["collapsed_nodes"]:
            self.settings["collapsed_nodes"].remove(parent_key)
            write_json(SETTINGS_PATH, self.settings)
        self.render_tasks()

    def cancel_subtask_draft(self) -> None:
        self.subtask_draft_parent_key = None
        self.subtask_draft_text.set("")
        self.render_tasks()

    def defer_cancel_empty_subtask_draft(
        self, entry: tk.Entry, parent_key: str
    ) -> None:
        self.root.after_idle(
            lambda: self.cancel_empty_subtask_draft(entry, parent_key)
        )

    def cancel_empty_subtask_draft(self, entry: tk.Entry, parent_key: str) -> None:
        if (
            self.subtask_draft_entry is entry
            and self.subtask_draft_parent_key == parent_key
            and not self.subtask_draft_text.get().strip()
        ):
            self.cancel_subtask_draft()

    def commit_subtask_draft(self) -> None:
        if self.settings["locked"]:
            return
        title = self.subtask_draft_text.get().strip()
        parent_key = self.subtask_draft_parent_key
        if not title or not parent_key:
            return
        parent_record = self.find_node(parent_key)
        if not parent_record or parent_record["depth"] >= 4:
            self.cancel_subtask_draft()
            return
        self.subtask_draft_parent_key = None
        self.subtask_draft_text.set("")
        root_task = parent_record["root"]
        parent = parent_record["node"]
        if root_task["source"] == "todo" and parent_record["depth"] == 0:
            self.render_tasks()
            self.run_async(
                lambda: self.graph.create_subtask(
                    self.settings["todo_list_id"], root_task["id"], title
                ),
                lambda _result: self.sync_remote(),
                "正在添加子任务...",
            )
        else:
            parent.setdefault("subtasks", []).append(local_subtask(title))
            self.save_nested_changes(root_task)
            self.render_tasks()

    def toggle_task(self, task: dict, completed: bool) -> None:
        if self.settings["locked"]:
            return
        if task["source"] == "todo":
            task["completed"] = completed
            self.remote_store.save(self.remote_tasks)
            self.render_tasks()
            self.run_async(
                lambda: self.graph.update_task_completed(
                    self.settings["todo_list_id"], task["id"], completed
                ),
                lambda _result: self.sync_remote(),
                "正在更新任务...",
            )
        else:
            task["completed"] = completed
            self.local_store.save(self.local_tasks)
            self.render_tasks()

    def toggle_subtask(self, root_task: dict, item: dict, completed: bool) -> None:
        if self.settings["locked"]:
            return
        if root_task["source"] == "todo" and item.get("source") == "todo_subtask":
            item["completed"] = completed
            self.save_nested_changes(root_task)
            self.render_tasks()
            self.run_async(
                lambda: self.graph.update_subtask_completed(
                    self.settings["todo_list_id"], root_task["id"], item["id"], completed
                ),
                lambda _result: self.sync_remote(),
                "正在更新子任务...",
            )
        else:
            item["completed"] = completed
            self.save_nested_changes(root_task)
            self.render_tasks()

    def delete_task(self, task: dict) -> None:
        if self.settings["locked"]:
            return
        if not messagebox.askyesno("删除任务", f"删除“{task['title']}”？", parent=self.root):
            return
        if task["source"] == "todo":
            self.run_async(
                lambda: self.graph.delete_task(self.settings["todo_list_id"], task["id"]),
                lambda _result: self.sync_remote(),
                "正在删除任务...",
            )
        else:
            self.local_tasks.remove(task)
            self.local_store.save(self.local_tasks)
            self.render_tasks()

    def delete_subtask(self, root_task: dict, parent: dict, item: dict) -> None:
        if self.settings["locked"]:
            return
        if root_task["source"] == "todo" and item.get("source") == "todo_subtask":
            self.run_async(
                lambda: self.graph.delete_subtask(
                    self.settings["todo_list_id"], root_task["id"], item["id"]
                ),
                lambda _result: self.sync_remote(),
                "正在删除子任务...",
            )
        else:
            parent["subtasks"].remove(item)
            self.save_nested_changes(root_task)
            self.render_tasks()

    def run_async(self, function, on_success=None, status: str | None = None) -> None:
        if status:
            self.status_text.set(status)

        def worker():
            try:
                result = function()
                self.worker_results.put(("ok", result, on_success))
            except Exception as exc:  # Keep UI responsive and report remote failures.
                self.worker_results.put(("error", exc, None))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_results(self) -> None:
        try:
            while True:
                kind, payload, callback = self.worker_results.get_nowait()
                if kind == "ok":
                    if callback:
                        callback(payload)
                elif isinstance(payload, GraphAuthRequired):
                    self.status_text.set(str(payload))
                else:
                    self.status_text.set("操作失败")
                    messagebox.showerror("操作失败", str(payload), parent=self.root)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_results)

    def sync_remote(self) -> None:
        if not self.settings["todo_list_id"] or not self.graph.configured:
            self.status_text.set("本地待办；可在设置中连接 Microsoft To Do")
            return
        if not self.graph.has_account:
            self.status_text.set("请在设置中登录 Microsoft To Do")
            return

        def completed(tasks):
            self.remote_tasks = self.merge_remote_nested(tasks)
            self.remote_store.save(self.remote_tasks)
            now = datetime.now().strftime("%H:%M")
            self.status_text.set(f"Microsoft To Do 已同步 {now}")
            self.render_tasks()

        self.run_async(
            lambda: self.graph.list_tasks(self.settings["todo_list_id"]),
            completed,
            "正在同步 Microsoft To Do...",
        )

    def request_sync(self) -> None:
        self.sync_remote()

    def schedule_sync(self) -> None:
        if self.refresh_job:
            self.root.after_cancel(self.refresh_job)
        delay = self.settings["refresh_minutes"] * 60 * 1000
        self.refresh_job = self.root.after(delay, self._periodic_sync)

    def _periodic_sync(self) -> None:
        self.sync_remote()
        self.schedule_sync()

    def open_settings(self) -> None:
        if self.settings_window and self.settings_window.winfo_exists():
            self.settings_window.lift()
            return
        window = tk.Toplevel(self.root)
        self.settings_window = window
        window.title("静默待办设置")
        window.geometry("470x660")
        window.minsize(440, 560)
        window.configure(bg="#f4f5f4")

        footer = tk.Frame(window, bg="#f4f5f4")
        footer.pack(side="bottom", fill="x", padx=18, pady=(4, 14))
        body = tk.Frame(window, bg="#f4f5f4")
        body.pack(fill="both", expand=True, padx=(18, 10), pady=(14, 0))
        settings_canvas = tk.Canvas(
            body, bg="#f4f5f4", bd=0, highlightthickness=0
        )
        settings_scroll = ttk.Scrollbar(
            body, orient="vertical", command=settings_canvas.yview
        )
        settings_canvas.configure(yscrollcommand=settings_scroll.set)
        settings_scroll.pack(side="right", fill="y")
        settings_canvas.pack(side="left", fill="both", expand=True)
        container = tk.Frame(settings_canvas, bg="#f4f5f4")
        container_window = settings_canvas.create_window(
            (0, 0), window=container, anchor="nw"
        )
        self.settings_canvas = settings_canvas
        self.settings_content = container

        def update_settings_region(_event=None):
            settings_canvas.configure(scrollregion=settings_canvas.bbox("all"))

        def resize_settings_content(event):
            content_width = min(event.width, 420)
            offset = max(0, (event.width - content_width) // 2)
            settings_canvas.coords(container_window, offset, 0)
            settings_canvas.itemconfigure(container_window, width=content_width)
            update_settings_region()

        def scroll_settings(event):
            settings_canvas.yview_scroll(int(-event.delta / 120), "units")
            return "break"

        container.bind("<Configure>", update_settings_region)
        settings_canvas.bind("<Configure>", resize_settings_content)
        window.bind("<MouseWheel>", scroll_settings)
        vars_ = {
            "auto_theme": tk.BooleanVar(value=self.settings["auto_theme"]),
            "theme_preset": tk.StringVar(
                value=THEME_PRESETS[self.settings["theme_preset"]]["label"]
            ),
            "opacity": tk.DoubleVar(value=self.settings["opacity"]),
            "font_family": tk.StringVar(value=self.settings["font_family"]),
            "font_size": tk.IntVar(value=self.settings["font_size"]),
            "auto_start": tk.BooleanVar(value=self.settings["auto_start"]),
            "refresh_minutes": tk.IntVar(value=self.settings["refresh_minutes"]),
            "background": tk.StringVar(value=self.settings["background"]),
            "foreground": tk.StringVar(value=self.settings["foreground"]),
            "muted": tk.StringVar(value=self.settings["muted"]),
            "accent": tk.StringVar(value=self.settings["accent"]),
            "client_id": tk.StringVar(value=self.settings["client_id"]),
            "tenant": tk.StringVar(value=self.settings["tenant"]),
            "todo_list": tk.StringVar(value=self.settings["todo_list_name"]),
        }
        window.vars_ = vars_

        def section(text):
            tk.Label(
                container,
                text=text,
                bg="#f4f5f4",
                fg="#33413e",
                font=("Microsoft YaHei UI", 10, "bold"),
                anchor="w",
            ).pack(fill="x", pady=(9, 4))

        def check(text, name):
            tk.Checkbutton(
                container,
                text=text,
                variable=vars_[name],
                bg="#f4f5f4",
                activebackground="#f4f5f4",
                font=("Microsoft YaHei UI", 9),
            ).pack(anchor="w")

        section("外观")
        theme_row = tk.Frame(container, bg="#f4f5f4")
        theme_row.pack(fill="x", pady=(0, 4))
        tk.Label(theme_row, text="颜色主题", bg="#f4f5f4", width=12, anchor="w").pack(
            side="left"
        )
        theme_combo = ttk.Combobox(
            theme_row,
            textvariable=vars_["theme_preset"],
            values=THEME_LABELS,
            state="readonly",
            width=22,
        )
        theme_combo.pack(side="left")
        preview = tk.Canvas(container, height=17, bd=0, highlightthickness=0)
        preview.pack(fill="x", pady=(0, 6))
        theme_combo.bind(
            "<<ComboboxSelected>>",
            lambda _event: self.choose_theme_preset(vars_, preview),
        )
        self.update_theme_preview(vars_, preview)
        tk.Label(container, text="窗口透明度", bg="#f4f5f4").pack(anchor="w")
        tk.Scale(
            container,
            variable=vars_["opacity"],
            from_=0.35,
            to=1,
            resolution=0.01,
            orient="horizontal",
            bg="#f4f5f4",
            highlightthickness=0,
        ).pack(fill="x")
        color_row = tk.Frame(container, bg="#f4f5f4")
        color_row.pack(fill="x", pady=(3, 4))
        for text, key in (
            ("背景色", "background"),
            ("文字色", "foreground"),
            ("辅助色", "muted"),
            ("强调色", "accent"),
        ):
            tk.Button(
                color_row,
                text=text,
                command=lambda k=key: self.choose_color(vars_, k, preview),
            ).pack(side="left", padx=(0, 7))

        font_row = tk.Frame(container, bg="#f4f5f4")
        font_row.pack(fill="x", pady=3)
        ttk.Combobox(
            font_row,
            textvariable=vars_["font_family"],
            values=("Microsoft YaHei UI", "Segoe UI", "微软雅黑", "Arial"),
            state="readonly",
            width=22,
        ).pack(side="left")
        tk.Spinbox(font_row, from_=8, to=20, textvariable=vars_["font_size"], width=5).pack(
            side="left", padx=10
        )

        section("行为")
        tk.Label(
            container,
            text="完成的任务自动进入主窗口的“已完成”清单。\n窗口固定在桌面底层，不会覆盖其他应用窗口。",
            bg="#f4f5f4",
            fg="#5b6664",
            anchor="w",
            justify="left",
        ).pack(fill="x")
        check("登录 Windows 后自动启动", "auto_start")
        refresh_row = tk.Frame(container, bg="#f4f5f4")
        refresh_row.pack(fill="x", pady=3)
        tk.Label(refresh_row, text="To Do 自动同步间隔（分钟）", bg="#f4f5f4").pack(
            side="left"
        )
        tk.Spinbox(
            refresh_row, from_=1, to=60, textvariable=vars_["refresh_minutes"], width=5
        ).pack(side="left", padx=8)

        section("Microsoft To Do 联动")
        tk.Label(
            container,
            text="需要在 Microsoft Entra 中注册桌面应用并授予 Tasks.ReadWrite 权限。",
            bg="#f4f5f4",
            fg="#5b6664",
            anchor="w",
        ).pack(fill="x")
        self.labeled_entry(container, "Client ID", vars_["client_id"])
        self.labeled_entry(container, "Tenant（个人账号用 common）", vars_["tenant"])
        auth_row = tk.Frame(container, bg="#f4f5f4")
        auth_row.pack(fill="x", pady=5)
        tk.Button(
            auth_row, text="登录并加载列表", command=lambda: self.login_from_settings(vars_)
        ).pack(side="left")
        tk.Button(auth_row, text="退出登录", command=self.logout_graph).pack(
            side="left", padx=8
        )
        self.list_combo = ttk.Combobox(
            container, textvariable=vars_["todo_list"], state="readonly"
        )
        self.list_combo.pack(fill="x", pady=(2, 3))
        self.list_combo.bind(
            "<<ComboboxSelected>>", lambda _event: self.choose_todo_list(vars_)
        )
        self.fill_list_combo(vars_)

        tk.Button(
            footer,
            text="保存并应用",
            command=lambda: self.save_settings_window(vars_, window),
        ).pack(side="right")
        tk.Button(footer, text="关闭组件", command=self.close).pack(side="left")

    def labeled_entry(self, parent, label: str, value: tk.StringVar) -> None:
        row = tk.Frame(parent, bg="#f4f5f4")
        row.pack(fill="x", pady=3)
        tk.Label(row, text=label, bg="#f4f5f4", width=24, anchor="w").pack(side="left")
        tk.Entry(row, textvariable=value).pack(side="left", fill="x", expand=True)

    def update_theme_preview(self, vars_: dict, preview: tk.Canvas) -> None:
        preview.configure(bg=vars_["background"].get())
        preview.delete("all")
        preview.create_rectangle(10, 5, 126, 12, fill=vars_["foreground"].get(), outline="")
        preview.create_rectangle(137, 5, 244, 12, fill=vars_["muted"].get(), outline="")
        preview.create_rectangle(255, 3, 436, 14, fill=vars_["accent"].get(), outline="")

    def choose_theme_preset(self, vars_: dict, preview: tk.Canvas) -> None:
        preset = THEME_KEYS_BY_LABEL.get(vars_["theme_preset"].get(), "custom")
        vars_["auto_theme"].set(preset == "wallpaper")
        definition = THEME_PRESETS[preset]
        for key in ("background", "foreground", "muted", "accent"):
            if key in definition:
                vars_[key].set(definition[key])
        self.update_theme_preview(vars_, preview)

    def choose_color(self, vars_: dict, key: str, preview: tk.Canvas) -> None:
        selected = colorchooser.askcolor(vars_[key].get(), parent=self.settings_window)[1]
        if selected:
            vars_[key].set(selected)
            vars_["theme_preset"].set(THEME_PRESETS["custom"]["label"])
            vars_["auto_theme"].set(False)
            self.update_theme_preview(vars_, preview)

    def save_auth_fields(self, vars_: dict) -> None:
        old_auth = (self.settings["client_id"], self.settings["tenant"])
        self.settings["client_id"] = vars_["client_id"].get().strip()
        self.settings["tenant"] = vars_["tenant"].get().strip() or "common"
        if old_auth != (self.settings["client_id"], self.settings["tenant"]):
            self.graph = GraphTodoClient(self.settings["client_id"], self.settings["tenant"])
            self.remote_tasks = []
        write_json(SETTINGS_PATH, self.settings)

    def login_from_settings(self, vars_: dict) -> None:
        self.save_auth_fields(vars_)

        def logged_in(name):
            self.status_text.set(f"{name} 已连接")
            self.load_todo_lists(vars_)

        self.run_async(self.graph.login, logged_in, "等待 Microsoft 登录...")

    def load_todo_lists(self, vars_: dict) -> None:
        def loaded(lists):
            self.todo_lists = lists
            self.fill_list_combo(vars_)
            if lists and not self.settings["todo_list_id"]:
                preferred = next(
                    (item for item in lists if item.get("wellknownListName") == "defaultList"),
                    lists[0],
                )
                self.settings["todo_list_id"] = preferred["id"]
                self.settings["todo_list_name"] = preferred["displayName"]
                vars_["todo_list"].set(preferred["displayName"])
                write_json(SETTINGS_PATH, self.settings)
            self.sync_remote()

        self.run_async(self.graph.list_task_lists, loaded, "正在加载 To Do 列表...")

    def fill_list_combo(self, vars_: dict) -> None:
        if not hasattr(self, "list_combo"):
            return
        names = [item.get("displayName", "未命名列表") for item in self.todo_lists]
        if self.settings["todo_list_name"] and self.settings["todo_list_name"] not in names:
            names.append(self.settings["todo_list_name"])
        self.list_combo.configure(values=names)
        vars_["todo_list"].set(self.settings["todo_list_name"])

    def choose_todo_list(self, vars_: dict) -> None:
        name = vars_["todo_list"].get()
        selected = next((item for item in self.todo_lists if item.get("displayName") == name), None)
        if selected:
            self.settings["todo_list_id"] = selected["id"]
            self.settings["todo_list_name"] = name
            write_json(SETTINGS_PATH, self.settings)
            self.sync_remote()

    def logout_graph(self) -> None:
        self.graph.logout()
        self.status_text.set("Microsoft To Do 已退出；保留最近同步任务")
        self.render_tasks()

    def save_settings_window(self, vars_: dict, window: tk.Toplevel) -> None:
        self.save_auth_fields(vars_)
        self.settings["theme_preset"] = THEME_KEYS_BY_LABEL.get(
            vars_["theme_preset"].get(), "custom"
        )
        self.settings["auto_theme"] = self.settings["theme_preset"] == "wallpaper"
        self.settings["opacity"] = vars_["opacity"].get()
        self.settings["font_family"] = vars_["font_family"].get()
        self.settings["font_size"] = vars_["font_size"].get()
        self.settings["refresh_minutes"] = vars_["refresh_minutes"].get()
        for key in ("background", "foreground", "muted", "accent"):
            self.settings[key] = vars_[key].get()
        requested_startup = vars_["auto_start"].get()
        try:
            self.set_auto_start(requested_startup)
            self.settings["auto_start"] = requested_startup
        except OSError as exc:
            messagebox.showerror("开机启动设置失败", str(exc), parent=window)
        self.settings = normalize_settings(self.settings)
        self.apply_appearance(refresh_theme=True)
        self.schedule_sync()
        window.destroy()

    def startup_command(self) -> str:
        if getattr(sys, "frozen", False):
            return f'"{sys.executable}"'
        return f'"{sys.executable}" "{Path(__file__).resolve()}"'

    def set_auto_start(self, enabled: bool) -> None:
        if winreg is None:
            raise OSError("仅 Windows 支持开机启动设置")
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
            if enabled:
                winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, self.startup_command())
            else:
                try:
                    winreg.DeleteValue(key, APP_NAME)
                except FileNotFoundError:
                    pass

    def wallpaper_path(self) -> Path | None:
        if winreg is None:
            return None
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\Desktop") as key:
                path, _ = winreg.QueryValueEx(key, "WallPaper")
            wallpaper = Path(path)
            return wallpaper if wallpaper.exists() else None
        except OSError:
            return None

    def wallpaper_palette(self) -> dict | None:
        if Image is None:
            return None
        path = self.wallpaper_path()
        if not path:
            return None
        try:
            self.root.update_idletasks()
            signature = (
                str(path),
                path.stat().st_mtime_ns,
                self.root.winfo_x(),
                self.root.winfo_y(),
                self.root.winfo_width(),
                self.root.winfo_height(),
            )
            with Image.open(path) as image:
                fitted = ImageOps.fit(
                    image.convert("RGB"),
                    (self.root.winfo_screenwidth(), self.root.winfo_screenheight()),
                )
                left = max(0, self.root.winfo_x())
                top = max(0, self.root.winfo_y())
                right = min(fitted.width, left + max(20, self.root.winfo_width()))
                bottom = min(fitted.height, top + max(20, self.root.winfo_height()))
                sample = fitted.crop((left, top, right, bottom)).resize((20, 20))
                average = tuple(int(value) for value in ImageStat.Stat(sample).mean[:3])
            self.last_wallpaper_signature = signature
        except (OSError, ValueError):
            return None
        dark = luminance(average) < 135
        background = blend(average, (16, 19, 21) if dark else (242, 243, 241), 0.3)
        foreground = blend(average, (244, 246, 245) if dark else (28, 36, 35), 0.83)
        muted = blend(average, foreground, 0.57)
        accent = blend(background, foreground, 0.32)
        return {
            "background": rgb_color(background),
            "foreground": rgb_color(foreground),
            "muted": rgb_color(muted),
            "accent": rgb_color(accent),
        }

    def schedule_theme_refresh(self) -> None:
        if self.theme_job:
            self.root.after_cancel(self.theme_job)
        self.theme_job = self.root.after(60000, self._periodic_theme_refresh)

    def _periodic_theme_refresh(self) -> None:
        if self.settings["auto_theme"]:
            self.apply_appearance(refresh_theme=True)
        self.schedule_theme_refresh()

    def close(self) -> None:
        self.save_geometry()
        self.root.destroy()


def main() -> None:
    ensure_data_dir()
    instance = SingleInstance()
    if not instance.acquire():
        return
    root = tk.Tk()
    try:
        QuietTodoWidget(root)
        root.mainloop()
    finally:
        instance.release()


if __name__ == "__main__":
    main()
