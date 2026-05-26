# 静默待办

一个适合放在 Windows 桌面的低存在感待办悬浮窗。它固定在桌面底层，不会覆盖普通应用窗口。

## 当前功能

- 无标题、无边框、小尺寸、任务栏隐藏式窗口；解锁时可拖动顶部空白区移动，并从右下角拉伸大小。
- 窗口持续置于普通应用下方；使用“显示桌面”导致窗口最小化时会自动恢复到桌面显示。
- 程序只允许一个实例运行，重复双击不会产生多个重叠弹窗。
- 提供小号锁头按钮；锁定前后布局保持一致，锁定后顶部的清单切换、设置、刷新与全部折叠/展开仍可使用，任务编辑、任务行折叠、拖动、滚动与窗口缩放等操作均失效，避免误操作。
- 设置中可直接选择自动融入壁纸、暮色灰蓝、雾松绿、暖砂岩、墨夜蓝等低对比度主题，也可自定义颜色。
- 设置内容较多时可在设置窗口中滚动查看，行为设置及 Microsoft To Do 联动配置不会被主题区域遮挡。
- 自动融入壁纸主题会读取当前 Windows 壁纸，并按窗口所在区域生成协调的背景、文字和强调色。
- 可调整背景色、文字色、辅助色、强调色、透明度、字体和字号；手动改色后自动切换为自定义主题。
- 长列表仅显示与主题同色系的细窄滚动提示块，不显示系统白色滚动条。
- 任务变化时暂停中途绘制并在完成后统一刷新；窗口只在需要时恢复桌面底层，不会随每次点击重复调整层级。
- 主窗口提供“待办”和“已完成”两个清单；勾选完成后任务移入已完成清单，可随时查看。
- 任务完成控件采用与文字高度协调且随文字居中的圆形按钮，不显示突兀的系统方框。
- 主任务和每一级带子项的任务均可折叠，折叠状态会随设置保存。
- 一级任务标题保持统一对齐，不会因是否包含子任务而左右跳动。
- 子任务按紧凑的小步幅逐级缩进，保留层级关系但避免列表被过度推向右侧。
- 顶部 `▸▸` / `▾▾` 图标可一键折叠或展开当前清单中的所有任务层级。
- 待办列表末尾提供无边框内联编辑行，鼠标悬浮到新增行时才显示任务圆点，直接输入主任务并按回车保存；“已完成”视图不显示新增入口。
- 任务与前三层子任务行尾均提供简洁的 `＋` 按钮；点击后在该层级末尾直接出现同样的无边框编辑行，回车保存、Esc 取消，空内容时点到别处也会自动取消，不再弹出对话框，最多可创建四级子任务。
- 解锁状态下可拖动 `≡` 手柄调整任务和每一级同级子任务的顺序，包括拖至同级列表末尾；本地子任务可拖到另一主任务下，或在另一子任务行上向右拖入标题区域，使其成为该任务的子任务。也可用星标置顶特定主任务。
- Microsoft To Do 同步得到的一级子任务可同级排序，但不能跨父任务拖动；其线上归属仍由 Microsoft To Do 保持。
- Microsoft To Do 任务的星标会同步为 To Do 的高重要性；自定义顺序只保存在本组件，不改变 To Do 客户端内的排序。
- Microsoft To Do 仅支持一级 checklist 子任务；在线任务的二至四级子任务保存在本组件本机，一级仍保持双向同步。
- 已同步的 To Do 任务会保存本地快照，包含完成状态；本地勾选或星标会先写入快照再尝试同步，退出登录或暂时离线时仍可查看最近状态。
- Microsoft To Do 登录后读取选定列表，并在窗口内新增、完成、删除任务及子任务。
- 可配置自动同步间隔、是否登录 Windows 后自动启动。
- 设置窗口横向拉大时，表单保持适中的最大宽度，不会把输入框和透明度滑块无限拉长。

## 直接运行

源码运行：

```powershell
python -m pip install -r .\requirements.txt
python .\todo_widget.py
```

生成无需安装 Python 的 EXE：

```powershell
powershell -ExecutionPolicy Bypass -File .\build-windows.ps1
```

生成文件位于 `dist\静默待办.exe`。首次使用时把 EXE 放在一个固定目录，再在设置中开启开机启动；移动 EXE 后需要重新开启该选项。

## 连接 Microsoft To Do

Microsoft To Do 通过 Microsoft Graph 读写。Microsoft 要求桌面应用先注册，不能仅凭账号密码直接接入。

1. 打开 [Microsoft Entra 管理中心](https://entra.microsoft.com/) 的“应用注册”，新建应用。
2. 支持的帐户类型选择包含你使用的账号类型；个人 Microsoft 账号应选择支持个人账号的类型。
3. 在“身份验证”中添加“移动和桌面应用程序”平台，重定向 URI 选择 `http://localhost`，并允许公共客户端流。
4. 在“API 权限”中添加 Microsoft Graph 的委托权限 `Tasks.ReadWrite`。
5. 复制应用程序（客户端）ID，在组件的 `···` 设置中填写 `Client ID`；个人账号的 Tenant 保持 `common`。
6. 点击“登录并加载列表”，在浏览器完成授权后选择需要显示的 To Do 列表。

新增任务时：已登录并选择列表则任务直接写入 Microsoft To Do；没有登录时则保存在本机 `%APPDATA%\QuietTodoWidget\local_tasks.json`。

Windows 上的 Microsoft To Do 客户端没有官方开放的本地读写接口可供本组件直接调用。受支持的双向同步方式是 Microsoft Graph；本组件会把最近同步结果同时保存到本机，供离线查看和保留完成状态。

## 数据位置

- 设置：`%APPDATA%\QuietTodoWidget\settings.json`
- 本地任务：`%APPDATA%\QuietTodoWidget\local_tasks.json`
- 最近同步的 Microsoft To Do 任务快照：`%APPDATA%\QuietTodoWidget\remote_tasks.json`
- Microsoft 登录缓存：`%APPDATA%\QuietTodoWidget\msal_cache.bin`（由当前 Windows 用户的 DPAPI 加密保护）

Microsoft To Do API 参考：

- [Microsoft Graph To Do API 概览](https://learn.microsoft.com/graph/api/resources/todo-overview?view=graph-rest-1.0)
- [任务列表的任务 API](https://learn.microsoft.com/graph/api/todotasklist-list-tasks?view=graph-rest-1.0)
- [子任务 checklistItem API](https://learn.microsoft.com/graph/api/resources/checklistitem?view=graph-rest-1.0)
