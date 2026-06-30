# Release Notes

## 2026.06.30

### Fixed

- Fixed a window recovery issue after monitor layout changes. If the saved widget position is outside the current visible screen area, the app now moves it back onto the available desktop at startup and saves the corrected position.

## 2026.06.25

### Added

- Added an archive view for root tasks. Archived tasks are hidden from the pending and completed views, and can be restored from the archive view.

### Changed

- Archive controls now appear only on root task rows; subtasks remain visible under their archived parent task.
- Archived root tasks keep their full subtask tree visible, including the existing fold and unfold behavior.
- The restore control in the archive view now uses a back-arrow icon for clearer intent.

## 2026.06.22

### Added

- Added Obsidian Markdown task sync. Users can choose a `.md` task file in settings, and the widget reads and writes standard `- [ ]` / `- [x]` task lines.
- Added inline title editing for tasks and subtasks. Click task text to edit in place, press Enter or leave the field to save, and press Esc to cancel.
- Added Microsoft To Do title update support for synced tasks and first-level checklist items.

### Changed

- New installs no longer enable Obsidian sync by default. Users choose their own Markdown file before syncing.
- Inline editing now matches the existing add-task input style: no popup, no focus border, no automatic text selection, and a neutral theme-aware selection color for manual text selection.
- Documentation now describes Obsidian setup, inline editing, data locations, and sync behavior without embedding local personal paths.

### Privacy

- The repository does not include local task data, Microsoft tokens, Obsidian vault contents, or user-specific absolute paths.
- Runtime data remains under `%APPDATA%\QuietTodoWidget\` unless the user explicitly chooses an Obsidian Markdown file.
