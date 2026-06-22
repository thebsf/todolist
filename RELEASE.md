# Release Notes

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
