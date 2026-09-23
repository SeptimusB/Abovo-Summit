# Backing up this Summit project and Codex work

Close Summit and Codex (including Codex CLI sessions) before taking a normal file-copy backup. This avoids copying workbooks mid-save or SQLite databases mid-transaction. A tested filesystem/VSS snapshot is an alternative. Retain multiple dated backups, including an off-machine copy.

## Essential locations on this machine

1. **C:/Repos/Abovo Summit/** — the complete solution, hidden `.git`, local `.codex`, `Library` masters, XML, documentation, tools and uncommitted work. `bin`/`obj` are usually rebuildable, but `obj` currently contains disposable benchmark and diagnostic evidence not committed to Git; include it if that evidence matters.
2. **C:/Users/jmwor/.codex/** — Codex sessions, archived sessions, attachments, automations, memories, skills/plugins, configuration, worktrees and local SQLite/history/state data. Back up the whole directory, including hidden files and any database WAL/SHM companions, rather than guessing a minimum subset. This location contains authentication material and private client content: encrypt/access-control the backup, and do not add it to GitHub.
3. **C:/Sandbox/** and the relevant source files under **D:/Downloads/** — business-plan originals, client test workbooks, saved results and documents live outside the repository. Also preserve important referenced material in D:/TEMP if no permanent copy exists. Git does not cover any of these automatically.

## Extra desktop state

For a fuller local desktop restoration, also retain:

- `C:/Users/jmwor/AppData/Roaming/Codex/`
- `C:/Users/jmwor/AppData/Local/Packages/OpenAI.Codex_2p2nqsd0c76g0/`

These folders were observed on this machine. They are additional app state, not substitutes for `.codex` and the project. The installed program itself can normally be reinstalled; DevExpress/Visual Studio dependencies and licences must be available for a rebuild.

Codex officially uses CODEX_HOME (normally `~/.codex`) for configuration and local state; database placement can be configured separately. No CODEX_HOME/SQLITE_HOME override was present in the inspected environment. Recheck these locations if configuration or Windows account changes. Sources: https://learn.chatgpt.com/docs/config-file/config-advanced and https://learn.chatgpt.com/docs/config-file/config-reference.

## Verify the backup

Confirm that the copied repository contains its `.git` history and uncommitted files; verify representative workbook hashes; then test restoration in a separate folder or machine without overwriting live data. Git push alone is not a complete project/conversation/client-file backup. No backup or restore has been performed as part of this guidance.
