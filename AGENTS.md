# Repository Guidelines

## Project Structure & Module Organization
Automation scenarios live in `_ps`, grouped by purpose (for example `test_*.ps1` for regression). Shared drivers and helpers sit under `_ps\_lib` and should be imported via `. "$PSScriptRoot\_lib\Common.ps1"`. Static documentation outputs are published to `_docs`, while `_json/ErrorCode.json` stores centralized error metadata consumed at runtime. Logs such as `_log\ai_*.log` are produced during runs; keep them trimmed when sharing work. Root-level assets, including `run.cmd`, support local execution flows.

## Build, Test, and Development Commands
Use `run.cmd` from the repository root to launch the default smoke scenario (`_ps\aaa.ps1`). Execute a specific script directly with `powershell.exe -ExecutionPolicy Bypass -File _ps\<script>.ps1`; prefer the `test_*` suite for coverage. When authoring new scenarios, always load `_ps\_lib\Common.ps1` first to wire environment setup, configuration, and logging.

## Coding Style & Naming Conventions
Stick to four-space indentation with braces on their own lines to align with existing class-style modules. Name classes and functions in PascalCase (`WebDriver`, `InitSession`), locals in camelCase (`sessionId`), and reserve ALL_CAPS for constants only. Scripts should follow `verb_TargetScenario.ps1` (e.g. `test_Login.ps1`). Keep comments concise, in English unless bilingual output is required.

## Testing Guidelines
Treat each `test_*` script as an executable spec: encapsulate setup/teardown inside `_ps\_lib` helpers and emit checkpoints with `$global:Common.WriteLog`. Validate new drivers with a smoke script plus at least one regression scenario. Capture illustrative artifacts under `_docs\pages` when the run surfaces UI or log insights worth sharing.

## Commit & Pull Request Guidelines
Mirror the short commit subjects seen historically (e.g. `mod file`, `fix loader`), keeping them under 72 characters and referencing issue IDs when applicable. Pull requests should describe scope, list touched scripts, and attach relevant console or log excerpts from the executed `test_*` runs. Document any `ErrorCode.json` changes so reviewers know to trigger `Common.LoadErrorCodes()` on deploy.

## Security & Configuration Tips
Store sensitive endpoints or credentials outside the repo and expose them via parameters or secure local stores. Before committing, prune bulky traces from `_log` to keep diffs focused. When touching shared configuration, confirm the change propagates through `Common.ps1` and downstream consumers.
