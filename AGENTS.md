# Repository Instructions

These instructions apply to the whole `threads-analytics` repository.

Reference basis: this file adapts local-only patterns from OpenCode project
rules (`AGENTS.md`), `itsinseong/value-for-fable` at
`f983f1300af803312b0fe51bc1328c93f49f0f0d`, and `code-yeongyu/lazycodex` at
`245fd8f45e37fe9b412ae57c1fb7cfbd672328b7`. It is repo-specific guidance, not
a request to install plugins, hooks, output styles, or global config.

## Project Shape

- This is a Python Threads API analytics tool. The main entry points are
  `auth.py`, `analyze.py`, `export_excel.py`, and `refresh_token.py`.
- `auth.py` opens a local HTTPS OAuth callback server and writes `ACCESS_TOKEN`
  and `USER_ID` into `.env`.
- `analyze.py` calls the Meta Threads API, updates `output/insights_cache.json`,
  and writes `output/analysis_*.json`.
- `export_excel.py` reads the newest `output/analysis_*.json` and writes
  `output/threads_analysis_*.xlsx`.
- The README and user-facing domain language are Korean. Preserve Korean copy
  unless the user asks to translate or rewrite it.

## Worktree Safety

- The worktree may contain user data, generated reports, and unrelated dirty
  files. Check `git status --short` before editing and do not revert, delete, or
  reformat files you were not asked to touch.
- Keep changes scoped. Do not modify `output/`, `.env`, local certificate files,
  `release/`, or untracked writing guides unless the task explicitly targets
  them.
- Do not run commands that refresh tokens, open OAuth flows, call the Meta API,
  or overwrite reports unless the user asks for that surface to be exercised.
- Use `rg`/`rg --files` for discovery. Batch independent reads when possible,
  but keep state-changing commands sequential.

## Secrets And Data

- `.env`, `localhost.crt`, `localhost.key`, and `output/` are intentionally
  gitignored. Treat them as sensitive local artifacts.
- Never print full access tokens, app secrets, user IDs, raw `.env` contents, or
  private analytics output in final answers. Redact secrets if you need to show
  evidence.
- `output/*.json`, `output/*.xlsx`, and `output/insights_cache.json` can contain
  private account analytics. Do not copy them into tracked docs or prompts.

## Python Workflow

- Use Python 3.8 or newer. On this machine, bare `python` can resolve to Python
  2.7, so prefer `python3` unless a virtual environment is already active.
- Use a virtual environment. The baseline install command is
  `python3 -m pip install -r requirements.txt`; if `uv` is already available,
  `uv pip install -r requirements.txt` is acceptable.
- For syntax-level verification, run:
  `python3 -m compileall analyze.py auth.py export_excel.py refresh_token.py`.
- If local tooling is already installed, prefer:
  `ruff check .` and
  `basedpyright analyze.py auth.py export_excel.py refresh_token.py`.
  Do not add new lint/type configuration just to satisfy a narrow task.
- For runtime QA, choose the cheapest faithful surface:
  `python3 export_excel.py` for workbook generation from existing local JSON,
  `python3 analyze.py` only when API access and runtime cost are intended, and
  `python3 auth.py` only when an OAuth browser flow is intended.

## Coding Guidance

- Read the relevant function and caller before editing. Preserve the simple
  script-oriented style unless a task truly needs a larger restructuring.
- Keep comments sparse and practical. Add comments only for constraints that
  are not obvious from code.
- Maintain request timeouts and defensive API handling when touching Threads
  API calls.
- Avoid broad refactors, dependency changes, and generated artifact churn for
  small fixes.

## Completion And Reporting

- Verify before claiming completion. Name the command or observable check that
  passed; if a check was skipped or failed, say exactly why.
- For meaningful investigation, implementation, debugging, configuration, or
  verification work, consult
  `/Users/solution/.claude/wiki/projects/threads-analytics.md` and related
  `sources/` or `lessons/` entries before trusting memory.
- If durable context is created, save it as a connected wiki note under
  `/Users/solution/.claude/wiki/sources/threads-analytics/` or another narrow
  existing source folder, then link it back from the project card, `index.md`,
  `log.md`, or a relevant lesson index.
- Verify live repo files before updating wiki memory; do not add broad wiki
  links only to increase link count.
- Respond in the user's language. For Korean prompts, answer in Korean while
  keeping command names, file paths, and API terms as-is.
- Final reports should start with the outcome, then concise evidence and any
  remaining caveats. Do not paste private local data or long unchanged file
  contents.
