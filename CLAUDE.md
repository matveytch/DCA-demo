# DCA — rules for Claude sessions

Talk to the user in Russian; code, comments and commit messages in English.

## Harness: where things are

Read at the start of every session, in this order:

1. This file.
2. [docs/requirements.md](docs/requirements.md) — what we build, for whom, rules, acceptance criteria, what we
   don't do. **Source of truth about the task.**
3. [docs/architecture.md](docs/architecture.md) — parts, where data lives, stack, boundaries.
4. [docs/tests.md](docs/tests.md) — how to run the tests, rules, acceptance criterion → test map.
5. [docs/deploy.md](docs/deploy.md) — install and update on the operator's Mac.
6. Open issues: `gh issue list`; the issue the user names: `gh issue view <n> --comments`.

Commands:

- Tests: `python3 -m pytest` (parallel; `-n 0` to debug one test).
- Program: `./dca <command>`. Never against real operator data from a Claude session: tests use `tmp_path`
  and `DCA_HOME`.
- Templates and Excel files: `python3 scripts/make_demo_templates.py` (then run the tests).

## Working rules

- Non-trivial change → show the plan (what and where) first and wait for «дальше». A change that fits one
  sentence → do it at once.
- Requirements outrank code. A new clarification from the user goes into docs/requirements.md first (with
  the date), then into code.
- Before new code: add or update the tests by the requirements (docs/tests.md); run the suite before and after.
- Done only after a real check by docs/tests.md. Tick an acceptance criterion only then. Never report «готово»
  in advance.
- A change of parts, data locations or boundaries updates docs/architecture.md in the same task.
- Secrets never go into code or these files (the backup password lives in the macOS Keychain).
- A risky step or a decision that is the user's → stop and ask; don't assume knowledge.
- Stop word: «Останься в плане, ничего не запускай» → only show what you would do and wait.

## Stay inside the repository

Sessions work **only inside this repository**. Do not read, list, search or write anything outside it.
If a task needs a file outside, say so and ask the user to copy it in or to do that step themselves.

- Enforced by the hook `.claude/hooks/stay_in_repo.py`, registered in `.claude/settings.json`
  (PreToolUse on Read, Write, Edit, MultiEdit, NotebookEdit, Glob, Grep, Bash). Tests: `tests/test_claude_hook.py`.
- The hook allows only: `/dev/null` & co, Claude Code's scratchpad `/tmp/claude-<uid>/`, and this
  project's folder in `~/.claude/projects/` (memory, saved tool output).
- The Bash check reads paths written in the command. Do not work around it: no computed paths, no
  scripts or programs that reach outside the repository.
- Never edit, disable or bypass the hook without the user's explicit OK in chat.

Real data (deals, people, companies) must never end up in templates, examples, tests or docs — use invented
values like those in `data/input_example_FILLED.xlsx`.

## Git

- Commit **and push** after every important change: a finished, working step (feature, bug fix, template or
  Excel change, test change, documentation update). Not every intermediate edit.
- Before committing: `python3 -m pytest` must be green. Read the summary line and check it says «passed».
  A pipe like `pytest … | tail -1 && git commit` commits even when pytest fails.
- One logical change per commit, message in English: what and why. Template and Excel changes get their own
  commit so they can be reverted alone.
- After pushing, tell the user the short hash + one line, so they can roll back to it.
- Never rewrite pushed history (`push --force`, `reset --hard` on `main`) without the user's explicit OK.

## Tasks live in GitHub Issues

- **Every large task is an issue.** Before creating it, show the user the full text; create it only after
  their OK. Sections: «Что делаем», «Зачем», «Границы», «Готово, когда» (checklist). Text in Russian.
- **A new session starts with the issue** the user names; stay inside its «Границы».
- **Link work to the issue:** a pull request body says `Closes #<n>`; commits say `Refs #<n>`.
- **Closing an issue leaves a short summary:** what was done, what was checked, what is left.

## Documentation

- Never delete or rewrite existing documentation text. When something is outdated, add the current version
  next to it (marked «актуально на <date>») and mark the old text «устарело — сохранено для истории».
- Every commit that changes behaviour, data layout, templates, Excel files, commands or rules is followed by
  bringing docs/ and README.md up to date — before the work is reported as done.
- New documentation is written in plain Russian: short sentences (≤ 20 words), active voice, one term per
  concept, instructions as plural imperatives («Запустите…»), lists for enumerations.
- Identifiers, file names, commands and program output stay as they are — not translated.
