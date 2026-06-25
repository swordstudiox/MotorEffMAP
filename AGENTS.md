# Agent Instructions

## Communication

- Always reply in Chinese.
- Clarify ambiguous requirements before making changes.
- For non-trivial work, write the plan to `tasks/todo.md` when that directory exists on the current branch.

## Git Branch Policy

This repository intentionally keeps two active local lines:

1. `main`
   - Must match `origin/main`.
   - Must not contain `tasks/` or `tests/`.
   - Used for the public release history and GitHub Releases.

2. `local/tasks-tests`
   - Must be based on `main`.
   - May contain `tasks/` and `tests/`.
   - Used for local task records, lessons, review notes, and regression tests.

When switching branches or preparing a release:

- Do not commit `tasks/` or `tests/` to `main`.
- Keep `main` and `origin/main` aligned unless explicitly asked otherwise.
- Put local-only task/test work on `local/tasks-tests`.
- If a lesson must be visible from both branch lines, write it in this `AGENTS.md` file, not only under `tasks/`.
- Use detailed Chinese commit messages that explain what changed, why it changed, and what was verified.

## Release Notes

- GitHub Release text should be detailed Chinese, including main fixes, behavior changes, compatibility notes, and verification.
- Updating a Git tag does not automatically update a GitHub Release entry.
- Git push credentials and GitHub Release API credentials are different paths; do not assume one implies access to the other.
- Verify proxy settings before remote operations. On this machine, `https.proxy=http://127.0.0.1:10808` has been verified for GitHub Git access.
