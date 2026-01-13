# Word Calendar Generator

Generates Word .docx calendar files, in the style of reference/*

## Commands

- `npm install` — install dependencies
- `npm run build` — build CLI to dist/cli.js
- `npm run build:web` — build browser bundle to web/calendar.bundle.js
- `npm run typecheck` — type check
- `npm run test:integration` — build, generate docx, verify XML content
- `npm run build && node dist/cli.js --start 2026-6 --months 3 -o test-output/test.docx` — generate test docx
- `uvx docx2pdf test-output/test.docx test-output/test.pdf` — convert docx to PDF for visual verification (uses Microsoft Word via automation)
- Browser integration via Chrome DevTools MCP
  - Use `mcp__chrome-devtools__new_page` tool to open web/index.html
  - Use `mcp__chrome-devtools__take_snapshot` to get element UIDs
  - Use `mcp__chrome-devtools__click` to click the Generate button
  - Use `mcp__chrome-devtools__list_console_messages` to check for errors

## Coding and interaction guidelines

This is a small compact minimal program. We will:
- KISS - Keep It Simple.
   - For error handling, we certainly want to handle all errors. But usually the best way of handling is a top-level exception handler that includes stacktrace; we should never attempt to recover-and-continue after an error unless we can document why such recovery is safe.
   - Avoid creating more top-level entities (files, functions, classes, dependencies), since each one requires documentation and invariants. It's usually preferable to keep functionality inline if it's only used once.
   - Functional style. We prefer constants, and prefer to avoid if blocks. It's better to have a single codepath that's used in all cases so as to reduce test burden.
   - For github, there will be no branches. This is a personal project. We will commit directly into 'master'
- DRY - Don't Repeat Yourself
   - For every function we write, we must evaluate whether it's needed, or whether some existing function should be made general-purpose.
- Invariants. I like provably correct code.
   - If there are two linked variables/fields (e.g. is_open and open_file_name), it's usually better to combine them into a single variable. If they must be kept separate, then (1) at declaration site we must document the invariant that relates them, (2) for every function that modifies them it must document how it restores the invariant.
   - If there is async re-entrancy, we must document what invariants are assumed.

For interaction,
- If the user asks a question "what is the difference", then the agent must give an answer, and must NOT proactively code.
- If the user explicitly asks for code, "Please implement this" or "Please make XYZ", only then should edits be made
