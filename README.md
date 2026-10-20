# VbeComplete · C# edition

*English · [简体中文](README.zh-CN.md)*

The same VBA-editor completion tool as the Python edition (`..\VbeComplete-PY`), rewritten in C#
as an independent project. The VBE never completes the names **you** declared; VbeComplete does.
It works with **Microsoft Office** (Excel / Word / PowerPoint / Access / Visio) and **WPS**, and
detects the host from the foreground window.

The C# rewrite exists for one reason: **it sends zero COM calls while you are not inside a VBE code
pane**, and its message loop costs nothing when idle.

## Features

- **Live candidates** for identifiers, declarations, procedures, modules, forms, controls, keywords
  and built-in functions (171 functions / 13 types / 79 keywords plus your own word bank).
- **Type icons, detail line, never steals focus** — the popup is fully owner-drawn and cannot take
  the caret away from the editor.
- **Smart acceptance**: `Tab` / `Enter` completes; `Sub` gets a space, `Function` gets `()` with the
  caret inside; the list also gets out of the way in comments, after a dot, and while you are
  declaring a new name.
- **Enter auto-indent**: block openers indent one level deeper and add `End If` / `Next` / `Loop`;
  `Else` / `Case` lines are re-aligned first.
- **Shift+Enter** starts a new line below without splitting the current one.
- **Backspace, two jobs**: deletes a whitespace-only line, or removes an empty pair `()` / `""`.
- **Auto-pairing** `(` `"` — full-width `（）` / `“”` under a Chinese-punctuation IME; the closing
  char steps over an existing one; a selection gets wrapped.
- **Line commands**: `Ctrl+/` comment toggle · `Ctrl+D` duplicate · `Ctrl+K` delete ·
  `Ctrl+Shift+Enter` insert above · `Ctrl+Shift+↑/↓` jump to the last declaration / closing line.
- **`Ctrl+Q` snippet library** — reads the same `snippets.db` as the Python edition (SQLite,
  fuzzy title search, insert at caret).
- **Adaptive polling**: 50 ms while typing, 150 ms when idle, 250 ms outside the VBE.

## Pros

- **Almost free when you are not writing VBA**: ~0.13 % of one core and 11.7 MB private, because
  no COM call is made outside the VBE code pane (the Python edition sits at ~2 % / 29 MB).
- Snappy: a single ~334 KB executable (framework-dependent), instant start, no interpreter warm-up.
- Shares the snippet database with the Python edition — switching between them loses no data.
- Every shortcut follows one rule: **if the action fails, the original key is replayed**, so keys
  are never silently swallowed.
- 205 self-checks plus a snippet-store check and two live probes keep the logic honest.

## Cons

- Needs the **.NET 10 Desktop Runtime**; Windows only.
- Requires *Trust access to the VBA project object model* (off by default), same as the Python edition.
- **Fewer features than the Python edition**: no watchdog / auto-restart on hang, no yielding to
  VBE's own member list, no parameter-info popup, no host enum constants.
- The global hook is mounted only while the VBE is focused, so there is no completion in other
  editors and a short delay when switching back.
- Live verification (`--carettest` / `--hktest`) requires an open VBA project and code window; with
  no host available those checks skip instead of running.
