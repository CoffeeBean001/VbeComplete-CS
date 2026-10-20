# VbeComplete · C# edition

*English · [简体中文](README.zh-CN.md)*

The same VBA-editor completion tool as the Python edition (`..\VbeComplete-PY`), rewritten in C#
as an independent project. The VBE never completes the names **you** declared; VbeComplete does.
It works with **Microsoft Office** (Excel / Word / PowerPoint / Access / Visio) and **WPS**, and
detects the host from the foreground window.

The C# rewrite exists for one reason: **it sends zero COM calls while you are not inside a VBE code
pane**, and its message loop costs nothing when idle.

## Features

- **Two editable word databases**: `keywords-cs.db` (79 VBA keywords) and `functions-cs.db`
  (171 functions + 13 types). Plain SQLite files — open them in DB Browser, edit them like a
  spreadsheet, restart to pick up the change. Both are created from built-in seeds on first run.
  **Enums are deliberately excluded** (`vbRed` / `xlUp` and friends are noise in a popup).
- **Live candidates** for identifiers, declarations, procedures, modules, forms, controls, and the
  keywords / functions from the two word databases.
- **Fuzzy matching** ranked as contiguous block > prefix > camel-hump initials > scattered, with
  **per-character highlighting** — typing `ds` lights up the `D` and `S` of `dataSheet`, not a whole
  slice of the name.
- **Chinese identifiers just work**: `Sub 汇总()` / `Dim 学生 As String` enter the pool, highlight
  per character, and the caret lands in the right place after `Tab` (verified live).
- **IDEA-style ordering**: match quality > first hit position > hit span > scope proximity
  (current-procedure locals > current module > other modules' Public) > recently used > shorter name.
- **Scope filtering**: a `Dim tmp` inside `Sub A` never shows up inside `Sub B`, and other modules'
  `Private` members never leak in — while other modules' `Public` procedures still autocomplete.
- **`Option Explicit` is honoured**: in a module that declares it, names you never `Dim`'d stay out
  of the pool (they would not compile anyway); without it those names are treated as implicitly
  declared and are offered normally. Host-declared objects (something seen as `control.member` or
  after `Set control =`) are exempt either way.
- **Silent where it should be**: after `As` (only a type name can follow), in the naming position of
  `Dim` / `Public` / `Private` / `Global`, and for the word you are currently typing.
- **Member completion after a dot**: once `Dim c As Cls1` or `Set c = New Cls1` is in the code,
  `c.` lists `Cls1`'s public members. If the type cannot be determined, it steps aside entirely
  instead of passing off the whole project pool as members.
- **Popup**: fixed width (30 characters, so it never jumps sideways), 15 rows max, per-character
  highlights, hover highlight, draggable scrollbar, and a detail line only when a name is truncated.
  Owner-drawn and it **never steals focus**.
- **Never fights the VBE for space**: it yields as soon as VBE's own member list appears, and moves
  below VBE's parameter-info popup instead of covering it.
- **`Tab` is the only accept key** — `Enter` still inserts a newline (a wrong completion costs too
  much). `Sub` gets a space, `Function` gets `()` with the caret inside. Silent inside comments and
  string literals.
- **Enter auto-indent**: block openers indent one level deeper and add `End If` / `Next` / `Loop`;
  `Else` / `Case` lines are re-aligned first.
- **Shift+Enter** starts a new line below without splitting the current one.
- **Backspace, two jobs**: deletes a whitespace-only line, or removes an empty pair `()` / `""`.
- **Auto-pairing** `(` `"` — full-width `（）` / `“”` under a Chinese-punctuation IME; the closing
  char steps over an existing one; a selection gets wrapped.
- **Line commands**: `Ctrl+/` comment toggle · `Ctrl+D` duplicate · `Ctrl+K` delete ·
  `Ctrl+Shift+Enter` insert above · `Ctrl+Shift+↑/↓` jump to the last declaration / closing line.
- **`Ctrl+Q` snippet library** — its own `snippets-cs.db` (SQLite). Two-line list on the left
  (title + first code line, per-character search highlighting), editor on the right, `Ctrl+N` new /
  `Ctrl+S` save / `Enter` or `Ctrl+Enter` to insert at the caret.
- **Adaptive polling**: 50 ms while typing, 150 ms when idle, 250 ms outside the VBE. Text changes
  caused by a paste or by another program only dismiss the popup — they never open it
  (`REQUIRE_KEY_TRIGGER`).

## Pros

- **Almost free when you are not writing VBA**: ~0.13 % of one core and 11.7 MB private, because
  no COM call is made outside the VBE code pane (the Python edition sits at ~2 % / 29 MB).
- Snappy: a single ~350 KB executable (framework-dependent), instant start, no interpreter warm-up.
- **Both word databases and the snippet database are plain SQLite** — edit them with any database
  tool; they are fully independent of the Python edition, sharing no file with it.
- Every shortcut follows one rule: **if the action fails, the original key is replayed**, so keys
  are never silently swallowed.
- 293 self-checks (fuzzy matching, scopes, `Option Explicit`, Chinese identifiers, member completion,
  highlight splitting) plus word-DB, snippet-store and off-screen rendering checks, and three live
  probes keep the logic honest.

## Cons

- Needs the **.NET 10 Desktop Runtime**; Windows only.
- Requires *Trust access to the VBA project object model* (off by default), same as the Python edition.
- **Fewer features than the Python edition**: no watchdog / auto-restart on hang, no parameter-info
  popup, no host enum constants, and member completion only understands `Dim ... As` and
  `Set ... = New` (it does no return-type inference).
- The word databases are read once at startup, so a change needs a restart (deliberate: it keeps
  "file changed underneath us" out of the running process).
- Project-wide identifiers are re-scanned every 2 s (`SCOPE_TTL`); with very many modules the first
  completion after a pause can take a few dozen milliseconds.
- The global hook is mounted only while the VBE is focused, so there is no completion in other
  editors and a short delay when switching back.
- Live verification (`--enginetest` / `--carettest` / `--hktest`) requires an open VBA project and
  code window; with no host available those checks skip instead of running.
