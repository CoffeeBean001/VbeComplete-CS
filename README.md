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
  **per-character emphasis** — typing `ds` bolds the `D` and `S` of `dataSheet`, not a whole slice
  of the name.
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
- **Silent where it should be**: after `As` **or `New`** (only a type name can follow — that includes
  `Set c = New |`, `For i As |` and parameter lists, not just `Dim` lines), in the naming position of
  `Dim` / `Public` / `Private` / `Global`, and for the word you are currently typing.
- **Class-module members are not bare names**: a `Public s1` inside `cls1` is reachable only as
  `c.s1` or from inside `cls1` itself — never as a bare name in another module,
  because `s1` alone would not compile there.
- **A dot belongs to the VBE**: after `student.`, `Range("A1").` or a `With` block's `.Size` it
  shows **no list of its own at all** — the VBE's own member list knows the real type and is the
  better one. `Ctrl+Space` still asks for our list on purpose; `MEMBER_POPUP=1` restores the old
  behaviour (narrowed to the declared type when `Dim c As Cls1` / `Set c = New Cls1` is in the
  code, full project pool otherwise).
- **Popup styled after the VBE's own list** (v139; fonts, colours and size updated in v140/v141):
  white background, black text, one icon per row, a 1 px raised border, and width that follows the
  longest candidate (640 px cap), 15 rows max; matched letters are **bold** rather than coloured.
  The font follows the VBE editor itself (reads `FontFace` / `FontHeight` from `VBA\7.1 → 7.0 → 6.0`)
  with `Light` / `Thin` weight suffixes stripped — thin weights at small sizes are exactly what
  looked "flat" — falling back to JetBrains Mono / Cascadia Mono / Consolas when that read fails or
  yields a stock face like `Courier New`. The size defaults to **one point smaller than the VBE**
  (14 pt editor → 13 pt list); `POPUP_FONT_SIZE = -2` goes two smaller, a positive value sets it
  outright. The selected row is a VBE-style **deep blue bar with white text** (`POPUP_SEL_BG`), and
  matched letters switch to a warm yellow on that bar. Owner-drawn and it **never steals focus**.
- **Popups only while you are actually typing**: a mouse click on a variable on another line (or an
  arrow-key move) only moves the caret and no longer pops the list — the check is "the caret line
  changed, so this is not typing", and typing never changes the line number.
- **Never fights the VBE for space**: it yields the moment VBE's own member list appears — even if
  ours is already on screen, it is taken down within one poll tick — and moves
  below VBE's parameter-info popup instead of covering it.
- **`Tab` is the only accept key** — `Enter` still inserts a newline (a wrong completion costs too
  much). `Sub` gets a space, `Function` gets `()` with the caret inside — and so do the functions
  from the word databases (`MsgBox` / `Split` …), even at the very start of a line. Type names
  (`String`) never get parens. Silent inside comments and string literals.
- **Enter auto-indent *and* auto-close**: block openers (`If…Then` / `For` / `Do` / `While` /
  `With` / `Select Case` / `Sub` / `Function` / `Property` / `Type` / `Enum`) indent the new line
  one level deeper and insert `End If` / `Next` / `Loop` / `Wend` / `End With` / `End Select` /
  `End Sub` … below it; `Else` / `Case` lines are re-aligned first. The closer decision follows the
  Python build's v80a rules: sibling code sitting right below no longer suppresses the closer
  (writing one `For`, then going back up to add another `For` above it is the common case), and it
  never invents a second closer when one already belongs to this block. Hitting Enter *immediately*
  after the last character is handled too: the hook only takes the key, the real decision happens on
  the UI thread with live COM data (and the newline is replayed if it turns out not to be a header);
  while an IME composition is active the key is never taken (there Enter commits the composition).
- **Shift+Enter** starts a new line below without splitting the current one.
- **Backspace, two jobs**: deletes a whitespace-only line, or removes an empty pair `()` / `""`.
- **Auto-pairing** `(` `"` — full-width `（）` / `“”` under a Chinese-punctuation IME; the closing
  char steps over an existing one; a selection gets wrapped.
- **Line commands**: `Ctrl+/` comment toggle — it fires VBE's own *Comment Block / Uncomment Block*
  command, so it costs one round-trip (~1 ms on a 40-line block, vs ~340 ms when it
  edited line by line) and leaves the selection intact; a multi-line selection is
  handled as a block. The numpad `/` and `Ctrl+Shift+/` count too ·
  `Ctrl+D` duplicate · `Ctrl+K` delete ·
  `Ctrl+Shift+Enter` insert above · `Ctrl+Shift+↑/↓` jump to the last declaration / closing line ·
  `Ctrl+Shift+L` **re-indent the whole module** (or just the selection when it spans lines). Only
  leading whitespace is rewritten — not a single character inside a line — so the worst case is a
  wrong indent, a few `Ctrl+Z` away.
- **`Ctrl+Q` snippet library** — its own `snippets-cs.db` (SQLite; the window title shows that
  file name), the only snippet file in the project. Two-line list on the left (title + first code
  line, per-character search highlighting), editor on the right, `Ctrl+N` new / `Ctrl+S` save /
  `Enter` or `Ctrl+Enter` to insert at the caret.
- **Adaptive polling**: 50 ms while typing, 150 ms when idle, 250 ms outside the VBE. Text changes
  caused by a paste or by another program only dismiss the popup — they never open it
  (`REQUIRE_KEY_TRIGGER`).
- **Tray icon + run at Windows startup**: a dark-blue V badge sits in the notification area
  (`app.ico` carries 16/20/24/32/40/48/64/128/256 frames, so it still reads at 16 px). The context
  menu has **"开机自动启动" (start with Windows)** — written to
  `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`, **no admin rights needed** — plus open-log
  and exit. On startup it also asks Windows to keep the icon always visible, so Win11 does not
  bury it in the `^` overflow.

## Pros

- **Almost free when you are not writing VBA**: ~0.13 % of one core and 11.7 MB private, because
  no COM call is made outside the VBE code pane (the Python edition sits at ~2 % / 29 MB).
- Snappy: a single ~350 KB executable (framework-dependent), instant start, no interpreter warm-up.
- **Both word databases and the snippet database are plain SQLite** — edit them with any database
  tool; they are fully independent of the Python edition, sharing no file with it.
- Every shortcut follows one rule: **if the action fails, the original key is replayed**, so keys
  are never silently swallowed.
- 363 self-checks (fuzzy matching, scopes, `Option Explicit`, Chinese identifiers, member completion,
  highlight splitting, call-site detection, `As`/`New` suppression, line-by-line re-indent) plus
  word-DB, snippet-store, icon/autostart and off-screen rendering checks, and seven live probes
  (including `Ctrl+Shift+L`) keep the logic honest.

## Cons

- Needs the **.NET 10 Desktop Runtime**; Windows only.
- Requires *Trust access to the VBA project object model* (off by default), same as the Python edition.
- **Fewer features than the Python edition**: no watchdog / auto-restart on hang, no parameter-info
  popup, no host enum constants, and the optional `MEMBER_POPUP=1` member list only understands
  `Dim ... As` and `Set ... = New` (it does no return-type inference).
- The word databases are read once at startup, so a change needs a restart (deliberate: it keeps
  "file changed underneath us" out of the running process).
- Autostart writes to the current user's registry; some security suites block that (the menu then
  reports a failure and the log says why). You can always drag the icon out of the `^` overflow
  yourself.
- Project-wide identifiers are re-scanned every 2 s (`SCOPE_TTL`); with very many modules the first
  completion after a pause can take a few dozen milliseconds.
- The global hook is mounted only while the VBE is focused, so there is no completion in other
  editors and a short delay when switching back.
- Live verification: `--enginetest` runs against a **synthetic project** (no host, no open code
  window), so it always really runs. `--paretest` / `--carettest` / `--hktest` do need an open VBA
  project and code window; with no host available those three skip instead of running.
  `--yieldprobe` samples the "is the VBE's own list on screen right now" check five times in a row —
  run it while a member list is up if the yielding ever looks wrong.
