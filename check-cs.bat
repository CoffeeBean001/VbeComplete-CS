@echo off
rem One-click self check. Twelve steps, cheapest first:
rem   1 pure logic           no COM, no desktop
rem   2 word db              creates/seeds the two word DBs in a temp folder
rem   3 snippet store        temp db, your own snippets are untouched
rem   4 offscreen UI preview two PNGs you can eyeball
rem   5 icon + autostart     embedded app.ico + HKCU Run entry, restored afterwards
rem   6 candidate rules       runs the engine against the REAL module text
rem   7 parens on accept      needs the VBE code window open
rem   8 caret landing        needs the VBE code window open
rem   9 hotkey commands      needs the VBE code window open
rem  10 enter auto-close     needs the VBE code window open (drives the REAL hook)
rem  11 save-reformat        needs the VBE code window open; presses REAL Ctrl+S,
rem                          the workbook is saved and restored afterwards
rem  12 module-level kw      read-only, real VBE backend + in-memory text copy
rem
rem Steps 6-9 say "no code module, skip" (instead of failing) when the code
rem window is closed -- there is no way to reach a code pane from COM then.
rem
rem NOTE: CRLF line endings, and no bare parenthesis in echo text.
setlocal
set "EXE=%~dp0dist\VbeComplete.exe"

if not exist "%EXE%" goto nobuild

echo.
echo [check] 1/12 pure logic - 587 assertions
"%EXE%" --selftest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0selftest.log' -Tail 4"
echo.

echo [check] 2/12 word db - keywords-cs.db / functions-cs.db, temp copies
"%EXE%" --dbtest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0dbtest.log' -Tail 4"
echo.

echo [check] 3/12 snippet store - temp db, the real one is untouched
"%EXE%" --snipetest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0snipetest.log' -Tail 4"
echo.

echo [check] 4/12 offscreen UI preview - popup-preview.png / snippet-preview.png
"%EXE%" --uitest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0uitest.log' -Tail 4"
echo.

echo [check] 5/12 icon + autostart - embedded app.ico, Run entry restored afterwards
"%EXE%" --startuptest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0startuptest.log' -Tail 8"
echo.

echo [check] 6/12 candidate rules on the real module - needs the VBE code window open
"%EXE%" --enginetest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0enginetest.log' -Tail 4"
echo.

echo [check] 7/12 parens after accept - needs the VBE code window open
"%EXE%" --paretest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0paretest.log' -Tail 6"
echo.

echo [check] 8/12 caret landing - needs the VBE code window open
"%EXE%" --carettest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0carettest.log' -Tail 4"
echo.

echo [check] 9/12 hotkey commands - needs the VBE code window open
"%EXE%" --hktest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0hktest.log' -Tail 4"
echo.
echo [check] 10/12 enter auto-close - drives the REAL WH_KEYBOARD_LL hook
"%EXE%" --enterprobe
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0enterprobe.log' -Tail 8"
echo.
echo [check] 11/12 save-reformat caret - REAL Ctrl+S, workbook saved and restored afterwards
"%EXE%" --fmtprobe
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0fmtprobe.log' -Tail 10"
echo.
echo [check] 12/12 module-level keywords - read-only, real VBE backend
"%EXE%" --hdrprobe
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0hdrprobe.log' -Tail 10"
echo.
echo Logs: selftest.log / dbtest.log / snipetest.log / uitest.log / startuptest.log /
echo       enginetest.log / paretest.log / carettest.log / hktest.log /
echo       enterprobe.log / fmtprobe.log / hdrprobe.log
pause
endlocal
exit /b 0

:nobuild
echo [check] dist\VbeComplete.exe not found - run build.bat first.
pause
endlocal
exit /b 1
