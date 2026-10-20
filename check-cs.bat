@echo off
rem One-click self check. Eight steps, cheapest first:
rem   1 pure logic           no COM, no desktop
rem   2 word db              creates/seeds the two word DBs in a temp folder
rem   3 snippet store        temp db, your own snippets are untouched
rem   4 offscreen UI preview two PNGs you can eyeball
rem   5 candidate rules       runs the engine against the REAL module text
rem   6 parens on accept      needs the VBE code window open
rem   7 caret landing        needs the VBE code window open
rem   8 hotkey commands      needs the VBE code window open
rem
rem The last four say "no code module, skip" (instead of failing) when the
rem code window is closed -- there is no way to reach a code pane from COM then.
rem
rem NOTE: CRLF line endings, and no bare parenthesis in echo text.
setlocal
set "EXE=%~dp0dist\VbeComplete.exe"

if not exist "%EXE%" goto nobuild

echo.
echo [check] 1/8  pure logic - 296 assertions
"%EXE%" --selftest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0selftest.log' -Tail 4"
echo.

echo [check] 2/8  word db - keywords-cs.db / functions-cs.db, temp copies
"%EXE%" --dbtest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0dbtest.log' -Tail 4"
echo.

echo [check] 3/8  snippet store - temp db, the real one is untouched
"%EXE%" --snipetest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0snipetest.log' -Tail 4"
echo.

echo [check] 4/8  offscreen UI preview - popup-preview.png / snippet-preview.png
"%EXE%" --uitest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0uitest.log' -Tail 4"
echo.

echo [check] 5/8  candidate rules on the real module - needs the VBE code window open
"%EXE%" --enginetest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0enginetest.log' -Tail 4"
echo.

echo [check] 6/8  parens after accept - needs the VBE code window open
"%EXE%" --paretest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0paretest.log' -Tail 6"
echo.

echo [check] 7/8  caret landing - needs the VBE code window open
"%EXE%" --carettest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0carettest.log' -Tail 4"
echo.

echo [check] 8/8  hotkey commands - needs the VBE code window open
"%EXE%" --hktest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0hktest.log' -Tail 4"
echo.
echo Logs: selftest.log / dbtest.log / snipetest.log / uitest.log /
echo       enginetest.log / paretest.log / carettest.log / hktest.log
pause
endlocal
exit /b 0

:nobuild
echo [check] dist\VbeComplete.exe not found - run build.bat first.
pause
endlocal
exit /b 1
