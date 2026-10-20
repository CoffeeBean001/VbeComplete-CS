@echo off
rem One-click self check. Seven steps, cheapest first:
rem   1 pure logic           no COM, no desktop
rem   2 word db              creates/seeds the two word DBs in a temp folder
rem   3 snippet store        temp db, your own snippets are untouched
rem   4 offscreen UI preview two PNGs you can eyeball
rem   5 candidate rules       runs the engine against the REAL module text
rem   6 caret landing        needs the VBE code window open
rem   7 hotkey commands      needs the VBE code window open
rem
rem The last three say "no code module, skip" (instead of failing) when the
rem code window is closed -- there is no way to reach a code pane from COM then.
rem
rem NOTE: CRLF line endings, and no bare parenthesis in echo text.
setlocal
set "EXE=%~dp0dist\VbeComplete.exe"

if not exist "%EXE%" goto nobuild

echo.
echo [check] 1/7  pure logic - 293 assertions
"%EXE%" --selftest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0selftest.log' -Tail 4"
echo.

echo [check] 2/7  word db - keywords-cs.db / functions-cs.db, temp copies
"%EXE%" --dbtest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0dbtest.log' -Tail 4"
echo.

echo [check] 3/7  snippet store - temp db, the real one is untouched
"%EXE%" --snipetest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0snipetest.log' -Tail 4"
echo.

echo [check] 4/7  offscreen UI preview - popup-preview.png / snippet-preview.png
"%EXE%" --uitest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0uitest.log' -Tail 4"
echo.

echo [check] 5/7  candidate rules on the real module - needs the VBE code window open
"%EXE%" --enginetest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0enginetest.log' -Tail 4"
echo.

echo [check] 6/7  caret landing - needs the VBE code window open
"%EXE%" --carettest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0carettest.log' -Tail 4"
echo.

echo [check] 7/7  hotkey commands - needs the VBE code window open
"%EXE%" --hktest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0hktest.log' -Tail 4"
echo.
echo Logs: selftest.log / dbtest.log / snipetest.log / uitest.log /
echo       enginetest.log / carettest.log / hktest.log
pause
endlocal
exit /b 0

:nobuild
echo [check] dist\VbeComplete.exe not found - run build.bat first.
pause
endlocal
exit /b 1
