@echo off
rem One-click self check: pure logic first, then the two real-machine checks.
rem
rem The last two need the VBE code window to be OPEN and focused -- there is no
rem way to reach a code pane from COM when the window is closed. If it is closed
rem they say "no code module, skip" instead of failing silently.
rem
rem NOTE: CRLF line endings, and no bare parenthesis in echo text.
setlocal
set "EXE=%~dp0dist\VbeComplete.exe"

if not exist "%EXE%" goto nobuild

echo.
echo [check] 1/3  pure logic - 103 assertions
"%EXE%" --selftest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0selftest.log' -Tail 4"
echo.

echo [check] 2/3  caret landing - needs the VBE code window open
"%EXE%" --carettest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0carettest.log' -Tail 4"
echo.

echo [check] 3/3  hotkey commands - needs the VBE code window open
"%EXE%" --hktest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0hktest.log' -Tail 4"
echo.
echo Logs: selftest.log / carettest.log / hktest.log in this folder.
pause
endlocal
exit /b 0

:nobuild
echo [check] dist\VbeComplete.exe not found - run build.bat first.
pause
endlocal
exit /b 1
