@echo off
setlocal

echo.
echo === FODE SESSION START ===
echo Current folder: %CD%
echo Time: %DATE% %TIME%
echo.

REM --- Ensure we are inside a git repo ---
git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
  echo ERROR: You are not inside a git repository.
  echo Fix: cd into your FODE repo folder first.
  exit /b 1
)

echo --- Git branch + status ---
git status -sb
echo.

REM --- .clasp.json safety check ---
if not exist ".clasp.json" (
  echo ERROR: .clasp.json missing in this folder.
  echo This machine is not configured for clasp.
  exit /b 1
)

git ls-files --error-unmatch .clasp.json >nul 2>&1
if not errorlevel 1 (
  echo ERROR: .clasp.json is tracked by git. This is dangerous.
  echo Run: git rm --cached .clasp.json
  exit /b 1
)

echo OK: Local clasp config present and safe.
echo.

echo Launching Codex...
echo.
Codex

endlocal