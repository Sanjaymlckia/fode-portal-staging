@echo off
setlocal

echo.
echo === FODE DEPLOY (guarded) ===
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

REM --- Enforce clean tree (no accidental deploys) ---
for /f %%A in ('git status --porcelain') do (
  echo ERROR: Working tree is NOT clean. Commit or stash before deploy.
  git status -sb
  exit /b 1
)

echo OK: Repo clean + clasp config safe.
echo.

REM --- Run deploy (agent) ---
if not exist "fode-agent.ps1" (
  echo ERROR: fode-agent.ps1 not found in this folder.
  echo Fix: run this from the folder that contains fode-agent.ps1.
  exit /b 1
)

echo Running: powershell -NoProfile -ExecutionPolicy Bypass -File ".\fode-agent.ps1" deploy
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File ".\fode-agent.ps1" deploy
set EC=%ERRORLEVEL%

echo.
if not "%EC%"=="0" (
  echo DEPLOY FAILED (exit code %EC%)
  exit /b %EC%
)

echo DEPLOY COMPLETE.
echo.

endlocal