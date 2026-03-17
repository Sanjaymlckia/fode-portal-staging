@echo off
setlocal

echo.
echo === FODE SESSION END ===
echo Current folder: %CD%
echo Time: %DATE% %TIME%
echo.

REM --- Ensure we are inside a git repo ---
git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
  echo Not inside a git repository.
  exit /b 1
)

echo --- Git branch + status ---
git status -sb
echo.

echo --- Uncommitted diff summary ---
git diff --stat
echo.

echo --- Deployment Reminder ---
echo - If you deployed, confirm you created a clasp version (rXXX).
echo - Confirm Admin + Student URLs are correct.
echo - Commit + push if changes are intentional.
echo - .clasp.json must remain local only.
echo.

endlocal