@echo off
cd /d D:\test-data-main

echo === git init ===
if not exist .git (
    git init
    if errorlevel 1 goto GITFAIL
) else (
    echo .git already exists, skip init
)

echo.
echo === git add ===
git add .
if errorlevel 1 goto GITFAIL

echo.
echo === git status ===
git status --short

echo.
echo === git commit ===
git -c i18n.commitencoding=utf-8 -c i18n.logoutputencoding=utf-8 commit -m "Initial commit: 单盘测试/表三/表四/段落清单/合并 + 分页符"
if errorlevel 1 (
    echo Commit may have failed or nothing to commit, continuing...
)

echo.
echo === git tag ===
git tag 20261007 2>nul
if errorlevel 1 (
    echo Tag 20261007 may already exist, skipping
)

echo.
echo === set default branch to main ===
git branch -M main
if errorlevel 1 goto GITFAIL

echo.
echo === git remote add origin ===
git remote remove origin 2>nul
git remote add origin git@github.com:husbmv/test-data.git
if errorlevel 1 goto GITFAIL

echo.
echo === git push -u origin main ===
git push -u origin main
if errorlevel 1 goto PUSHFAIL

echo.
echo === git push tag ===
git push origin 20261007
if errorlevel 1 goto PUSHFAIL

echo.
echo === git log ===
git log --oneline
echo.
echo === git tag list ===
git tag -l
echo.
echo === git remote -v ===
git remote -v

echo.
echo === DONE ===
pause
exit /b 0

:GITFAIL
echo.
echo [ERROR] git command failed. Make sure git is installed and in PATH.
pause
exit /b 1

:PUSHFAIL
echo.
echo [ERROR] git push failed.
echo Possible reasons:
echo   1. SSH key not configured. Run: ssh -T git@github.com
echo   2. Remote repo already has commits. Run: git pull origin main --rebase
echo   3. Network issue.
pause
exit /b 1
