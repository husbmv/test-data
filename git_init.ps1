# Git init + push to GitHub
# Run: powershell -ExecutionPolicy Bypass -File .\git_init.ps1

$ErrorActionPreference = 'Continue'
Set-Location 'D:\test-data-main'

function Step($msg) {
    Write-Host ""
    Write-Host "=== $msg ===" -ForegroundColor Cyan
}

Step "git init"
if (-not (Test-Path .git)) {
    git init
} else {
    Write-Host ".git already exists, skip"
}

Step "git add ."
git add .

Step "git status"
git status --short

Step "git commit"
git -c i18n.commitencoding=utf-8 -c i18n.logoutputencoding=utf-8 commit -m "Initial commit: pagination fix for all sheets"
if ($LASTEXITCODE -ne 0) {
    Write-Host "commit failed or nothing to commit, continuing..." -ForegroundColor Yellow
}

Step "git tag 20261007"
git tag 20261007 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "tag may already exist, skip" -ForegroundColor Yellow
}

Step "branch -> main"
git branch -M main

Step "add remote origin"
git remote remove origin 2>$null
git remote add origin git@github.com:husbmv/test-data.git
git remote -v

Step "push main"
git push -u origin main
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "[ERROR] push main failed." -ForegroundColor Red
    Write-Host "If remote has commits, run:" -ForegroundColor Yellow
    Write-Host "  git pull origin main --rebase --allow-unrelated-histories" -ForegroundColor Yellow
    Write-Host "  git push -u origin main" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Step "push tag"
git push origin 20261007

Step "log"
git log --oneline
git tag -l

Write-Host ""
Write-Host "=== DONE ===" -ForegroundColor Green
Read-Host "Press Enter to exit"
