param(
  [Parameter(Mandatory=$false)]
  [string]$Branch = "master"
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "=== Синхронизация проекта с GitHub ===" -ForegroundColor Cyan
Write-Host "Репозиторий: https://github.com/vlddzhus/exel_agent.git" -ForegroundColor Cyan
Write-Host "Ветка: $Branch" -ForegroundColor Cyan
Write-Host ""

# Проверка, что мы в git-репозитории
if (-not (Test-Path (Join-Path $RepoRoot ".git"))) {
  Write-Host "Ошибка: не найден .git в $RepoRoot" -ForegroundColor Red
  Write-Host "Убедись, что скрипт запускается из корня проекта." -ForegroundColor Red
  exit 1
}

$git = Get-Command "git.exe" -ErrorAction SilentlyContinue
if (-not $git) {
  Write-Host "Ошибка: git не найден. Установи Git: winget install Git.Git" -ForegroundColor Red
  exit 1
}

try {
  Set-Location -LiteralPath $RepoRoot

  Write-Host "[1/4] Проверка текущего статуса..." -ForegroundColor Yellow
  $status = & git status --porcelain
  if ($status) {
    Write-Host "Есть незакоммиченные изменения:" -ForegroundColor Yellow
    Write-Host $status
    $answer = Read-Host "Спрятать их (stash) перед pull? [Y/n]"
    if ($answer -ne "n") {
      & git stash push -m "auto-stash перед sync"
      Write-Host "Изменения спрятаны в stash." -ForegroundColor Green
    }
  }

  Write-Host "[2/4] Переключение на ветку $Branch..." -ForegroundColor Yellow
  & git checkout $Branch 2>$null

  Write-Host "[3/4] Pull с GitHub..." -ForegroundColor Yellow
  & git pull origin $Branch --ff-only
  if ($LASTEXITCODE -ne 0) {
    Write-Host "Fast-forward не удался. Попробую merge..." -ForegroundColor Yellow
    & git pull origin $Branch
  }

  Write-Host "[4/4] Установка зависимостей..." -ForegroundColor Yellow
  Write-Host "  -> npm install (корень)..." -ForegroundColor Gray
  & npm install --no-audit --no-fund 2>&1 | Out-Null
  Write-Host "  -> npm install (backend)..." -ForegroundColor Gray
  & npm install --prefix backend --no-audit --no-fund 2>&1 | Out-Null

  Write-Host ""
  Write-Host "=== Готово! Проект актуален. ===" -ForegroundColor Green
  Write-Host "Коммит: $((& git log --oneline -1))" -ForegroundColor Cyan
}
catch {
  Write-Host "Ошибка: $_" -ForegroundColor Red
  exit 1
}
