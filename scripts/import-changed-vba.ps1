<#
.SYNOPSIS
    変更された VBA ソースファイルから対応する .xlsm を特定してインポートする。

.DESCRIPTION
    lefthook の post-commit / post-checkout / post-merge から呼ばれる。
    引数でファイルパスが渡されない場合は、直前のコミットの差分から
    src/ 配下の変更ファイルを自動検出する。

    pre-commit で xlsm からエクスポート済みの場合はスキップする
    （xlsm を正としての循環防止）。

.PARAMETER Files
    変更されたファイルパス（省略可。省略時は git diff で自動検出）。

.EXAMPLE
    .\import-changed-vba.ps1
    .\import-changed-vba.ps1 "src/MyWorkbook/Module1.bas"
#>
param (
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Files
)

$repoRoot = Split-Path $PSScriptRoot -Parent
$importScript = Join-Path $PSScriptRoot "import-vba.ps1"
$gitDir = (git -C $repoRoot rev-parse --git-dir 2>$null)
$skipMarker = if ($gitDir) { Join-Path $gitDir "vba-exported" } else { Join-Path $repoRoot ".git" "vba-exported" }

# --- 循環防止: pre-commit でエクスポート済みならスキップ ---
if (Test-Path $skipMarker) {
    Remove-Item $skipMarker -Force
    Write-Host "スキップ: xlsm からエクスポート済み（xlsm を正として src/ を更新済み）"
    exit 0
}

# --- 引数がない場合は git diff で自動検出 ---
if (-not $Files -or $Files.Count -eq 0) {
    $Files = @(git -C $repoRoot diff --name-only HEAD~1 HEAD -- "src/" 2>$null)
    if (-not $Files -or $Files.Count -eq 0) {
        Write-Host "対象ファイルなし（スキップ）"
        exit 0
    }
    Write-Host "git diff から検出: $($Files.Count) ファイル"
}

# --- 変更ファイルから対応するブック名を抽出（重複排除） ---
$workbookNames = @{}

foreach ($file in $Files) {
    if ($file -match '^src[/\\]([^/\\]+)[/\\]') {
        $workbookNames[$Matches[1]] = $true
    }
}

if ($workbookNames.Count -eq 0) {
    Write-Host "インポート対象なし（src/ 配下の変更ファイルが見つかりません）"
    exit 0
}

# --- 各ブックに対してインポート実行 ---
$xlsmPaths = @()

foreach ($name in $workbookNames.Keys) {
    $xlsmPath = Join-Path $repoRoot "$name.xlsm"

    if (-not (Test-Path $xlsmPath)) {
        Write-Host "警告: 対応する .xlsm が見つかりません: $xlsmPath" -ForegroundColor Yellow
        continue
    }

    $xlsmPaths += $xlsmPath
}

if ($xlsmPaths.Count -eq 0) {
    Write-Host "インポート対象の .xlsm が見つかりません"
    exit 0
}

Write-Host "=== VBA インポート開始 ==="
Write-Host "対象ブック数: $($xlsmPaths.Count)"

& $importScript @xlsmPaths
