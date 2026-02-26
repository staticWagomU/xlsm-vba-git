<#
.SYNOPSIS
    Excel VBA マクロのソースコードをエクスポートする。

.DESCRIPTION
    指定された .xlsm ファイルから VBA コンポーネント（標準モジュール、クラスモジュール、
    フォーム、ドキュメントモジュール）をテキストファイルとしてエクスポートする。
    出力先は src/<ブック名>/ ディレクトリ。

.PARAMETER Path
    エクスポート対象の .xlsm ファイルパス（1つ以上）。

.EXAMPLE
    .\export-vba.ps1 MyWorkbook.xlsm
    .\export-vba.ps1 Book1.xlsm Book2.xlsm
#>
param (
    [Parameter(Mandatory = $true, Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Path
)

function Get-VBComponentTypeExtension {
    param([int]$Type)

    switch ($Type) {
        1   { return "bas" }   # 標準モジュール
        2   { return "cls" }   # クラスモジュール
        3   { return "frm" }   # ユーザーフォーム
        100 { return "cls" }   # ドキュメントモジュール (ThisWorkbook, Sheet等)
        default { return $null }
    }
}

function Export-VBAFromWorkbook {
    param([string]$XlsmPath)

    $xl = $null
    $wb = $null

    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $xl.AutomationSecurity = 3  # msoAutomationSecurityForceDisable

        Write-Host "ファイルを開いています: $XlsmPath"
        $wb = $xl.Workbooks.Open($XlsmPath)

        $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($XlsmPath)
        $repoRoot = Split-Path $PSScriptRoot -Parent
        $exportDir = Join-Path $repoRoot "src" $fileNameWithoutExt

        if (-not (Test-Path $exportDir)) {
            New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
        }

        Write-Host "VBAプロジェクトにアクセスしています..."
        $components = $wb.VBProject.VBComponents
        $count = $components.Count
        Write-Host "VBAコンポーネント数: $count"

        if ($count -eq 0) {
            Write-Host "警告: エクスポート対象のVBAコンポーネントが見つかりません。"
            return
        }

        $exported = 0
        $components | ForEach-Object {
            $extension = Get-VBComponentTypeExtension $_.Type
            if (-not $extension) {
                return  # ForEach-Object 内の return は continue 相当
            }

            # ドキュメントモジュール (Type=100) はコードがある場合のみエクスポート
            if ($_.Type -eq 100) {
                $lineCount = $_.CodeModule.CountOfLines
                if ($lineCount -eq 0) {
                    return
                }
            }

            $path = Join-Path $exportDir "$($_.Name).$extension"
            $_.Export($path)
            Write-Host "  エクスポート: $($_.Name).$extension"
            $exported++
        }

        Write-Host "完了: $exported 個のコンポーネントをエクスポートしました → $exportDir"
    }
    catch {
        Write-Host "エラーが発生しました: $_" -ForegroundColor Red
        if ($_.Exception.Message -match "programmatic access") {
            Write-Host ""
            Write-Host "【対処法】Excel のトラスト センター設定で" -ForegroundColor Yellow
            Write-Host "「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にしてください。" -ForegroundColor Yellow
            Write-Host "  ファイル → オプション → トラスト センター → トラスト センターの設定 → マクロの設定" -ForegroundColor Yellow
        }
        $script:hasError = $true
    }
    finally {
        if ($wb) {
            $wb.Close($false)
        }
        if ($xl) {
            $xl.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# --- メイン処理 ---

$script:hasError = $false

foreach ($filePath in $Path) {
    # 相対パスを絶対パスに変換
    if (-not [System.IO.Path]::IsPathRooted($filePath)) {
        $filePath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $filePath))
    }

    if ([System.IO.Path]::GetExtension($filePath).ToLower() -ne ".xlsm") {
        Write-Host "スキップ: .xlsm ファイルではありません → $filePath" -ForegroundColor Yellow
        continue
    }

    if (-not (Test-Path $filePath)) {
        Write-Host "スキップ: ファイルが見つかりません → $filePath" -ForegroundColor Yellow
        continue
    }

    Export-VBAFromWorkbook -XlsmPath $filePath
}

if ($script:hasError) {
    exit 1
}
