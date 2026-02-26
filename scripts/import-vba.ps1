<#
.SYNOPSIS
    VBA ソースコードを Excel ワークブックにインポートする。

.DESCRIPTION
    src/<ブック名>/ ディレクトリにある .bas/.cls/.frm ファイルを
    指定された .xlsm ファイルの VBA プロジェクトにインポートする。
    ドキュメントモジュール（ThisWorkbook, Sheet等）はコード差し替え、
    標準モジュール・クラスモジュールは削除→再インポートで処理する。

.PARAMETER Path
    インポート先の .xlsm ファイルパス（1つ以上）。

.EXAMPLE
    .\import-vba.ps1 MyWorkbook.xlsm
#>
param (
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Path
)

if (-not $Path -or $Path.Count -eq 0) {
    Write-Host "対象ファイルなし（スキップ）"
    exit 0
}

$sjis = [System.Text.Encoding]::GetEncoding(932)

function Get-VBComponentName {
    param([string]$FilePath)

    $lines = [System.IO.File]::ReadAllLines($FilePath, $sjis)
    foreach ($line in $lines) {
        if ($line -match 'Attribute VB_Name = "(.+)"') {
            return $Matches[1]
        }
    }
    return [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
}

function Get-CodeFromFile {
    param([string]$FilePath)

    $lines = [System.IO.File]::ReadAllLines($FilePath, $sjis)
    $codeStartIndex = 0
    $inBeginBlock = $false

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]

        # VERSION 行をスキップ
        if ($line -match "^VERSION ") {
            $codeStartIndex = $i + 1
            continue
        }

        # BEGIN...END ブロックをスキップ
        if ($line -match "^BEGIN$") {
            $inBeginBlock = $true
            $codeStartIndex = $i + 1
            continue
        }
        if ($inBeginBlock) {
            $codeStartIndex = $i + 1
            if ($line -match "^END$") {
                $inBeginBlock = $false
            }
            continue
        }

        # Attribute 行をスキップ
        if ($line -match "^Attribute ") {
            $codeStartIndex = $i + 1
            continue
        }

        # ヘッダでない行に到達したら終了
        break
    }

    if ($codeStartIndex -ge $lines.Count) {
        return ""
    }

    $codeLines = $lines[$codeStartIndex..($lines.Count - 1)]
    return ($codeLines -join "`r`n")
}

function Import-VBAToWorkbook {
    param([string]$XlsmPath)

    $xl = $null
    $wb = $null

    try {
        # ソースディレクトリの特定
        $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($XlsmPath)
        $repoRoot = Split-Path $PSScriptRoot -Parent
        $srcDir = Join-Path $repoRoot "src" $fileNameWithoutExt

        if (-not (Test-Path $srcDir)) {
            Write-Host "エラー: ソースディレクトリが見つかりません: $srcDir" -ForegroundColor Red
            $script:hasError = $true
            return
        }

        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $xl.AutomationSecurity = 3  # msoAutomationSecurityForceDisable

        Write-Host "ファイルを開いています: $XlsmPath"
        $wb = $xl.Workbooks.Open($XlsmPath)

        Write-Host "ソースディレクトリ: $srcDir"

        $sourceFiles = Get-ChildItem -Path $srcDir -File | Where-Object {
            $_.Extension -in ".bas", ".cls", ".frm"
        }

        if ($sourceFiles.Count -eq 0) {
            Write-Host "警告: インポート対象のソースファイルが見つかりません。"
            return
        }

        Write-Host "インポート対象ファイル数: $($sourceFiles.Count)"

        $components = $wb.VBProject.VBComponents
        $imported = 0

        foreach ($file in $sourceFiles) {
            $componentName = Get-VBComponentName $file.FullName
            Write-Host "  処理中: $($file.Name) (コンポーネント名: $componentName)"

            # 既存コンポーネントの確認
            $existing = $null
            try {
                $existing = $components.Item($componentName)
            }
            catch {
                $existing = $null
            }

            if ($existing -and $existing.Type -eq 100) {
                # ドキュメントモジュール (ThisWorkbook, Sheet等) はコード差し替え
                Write-Host "    ドキュメントモジュール → コード差し替え"
                $codeModule = $existing.CodeModule

                # 既存コードを全削除
                if ($codeModule.CountOfLines -gt 0) {
                    $codeModule.DeleteLines(1, $codeModule.CountOfLines)
                }

                # ファイルからコード部分のみ取得して注入
                $code = Get-CodeFromFile $file.FullName
                if ($code -and $code.Trim().Length -gt 0) {
                    $codeModule.AddFromString($code)
                }
            }
            else {
                # 標準モジュール・クラスモジュール・フォーム → 削除して再インポート
                if ($existing) {
                    Write-Host "    既存コンポーネントを削除"
                    $components.Remove($existing)
                }

                Write-Host "    インポート中: $($file.FullName)"
                $components.Import($file.FullName) | Out-Null
            }

            $imported++
        }

        # ブックを保存
        Write-Host "ブックを保存しています..."
        $wb.Save()

        Write-Host "完了: $imported 個のコンポーネントをインポートしました ← $srcDir"
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

    Import-VBAToWorkbook -XlsmPath $filePath
}

if ($script:hasError) {
    exit 1
}
