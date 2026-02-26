# サンプルワークブックの作成手順

動作確認用の `.xlsm` ファイルを手動で作成する手順。

## 手順

1. Excel を起動し、新しいブックを作成する
2. **ファイル → 名前を付けて保存** で、ファイルの種類を **Excel マクロ有効ブック (*.xlsm)** に変更する
3. `SampleMacro.xlsm` としてこの `sample/` ディレクトリに保存する
4. **Alt + F11** で VBA エディタを開く
5. **挿入 → 標準モジュール** で `Module1` を作成する
6. 以下のコードを貼り付ける:

```vb
Option Explicit

Sub SayHello()
    MsgBox "Hello from VBA!"
End Sub

Function Add(a As Long, b As Long) As Long
    Add = a + b
End Function
```

7. **Ctrl + S** で保存する
8. VBA エディタを閉じ、Excel を閉じる

## 動作確認

```powershell
# 手動エクスポート
pwsh scripts/export-vba.ps1 sample/SampleMacro.xlsm

# src/SampleMacro/Module1.bas が生成されたことを確認
ls src/SampleMacro/

# Git でコミットしてみる（lefthook が自動エクスポートする）
git add sample/SampleMacro.xlsm
git commit -m "test: サンプルワークブックを追加"

# src/ も一緒にコミットされていることを確認
git log --stat -1
```
