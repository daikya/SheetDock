# SheetDock Windowsビルド・確認

## 現在の状態（2026-10-20）

公開済みの正式版は1.1.0です。検証済み開発版1.2.0.dev2はWindowsソース894件成功（90.24秒）、ビルド時894件成功（89.43秒）、exe実機1～8・保存失敗ロック確認まで完了しました。元CSV／XLSXは不変、開発ZIPのSHA-256も一致しています。

現在のソースは正式リリース準備用の1.2.0です。機能は検証済みdev2と同じで、Versionと配布説明だけを更新しました。次は正式ZIPを新しくビルドし、以下の「正式配布物確認1～6」を行います。公開済み1.0.0・1.1.0と検証済みdev2の成果物は維持します。[Phase3_Review.md](Phase3_Review.md)に確認結果・制限を記録しています。

## ビルド

SheetDockと独立検証画面を終了し、プロジェクト直下で`.venv312`を有効にして実行します。

```powershell
python --version
python -m pip install -e ".[test,build]"
python build_release.py
```

Windowsの64bit Python 3.12を使います。新しいPowerShellなら先に`.\.venv312\Scripts\Activate.ps1`を実行してください。

ビルド処理は通常894件のpytest成功後にPyInstallerを起動します。依存はtksheet 7.6.0、openpyxl 3.1.5、et-xmlfile 2.0.0、PyInstaller 6.22.3。フォルダー形式、コンソールなし、UPXなしです。比較用probesはexeに含めません。

## 出力と展開

```text
dist/SheetDock-1.2.0-windows-x64.zip
dist/SheetDock-1.2.0-windows-x64.zip.sha256
```

ZIP内の先頭フォルダーはSheetDock-1.2.0-windows-x64です。同梱はSheetDock.exe、_internal、LICENSE、THIRD_PARTY_NOTICES.md、DISTRIBUTION_README.md、licenses、BUILD_INFO.jsonです。

ZIPを新しい場所へ完全に展開してSheetDock.exeを起動します。ZIP内から直接起動せず、exeだけを移動しないでください。Pythonで起動した本体や旧exeを終了し、今回展開したexeで確認します。ソース版とexe版は同じLOCALAPPDATAの設定を使います。既存設定を削除する必要はありません。

READMEの`{{VERSION}}`はビルド時に実際のVersionへ差し替わります。ユーザーのCSV・XLSX、settings.json、ログ、tests、probes、仮想環境は同梱しません。Python/Tcl/Tk/tksheet/PyInstaller/openpyxl/et-xmlfileのライセンス原文を同梱します。Tcl/Tkの原文がローカルになければ実行時Versionに対応する公式リリースから取得し、取得元・Version・SHA-256をlicenses/SOURCES.jsonへ記録します。取得・検証の失敗はZIP完成前に中止します。既存ZIPは新しい成果物の検証完了まで置換しません。

## 正式配布物確認1～6（今回）

開発版で機能の詳細確認は完了しています。今回は正式ZIPを新しい場所へ完全に展開し、旧exeを終了して、正式版のexeで以下を確認してください。サンプルや設定を削除・再生成する必要はありません。元サンプルへは保存せず、未使用名の検証用コピーを使います。

```powershell
$releaseCsvSourceHash = (Get-FileHash probes/sort_filter_samples/sort_basic.csv -Algorithm SHA256).Hash
$releaseXlsxSourceHash = (Get-FileHash probes/xlsx_sort_filter_samples/sort_basic.xlsx -Algorithm SHA256).Hash
```

1. **Version・起動**：タイトルと配布READMEが1.2.0で開発版表記や`{{VERSION}}`が残らないこと。アイコンが表示され、コンソールなしで起動すること。
2. **同梱物**：BUILD_INFO.jsonのsheetdockが1.2.0、依存Versionが上記の固定値であること。LICENSE・THIRD_PARTY_NOTICES.md・配布README・licensesの原文とSOURCES.jsonがあり、ユーザーのデータ・設定・ログが混入していないこと。
3. **CSV**：UTF-8／CP932を開き、日本語・先頭ゼロを確認。sort_basic.csvで数値順の並べ替え・Undo/Redo・東京の絞り込み・本文編集が動くこと。絞り込み中と0件表示の別名保存後、コピーを開き直すと全行と保存形式を保持すること。
4. **XLSX**：sort_basic.xlsxで型順の並べ替え・Undo/Redo・東京の絞り込み・F2直接編集が動くこと。別名保存・再読込で全行・他シート・型・書式・行高・列幅を保持し、絞り込み中のCSV出力が全行を含むこと。コピー上の上書き保存・バックアップと、数式・結合・読み取り専用の編集保護も確認すること。
5. **混在と終了**：CSVとXLSXで編集・履歴・条件が混ざらず、未保存の保存／破棄／キャンセルが動くこと。入力中の保存・タブ切替も元のセルへ確定し、再起動で画面位置等は復元され、フィルター条件は解除されること。
6. **元ファイル・ZIPハッシュ**：下記で元CSV／XLSXが両方True、正式ZIPのSHA-256と.sha256が一致すること。別PCやExcel等との個別環境を追加確認した場合は使用環境も記録すること。

```powershell
(Get-FileHash probes/sort_filter_samples/sort_basic.csv -Algorithm SHA256).Hash -eq $releaseCsvSourceHash
(Get-FileHash probes/xlsx_sort_filter_samples/sort_basic.xlsx -Algorithm SHA256).Hash -eq $releaseXlsxSourceHash
Get-FileHash .\dist\SheetDock-1.2.0-windows-x64.zip -Algorithm SHA256
Get-Content .\dist\SheetDock-1.2.0-windows-x64.zip.sha256
```

以下の詳細1～8とロック確認はdev2で確認済みです。機能変更や不具合修正が発生した場合は、関係する項目を再確認します。

## 機能の詳細確認1～8（修正時の再確認用）

既存の`probes/sort_filter_samples`と`probes/xlsx_sort_filter_samples`を使用し、サンプルを再生成しないでください。元サンプルは上書きせず、保存・ロック確認には別名保存した検証用コピーを使います。元からやり直す場合はタブを閉じて破棄し、元ファイルを開き直します。

exe起動前に、プロジェクト直下のPowerShellで比較元を記録します。

```powershell
$exeCsvSourceHash = (Get-FileHash probes/sort_filter_samples/sort_basic.csv -Algorithm SHA256).Hash
$exeXlsxSourceHash = (Get-FileHash probes/xlsx_sort_filter_samples/sort_basic.xlsx -Algorithm SHA256).Hash
```

1. **起動・同梱物**：タイトルと配布READMEが1.2.0で正式版と分かり、`{{VERSION}}`が残らないこと。アイコン・コンソールなし起動、BUILD_INFO.jsonのsheetdock・依存Version、各ライセンス原文とSOURCES.jsonを確認すること。個人データ・設定・ログがZIPに混入していないこと。
2. **CSV並べ替え・フィルター**：sort_basic.csvをUTF-8・カンマ・見出しありで開く。列2の文字列昇順は00005→00123→00002→00003→00004、数値昇順は00005→00002→00003→00123→00004。降順・同値順・空セル末尾を保ち、1回でUndo/Redoできること。列3・東京は本文3行、空条件と0件表示も可能で、絞り込みだけでは未保存にならないこと。CP932の既存サンプルも開けること。
3. **CSV編集・全行保存**：東京の行を直接編集し、条件から外れると非表示になり、ボタン・Ctrl+Z／Ctrl+Shift+Zで戻せること。非連続表示行への範囲貼付けが1操作で、範囲超過は全体を拒否すること。入力中の保存と0件表示での別名保存が非表示行を含む全行を保持し、コピーを開き直して日本語・先頭ゼロ・改行・文字コード情報を確認できること。絞り込み中の行列操作等は解除を要求すること。
4. **XLSX並べ替えと保護**：sort_basic.xlsxの列2・見出しあり昇順は00005→00002→00003→00123→00004。降順・同値順・空セル末尾を保ち、値・型・書式・行高も行に追従すること。列幅・Otherは不変で、Undo/Redoが動くこと。formula・merged・readonly・mixed・hidden・errorの並べ替えは理由を示して中止し、数式・結合範囲・エラーセルは編集できないこと。読み取り専用でもフィルターは使えること。
5. **XLSXフィルター・直接入力**：列5・東京で本文行番号2・4・6と見出し行1を表示。F2／ダブルクリックでA4を編集し、E4を北海道へ変えると表示から外れ、Undo/Redoで再評価すること。数値の編集開始は書式を外した元の値、abcは拒否。E2起点に2行×2列を貼り付けると元の2行と4行が対象で、隠れた3行と5行は不変。見出し編集は無効、コピーは表示値のみ。行列操作・検索置換・シート構造変更・行高変更の制限も働くこと。
6. **XLSX全ブック保存・失敗保護**：フィルター中と本文0件で、未使用名へ別名保存する。開き直すと全本文・見出し・Other・型・書式・寸法が残り、条件は解除されること。入力中保存、検証用コピーの上書き保存・バックアップも確認すること。CSV出力は絞り込みによらず選択シート全行を含み、元XLSXの未保存状態を変えないこと。コピーをロックして保存を失敗させても未保存編集・履歴・条件を保ち、ロック解除後に保存できること。外部変更時の上書き／再読込／キャンセルと既存保存先への別名保存拒否も確認すること。
7. **混在タブ・条件・終了**：CSVと複数XLSXを混在させ、条件と履歴がファイル・シート別であること。入力中のファイル／シート切替は元のセルへ確定し、不正入力では移動を止めること。文字入力欄のCtrl+Zで表のモデル履歴を操作しないこと。未保存の保存／破棄／キャンセル、画面位置・最近使ったファイル・列幅の復元が動くこと。再起動後は条件を保持せず、過去の行列・シート構造を履歴復元して範囲が変わった場合も条件を解除すること。
8. **ハッシュ・環境の記録**：下記の元サンプル比較が両方Trueで、ZIPのSHA-256が.sha256と一致すること。Python未導入の別Windows PCやExcel等が使える場合は、起動・CSV/XLSX読込や日本語・空セル・引用符・セル内改行を含む相互コピーも確認し、使用環境を記録すること。実施しない環境は未確認として扱うこと。

```powershell
(Get-FileHash probes/sort_filter_samples/sort_basic.csv -Algorithm SHA256).Hash -eq $exeCsvSourceHash
(Get-FileHash probes/xlsx_sort_filter_samples/sort_basic.xlsx -Algorithm SHA256).Hash -eq $exeXlsxSourceHash
Get-FileHash .\dist\SheetDock-1.2.0-windows-x64.zip -Algorithm SHA256
Get-Content .\dist\SheetDock-1.2.0-windows-x64.zip.sha256
```

貼付けに使う2行×2列の例です。XLSXのE2:F2とE4:F4を確認する際に使えます。CSVでは列構成が違うため、数量と区分を対象に数値と区分の組を使ってください。

```text
大阪	first
京都	second
```

### 保存失敗のロック確認

先に別名保存で検証用コピーを作ってください。別のPowerShellで、その保存済みコピーの実際のフルパスを入力します。元サンプルはロックしません。

```powershell
$lockPath = Read-Host '保存済み検証用コピーのフルパス'
$fileLock = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
```

exe画面でコピーを編集して上書き保存し、エラー後も編集・履歴・条件が残ることを確認します。確認後は同じPowerShellで解除し、もう一度保存します。

```powershell
if ($null -ne $fileLock) { $fileLock.Dispose(); $fileLock = $null }
```

## 確認後

正式ビルドログ、配布物確認1～6、元ファイル比較・ZIPハッシュの結果を記録します。問題がなければGitの差分とコミット対象を確認し、コミット・push、v1.2.0のReleaseへ正式ZIPと.sha256の添付・公開を行います。開発ZIPの改名だけで正式化しません。このビルドスクリプトはGitHubへのアップロードや公開を行いません。

開発版dev2のWindows exeは確認済みです。正式1.2.0のビルド・配布物確認・公開は未実施です。別PC・Excel等との個別環境も未確認であり、開発PCでの総合確認と区別します。
