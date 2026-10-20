# SheetDock 1.1.0 Windows正式版ビルド

2026-10-20、dev12のWindowsソース版747件成功（66.26秒）・実機確認1～8、およびWindowsビルド時747件成功（66.53秒）・配布ZIP生成・exe実機確認の完了を確認しました。以下は機能を変更せずVersionと配布文書を1.1.0へ揃えた正式版のビルド手順です。比較用GUIプローブはexeに含めません。

ソースは1.1.0へ変更済みです。このVersionで改めてビルド・配布物確認を行います。既存dev12 ZIPのリネームでは代用しません。公開済み1.0.0は維持します。要件照合と正式化の手順は[Phase2_Review.md](Phase2_Review.md)を参照してください。

## ビルド

Windowsの64bit Python 3.12環境で、プロジェクトのルートから実行します。
新しいPowerShellでは仮想環境を有効にしてください。

```powershell
.\.venv312\Scripts\Activate.ps1
python --version
python -m pip install -e ".[test,build]"
python build_release.py
```

SheetDockや前のexeを終了してからビルドします。
ビルド処理は747件のpytestを実行し、成功した場合だけPyInstallerを起動します。
依存はtksheet 7.6.0、openpyxl 3.1.5、et-xmlfile 2.0.0、PyInstaller 6.22.3です。
フォルダー形式、コンソールなし、UPXなしでビルドします。

## 出力

```text
dist/SheetDock-1.1.0-windows-x64.zip
dist/SheetDock-1.1.0-windows-x64.zip.sha256
```

ZIP内の先頭フォルダーはSheetDock-1.1.0-windows-x64です。
同梱：SheetDock.exe、_internal、LICENSE、THIRD_PARTY_NOTICES.md、DISTRIBUTION_README.md、licenses、BUILD_INFO.json。

配布READMEのVersionは実際のビルドVersionに差し替えます。ソースのDISTRIBUTION_README.mdにある`{{VERSION}}`は差込用の印です。
ユーザーのCSV・XLSX、settings.json、ログ、tests、probes、仮想環境は同梱しません。
確認用サンプルはソース側にある既存ファイルを使います。

ライセンスは実際のビルド環境から収集します。Python/Tcl/Tk/tksheet/PyInstaller/openpyxl/et-xmlfileの文書を同梱します。
Tcl/Tkの原文がローカルにない場合、実行時Versionに対応する公式リリースの原文を取得します。
取得失敗はexeビルド前に中止します。取得元とSHA-256はlicenses/SOURCES.jsonに残します。インターネット接続が必要になる場合があります。

新しいビルド・ライセンス収集・ZIP検証が完了するまで、同じ名前の既存ZIPを置換しません。
正式版1.1.0のZIPは1.0.0・dev12とは別名です。

## 今回の正式配布物確認1～6

dev12の機能確認は完了しています。今回は機能変更がないため、正式Versionと配布物の確認、および主要操作の確認を行います。下の詳細1～8は不具合が出た場合や再確認時に利用できます。

1. **Version・起動**：正式ZIPを新しいフォルダーへ完全に展開し、SheetDock.exeを起動する。タイトルと配布READMEが1.1.0で、開発版表記や`{{VERSION}}`が残らないこと。アイコンとコンソールなし起動も確認する。
2. **同梱物**：`BUILD_INFO.json`のsheetdockが1.1.0、tksheetが7.6.0、openpyxlが3.1.5、et-xmlfileが2.0.0、pyinstallerが6.22.3であること。LICENSE、THIRD_PARTY_NOTICES.md、licenses内の各ライセンスがあり、手元のデータ・設定・ログが混入していないこと。
3. **CSV**：サンプルのコピーを使い、UTF-8とCP932を開く。日本語・先頭ゼロを確認し、セル編集・Undo/Redo・別名保存・再読込ができること。
4. **XLSX**：サンプルのコピーを開き、直接編集・入力中の保存・Undo/Redo・シート切替・別名保存・再読込ができること。数式セル・結合範囲・読み取り専用が編集できず、CSV出力ができること。
5. **共通操作**：CSVとXLSXを混在させ、編集内容と履歴が別タブに漏れないこと。未保存のまま閉じると保存・破棄・キャンセルが選べ、再起動後に画面位置が復元されること。
6. **ハッシュ**：下記Get-FileHashの値と`.zip.sha256`が一致すること。別PCで実施する場合はPythonの有無も記録する。別PCやExcel等を使わない場合は、その環境での動作を未確認として扱う。

## exe版の詳細確認1～8

ZIPを新しいフォルダーへ完全に展開します。ZIP内で直接実行せず、exeだけを切り離さないでください。
ソース版とexe版は同じLOCALAPPDATAの設定を使います。既存設定を消す必要はありません。

1. **起動と同梱物**：SheetDock.exeがコンソールなしで起動し、タイトルに1.1.0とアイコンが表示されること。配布READMEにも同じVersionが入り、`{{VERSION}}`や「XLSXは未対応」という旧説明が残らないこと。BUILD_INFO.jsonにopenpyxl・et-xmlfileのVersion、licensesに両ライセンスがあること。
2. **CSV**：UTF-8／CP932のCSVを開き、日本語・先頭ゼロ・セル内改行、セル編集・行列操作、列幅のUndo/Redoが動くこと。上書き・別名保存と再読込ができること。
3. **XLSX表示とシート**：basic.xlsx、display_formats.xlsx、sheet_operations.xlsxを開き、表示書式・列幅・シートタブがソース版どおりであること。条件が許すブックでシート追加・名前変更・削除とUndo/Redoが動くこと。XLSXの「行追加」「列追加」、追加セルの編集・貼付け、空の追加範囲を含む保存・再読込も確認すること。
4. **XLSX編集・検索・コピー**：通常セルのF2・ダブルクリック・文字キーでの直接編集、Enter／Tab確定とEsc取消、入力中と確定後のUndo/Redo、表示範囲への貼付け、前後検索、1件／一括置換とUndo/Redoが動くこと。Excel等が使える場合は日本語・空セル・引用符・セル内改行を含む2行×2列以上の相互コピーも確認すること。Excel側で先頭ゼロを確認する列は文字列書式にすること。
5. **保存と外部変更**：XLSXの検証用コピーを新規名へ別名保存→再編集→上書きし、バックアップ設定、列幅、表示書式、再読込が動くこと。外部変更時に上書き・再読込・キャンセルを選べること。既存のファイルロック手順で保存を失敗させ、元ファイル・未保存編集・Undoが残ること。
6. **保護とCSV出力**：数式・結合範囲・エラーの編集制限と、validation_readonly.xlsxの読み取り専用判定が働くこと。選択シートのCSV出力ができ、元ブックの未保存状態が変わらないこと。数式出力は初期中止・明示選択時のみ許可されること。
7. **共通操作と再起動**：CSVとXLSXを混在させ、保存・Undo/Redoが選択ファイルにだけ作用すること。未保存タブの保存／破棄／キャンセル、画面位置・サブ画面位置・最近使ったファイル・CSV列幅の復元も確認すること。XLSX列幅は保存したブックから復元されること。XLSX入力中の保存・ファイル／シート切替・閉じるで先に入力を確定し、不正入力では操作を止めて元のタブへ戻ること。
8. **ハッシュ・別PC**：Get-FileHashのSHA-256と同梱.sha256の値が一致すること。可能ならPython未導入の別Windows PCでも起動し、CSVとXLSXを開けること。別PCがない場合は「未確認」と記録し、開発PCでの起動と区別すること。

```powershell
Get-FileHash .\dist\SheetDock-1.1.0-windows-x64.zip -Algorithm SHA256
Get-Content .\dist\SheetDock-1.1.0-windows-x64.zip.sha256
```

dev12のビルド・exe確認は完了しています。別PC起動とExcel等との相互コピーは個別環境の報告がないため、実証済みとは記録していません。このスクリプトはGitHubへのアップロードや公開を行いません。
Linuxの自動テストはWindows exeのビルド成功を意味しません。
