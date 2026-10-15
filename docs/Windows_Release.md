# SheetDock 1.0.0 Windowsビルド

Python 3.12の64bit仮想環境で、プロジェクトのルートから実行します。

```powershell
python -m pip install tksheet==7.6.0 pytest pyinstaller==6.22.3
python -m pytest -q
python -m sheetdock
```

ソース版の画面タイトルがSheetDock 1.0.0であることを確認し、アプリを終了します。

```powershell
python build_release.py
```

ビルド処理はpytestをもう一度実行し、成功した場合にだけPyInstallerを起動します。フォルダー形式、コンソールなし、UPXなしでビルドします。

## 出力

```text
dist/SheetDock-1.0.0-windows-x64.zip
dist/SheetDock-1.0.0-windows-x64.zip.sha256
```

ZIP内の先頭フォルダー：SheetDock-1.0.0-windows-x64。

同梱：SheetDock.exe、_internal、LICENSE、THIRD_PARTY_NOTICES.md、DISTRIBUTION_README.md、licenses、BUILD_INFO.json。

ユーザーのCSV、settings.json、ログ、tests、probes、仮想環境は同梱しません。licensesは実際のビルド環境から集めます。Python/Tcl/Tk/tksheet/PyInstallerの原文が見つからない場合、実行時のTcl/TkのVersionに対応する公式リリースの文書を取得します。取得失敗はexeビルド前に中止します。取得元とSHA-256はlicenses/SOURCES.jsonに残します。インターネット接続が必要になる場合があります。

以前の配布ZIPは、新しいビルド・ライセンス収集・ZIP検証が完了するまで置換しません。

## exe版の確認

ZIPを新しいフォルダーへ展開します。ZIP内で直接実行せず、exeだけを切り離さないでください。

1. SheetDock.exeがコンソールを出さずに起動し、Version 1.0.0と新しいアイコンを表示する。
2. UTF-8／CP932のCSVを開き、日本語・先頭ゼロ・セル内改行を確認する。
3. 複数タブ、行列操作、コピー・貼り付け、検索・置換、Undo/Redoを確認する。
4. 上書き・別名保存、バックアップ設定、外部変更確認が動作する。
5. ウィンドウ・サブ画面位置、列幅、最近使ったファイルを再起動後に復元する。
6. エラー・警告ログと大容量CSV警告を確認する。
7. Python未導入環境での起動は、可能なら別のWindowsマシンで確認する。同じ開発PCでの起動だけでは未導入環境の確認完了とはしない。

ソース版とexe版は同じLOCALAPPDATAの設定を使います。既存設定を消す必要はありません。

exe版の確認後に、SHA-256と配布物を確認してGitHub Releaseへ進みます。このスクリプトはGitHubへのアップロードや公開を行いません。

## ソース版確認の記録

2026年10月13日：235件のpytest、段階ごとのWindows確認、保存失敗時の保護、本体最終版でのExcel相互コピーの6項目が成功。
2026年10月15日：Version 1.0.0・MIT・配布処理を追加。exe版のWindows確認はこれから行う。

ビルドは実行OSに依存します。Linux上でWindows exeを作成したことにはしません。

参照： https://pyinstaller.org/en/v6.22.3/operating-mode.html
