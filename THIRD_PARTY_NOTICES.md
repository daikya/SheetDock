# Third-party notices

SheetDock本体はMITライセンスです。同梱する依存ソフトウェアのライセンスは、それぞれの条件が適用されます。

| ソフトウェア | 用途 | ライセンス情報 |
| --- | --- | --- |
| CPython / Python標準ライブラリ | Python実行環境 | PSF License等。ビルド元PythonのLICENSE.txtを同梱 |
| Tcl / Tk | TkinterのGUI実行環境 | ビルド元の原文、または同じVersionの公式リリースの原文を同梱 |
| tksheet 7.6.0 | 表編集ウィジェット | MIT。インストール済み配布物のライセンス原文を同梱 |
| PyInstaller 6.22.3 | exeのビルド・起動部分 | GPL例外・Apache-2.0等。インストール済み配布物のライセンス文書を同梱 |

具体的なライセンス原文は配布ZIPのlicensesフォルダーにあります。licenses/INDEX.txtには同梱した文書の一覧、BUILD_INFO.jsonにはビルド元のPython・tksheet・PyInstallerのVersionを記録します。Python/Tcl/Tkのソースコードは変更していません。

pytestはテスト専用で、実行時には同梱しません。SheetDock 1.0.0ではopenpyxlを使用しません。

参照：

- https://docs.python.org/3.12/license.html
- https://pyinstaller.org/en/v6.22.3/license.html
- https://github.com/ragardner/tksheet

ローカルに原文がないTcl/Tkは同じVersionの公式リリースから取得します。公式取得分の出典・Version・SHA-256はlicenses/SOURCES.jsonに記録します。必要な文書を確保できない場合はビルドを中止します。
