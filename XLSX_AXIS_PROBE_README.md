# XLSX行列挿入・削除 技術検証

独立検証は590件とWindows実機確認1～6を完了しました。dev7からは本体に組み込んだ処理を呼び出す比較用CLIです。本体GUIの手順はdocs/Phase2_Integration.md末尾を参照してください。
dev12の通常テストは本体747件です。この比較CLIの旧37件を合わせて実行する場合は784件になります。

## 今回の範囲

- 既存範囲内で、複数行・複数列を挿入／削除します。位置はExcelと同じ1始まりです。
- 対象シート全体が1000行×50列以内で、操作後も同じ範囲・ブック全体10万セル以内の場合だけ扱います。1回の件数は1～100です。
- セルの値・型・基本書式、列幅、行高を対応する位置へ移します。挿入部分は空セル・標準の行高／列幅で、隣接書式の自動コピーは行いません。
- 削除範囲内のセル・行高・列幅定義を削除し、後続を詰めます。少なくとも1行・1列を残します。
- 別シートを含め数式が1つでもあるブック、対象シートに結合があるブック、読み取り専用ブックは操作を中止します。
- 対象シートの非表示・高さ0・幅0・グループ化、複数列をまとめた定義等も今回対象外です。他シートの結合は変更しません。
- 末尾への空行列追加は対象外です。完全に空の末尾行列の数はXLSXへの保存・再読込で保持する対象ではありません。

変更は複製したブックで行い、セルの移動結果と他シートの不変、保存・再読込の一致を検証してから履歴へ反映します。
CLIではさらにUndo→元の状態、Redo→操作後の状態を自動検証します。
結果は新しいファイル名へだけ保存し、元ファイルや既存の出力は上書きしません。
失敗時には元ファイルを変更せず、理由を表示して中止します。

これは限定した検証で、全Excel機能を保持できる保証ではありません。
検証対象を広げる場合は、数式・結合・名前定義・テーブル等への対応を別途確認する必要があります。

## 実行

既存のPython 3.12仮想環境.venv312を使い、SheetDockのルートで実行します。

```powershell
python -m pytest tests probes/test_xlsx_axis_probe.py -q --tb=short
python probes/xlsx_axis_probe.py --make-samples probes/xlsx_axis_samples
```

上記の比較テスト込みの期待結果は784件成功です。通常の`python -m pytest`は本体747件です。
サンプル生成は初回のみで、同名ファイルが1つでもあれば上書きせず中止します。

### 4操作

すべて元のaxis_basic.xlsxから別々のコピーを作ります。既存の同名出力がある場合は未使用名へ変えてください。

```powershell
python probes/xlsx_axis_probe.py probes/xlsx_axis_samples/axis_basic.xlsx --action insert-rows --index 2 --count 2 --save-as probes/xlsx_axis_samples/insert_rows.xlsx
python probes/xlsx_axis_probe.py probes/xlsx_axis_samples/axis_basic.xlsx --action delete-rows --index 2 --count 2 --save-as probes/xlsx_axis_samples/delete_rows.xlsx
python probes/xlsx_axis_probe.py probes/xlsx_axis_samples/axis_basic.xlsx --action insert-columns --index 2 --count 2 --save-as probes/xlsx_axis_samples/insert_columns.xlsx
python probes/xlsx_axis_probe.py probes/xlsx_axis_samples/axis_basic.xlsx --action delete-columns --index 2 --count 2 --save-as probes/xlsx_axis_samples/delete_columns.xlsx
```

成功すると「行列操作・Undo/Redo・保存再読込の検証完了」と表示します。

### 操作を中止するケース

```powershell
python probes/xlsx_axis_probe.py probes/xlsx_axis_samples/axis_formula.xlsx --action insert-rows --index 2 --save-as probes/xlsx_axis_samples/rejected_formula.xlsx
python probes/xlsx_axis_probe.py probes/xlsx_axis_samples/axis_merged.xlsx --action insert-rows --index 2 --save-as probes/xlsx_axis_samples/rejected_merged.xlsx
python probes/xlsx_axis_probe.py probes/xlsx_axis_samples/axis_readonly.xlsx --action insert-rows --index 2 --save-as probes/xlsx_axis_samples/rejected_readonly.xlsx
```

この3つは「操作中止」と理由が出るのが正常です。終了コードは1ですが、今回の確認では意図した拒否であり、出力XLSXは作りません。

## 実機確認1～6

本体GUIを起動し、生成したXLSXを開いてセルを確認します。dev7ではGUIにも行列操作ボタンがありますが、この文書ではCLIの出力ファイルを確認します。
dev8では行高も本体GUIへ表示します。太字の見た目は未対応なので、Excel等があればそちらでも確認します。

1. **挿入**：4操作がすべて検証完了になること。insert_rows.xlsxの2～3行目が空でA4=`00123`、A5=`00456`になること。insert_columns.xlsxのB～C列が空で、A2=`00123`、D2=`3`、E2が日付、F2がTRUEになること。
2. **削除**：delete_rows.xlsxではA2=`00789`になること。delete_columns.xlsxはA列がコード、B列が真偽値になり、数量・日付の列が消えること。
3. **型・書式・行高列幅**：先頭ゼロ、数値・日時・真偽値の型、A列の色と背景が保たれること。列挿入では元B列の幅がD列へ、元C列の幅がE列へ移ること。Excel等があれば、行挿入で32の行高が2行目から4行目へ移り、コードの太字も保持されることを確認すること。Excel等がなければ、行高・太字の目視だけは未確認として記録します（自動テストでは検証済み）。
4. **保護**：数式・結合・読み取り専用の3ケースが操作中止になり、rejected_*.xlsxが作成されないこと。数式は操作対象とは別のNotesシートにありますが、その場合も止まること。
5. **元ファイルと他シート**：axis_basic.xlsxが元の4行×4列のまま、各出力のNotesが「他シートは変更しない」のままであること。既存の出力ファイルを指定して再実行すると中止し、その出力も変わらないこと。
6. **保存・履歴検証**：4操作すべてでUndo/Redo・保存再読込の検証完了が出ること。生成ファイルを本体またはExcel等で開けること。本体dev6の通常の編集・保存も従来どおり使えること。

この段階では、CLI内の履歴を検証しています。CLIが出力したファイルを本体GUIで開き直しても、CLIのUndo履歴は引き継ぎません。
