# CSV並べ替え・フィルター：独立技術検証

このCLIは比較用です。1.2.0.dev1でCSV本体GUIへ接続しました。今回は通常の本体GUIを確認します。
並べ替えは行全体を移動してUndo/Redo対象とし、フィルターは表示対象の行番号だけを返します。詳細は[docs/Phase3_Validation.md](docs/Phase3_Validation.md)を参照してください。

2026-10-20：Windowsで778件成功（64.75秒）、以下の確認結果と元ファイル・既存出力のハッシュ不変を確認しました。次の[独立GUI検証](CSV_SORT_FILTER_GUI_PROBE_README.md)では、表示行の編集・範囲貼付け・全行保存を確認します。以下のCLI手順は引き続き比較に利用できます。

## 実行環境

Windowsの`.venv312`を有効にし、SheetDock直下で実行します。追加パッケージはありません。

```powershell
python -m pytest tests probes/test_csv_sort_filter_probe.py -q --tb=short
```

期待値は **809件成功**（本体805＋比較プローブ4）。通常の`python -m pytest -q`は本体805件です。モデル27件は本体へ移し、CLIと別名コピーの比較4件を残しました。

## Windows確認1～6

### 1. サンプル生成

```powershell
python probes/csv_sort_filter_probe.py --make-samples probes/sort_filter_samples
$sourceHash = (Get-FileHash probes/sort_filter_samples/sort_basic.csv -Algorithm SHA256).Hash
```

sort_basic.csvとsort_invalid.csvの2ファイルを生成します。同名があれば中止するので、再実行時は別フォルダーを指定してください。UTF-8 BOM・CRLFのサンプルです。元サンプルは変更せずに使います。

### 2. 文字列順と別名保存

```powershell
python probes/csv_sort_filter_probe.py probes/sort_filter_samples/sort_basic.csv --sort-column 2 --save-as probes/sort_filter_samples/text_copy.csv
```

見出しが先頭に残り、コード順が00005→00123→00002→00003→00004になること。文字列比較なので数量10が2より前です。空セルは最後、同値2の2行は元の順序のままです。Undo/Redoと別名保存・再読込の検証完了が表示されること。先頭ゼロ・日本語・セル内改行・引用符が残ること。

### 3. 数値順と降順

```powershell
python probes/csv_sort_filter_probe.py probes/sort_filter_samples/sort_basic.csv --sort-column 2 --mode number --save-as probes/sort_filter_samples/number_copy.csv
python probes/csv_sort_filter_probe.py probes/sort_filter_samples/sort_basic.csv --sort-column 2 --mode number --descending
```

昇順のコードは00005→00002→00003→00123→00004、降順は00123→00002→00003→00005→00004です。降順でも同値順序を保ち、空セルは最後です。並べ替え後は未保存「あり」。このCLIの別名保存は検証用コピーの出力で、編集モデルの保存済み基準は変更しません。

### 4. フィルターと無変更

```powershell
python probes/csv_sort_filter_probe.py probes/sort_filter_samples/sort_basic.csv --filter-column 3 --query 東京
python probes/csv_sort_filter_probe.py probes/sort_filter_samples/sort_basic.csv --filter-column 2 --operator empty
python probes/csv_sort_filter_probe.py probes/sort_filter_samples/sort_basic.csv --filter-column 3 --operator equals --query 該当なし
python probes/csv_sort_filter_probe.py probes/sort_filter_samples/sort_basic.csv
```

順に、見出し＋東京の3行、見出し＋00004の1行、見出しだけ、見出し＋全5行が表示されること。すべて未保存「なし」です。表示の「データ行N」はフィルター適用前の現在のデータ位置（見出し込み1始まり）です。並べ替えも指定した場合は、並べ替え後のデータ位置になります。

### 5. 不正数値・フィルター中の保存を拒否

```powershell
python probes/csv_sort_filter_probe.py probes/sort_filter_samples/sort_invalid.csv --sort-column 2 --mode number --save-as probes/sort_filter_samples/rejected_number.csv
python probes/csv_sort_filter_probe.py probes/sort_filter_samples/sort_basic.csv --filter-column 3 --query 東京 --save-as probes/sort_filter_samples/rejected_filter.csv
Test-Path probes/sort_filter_samples/rejected_number.csv
Test-Path probes/sort_filter_samples/rejected_filter.csv
```

前者は数値として比較できない行、後者はフィルター解除が必要という操作中止になり、どちらもFalseになること。これらのコマンドの終了コード1は想定した拒否です。元データは変更しません。

### 6. 既存出力と元サンプルを保護

```powershell
$copyHash = (Get-FileHash probes/sort_filter_samples/number_copy.csv -Algorithm SHA256).Hash
python probes/csv_sort_filter_probe.py probes/sort_filter_samples/sort_basic.csv --sort-column 2 --mode number --save-as probes/sort_filter_samples/number_copy.csv
(Get-FileHash probes/sort_filter_samples/number_copy.csv -Algorithm SHA256).Hash -eq $copyHash
(Get-FileHash probes/sort_filter_samples/sort_basic.csv -Algorithm SHA256).Hash -eq $sourceHash
```

既存出力を拒否し、両方Trueになること。サンプルフォルダーはGitの対象外です。このCLIは元ファイルを上書きしません。XLSX並べ替え・フィルターはまだ未対応です。本体の確認はdocs/Phase3_Validation.md末尾を参照してください。
