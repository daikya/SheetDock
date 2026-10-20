# XLSX末尾追加の独立プローブ

2026-10-20。独立検証はWindowsで695件成功（59.90秒）と確認1～6を完了しました。dev11で追加処理を本体へ移しました。このCLIは本体と同じ処理を呼び出す比較用です。

本体GUIの「行追加」「列追加」の確認は[docs/Phase2_Integration.md](docs/Phase2_Integration.md)dev11確認1～6の節を使います。以下はCLI単体の比較手順として残します。

## 検証した方式

値も書式もないセルを`sheet.cell()`で作るだけでは、保存・再読込でその空の末尾範囲が消えます。
そこで、追加行には標準の行高、追加列には標準の列幅を明示した定義をXLSX内に保存します。値・仮の文字列・余分なセル書式は入れません。JSONも作りません。

表示範囲は、値・セル書式・結合による使用範囲と、行列定義の最大位置を合わせて算出します。描画時に確保した空セルは数えません。通常のopenpyxl保存を繰り返しても、この方式で追加範囲を復元できます。一方、`max_row/max_column`自体は空行列の分だけ増えるわけではありません。CLIは両方の範囲を表示します。

- 追加後の行列定義、値、型、セル書式、結合、他シートを比較してから変更を確定します。
- 1回の追加を1回のUndo/Redoにします。別名保存・再読込でも内容と表示範囲を比較します。保存失敗時は元ファイル・編集中の追加・履歴を残します。
- 追加前の範囲外にある行高・列幅の定義も尊重し、その外側へ追加します。既存定義を上書きしません。
- 新規の行高はシート標準値、未指定なら15ポイントです。列幅はシート標準値を使い、未指定なら既存の列幅換算（数字幅7ピクセルを仮定）を利用します。最後のデータ行・列の書式をコピーしません。
- 空のシートは1行×1列から始めます。2行追加すれば表示範囲は3行×1列です。
- 数式がブック内にある場合、対象シートに結合・非表示・グループ等がある場合、読み取り専用の場合は追加を拒否します。1回1～100件、対象シート1000行×50列、表示範囲を含めたブック全体10万セルまでです。
- Excelのグリッドそのものを増やす処理ではなく、SheetDockで表示・編集対象にする範囲を保持する検証です。他のソフトが行列定義を削除・正規化した後まで範囲保持を保証するものではありません。

## dev11の本体接続

表示範囲とセルの使用範囲を分けて算出し、追加セルの編集・貼り付け・寸法変更・行列操作に接続しました。表示するために空セルを取得しただけでは表示範囲や未保存状態を変えません。検索は値を対象とし、CSV出力は行高・列幅だけの空余白を除きます。追加範囲に値を入れればその位置まで出力します。
GUIは標準フォントを測って列幅を換算し、CLIは既定の数字幅7ピクセルで換算します。行列定義を外部ソフトが削除した後まで範囲保持を保証するものではありません。
モデルテスト28件を`tests/test_xlsx_append.py`へ移し、本体接続の追加テストも通常pytestへ含めました。`probes/test_xlsx_append_probe.py`はCLI比較4件のみです。本体からprobesは参照しません。

## Windowsでの確認1～6

プロジェクトルートで`.venv312`を使います。確認のためにexeを再ビルドする必要はありません。

### 1. テストとサンプル生成

```powershell
python -m pytest tests probes/test_xlsx_append_probe.py -q --tb=short
python probes/xlsx_append_probe.py --make-samples probes/xlsx_append_samples
```

比較実行の期待結果は**751件成功（本体747＋CLI比較4）**です。通常の`python -m pytest`は本体747件です。
サンプルはbasic、empty、formula、merged、readonlyの5ファイルです。同名サンプルがある場合は上書きせず中止します。再実行時は別のサンプルフォルダーを指定してください。

### 2. 空行を2行追加

```powershell
$sourceHash = (Get-FileHash probes/xlsx_append_samples/append_basic.xlsx -Algorithm SHA256).Hash
python probes/xlsx_append_probe.py probes/xlsx_append_samples/append_basic.xlsx --axis rows --count 2 --save-as probes/xlsx_append_samples/rows_copy.xlsx
```

表示範囲は**4行×2列**、セル範囲は2行×2列。「末尾追加・Undo/Redo・保存再読込の検証完了」が出ることを確認します。先頭ゼロ・型・元の書式・他シートも自動比較しています。

### 3. 空列を2列追加

```powershell
python probes/xlsx_append_probe.py probes/xlsx_append_samples/append_basic.xlsx --axis columns --count 2 --save-as probes/xlsx_append_samples/columns_copy.xlsx
```

表示範囲は**2行×4列**、セル範囲は2行×2列。同じ検証完了メッセージを確認します。

### 4. 再読込したファイルへの追加と空シート

```powershell
python probes/xlsx_append_probe.py probes/xlsx_append_samples/rows_copy.xlsx --axis columns --count 2 --save-as probes/xlsx_append_samples/both_copy.xlsx
python probes/xlsx_append_probe.py probes/xlsx_append_samples/append_empty.xlsx --axis rows --count 2 --save-as probes/xlsx_append_samples/empty_copy.xlsx
```

表示範囲は順に**4行×4列**、**3行×1列**。両方で検証完了することを確認します。

### 5. 保護対象の拒否と元ファイル不変

```powershell
python probes/xlsx_append_probe.py probes/xlsx_append_samples/append_formula.xlsx --axis rows --save-as probes/xlsx_append_samples/rejected_formula.xlsx
python probes/xlsx_append_probe.py probes/xlsx_append_samples/append_merged.xlsx --axis rows --save-as probes/xlsx_append_samples/rejected_merged.xlsx
python probes/xlsx_append_probe.py probes/xlsx_append_samples/append_readonly.xlsx --axis rows --save-as probes/xlsx_append_samples/rejected_readonly.xlsx
(Get-FileHash probes/xlsx_append_samples/append_basic.xlsx -Algorithm SHA256).Hash -eq $sourceHash
```

3件とも理由付きの「操作中止」、出力ファイルなしを確認します。最後のハッシュ比較はTrueです。拒否時の終了コード1は想定動作です。

### 6. 既存出力の上書き拒否と件数上限

```powershell
$copyHash = (Get-FileHash probes/xlsx_append_samples/rows_copy.xlsx -Algorithm SHA256).Hash
python probes/xlsx_append_probe.py probes/xlsx_append_samples/append_basic.xlsx --axis rows --count 2 --save-as probes/xlsx_append_samples/rows_copy.xlsx
(Get-FileHash probes/xlsx_append_samples/rows_copy.xlsx -Algorithm SHA256).Hash -eq $copyHash
python probes/xlsx_append_probe.py probes/xlsx_append_samples/append_basic.xlsx --axis rows --count 101 --save-as probes/xlsx_append_samples/rejected_count.xlsx
```

既存出力の上書き拒否、ハッシュ比較True、101件の追加拒否と出力なしを確認します。本体GUIの追加操作は、上記のdev11実機確認を参照してください。
