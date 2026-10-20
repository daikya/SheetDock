# XLSX並べ替え・フィルター：独立CLI検証

CSV本体1.2.0.dev1のWindows確認完了を受け、XLSXの並べ替え・フィルターを独立プローブで検証します。独立CLIとしてWindows確認を完了しました。現在は1.2.0.dev2の本体GUIへ接続しています。追加パッケージはありません。

Windowsで851件成功（78.44秒）、確認1～6と専用ハッシュ変数での元ファイル不変を確認済みです。独立GUIも確認完了です。続きは[本体実機確認1～8](docs/Phase3_Validation.md)です。以下はCLIの比較用手順として残します。

## 今回の範囲

| 操作 | 仕様 |
| --- | --- |
| 並べ替え | 選択した1列をキーにシートの行全体を移動。見出しありなら1行目を除外 |
| 比較 | 既存の型に従う。数値・文字列・真偽値・日時・時刻・時間間隔。異なる型が混在したキー列は拒否。文字列の数値化はしない |
| 順序 | 昇順／降順。同値の元の順序を保つ。空セルはどちらも最後 |
| 保持 | 行の全セルの値・Excel型・セル書式と行高・行の書式定義を移動。列幅・他シート・論理範囲は維持 |
| 検証 | コピーしたブックで操作結果を期待値と照合し、保存・再読込も一致した後だけ編集中モデルを更新。1回でUndo/Redo |
| 並べ替えの制限 | 数式を含むブック、対象シートの結合・エラー値・非表示／グループ行列・標準行非表示、読み取り専用は拒否 |
| フィルター | 1列の部分一致・完全一致・空・空以外。条件は表示用の本文行番号だけを返し、見出しを含めない。大文字小文字を区別 |
| 比較する文字 | 表示書式を付ける前の値。真偽値はTRUE/FALSE、日時はISO形式、数式は式文字列。数式の計算結果では比較しない |
| 非破壊性 | 値・型・書式・Excelの非表示設定・未保存状態・履歴を変えない。既存の非表示行は表示対象から除外。標準行非表示は今回対象外 |
| 保存 | 並べ替え後の全ブックを未使用名へ別名保存し、再読込比較。元ファイル・既存出力は上書きしない。フィルター指定と保存の同時実行は今回拒否 |

対象シート全体1000行×50列以内、ブック全体10万セル以内です。フィルターは読み取り専用・数式・結合セルのあるブックでも表示検証できます。Excelのオートフィルター設定を追加・変更する機能ではありません。条件はシート別のメモリ上にだけ保持します。範囲部分だけの並べ替え・複数キー・型混在時の優先順は今回含めません。

## Windows確認1～6

プロジェクト直下で`.venv312`を利用します。元サンプルは変更しないでください。

### 1. テストとサンプル生成

```powershell
python -m pytest tests probes/test_xlsx_sort_filter_probe.py -q --tb=short
python probes/xlsx_sort_filter_probe.py --make-samples probes/xlsx_sort_filter_samples
$xlsxSourceHash = (Get-FileHash probes/xlsx_sort_filter_samples/sort_basic.xlsx -Algorithm SHA256).Hash
```

現在の期待値は **940件成功**（本体894＋独立46）。独立CLIのWindows確認時は851件でした。サンプルはbasic・formula・merged・readonly・mixed・hidden・errorの7ファイル。同名があれば生成を中止するため、再実行時は別フォルダーを指定します。

basicのDataは、コード・数量・日付・有効・区分・メモの6列です。本文のコードは00123、00002、00003、00004、00005、数量は10、2、2、空、-1です。各行に異なる行高や書式を設定し、Otherシートも含めています。

### 2. 数値順・Undo/Redo・別名保存

```powershell
python probes/xlsx_sort_filter_probe.py probes/xlsx_sort_filter_samples/sort_basic.xlsx --sort-column 2 --save-as probes/xlsx_sort_filter_samples/number_copy.xlsx
```

コードが00005→00002→00003→00123→00004の順に表示され、並べ替え・Undo/Redo・別名保存・再読込の検証完了になること。空セルは最後、数量2の元の順序を保ちます。値・型・セル書式・行高・行定義・列幅・他シートはプログラム内で比較します。既存のSheetDockでコピーを開いて表示を確認しても構いませんが、XLSXの並べ替えボタンはまだ追加していません。

### 3. 降順・文字列・真偽値

```powershell
python probes/xlsx_sort_filter_probe.py probes/xlsx_sort_filter_samples/sort_basic.xlsx --sort-column 2 --descending
python probes/xlsx_sort_filter_probe.py probes/xlsx_sort_filter_samples/sort_basic.xlsx --sort-column 1
python probes/xlsx_sort_filter_probe.py probes/xlsx_sort_filter_samples/sort_basic.xlsx --sort-column 4
```

数量の降順は00123→00002→00003→00005→00004。コードの文字列昇順は00002→00003→00004→00005→00123。真偽値はFALSE→TRUEで、コードが00002→00004→00123→00003→00005です。元のセル値・型を保ちます。

### 4. フィルター・読み取り専用

```powershell
python probes/xlsx_sort_filter_probe.py probes/xlsx_sort_filter_samples/sort_basic.xlsx --filter-column 5 --query 東京
python probes/xlsx_sort_filter_probe.py probes/xlsx_sort_filter_samples/sort_basic.xlsx --filter-column 2 --operator empty
python probes/xlsx_sort_filter_probe.py probes/xlsx_sort_filter_samples/sort_basic.xlsx --filter-column 5 --operator equals --query 該当なし
python probes/xlsx_sort_filter_probe.py probes/xlsx_sort_filter_samples/sort_readonly.xlsx --filter-column 5 --query 東京
```

順に本文行番号[2, 4, 6]、[5]、[]、[2, 4, 6]です。すべて未保存「なし」、フィルターの非破壊検証完了になること。表示行番号は1始まりで、見出しを除きます。読み取り専用ブック自体は保存できません。

### 5. 未対応の並べ替えを拒否

```powershell
foreach ($case in 'formula', 'merged', 'readonly', 'mixed', 'hidden', 'error') {
    python probes/xlsx_sort_filter_probe.py ("probes/xlsx_sort_filter_samples/sort_{0}.xlsx" -f $case) --sort-column 2 --save-as ("probes/xlsx_sort_filter_samples/rejected_{0}.xlsx" -f $case)
}
@(Get-ChildItem probes/xlsx_sort_filter_samples -Filter 'rejected_*.xlsx').Count -eq 0
```

それぞれ別シートを含む数式、結合、読み取り専用、数値と文字列の型混在、非表示行、既存エラー値の理由で操作中止になること。出力が作られず、最後はTrueになります。これらのコマンドの終了コード1は想定した拒否です。

### 6. 既存出力・元ファイルを保護

```powershell
$copyHash = (Get-FileHash probes/xlsx_sort_filter_samples/number_copy.xlsx -Algorithm SHA256).Hash
python probes/xlsx_sort_filter_probe.py probes/xlsx_sort_filter_samples/sort_basic.xlsx --sort-column 2 --save-as probes/xlsx_sort_filter_samples/number_copy.xlsx
(Get-FileHash probes/xlsx_sort_filter_samples/number_copy.xlsx -Algorithm SHA256).Hash -eq $copyHash
python probes/xlsx_sort_filter_probe.py probes/xlsx_sort_filter_samples/sort_basic.xlsx --filter-column 5 --query 東京 --save-as probes/xlsx_sort_filter_samples/rejected_filter.xlsx
Test-Path probes/xlsx_sort_filter_samples/rejected_filter.xlsx
(Get-FileHash probes/xlsx_sort_filter_samples/sort_basic.xlsx -Algorithm SHA256).Hash -eq $xlsxSourceHash
```

既存出力を拒否し、コピーのハッシュはTrue。フィルターと保存の同時指定も拒否し、Test-PathはFalse。元サンプルのハッシュはTrueになること。生成フォルダーはGit対象外です。このCLIは比較検証用です。本体GUIの確認はPhase3_Validation.md末尾を参照してください。exeビルドはまだ行いません。
