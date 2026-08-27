# Excel VBA 中級研修　Scripting.Dictionary 実務活用ドリル

`Dictionary実務活用ドリル.xlsm` は、Scripting.Dictionary を使うと「断然効率的になる」処理を
8 例そろえた研修用ブックです。モデルデータとマクロが入っており、そのまま実行して結果を確認できます。

## 収録内容

| No | シート | テーマ | 実行マクロ |
|----|--------|--------|-----------|
| 1 | 01_重複排除 | 重複を除いたユニーク一覧 | `M01_UniqueList` |
| 2 | 02_キー別集計 | キー別に件数・数量・金額を同時集計 | `M02_GroupSum` |
| 3 | 03_マスタ照合 | VLOOKUP／Find の繰り返しを置き換える | `M03_MasterLookup` |
| 4 | 04_2表突合 | 前月・当月の差分（新規／削除／変更／一致） | `M04_CompareTwoLists` |
| 5 | 05_クロス集計 | 複合キーで支店×商品のマトリクス作成 | `M05_CrossTab` |
| 6 | 06_明細グループ化 | 値に Collection を持たせて明細を束ねる／シート分割 | `M06_GroupDetail` / `M06_SplitToSheets` / `M06_DeleteSplitSheets` |
| 7 | 07_採番連番 | キーごとの枝番採番と「未登録キーを読むと増える」落とし穴 | `M07_SerialNumber` / `M07_KeyPitfall` |
| 8 | 08_速度比較 | COUNTIF 方式と Dictionary 方式の実測比較 | `M08_Benchmark` |

- `00_目次` … 例題一覧、Dictionary 早見表、つまずきやすい 5 点
- `09_コード全文` … 全モジュールのコード（印刷・回覧用）
- 例 1〜7 をまとめて実行する `RunAll`（M00_Common）も用意

## 使い方

1. ブックを開き、警告バーが出たら［コンテンツの有効化］をクリック
   （ダウンロードしたファイルは、右クリック →［プロパティ］→［ブロックの解除］が必要な場合があります）
2. 各シートを開き、`Alt`+`F8` → 対象マクロを選んで実行
3. コードは `Alt`+`F11`（標準モジュール M00〜M08）で確認

コード中の目印:

- `'★定義：` … Dictionary を用意している行
- `'★書込：` … 登録・更新している行
- `'★読込：` … 取り出し・存在確認している行

参照設定は不要です（`CreateObject("Scripting.Dictionary")` の遅延バインディング）。

## ソースと再生成

- `vba/*.bas` … 各モジュールのソース（VBE へのインポートも可能）
- `tools/build_xlsm.py` … ブック生成スクリプト
  （`vba/*.bas` から vbaProject.bin を組み立て、マクロ入り .xlsm を書き出す）

```bash
pip install openpyxl
python3 tools/build_xlsm.py            # Dictionary実務活用ドリル.xlsm を再生成
```

`tools/cfb.py`（OLE 複合ファイル書き出し）と `tools/ovba.py`（MS-OVBA 圧縮）は
その補助モジュールです。
