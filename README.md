# Excel 2016 互換チェック ツール

`excel2016_compat_check.bat` と `excel2016_compat_check.py` で、ブック内の数式から Excel 2016 に存在しない関数や、環境依存の関数の使用箇所を検出して Markdown レポートを出力します。

## 特長
- ドラッグ&ドロップで簡単実行（複数ファイル対応）
- Windows の `python` / `py -3` / WSL の `python3` を自動検出
- 必要なパッケージ `openpyxl` を自動インストール（ネット接続が必要）
- `_xlfn.XLOOKUP`、`@XLOOKUP` などのプレフィックスを正規化して検出

## 前提 / 要件
- Windows 環境（バッチでの実行を想定）
  - いずれかで Python が使用可能であること
    - Windows の `python` または `py -3`
    - もしくは WSL に `python3`（WSL が有効な場合）
- ネットワーク接続（初回の `openpyxl` 自動インストール時）

## 使い方
### 1) ドラッグ&ドロップ（推奨）
1. 対象の `.xlsx` / `.xlsm` ファイル（複数可）を `excel2016_compat_check.bat` にドラッグ&ドロップします。
2. 同じフォルダに `<元ファイル名>_2016_compat_report.md` が出力されます。

### 2) コマンドラインから（Windows cmd）
```bat
C:\> C:\path\to\excel2016_compat_check.bat "C:\path\to\book.xlsx"
```

### 3) 直接 Python で実行
```bat
C:\> python C:\path\to\excel2016_compat_check.py "C:\path\to\book.xlsx"
```
複数ファイルを続けて指定できます。

## 出力（Markdown レポート）
- 出力ファイル名: `<元ファイル名>_2016_compat_report.md`
- 記載内容:
  - 検出サマリ（エラー/注意の件数）
  - エラー: Excel 2016 に存在しない関数の使用箇所（シート・セル番地・数式要約）
  - 注意: Excel 2016 では環境依存の関数の使用箇所
  - 参考: 検出関数ごとの代替案

## 検出対象の例
- 2016に存在しない関数（エラー）
  - 例: `XLOOKUP`, `XMATCH`, `FILTER`, `UNIQUE`, `SORT`, `LET`, `LAMBDA`, `TEXTSPLIT`, ほか
- 2016で環境依存の関数（注意）
  - 例: `CONCAT`, `TEXTJOIN`, `IFS`, `SWITCH`, `MAXIFS`, `MINIFS`

※ 検出は英語関数名ベースです。`_xlfn.` プレフィックスや先頭の `@` は正規化してから判定します。

## 制限事項
- セルの数式と「定義名（Named Formula）」のみが対象
  - 条件付き書式、データの入力規則、ピボット、Power Query など内部の式は対象外
- `.xlsb` の式抽出は未対応（必要な場合は `.xlsx` へ保存してから実行）
- 関数名のローカライズ（日本語関数名など）には非対応（英語関数名で判定）

## トラブルシューティング
- 「Python が見つかりません」
  - Windows に Python をインストールするか、WSL の `python3` を用意してください。
  - `py -3 -V` または `python -V` で動作確認できます。
- 文字化けする
  - バッチは UTF-8（`chcp 65001`）で出力します。Windows Terminal の使用を推奨します。
  - 直接 Python 実行でも `utf-8` で出力するよう調整しています。
- パッケージがインストールできない
  - ネットワーク制限がある場合は、事前に `pip install openpyxl` を行ってください。

## フォルダ構成
- `excel2016_compat_check.bat` … D&D 実行用バッチ
- `excel2016_compat_check.py` … 本体スクリプト
- `2016_365_test.xlsx` … テスト用サンプル

## 変更履歴
- 2025-08-22
  - D&D の堅牢化（複数ファイル、スペース/日本語パス対応）
  - `python` / `py -3` / `WSL python3` 自動検出
  - 出力の文字化け対策（UTF-8 固定）
  - 関数検出の正規化強化（`_xlfn.*`, `@*` を考慮）

