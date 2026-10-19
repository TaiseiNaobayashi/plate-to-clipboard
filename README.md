# plate-to-clipboard

プレートリーダー (BioTek Gen5 / Synergy HTX) が出力した xlsx から 96 ウェルの値を探し出し、
TSV 形式でクリップボードにコピーする Windows 用ツール。

```
A1	0.035
A2	0.07
...
H12	0.11
```

1 ファイルに複数プレートがある場合は、ウェルごとにプレートを横に並べ、見出し行を付ける。

```
Well	Plate 1 - Sheet1	Plate 2 - Sheet1
A1	0.035	0.105
```

## 使い方

- ウィンドウに xlsx をドラッグ&ドロップ(またはクリックして選択)
- exe に xlsx を直接ドラッグ&ドロップしてもよい
- ウィンドウに今クリップボードに入っているファイル名と履歴が表示される(履歴はダブルクリックで再コピー)

全シートから「1〜12 の列見出し + その左列に A〜H の行見出し」がある領域を探すので、
結果の行位置がずれていても動く。

## ビルド

```bash
pip install openpyxl tkinterdnd2 pyinstaller
pyinstaller --onefile --noconsole --collect-all tkinterdnd2 plate_to_clipboard.py
```
