# plate-to-clipboard

A Windows tool that finds the 96-well values in an xlsx exported by a plate reader
(BioTek Gen5 / Synergy HTX) and copies them to the clipboard as TSV.

```
A1	0.035
A2	0.07
...
H12	0.11
```

If one file contains several plates, each plate gets its own column next to the well names, with a header row.

```
Well	Plate 1 - Sheet1	Plate 2 - Sheet1
A1	0.035	0.105
```

## Usage

- Drag & drop an xlsx onto the window (or click it to choose a file)
- You can also drop an xlsx directly onto the exe
- The window shows which file is currently on the clipboard, plus a history (double-click an entry to copy it again)

The tool searches every sheet for a block with column headers 1–12 and row labels A–H in the column to their left,
so it still works when the results start on a different row.

## Build

```bash
pip install openpyxl tkinterdnd2 pyinstaller
pyinstaller --onefile --noconsole --collect-all tkinterdnd2 plate_to_clipboard.py
```
