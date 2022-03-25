# Excel VBA: Convert Coordinates from dms to decimal

Many of my written coordinates are in the "dms" Format

```
44°43'30,2"N	5°40'47,2"E
```

but many applications expect a decimal representation like

```
44,725	5,679722222
```

I found an Excel VBA script (from this Blogpost by Glen Bambrick)[https://glenbambrick.com/2015/08/16/dms-to-dd-excel/], but it gave broken coordinates for me. This repository fixes it.

## Usage

Open the Excel Spreadsheet with your DMS Coordinates, then open *Visual Basic for Applications* by pressing <kbd>Alt</kbd> + <kbd>F11</kbd>

Choose File > Import (<kbd>Ctrl</kbd> + <kbd>M</kbd>) and choose `ConvertToDecimal.bas`. Close *Visual Basic for Applications*.

You can now use `ConvertToDecimal()` as a Formula in Excel:

![Screenshot](/Readme.md-attachments/Screenshot.png)