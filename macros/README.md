# LibreOffice macros
Macros and extenstions for LibreOffice/OpenOffice

## FormulaFlatten.py
If the current active cell has a formula in it, then the precedents are
expanded until it reaches a cell that only has a value in it, a reference
to cell in another spreadsheet or a range. Those last two limitations are
probably a lack of imagination on my part, but it feels like a natural place
to stop the expansion.

The motivation is to summarize a set of calculations that's grown organically.
The 'flattened' formula can then be simplified or refactored or just copied
elsewhere for use.

Being cowardly, the resulting formula appears in a message box so the source
spreadsheet is left unchanged.

### Installation
Depends on your setup, but on Linux it can just be copied to:
~/.config/libreoffice/4/user/Scripts/python

But only _after_ you're comfortable it contains nothing nefarious. If you
value your system, data and/or privacy, please don't stuff things into your
macro folder on faith.

### Testing/Support
Created and tested on Ubuntu 16.04 against LibreOffice 6.3

