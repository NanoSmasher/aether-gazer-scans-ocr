# aether-gazer-scans-ocr
Extracting scan information (unit pulls) for Aether Gazer through optical character recognition  

## Requirements

You need to the game Aether Gazer from the play store, and currently the only method this scanner uses for screenshtos is Bluestacks 4/5.

Along with all libraries listed in requirements.txt, users will also need to [install tesseract](https://github.com/tesseract-ocr/tesseract#installing-tesseract).

The following may need to be changed if you've installed tesseract to a different location:
```
tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
```

## Miscellaneous

Recommended to have as large of a screenshot as possible otherwise the OCR may spit out weird results.

## To Do:

[ ] Better interface
[ ] Error correction
[ ] Add other image methods