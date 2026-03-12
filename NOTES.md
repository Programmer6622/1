# Project Notes (MyTest Bot)

## How to run
- Install deps: `pip install -r requirements.txt`
- Token in `d:\Abror\Project\.env`:
  - `TELEGRAM_BOT_TOKEN=YOUR_TOKEN`
- Start bot: `python bot.py`
- CLI test:
  - `python -m mytest_bot.cli "student 1.docx"`
  - `python -m mytest_bot.cli "student 3.pdf"`

## Supported formats
- `.pdf` (including yellow highlight answers)
- `.docx`
- `.xlsx`, `.xlsm`
- `.txt`, `.csv`
- images: `.jpg`, `.jpeg`, `.png`, `.tif`, `.tiff`, `.bmp` (OCR)

## Not supported (intentional)
- `.doc` -> ask user to save as `.docx`
- `.xls` -> ask user to save as `.xlsx`

## Important requirements
- OCR needs Tesseract installed and on PATH:
  - check: `tesseract --version`

## PDF parsing improvements already added
- Yellow highlight answers detected
- Column-aware reading for multi-column PDFs
- Avoid fake columns with very few words
- Continuation lines appended to the last option
- Numeric options supported, e.g. `1) [+] Text`

## Known limitations
- No 100% guarantee for all layouts
- Scan PDFs (image-only) need OCR (Tesseract)
- Some unusual multi-column layouts may still need tuning
