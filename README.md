# MyTest Telegram Bot (converter only)

This project converts student test files into a simple MyTest text format:

```
#1 Question text
+correct answer
-wrong answer
-wrong answer

#2 Another question
...
```

## What is included
- Text extraction for pdf, docx, txt, csv, images (OCR), and Excel (xlsx).
- Heuristics to parse common test layouts with marked correct answers.
- A Telegram bot that accepts files and returns the converted output.
- A local CLI for quick testing.

## Install
```
pip install -r requirements.txt
```

Optional system tools:
- OCR: install Tesseract.

## Local conversion
```
python -m mytest_bot.cli "student 1.docx"
python -m mytest_bot.cli "student 3.pdf" -o out.mytest.txt
```

## Run the bot
```
set TELEGRAM_BOT_TOKEN=YOUR_TOKEN
python -m mytest_bot.telegram_bot
```

Or run the shortcut:
```
python bot.py
```

Or use a `.env` file in the project root:
```
TELEGRAM_BOT_TOKEN=YOUR_TOKEN
```

You can also limit file size:
```
set MAX_FILE_MB=20
```

## Notes
- The parser expects the correct answer to be marked (A#, A*, +, or "Javob: A").
- For PDFs, highlighted answers are detected (yellow marker highlight).
- .doc and .xls are not supported; save them as .docx and .xlsx.
- If parsing fails for a file, add more examples and extend the parser rules.
