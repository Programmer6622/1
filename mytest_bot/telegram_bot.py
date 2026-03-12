import logging
import os
import tempfile
from pathlib import Path
from urllib.parse import urlparse

from telegram import InputFile, Update
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters

from .converter import convert_file_with_report, extract_raw_text
from .errors import ConversionError

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None

if load_dotenv:
    load_dotenv()


MAX_FILE_MB = int(os.environ.get("MAX_FILE_MB", "20"))
EXPORT_RAW_TEXT = os.environ.get("EXPORT_RAW_TEXT", "0").strip().lower() in {
    "1",
    "true",
    "yes",
    "on",
}
ALLOWED_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".txt",
}


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Send a test file (pdf, docx, or txt)."
    )


async def handle_batch(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text("Starting batch conversion of files in 'file' directory...")
    file_dir = Path("d:/Abror/Project/file")
    if not file_dir.exists() or not file_dir.is_dir():
        await update.message.reply_text("Directory 'file' not found.")
        return
        
    files = [f for f in file_dir.iterdir() if f.is_file() and f.suffix.lower() in ALLOWED_EXTENSIONS]
    if not files:
        await update.message.reply_text("No supported files found in 'file' directory.")
        return
        
    await update.message.reply_text(f"Found {len(files)} files to process.")
    
    for input_path in files:
        await update.message.reply_text(f"Processing: {input_path.name}")
        raw_text = None
        if EXPORT_RAW_TEXT:
            try:
                raw_text = extract_raw_text(input_path)
            except ConversionError:
                raw_text = None
        try:
            output_text, report_text = convert_file_with_report(input_path)
        except ConversionError as exc:
            await update.message.reply_text(f"Conversion failed for {input_path.name}: {exc}")
            continue
        
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir) / input_path.name
            
            output_path = tmp_path.with_suffix(".mytest.txt")
            output_path.write_text(output_text, encoding="utf-8")
            with output_path.open("rb") as handle:
                await update.message.reply_document(
                    document=InputFile(handle, filename=output_path.name)
                )
                
            if report_text:
                report_path = tmp_path.with_name(f"{tmp_path.stem}.report.txt")
                report_path.write_text(report_text, encoding="utf-8")
                with report_path.open("rb") as handle:
                    await update.message.reply_document(
                        document=InputFile(handle, filename=report_path.name),
                        caption=f"Conversion report for {input_path.name}",
                    )
            
            if raw_text:
                raw_path = tmp_path.with_name(f"{tmp_path.stem}.raw.txt")
                raw_path.write_text(raw_text, encoding="utf-8")
                with raw_path.open("rb") as handle:
                    await update.message.reply_document(
                        document=InputFile(handle, filename=raw_path.name),
                        caption=f"Extracted raw text for {input_path.name}",
                    )
                    
    await update.message.reply_text("Batch conversion completed.")


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    doc = update.message.document
    if not doc:
        return
    if doc.file_size and doc.file_size > MAX_FILE_MB * 1024 * 1024:
        await update.message.reply_text("File is too large.")
        return

    file_name = doc.file_name or "file"
    ext = Path(file_name).suffix.lower()
    if ext in {".doc", ".xls"}:
        await update.message.reply_text(
            ".doc and .xls are not supported. Please save as .docx."
        )
        return
    if ext and ext not in ALLOWED_EXTENSIONS:
        await update.message.reply_text(
            "Supported formats: pdf, docx, and txt."
        )
        return

    progress_message = await update.message.reply_text("File processing... 0%")

    async def set_progress(percent: int, status: str = "File processing") -> None:
        try:
            await progress_message.edit_text(f"{status}... {percent}%")
        except Exception:
            pass

    await set_progress(10, "Downloading file")
    tg_file = await context.bot.get_file(doc.file_id)
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = Path(tmpdir) / file_name
        await tg_file.download_to_drive(custom_path=str(input_path))
        await set_progress(25)
        raw_text = None
        if EXPORT_RAW_TEXT:
            try:
                raw_text = extract_raw_text(input_path)
            except ConversionError:
                raw_text = None
        await set_progress(55)
        try:
            output_text, report_text = convert_file_with_report(input_path)
        except ConversionError as exc:
            await set_progress(100, "Processing failed")
            await update.message.reply_text(f"Conversion failed: {exc}")
            if raw_text:
                raw_path = input_path.with_name(f"{input_path.stem}.raw.txt")
                raw_path.write_text(raw_text, encoding="utf-8")
                with raw_path.open("rb") as handle:
                    await update.message.reply_document(
                        document=InputFile(handle, filename=raw_path.name),
                        caption="Extracted raw text",
                    )
            return
        await set_progress(80, "Preparing output")
        output_path = input_path.with_suffix(".mytest.txt")
        output_path.write_text(output_text, encoding="utf-8")
        with output_path.open("rb") as handle:
            await update.message.reply_document(
                document=InputFile(handle, filename=output_path.name)
            )
        if report_text:
            report_path = input_path.with_name(f"{input_path.stem}.report.txt")
            report_path.write_text(report_text, encoding="utf-8")
            with report_path.open("rb") as handle:
                await update.message.reply_document(
                    document=InputFile(handle, filename=report_path.name),
                    caption="Conversion report",
                )
        if raw_text:
            raw_path = input_path.with_name(f"{input_path.stem}.raw.txt")
            raw_path.write_text(raw_text, encoding="utf-8")
            with raw_path.open("rb") as handle:
                await update.message.reply_document(
                    document=InputFile(handle, filename=raw_path.name),
                    caption="Extracted raw text",
                )
        await set_progress(100, "Completed")


def main() -> None:
    if load_dotenv:
        load_dotenv()
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        raise SystemExit("TELEGRAM_BOT_TOKEN is not set.")
    webhook_flag = os.environ.get("WEBHOOK", "").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }
    force_polling = os.environ.get("FORCE_POLLING", "").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }
    render_external_url = os.environ.get("RENDER_EXTERNAL_URL", "").strip()
    webhook_url = os.environ.get("WEBHOOK_URL", "").strip()
    use_webhook = not force_polling and (
        webhook_flag or bool(webhook_url) or bool(render_external_url)
    )
    if not webhook_url and use_webhook:
        webhook_url = render_external_url
    webhook_secret = os.environ.get("WEBHOOK_SECRET", "").strip() or None
    listen_host = os.environ.get("WEBHOOK_LISTEN", "0.0.0.0").strip()
    port = int(os.environ.get("PORT", os.environ.get("WEBHOOK_PORT", "10000")))
    url_path = os.environ.get("WEBHOOK_PATH", "telegram/webhook").strip()
    if url_path.startswith("http://") or url_path.startswith("https://"):
        parsed_path = urlparse(url_path).path.strip("/")
        url_path = parsed_path or "telegram/webhook"
    url_path = url_path.lstrip("/")
    if use_webhook and not webhook_url:
        raise SystemExit(
            "Webhook mode is enabled but WEBHOOK_URL is not set "
            "(or RENDER_EXTERNAL_URL is missing)."
        )
    webhook_url_full = webhook_url.rstrip("/")
    if webhook_url:
        parsed_webhook = urlparse(webhook_url)
        has_path = bool(parsed_webhook.path and parsed_webhook.path.strip("/"))
        if not has_path:
            webhook_url_full = f"{webhook_url_full}/{url_path}"

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    app = Application.builder().token(token).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("batch", handle_batch))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    if use_webhook:
        logging.info("Starting in webhook mode: %s", webhook_url_full)
        app.run_webhook(
            listen=listen_host,
            port=port,
            url_path=url_path,
            webhook_url=webhook_url_full,
            secret_token=webhook_secret,
            drop_pending_updates=True,
        )
    else:
        logging.info("Starting in polling mode")
        app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
