import asyncio
import logging
import os
import re
import threading
from collections import defaultdict
from io import BytesIO
from datetime import date

from flask import Flask, request, jsonify
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler, MessageHandler,
    ContextTypes, filters
)
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

logging.basicConfig(
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    level=os.getenv("LOG_LEVEL", "INFO")
)
log = logging.getLogger("qada-bot")

BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
GEMINI_API_KEY = os.environ["GEMINI_API_KEY"]
PUBLIC_BASE_URL = os.environ["PUBLIC_BASE_URL"].rstrip("/")
WEBHOOK_SECRET = os.environ["WEBHOOK_SECRET"]
MODEL = os.getenv("GEMINI_MODEL", "gemini-3.7-flash")
DAILY_LIMIT = int(os.getenv("DAILY_LIMIT", "8"))
MAX_FILE_MB = int(os.getenv("MAX_FILE_MB", "12"))

client = genai.Client(api_key=GEMINI_API_KEY)
app = Flask(__name__)
application = Application.builder().token(BOT_TOKEN).build()

# RAM-only state: it is deliberately not persisted. A restart clears pending uploads.
sessions = {}
usage = defaultdict(int)
usage_day = date.today().isoformat()
loop = None
telegram_app_ready = threading.Event()


def main_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🖼 استخراج النص من صورة", callback_data="ocr")],
        [InlineKeyboardButton("📝 تعبئة قالب Word", callback_data="word")],
        [InlineKeyboardButton("ℹ️ طريقة الاستخدام", callback_data="help")]
    ])


def session_for(user_id):
    return sessions.setdefault(user_id, {"mode": None})


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.effective_message.reply_text(
        "مرحباً بك في المساعد الذكي للمحررات القضائية.\n"
        "اختر الخدمة من الأزرار أدناه.\n\n"
        "تنبيه: راجع النص المستخرج وملف Word قبل استخدامهما رسمياً.",
        reply_markup=main_menu()
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.effective_message.reply_text(
        "1) استخراج النص: اختر الزر ثم أرسل صورة أو ملف صورة.\n"
        "2) تعبئة Word: اختر الزر ثم أرسل قالب DOCX، وبعدها أجب عن الحقول التي تظهر.\n"
        "3) أرسل نص المذكرة بعد إدخال البيانات.\n\n"
        "العلامات المدعومة في القالب مثل:\n"
        "#هنا اسم المحكمة#، #هنا الموضوع#، #هنا اسم المدعي#، "
        "#هنا اسم المدعي عليه#، #هنا اسم القاضي#، #تبدأ كتابة من هنا#.\n\n"
        "أوامر التنسيق تكون في سطر مستقل بين #...#، مثال:\n"
        "الفقرة المراد تنسيقها\n#اجعل السطر السابق باللون الأحمر#\n\n"
        "لا ترسل مستندات سرية أو بيانات قضايا حساسة إلى خدمة الذكاء الاصطناعي المجانية."
    )


async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    s = session_for(q.from_user.id)
    if q.data == "ocr":
        s.clear()
        s["mode"] = "ocr_wait_image"
        await q.message.reply_text("أرسل صورة الكتابة العربية الآن (كصورة أو كملف).")
    elif q.data == "word":
        previous_ocr = s.get("last_ocr")
        s.clear()
        if previous_ocr:
            s["last_ocr"] = previous_ocr
        s["mode"] = "word_wait_template"
        await q.message.reply_text("أرسل قالب Word بصيغة DOCX. سأقرأ العلامات الموجودة داخله.")
    elif q.data == "help":
        await help_command(update, context)


def count_request(user_id):
    global usage_day
    today = date.today().isoformat()
    if today != usage_day:
        usage.clear()
        usage_day = today
    if usage[user_id] >= DAILY_LIMIT:
        return False
    usage[user_id] += 1
    return True


async def download_telegram_file(message, context):
    tg_file = None
    if message.photo:
        tg_file = await context.bot.get_file(message.photo[-1].file_id)
    elif message.document:
        tg_file = await context.bot.get_file(message.document.file_id)
    if not tg_file:
        return None
    bio = BytesIO()
    await tg_file.download_to_memory(out=bio)
    data = bio.getvalue()
    if len(data) > MAX_FILE_MB * 1024 * 1024:
        raise ValueError(f"حجم الملف أكبر من الحد المسموح ({MAX_FILE_MB} ميجابايت).")
    return data


async def gemini_ocr(image_bytes, mime_type):
    prompt = """
أنت نظام نسخ حرفي للكتابة العربية اليدوية، وقد تكون بخط الرقعة.
انسخ ما يظهر في الصورة فقط، من أعلى الصفحة إلى أسفلها ومن اليمين إلى اليسار.
لا تلخص، لا تصحح الإملاء، لا تخمن الكلمات غير المقروءة ولا تضف معلومات.
حافظ على فواصل الأسطر والفقرات قدر الإمكان.
ضع [غير واضح] مكان أي كلمة لا تستطيع قراءتها بثقة.
أخرج النص المستخرج فقط دون مقدمة أو شرح.
"""
    response = client.models.generate_content(
        model=MODEL,
        contents=[
            types.Part.from_bytes(data=image_bytes, mime_type=mime_type),
            prompt
        ],
        config=types.GenerateContentConfig(temperature=0.1)
    )
    return (response.text or "").strip()


def iter_paragraphs(parent):
    # Paragraphs at document body level
    for p in parent.paragraphs:
        yield p
    for table in parent.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from iter_paragraphs(cell)


def all_paragraphs(doc):
    yield from iter_paragraphs(doc)
    for section in doc.sections:
        for part in (section.header, section.footer, section.first_page_header,
                     section.first_page_footer, section.even_page_header,
                     section.even_page_footer):
            yield from iter_paragraphs(part)


def replace_text_in_paragraph(paragraph, old, new):
    full = paragraph.text
    if old not in full:
        return False
    updated = full.replace(old, new)
    if paragraph.runs:
        paragraph.runs[0].text = updated
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(updated)
    return True


FIELD_MAP = [
    (("اسم المحكمة", "المحكمة"), "اسم المحكمة"),
    (("الموضوع",), "الموضوع"),
    (("اسم المدعي عليه", "اسم المدعى عليه", "المدعي عليه", "المدعى عليه"), "اسم المدعى عليه"),
    (("اسم المدعي", "المدعي"), "اسم المدعي"),
    (("اسم القاضي", "القاضي"), "اسم القاضي"),
]


def marker_label(marker):
    inside = marker.strip("#").strip()
    for aliases, label in FIELD_MAP:
        if any(alias in inside for alias in aliases):
            return label
    if "تبدأ كتابة" in inside or "ابدأ كتابة" in inside:
        return None
    return inside or None


def find_markers(doc):
    found = []
    seen = set()
    for p in all_paragraphs(doc):
        for marker in re.findall(r"#[^#\n]{2,100}#", p.text):
            if marker not in seen:
                seen.add(marker)
                found.append(marker)
    return found


def insert_after(paragraph, text):
    # Insert one or more new paragraphs after the anchor, preserving the rest of the template.
    lines = text.splitlines() or [""]
    anchor = paragraph
    for line in lines:
        new_p = OxmlElement("w:p")
        anchor._p.addnext(new_p)
        new_para = Paragraph(new_p, paragraph._parent)
        new_para.style = paragraph.style
        new_para.add_run(line)
        anchor = new_para


def apply_commands_and_clean(text):
    """Commands are standalone #...# lines. Returns clean text and formatting actions."""
    actions = []
    out = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("#") and stripped.endswith("#") and len(stripped) > 2:
            cmd = stripped[1:-1].strip().lower()
            refers_to_previous = any(x in cmd for x in ("السطر السابق", "هذا السطر", "ذا السطر", "السطر ذا"))
            if refers_to_previous and ("أحمر" in cmd or "احمر" in cmd or "بالأحمر" in cmd):
                actions.append({"target": max(0, len(out)-1), "type": "red"})
            elif refers_to_previous and ("عريض" in cmd or "غامق" in cmd or "bold" in cmd):
                actions.append({"target": max(0, len(out)-1), "type": "bold"})
            elif refers_to_previous and ("توسيط" in cmd or "الوسط" in cmd or "منتصف" in cmd):
                actions.append({"target": max(0, len(out)-1), "type": "center"})
            elif "صفحة جديدة" in cmd or "فاصل صفحة" in cmd:
                actions.append({"target": len(out), "type": "page_break"})
            elif refers_to_previous and ("تحته خط" in cmd or "تسطير" in cmd):
                actions.append({"target": max(0, len(out)-1), "type": "underline"})
            else:
                # Do not write command syntax into the legal document.
                continue
        else:
            out.append(line)
    return "\n".join(out).strip(), actions


def add_content_at_anchor(doc, content, actions):
    anchor_marker = None
    for p in all_paragraphs(doc):
        for m in re.findall(r"#[^#\n]{2,100}#", p.text):
            if "تبدأ كتابة" in m or "ابدأ كتابة" in m:
                anchor_marker = (p, m)
                break
        if anchor_marker:
            break

    if not anchor_marker:
        raise ValueError("لم أجد علامة بداية الكتابة (#تبدأ كتابة من هنا#) داخل القالب.")

    paragraph, marker = anchor_marker
    replace_text_in_paragraph(paragraph, marker, "")
    lines = content.splitlines() or [""]
    inserted = []
    current = paragraph
    for line in lines:
        new_p = OxmlElement("w:p")
        current._p.addnext(new_p)
        p = Paragraph(new_p, paragraph._parent)
        p.style = paragraph.style
        p.add_run(line)
        inserted.append(p)
        current = p

    for action in actions:
        idx = action["target"]
        if action["type"] == "page_break":
            idx = min(idx, len(inserted)-1)
            inserted[idx].paragraph_format.page_break_before = True
            continue
        if not inserted:
            continue
        idx = min(max(idx, 0), len(inserted)-1)
        p = inserted[idx]
        if action["type"] == "red":
            for r in p.runs:
                r.font.color.rgb = RGBColor(255, 0, 0)
        elif action["type"] == "bold":
            for r in p.runs:
                r.bold = True
        elif action["type"] == "underline":
            for r in p.runs:
                r.underline = True
        elif action["type"] == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def build_docx(template_bytes, values, body_text):
    doc = Document(BytesIO(template_bytes))
    # Replace all field markers, including those in tables, headers and footers.
    for p in list(all_paragraphs(doc)):
        for marker in re.findall(r"#[^#\n]{2,100}#", p.text):
            label = marker_label(marker)
            if label and label in values:
                replace_text_in_paragraph(p, marker, values[label])

    clean_text, actions = apply_commands_and_clean(body_text)
    add_content_at_anchor(doc, clean_text, actions)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


async def handle_photo_or_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    s = session_for(user_id)
    msg = update.effective_message

    if not count_request(user_id):
        await msg.reply_text(f"وصلت إلى الحد اليومي التجريبي ({DAILY_LIMIT} طلبات). حاول غداً.")
        return

    try:
        data = await download_telegram_file(msg, context)
    except Exception as e:
        await msg.reply_text(str(e))
        return

    if not data:
        await msg.reply_text("أرسل صورة أو ملفاً من فضلك.")
        return

    if s.get("mode") == "ocr_wait_image":
        mime = "image/jpeg"
        if msg.document and msg.document.mime_type:
            mime = msg.document.mime_type
        await msg.reply_text("جاري قراءة الصورة. قد تستغرق العملية بعض الوقت...")
        try:
            text = await asyncio.to_thread(gemini_ocr, data, mime)
            s["last_ocr"] = text
            s["mode"] = "idle"
            await msg.reply_text(
                "النص المستخرج (راجعه وصحح أي كلمة غير واضحة):\n\n" + (text[:3900] or "لم أتمكن من استخراج نص."),
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("استخدام النص في قالب Word", callback_data="use_ocr")]
                ])
            )
        except Exception:
            log.exception("OCR failed")
            await msg.reply_text("تعذر استخراج النص. قد تكون حصة Gemini انتهت أو الصورة غير واضحة.")
        return

    if s.get("mode") == "word_wait_template":
        if not msg.document or not (msg.document.file_name or "").lower().endswith(".docx"):
            await msg.reply_text("من فضلك أرسل قالباً بامتداد DOCX كملف، وليس صورة.")
            return
        try:
            doc = Document(BytesIO(data))
            markers = find_markers(doc)
            if not markers:
                await msg.reply_text("لم أجد علامات #...# في القالب. تأكد أن العلامات موجودة كنص داخل ملف DOCX.")
                return
            s["template"] = data
            s["values"] = {}
            s["fields"] = [(m, marker_label(m)) for m in markers if marker_label(m)]
            s["field_index"] = 0
            s["mode"] = "word_collect_fields"
            await ask_next_field(msg, s)
        except Exception:
            log.exception("Template parse failed")
            await msg.reply_text("تعذر فتح القالب. تأكد أنه ملف DOCX صالح وغير محمي بكلمة مرور.")


async def ask_next_field(message, s):
    fields = s["fields"]
    i = s["field_index"]
    if i < len(fields):
        marker, label = fields[i]
        await message.reply_text(f"أدخل {label} (الموضع: {marker})")
    else:
        s["mode"] = "word_wait_body"
        await message.reply_text(
            "تم جمع البيانات الأساسية.\n"
            "الآن أرسل نص المذكرة كاملاً، أو اضغط استخدام النص المستخرج إذا كنت قد استخرجته سابقاً."
        )


async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    s = session_for(user_id)
    msg = update.effective_message
    text = msg.text or ""

    if s.get("mode") == "word_collect_fields":
        marker, label = s["fields"][s["field_index"]]
        s["values"][label] = text.strip()
        s["field_index"] += 1
        await ask_next_field(msg, s)
        return

    if s.get("mode") == "word_wait_body":
        await finish_word(msg, s, text)
        return

    await msg.reply_text("اختر إحدى الخدمتين من القائمة:", reply_markup=main_menu())


async def finish_word(message, s, body):
    try:
        result = await asyncio.to_thread(
            build_docx, s["template"], s["values"], body
        )
        await message.reply_document(
            document=BytesIO(result),
            filename="المذكرة_القضائية.docx",
            caption="تم إنشاء الملف. يرجى مراجعته والتأكد من صحة الأسماء والتنسيق قبل استخدامه."
        )
        s.clear()
        s["mode"] = None
        await message.reply_text("هل تريد تنفيذ خدمة أخرى؟", reply_markup=main_menu())
    except Exception:
        log.exception("DOCX build failed")
        await message.reply_text("تعذر إنشاء الملف. تحقق من وجود علامة #تبدأ كتابة من هنا# في القالب.")


async def callback_use_ocr(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    s = session_for(q.from_user.id)
    if not s.get("last_ocr"):
        await q.message.reply_text("لا يوجد نص مستخرج محفوظ في الجلسة الحالية.")
        return
    # The template flow must already have collected all its fields.
    if s.get("mode") != "word_wait_body":
        await q.message.reply_text("ابدأ أولاً خدمة تعبئة Word وأكمل بيانات القالب، ثم استخدم النص المستخرج.")
        return
    await finish_word(q.message, s, s["last_ocr"])


async def post_init(application):
    global loop
    loop = asyncio.get_running_loop()
    url = f"{PUBLIC_BASE_URL}/telegram/{WEBHOOK_SECRET}"
    await application.bot.set_webhook(url=url, allowed_updates=Update.ALL_TYPES)
    log.info("Webhook configured: %s/telegram/...", PUBLIC_BASE_URL)


application.add_handler(CommandHandler("start", start))
application.add_handler(CommandHandler("help", help_command))
application.add_handler(CallbackQueryHandler(callback_use_ocr, pattern="^use_ocr$"))
application.add_handler(CallbackQueryHandler(button_handler))
application.add_handler(MessageHandler(filters.PHOTO | filters.Document.ALL, handle_photo_or_document))
application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))

# Flask webhook endpoint; PTB processes updates on its own asyncio loop.
def run_telegram():
    global loop
    async def runner():
        global loop
        await application.initialize()
        await application.start()
        loop = asyncio.get_running_loop()
        await application.bot.set_webhook(
            url=f"{PUBLIC_BASE_URL}/telegram/{WEBHOOK_SECRET}",
            allowed_updates=Update.ALL_TYPES
        )
        telegram_app_ready.set()
        await asyncio.Event().wait()
    asyncio.run(runner())


@app.get("/")
def health():
    return jsonify({"status": "ok", "service": "telegram-legal-docx-bot"})


@app.post("/telegram/<secret>")
def telegram_webhook(secret):
    if secret != WEBHOOK_SECRET:
        return "forbidden", 403
    if loop is None or not telegram_app_ready.is_set():
        return "starting", 503
    try:
        payload = request.get_json(force=True)
        update = Update.de_json(payload, application.bot)
        future = asyncio.run_coroutine_threadsafe(
            application.process_update(update), loop
        )
        # Do not block Telegram waiting for the model response.
        return "ok", 200
    except Exception:
        log.exception("Webhook processing error")
        return "error", 500


# Start the Telegram event loop when Gunicorn imports this module.
# Render is configured with one worker to avoid duplicate webhook consumers.
_bot_thread = threading.Thread(target=run_telegram, daemon=True)
_bot_thread.start()

if __name__ == "__main__":
    port = int(os.getenv("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
