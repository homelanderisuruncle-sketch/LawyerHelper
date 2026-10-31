import os
import io
import json
import logging
from datetime import datetime
from PIL import Image
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler, 
    ContextTypes, filters, ConversationHandler
)
from google import genai
from google.genai import types
from docxtpl import DocxTemplate
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# إعداد التسجيل لمتابعة الأخطاء واستهلاك التوكنز
logging.basicConfig(format= '%(asctime)s - %(name)s - %(levelname)s - %(message)s' , level=logging.INFO)

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

# إعداد عميل Gemini API
client = genai.Client(api_key=GEMINI_API_KEY)

# حالات المحادثة
WAITING_IMAGE = 1
WAITING_INFO = 2

# أزرار القائمة الرئيسية
MAIN_KEYBOARD = [
    ["📝 استخراج النص", "💼 حقيبتك"],
    ["📄 إنشاء ملف Word", "❌ إنهاء الجلسة"]
]

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("texts", [])
    context.user_data.setdefault("bag", {})

    markup = ReplyKeyboardMarkup(MAIN_KEYBOARD, resize_keyboard=True)
    await update.message.reply_text(
        "أهلاً بك في المساعد القانوني الذكي ⚖️\nاختر من القائمة أدناه للبدء:",
        reply_markup=markup
    )

# ==========================================
# AGENT 1: استخراج النص وقراءة خط الرقعة
# ==========================================

async def start_ocr(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📸 أرسل صورة الورقة (بخط الرقعة أو اليدوي):")
    return WAITING_IMAGE

async def process_image(update: Update, context: ContextTypes.DEFAULT_TYPE):
    photo = update.message.photo[-1]
    file = await photo.get_file()
    img_bytes = await file.download_as_bytearray()
    image = Image.open(io.BytesIO(img_bytes))

    await update.message.reply_text("جاري قراءة خط الرقعة وتحليل المستند... ⏳")

    prompt = """
    أنت وكيل قانوني متخصص في قراءة الخط العربي اليدوي وخط الرقعة.
    قم باستخراج النص الموجود في الصورة كاملاً بدقة عالية مع الحفاظ على المصطلحات القانونية.
    أخرج النص الصافي المباشر فقط بدون مقدمات أو تعليقات.
    """

    config = types.GenerateContentConfig(
        temperature=0.1,
        max_output_tokens=2000
    )

    try:
        response = client.models.generate_content(
            model= 'gemini-2.5-flash' ,
            contents=[image, prompt],
            config=config
        )

        # طباعة استهلاك التوكنز في السجلات (Logs)
        if hasattr(response,  'usage_metadata' ) and response.usage_metadata:
            logging.info(f"📊 [Agent 1 Tokens] Input: {response.usage_metadata.prompt_token_count} | Output: {response.usage_metadata.candidates_token_count}")

        extracted_text = response.text.strip()
        context.user_data["texts"].append(extracted_text)
        
        text_num = len(context.user_data["texts"])
        markup = ReplyKeyboardMarkup(MAIN_KEYBOARD, resize_keyboard=True)
        await update.message.reply_text(
            f"✅ **تم حفظ النص رقم ({text_num}):**\n\n{extracted_text}",
            reply_markup=markup
        )
    except Exception as e:
        logging.error(f"Error in Agent 1: {e}")
        await update.message.reply_text(f"❌ حدث خطأ أثناء قراءة الصورة: {e}")

    return ConversationHandler.END

# ==========================================
# AGENT 2: دمج القالب وتطبيق تنسيقات الهاشتاج
# ==========================================

async def start_word(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texts = context.user_data.get("texts", [])
    if not texts:
        await update.message.reply_text("⚠️ لا توجد نصوص مستخرجة! قم بتصوير واستخراج صور أولاً.")
        return ConversationHandler.END

    await update.message.reply_text(
        "📝 أدخل بيانات القضية مفصولة بفواصل (,) بالترتيب التالي:\n\n"
        "`اسم المحكمة, اسم القاضي, المدعي, المدعى عليه`\n\n"
        "*(ملاحظة: سيتم إضافة تاريخ اليوم تلقائياً)*",
        parse_mode="Markdown"
    )
    return WAITING_INFO

def parse_formatting_instructions(text_content):
    """
    استخدام نموذج gemini-2.5-pro بأعلى دقة لتحليل النصوص واستخراج أوامر التنسيق بين الهاشتاجات #أمر#
    """
    # إعدادات أقصى دقة ومنع التخمين
    config = types.GenerateContentConfig(
        temperature=0.0,
        max_output_tokens=4000,
        response_mime_type="application/json"  # إجبار الموديل على إخراج JSON نقي 100%
    )

    prompt = f"""
    أنت خبير قانوني ومساعد دقيق جداً في تنسيق مستندات Microsoft Word.
    
    المطلوب منك:
    1. اقرأ النص التالي بعناية.
    2. ابحث عن أي كلمات أو عبارات محاطة بعلامات الهاشتاج مثل #أمر أو تنسيق# أو ##أمر##.
    3. استخرج أمر التنسيق المطلوب (مثل: أحمر، أزرق، أخضر، عريض، كبر الخط، خط صغير).
    4. طبّق التنسيق على النص المكتوب داخل علامات الهاشتاج فقط.
    5. أزل علامات الهاشتاج والأوامر من النص النهائي الصافي.
    
    النص المدخل:
    "{text_content}"
    
    أخرج الناتج كـ JSON Array يحتوي على مصفوفة من العناصر بالتنسيق التالي حصراً:
    [
        {{
            "text": "النص الصافي بعد حذف علامات الهاشتاج وكلمات الأمر",
            "bold": true/false,
            "color": "red/blue/green/black",
            "size": 14
        }}
    ]
    """

    try:
        response = client.models.generate_content(
            model= 'gemini-2.5-pro' ,  # استخدام الفئة المحترفة والأعلى دقة
            contents=[prompt],
            config=config
        )

        if hasattr(response,  'usage_metadata' ) and response.usage_metadata:
            logging.info(f"📊 [Agent 2 Tokens] Input: {response.usage_metadata.prompt_token_count} | Output: {response.usage_metadata.candidates_token_count}")

        return json.loads(response.text.strip())
    except Exception as e:
        logging.error(f"Error in formatting parsing: {e}")
        return [{"text": text_content, "bold": False, "color": "black", "size": 13}]


def generate_word_from_template(info_parts, texts_list):
    """دمج البيانات المستخرجة بدقة داخل القالب template.docx"""
    doc = DocxTemplate("template.docx")
    subdoc = doc.new_subdoc()

    # معالجة كل نص مضاف وتطبيق التنسيقات عليه
    for idx, text in enumerate(texts_list, 1):
        if len(texts_list) > 1:
            subdoc.add_heading(f"المستند رقم ({idx}):", level=2)
            
        p = subdoc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        formatted_runs = parse_formatting_instructions(text)

        for item in formatted_runs:
            run = p.add_run(item.get("text", ""))
            
            # تطبيق العريض (Bold)
            if item.get("bold"):
                run.bold = True
                
            # تطبيق حجم الخط
            if item.get("size"):
                run.font.size = Pt(int(item.get("size")))
            else:
                run.font.size = Pt(13)

            # تطبيق الألوان
            color = str(item.get("color", "")).lower()
            if "red" in color or "أحمر" in color:
                run.font.color.rgb = RGBColor(200, 0, 0)
            elif "blue" in color or "أزرق" in color:
                run.font.color.rgb = RGBColor(0, 102, 204)
            elif "green" in color or "أخضر" in color:
                run.font.color.rgb = RGBColor(0, 153, 76)
            else:
                run.font.color.rgb = RGBColor(0, 0, 0)

    # استخراج البيانات المطلوبة للقالب
    court = info_parts[0] if len(info_parts) > 0 else "غير محدد"
    subject = info_parts[1] if len(info_parts) > 1 else "طلب تقديم مستندات"
    plaintiff = info_parts[2] if len(info_parts) > 2 else "غير محدد"
    defendant = info_parts[3] if len(info_parts) > 3 else "غير محدد"
    judge = info_parts[4] if len(info_parts) > 4 else "غير محدد"

    # ربط المتغيرات مع خانات الـ Jinja2 المضافة بالـ template.docx
    context_data = {
        'court' : court,
        'subject' : subject,
        'plaintiff' : plaintiff,
        'defendant' : defendant,
        'judge' : judge,
        'body_content' : subdoc
    }

    doc.render(context_data)

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream, f"قضية_{court}.docx"

async def process_word_creation(update: Update, context: ContextTypes.DEFAULT_TYPE):
    info_text = update.message.text
    info_parts = [i.strip() for i in info_text.split(' , ')]
    texts_list = context.user_data.get("texts", [])

    await update.message.reply_text("جاري دمج النص مع القالب وتطبيق الأوامر... ⏳")

    try:
        doc_stream, filename = generate_word_from_template(info_parts, texts_list)
        doc_bytes = doc_stream.getvalue()

        await update.message.reply_document(
            document=io.BytesIO(doc_bytes),
            filename=filename,
            caption="📄 تم إنشاء المستند باستخدام القالب بنجاح!"
        )

        context.user_data["bag"][filename] = doc_bytes

    except Exception as e:
        logging.error(f"Error creating Word document: {e}")
        await update.message.reply_text(f"❌ حدث خطأ أثناء الدمج: {e}\n(تأكد من وجود ملف template.docx في السيرفر)")

    return ConversationHandler.END

# ==========================================
# الخدمات الجانبية
# ==========================================

async def view_bag(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bag = context.user_data.get("bag", {})
    if not bag:
        await update.message.reply_text("💼 حقيبتك فارغة حالياً.")
        return

    await update.message.reply_text(f"💼 تحتوي حقيبتك على {len(bag)} ملف(ات):")
    for filename, doc_bytes in bag.items():
        await update.message.reply_document(
            document=io.BytesIO(doc_bytes),
            filename=filename
        )

async def clear_session(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["texts"] = []
    markup = ReplyKeyboardMarkup(MAIN_KEYBOARD, resize_keyboard=True)
    await update.message.reply_text("❌ تم إنهاء الجلسة ومسح النصوص المؤقتة.", reply_markup=markup)

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    markup = ReplyKeyboardMarkup(MAIN_KEYBOARD, resize_keyboard=True)
    await update.message.reply_text("تم إلغاء العملية.", reply_markup=markup)
    return ConversationHandler.END

def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()

    ocr_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^📝 استخراج النص$"), start_ocr)],
        states={WAITING_IMAGE: [MessageHandler(filters.PHOTO, process_image)]},
        fallbacks=[CommandHandler("cancel", cancel)]
    )

    word_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^📄 إنشاء ملف Word$"), start_word)],
        states={WAITING_INFO: [MessageHandler(filters.TEXT & ~filters.COMMAND, process_word_creation)]},
        fallbacks=[CommandHandler("cancel", cancel)]
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(ocr_handler)
    app.add_handler(word_handler)
    app.add_handler(MessageHandler(filters.Regex("^💼 حقيبتك$"), view_bag))
    app.add_handler(MessageHandler(filters.Regex("^❌ إنهاء الجلسة$"), clear_session))

    print("البوت يعمل بنجاح مع القالب والدقة المحسنة وتتبع التوكنز...")
    app.run_polling()

if __name__ == "__main__":
    main()
