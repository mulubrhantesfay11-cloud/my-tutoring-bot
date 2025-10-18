import logging
from pathlib import Path
import sqlite3
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputFile
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ConversationHandler,
    CallbackQueryHandler, ContextTypes, filters
)

# ---------- CONFIG ----------
TOKEN = "7642787472:AAGV4KwuQoJXBxxajQWnVxBbZYqupHFqbrw"
ADMIN_IDS = [5988494903, 5872954068, 5813730985]  # Add all admin Telegram IDs here

# ---------- Logging ----------
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ---------- Data directories ----------
DATA_DIR = Path("data")
FILES_DIR = DATA_DIR / "uploads"
FILES_DIR.mkdir(parents=True, exist_ok=True)

# ---------- States ----------
(
    LANG_CHOICE, MENU,
    STU_NAME, STU_GRADE, STU_PHONE, STU_ADDRESS, STU_PREF_GENDER,
    TUT_NAME, TUT_GENDER, TUT_PHONE, TUT_GRADES, TUT_PROFILE_UPLOAD,
    CONTACT_ADMIN, ADMIN_MENU
) = range(14)

# ---------- Storage ----------
user_lang = {}

# ---------- Language texts ----------
TEXTS = {
    "en": {
        "start": "Welcome! Please choose your language:",
        "menu": "Choose an option:",
        "student": "Apply as Student/Family",
        "tutor": "Apply as Tutor",
        "contact_admin": "Contact Admin",
        "change_lang": "🌐 Change Language",
        "cancel": "Cancel",
        "admin": "Admin Panel",
        "stu_done": "✅ Student/Family application submitted! for more information contact us at @Mentorxham    ",
        "tut_done": "✅ Tutor application submitted! for more information contact us at @Mentorxham    ",
        "not_admin": "⛔ You are not authorized to access admin features.",
        "new_contact": "📩 New message from {name} (ID: {user_id}):",
        "menu_header": "Main Menu",
        "enter_name": "Enter full name:",
        "enter_grade": "Enter grade:",
        "enter_phone": "Enter phone number:",
        "enter_address": "Enter address:",
        "enter_pref_gender": "Preferred tutor's gender (Male/Female/any):",
        "enter_tut_name": "Enter full name:",
        "enter_tut_gender": "Enter gender (Male/Female):",
        "enter_tut_phone": "Enter phone number:",
        "enter_tut_grades": "Enter grades you can teach (1-4/5-8/9-12/any grade):",
        "upload_profile": "Please upload your academic profile (e.g., Grade 12 result) as a file or photo through @Mentorxham, and type 'done' here after completing the upload.",
        "send_to_admin": "You can ask any question or send a file/photo",
        "contact_admin_done": "✅ Your message has been sent to the admin. for more information contact us at @Mentorxham   "
    },
    "ti": {
        "start": "እንኳዕ ብደሓን መጻእኹም ፤ ቋንቋ ይምረፁ።",
        "menu": "ማውጫ፡",
        "student": "ከም ተመሃሮ/ቤተሰብ ተመዝገብ",
        "tutor": "ከም መምህር ተመዝገብ",
        "contact_admin": "ናብ admin ጽሓፍ",
        "change_lang": "🌐 ቋንቋ ቀይር",
        "cancel": "አቋርጽ",
        "admin": "ናይ admin መደብ",
        "stu_done": "✅ ናይ ተመሃሮ ምዝገባ ብ ትክክል ወዲኦም አለዉ! ን ዝበለጸ ሓበሬታ በዚ @Mentorxham ይርከቡና።",
        "tut_done": "✅ ናይ መምህር ምዝገባ ብ ትክክል ወዲኦም አለዉ! ን ዝበለጸ ሓበሬታ በዚ @Mentorxham ይርከቡና።",
        "not_admin": "⛔ ይቅርታ admin ኣይኮኑን።",
        "new_contact": "📩 ሓደ መልእኽቲ ካብ {name} (ID: {user_id}):",
        "menu_header": "ዋና ማውጫ",
        "enter_name": "ሙሉእ ስም የእትዉ።",
        "enter_grade": "ክፍሊ የእትዉ።",
        "enter_phone": "ስልኪ ቁጽሪ የእትዉ።",
        "enter_address": "አድራሻ የእትዉ።",
        "enter_pref_gender": "ናይ መምህር ፆታ ይምረጹ (ተባ/አን/ዝኾነ):",
        "enter_tut_name": "ሙሉእ ስም የእትዉ።",
        "enter_tut_gender": "ፆታ የእትዉ (ተባ/አን):",
        "enter_tut_phone": "ስልኪ ቁጽሪ የእትዉ።",
        "enter_tut_grades": "ከምህርዎም ዝኸሉ ክፍሊታት የአትዉ (1-4/5-8/9-12/ዝኾነ)",
        "upload_profile": "አካዳሚ ፕሮፋይሎም  (አብነት - ናይ 12 ክፍሊ ውጽኢት)(ፋይል/ፎቶ) ናብ @Mentorxham ይልኣኹ :ምስ ወድኡ ኣብዚ 'done' ኢሎም ይልአኹ።",
        "send_to_admin": "ዘለዎም መልእኽቲ ወይ ሕቶ ናብ admin ይልአኹ",
        "contact_admin_done": "✅ መልእኽቶም ብ ትኽኽል ናብ admin ተላኢኹ። ን ዝበለጸ ሓበሬታ በዚ @Mentorxham ይርከቡና።"
    }
}

# ---------- SQLite Setup ----------
DB_FILE = DATA_DIR / "bot_data.sqlite3"
conn = sqlite3.connect(DB_FILE)
c = conn.cursor()
c.execute('''
CREATE TABLE IF NOT EXISTS students (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    grade TEXT,
    phone TEXT,
    address TEXT,
    pref_gender TEXT
)
''')
c.execute('''
CREATE TABLE IF NOT EXISTS tutors (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    gender TEXT,
    phone TEXT,
    grades TEXT,
    profile_path TEXT
)
''')
conn.commit()
conn.close()

# ---------- Helper Functions ----------
def get_main_menu_keyboard(lang="en"):
    keyboard = [
        [InlineKeyboardButton(TEXTS[lang]["student"], callback_data="student"),
         InlineKeyboardButton(TEXTS[lang]["tutor"], callback_data="tutor")],
        [InlineKeyboardButton(TEXTS[lang]["contact_admin"], callback_data="contact_admin")],
        [InlineKeyboardButton(TEXTS[lang]["change_lang"], callback_data="change_lang")],
        [InlineKeyboardButton(TEXTS[lang]["cancel"], callback_data="cancel")]
    ]
    return InlineKeyboardMarkup(keyboard)

async def return_to_main_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lang = user_lang.get(update.effective_user.id, "en")
    if update.message:
        await update.message.reply_text(TEXTS[lang]["menu"], reply_markup=get_main_menu_keyboard(lang))
    elif update.callback_query:
        await update.callback_query.edit_message_text(TEXTS[lang]["menu"], reply_markup=get_main_menu_keyboard(lang))
    context.user_data.clear()
    return MENU

async def forward_to_admins(update: Update, context: ContextTypes.DEFAULT_TYPE, message_text, file_path=None):
    for admin_id in ADMIN_IDS:
        msg = TEXTS[user_lang.get(update.effective_user.id,"en")]["new_contact"].format(
            name=update.effective_user.full_name,
            user_id=update.effective_user.id
        )
        await context.bot.send_message(admin_id, msg)
        if message_text:
            await context.bot.send_message(admin_id, message_text)
        if file_path:
            await context.bot.send_document(admin_id, InputFile(file_path))

# ---------- Start / Language ----------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [[InlineKeyboardButton("English", callback_data="lang_en"),
                 InlineKeyboardButton("ትግርኛ", callback_data="lang_ti")]]
    await update.message.reply_text(TEXTS["en"]["start"], reply_markup=InlineKeyboardMarkup(keyboard))
    return LANG_CHOICE

async def lang_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    user_lang[user_id] = "en" if query.data=="lang_en" else "ti"
    await query.edit_message_text(TEXTS[user_lang[user_id]]["menu"], reply_markup=get_main_menu_keyboard(user_lang[user_id]))
    return MENU

# ---------- Menu Handler ----------
async def menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    lang = user_lang.get(update.effective_user.id, "en")

    if data == "cancel":
        return await return_to_main_menu(update, context)
    if data == "change_lang":
        keyboard = [[InlineKeyboardButton("English", callback_data="lang_en"),
                     InlineKeyboardButton("ትግርኛ", callback_data="lang_ti")]]
        await query.edit_message_text(TEXTS[lang]["start"], reply_markup=InlineKeyboardMarkup(keyboard))
        return LANG_CHOICE

    if data == "student":
        await query.edit_message_text(TEXTS[lang]["enter_name"], reply_markup=None)
        return STU_NAME
    if data == "tutor":
        await query.edit_message_text(TEXTS[lang]["enter_tut_name"], reply_markup=None)
        return TUT_NAME
    if data == "contact_admin":
        await query.edit_message_text(TEXTS[lang]["send_to_admin"], reply_markup=None)
        return CONTACT_ADMIN
    if data == "admin":
        if query.from_user.id not in ADMIN_IDS:
            await query.edit_message_text(TEXTS[lang]["not_admin"], reply_markup=get_main_menu_keyboard(lang))
            return MENU
        msg = "Admin Panel:\n- Search\n- Match\n- Export"
        await query.edit_message_text(msg, reply_markup=None)
        return ADMIN_MENU

    return MENU

# ---------- Student Handlers ----------
async def stu_name(update, context):
    lang = user_lang.get(update.effective_user.id, "en")
    context.user_data["stu_name"] = update.message.text
    await update.message.reply_text(TEXTS[lang]["enter_grade"])
    return STU_GRADE

async def stu_grade(update, context):
    lang = user_lang.get(update.effective_user.id, "en")
    context.user_data["stu_grade"] = update.message.text
    await update.message.reply_text(TEXTS[lang]["enter_phone"])
    return STU_PHONE

async def stu_phone(update, context):
    lang = user_lang.get(update.effective_user.id, "en")
    context.user_data["stu_phone"] = update.message.text
    await update.message.reply_text(TEXTS[lang]["enter_address"])
    return STU_ADDRESS

async def stu_address(update, context):
    lang = user_lang.get(update.effective_user.id, "en")
    context.user_data["stu_address"] = update.message.text
    await update.message.reply_text(TEXTS[lang]["enter_pref_gender"])
    return STU_PREF_GENDER

async def stu_pref_gender(update, context):
    lang = user_lang.get(update.effective_user.id, "en")
    context.user_data["stu_pref_gender"] = update.message.text
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        INSERT INTO students (name, grade, phone, address, pref_gender)
        VALUES (?, ?, ?, ?, ?)
    ''', (context.user_data['stu_name'], context.user_data['stu_grade'],
          context.user_data['stu_phone'], context.user_data['stu_address'],
          context.user_data['stu_pref_gender']))
    conn.commit()
    conn.close()
    await update.message.reply_text(TEXTS[lang]["stu_done"], reply_markup=get_main_menu_keyboard(lang))
    msg = f"Student/Family Application:\nName: {context.user_data['stu_name']}\nGrade: {context.user_data['stu_grade']}\nPhone: {context.user_data['stu_phone']}\nAddress: {context.user_data['stu_address']}\nPreferred Tutor Gender: {context.user_data['stu_pref_gender']}"
    await forward_to_admins(update, context, msg)
    context.user_data.clear()
    return MENU

# ---------- Tutor Handlers ----------
async def tut_name(update, context):
    lang = user_lang.get(update.effective_user.id, "en")
    context.user_data["tut_name"] = update.message.text
    await update.message.reply_text(TEXTS[lang]["enter_tut_gender"])
    return TUT_GENDER

async def tut_gender(update, context):
    lang = user_lang.get(update.effective_user.id, "en")
    context.user_data["tut_gender"] = update.message.text
    await update.message.reply_text(TEXTS[lang]["enter_tut_phone"])
    return TUT_PHONE

async def tut_phone(update, context):
    lang = user_lang.get(update.effective_user.id, "en")
    context.user_data["tut_phone"] = update.message.text
    await update.message.reply_text(TEXTS[lang]["enter_tut_grades"])
    return TUT_GRADES

async def tut_grades(update, context):
    lang = user_lang.get(update.effective_user.id, "en")
    context.user_data["tut_grades"] = update.message.text
    await update.message.reply_text(TEXTS[lang]["upload_profile"])
    return TUT_PROFILE_UPLOAD

async def tut_profile_upload(update, context):
    lang = user_lang.get(update.effective_user.id, "en")
    file_path = None
    if update.message.document:
        file = await update.message.document.get_file()
        file_path = FILES_DIR / f"{update.message.document.file_unique_id}_{update.message.document.file_name}"
        await file.download_to_drive(file_path)
    elif update.message.photo:
        photo = update.message.photo[-1]
        file = await photo.get_file()
        file_path = FILES_DIR / f"{photo.file_unique_id}.jpg"
        await file.download_to_drive(file_path)

    context.user_data["tut_profile"] = str(file_path) if file_path else None
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        INSERT INTO tutors (name, gender, phone, grades, profile_path)
        VALUES (?, ?, ?, ?, ?)
    ''', (context.user_data['tut_name'], context.user_data['tut_gender'],
          context.user_data['tut_phone'], context.user_data['tut_grades'],
          context.user_data['tut_profile']))
    conn.commit()
    conn.close()

    await update.message.reply_text(TEXTS[lang]["tut_done"], reply_markup=get_main_menu_keyboard(lang))
    msg = f"Tutor Application:\nName: {context.user_data['tut_name']}\nGender: {context.user_data['tut_gender']}\nPhone: {context.user_data['tut_phone']}\nGrades: {context.user_data['tut_grades']}"
    await forward_to_admins(update, context, msg, file_path)
    context.user_data.clear()
    return MENU

# ---------- Contact Admin ----------
async def contact_admin(update, context):
    lang = user_lang.get(update.effective_user.id, "en")
    text = update.message.text
    file_path = None
    if update.message.document:
        file = await update.message.document.get_file()
        file_path = FILES_DIR / f"{update.message.document.file_unique_id}_{update.message.document.file_name}"
        await file.download_to_drive(file_path)
    elif update.message.photo:
        photo = update.message.photo[-1]
        file = await photo.get_file()
        file_path = FILES_DIR / f"{photo.file_unique_id}.jpg"
        await file.download_to_drive(file_path)
    await forward_to_admins(update, context, text, file_path)
    await update.message.reply_text(TEXTS[lang]["contact_admin_done"], reply_markup=get_main_menu_keyboard(lang))
    context.user_data.clear()
    return MENU

# ---------- Admin Export ----------
async def export_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ You are not authorized.")
        return
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT * FROM students")
    students_data = c.fetchall()
    c.execute("SELECT * FROM tutors")
    tutors_data = c.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Students"
    ws1.append(["ID","Name","Grade","Phone","Address","Preferred Gender"])
    for row in students_data:
        ws1.append(row)
    ws2 = wb.create_sheet("Tutors")
    ws2.append(["ID","Name","Gender","Phone","Grades","Profile Path"])
    for row in tutors_data:
        ws2.append(row)

    export_file = DATA_DIR / "export.xlsx"
    wb.save(export_file)
    await update.message.reply_document(InputFile(export_file))

# ---------- Cancel Handler ----------
async def cancel(update, context):
    return await return_to_main_menu(update, context)

# ---------- Main for Render/Vercel Deployment ----------
from fastapi import FastAPI, Request
from telegram import Update
import asyncio

# --- Initialize FastAPI ---
app = FastAPI()

# --- Initialize Telegram Bot Application ---
application = Application.builder().token(TOKEN).build()

conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            LANG_CHOICE: [CallbackQueryHandler(lang_choice)],
            MENU: [CallbackQueryHandler(menu_handler)],
            STU_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, stu_name)],
            STU_GRADE: [MessageHandler(filters.TEXT & ~filters.COMMAND, stu_grade)],
            STU_PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, stu_phone)],
            STU_ADDRESS: [MessageHandler(filters.TEXT & ~filters.COMMAND, stu_address)],
            STU_PREF_GENDER: [MessageHandler(filters.TEXT & ~filters.COMMAND, stu_pref_gender)],
            TUT_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, tut_name)],
            TUT_GENDER: [MessageHandler(filters.TEXT & ~filters.COMMAND, tut_gender)],
            TUT_PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, tut_phone)],
            TUT_GRADES: [MessageHandler(filters.TEXT & ~filters.COMMAND, tut_grades)],
            TUT_PROFILE_UPLOAD: [MessageHandler(filters.ALL & ~filters.COMMAND, tut_profile_upload)],
            CONTACT_ADMIN: [MessageHandler(filters.ALL & ~filters.COMMAND, contact_admin)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        per_user=True,
        per_chat=True,
        allow_reentry=True
    )

    app.add_handler(conv_handler)
    app.add_handler(CommandHandler("export", export_data))

# --- Webhook Endpoint ---
@app.post("/webhook")
async def webhook(request: Request):
    """Handles incoming updates from Telegram via webhook."""
    try:
        data = await request.json()
        update = Update.de_json(data, application.bot)
        await application.process_update(update)
    except Exception as e:
        print(f"Error handling update: {e}")
        return {"ok": False, "error": str(e)}
    return {"ok": True}

# --- Health Check Endpoint ---
@app.get("/")
def home():
    return {"status": "✅ Bot is running successfully on Render/Vercel"}

# --- Local Test Mode (optional) ---
if __name__ == "__main__":
    import uvicorn
    print("🚀 Starting local server at http://127.0.0.1:8000 ...")
    uvicorn.run("main:app", host="0.0.0.0", port=8000)


