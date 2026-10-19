import os
import sys
import json
import html
import io
import zipfile
import urllib.parse
import sqlite3
from google import genai
from pptx import Presentation
from pptx.util import Pt
from docx import Document
from docx.shared import Inches
from PIL import Image

# Задаем кодировку UTF-8
os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["LANG"] = "en_US.UTF-8"

import asyncio
import logging

from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import (
    ReplyKeyboardMarkup, 
    KeyboardButton, 
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    BotCommand,
    WebAppInfo,
    MenuButtonDefault,
    BufferedInputFile
)

# Импорты для FastAPI веб-сервера
from fastapi import FastAPI
import uvicorn
from threading import Thread

BOT_TOKEN = "8801689561:AAFgWYikHCNmJwrAM8RfZELM36laUyU4Prc"
GEMINI_API_KEY = "AQ.Ab8RN6IVvQz8x1XDCPWPFTy0SrcKpDEosTp9VTLKOMVpqs_qRQ"
ADMIN_ID = 762012445                  

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
ai_client = genai.Client(api_key=GEMINI_API_KEY)

# --- РАБОТА С БАЗОЙ ДАННЫХ (SQLite) ---
DB_NAME = "bot_database.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            balance INTEGER DEFAULT 3000
        )
    ''')
    conn.commit()
    conn.close()

def get_user_balance(user_id: int) -> int:
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('SELECT balance FROM users WHERE user_id = ?', (user_id,))
    row = cursor.fetchone()
    if row is None:
        # Если пользователя нет в базе, даем начальный бонус 3000 сум
        cursor.execute('INSERT OR REPLACE INTO users (user_id, balance) VALUES (?, ?)', (user_id, 3000))
        conn.commit()
        balance = 3000
    else:
        balance = row[0]
    conn.close()
    return balance

def update_user_balance(user_id: int, new_balance: int):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('INSERT OR REPLACE INTO users (user_id, balance) VALUES (?, ?)', (user_id, new_balance))
    conn.commit()
    conn.close()
# -------------------------------------

# Настройка FastAPI приложения для Render
app = FastAPI()

@app.get("/")
def index():
    return {"status": "Bot is running online!"}

def run_fastapi():
    port = int(os.environ.get("PORT", 10000))
    uvicorn.run(app, host="0.0.0.0", port=port)

waiting_for_receipt = set()
waiting_for_pdf_photos = set()
user_pdf_images = {}

waiting_for_zip_files = set()
user_zip_files = {}

waiting_for_edit_file = set()

def get_main_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="PDF yaratish 📕"), KeyboardButton(text="Fayllarni zip qilish 🗂")],
            [
                KeyboardButton(text="Taqdimot tayyorlash 🖼", web_app=WebAppInfo(url="https://farruh707.github.io/ai-presentation-bot/index.html")), 
                KeyboardButton(text="Akademik Slayd yaratish 🎓")
            ],
            [KeyboardButton(text="Media Slaydlar 🎬")],
            [KeyboardButton(text="pptx ➔ pdf"), KeyboardButton(text="Slaydni tahrirlash 🛠")],
            [
                KeyboardButton(text="Word tahrirlash 📝"), 
                KeyboardButton(
                    text="Titul varag'i 📄", 
                    web_app=WebAppInfo(url="https://farruh707.github.io/ai-presentation-bot/titul.html")
                )
            ],
            [KeyboardButton(text="Obyektivka 📎")]
        ],
        resize_keyboard=True
    )

async def set_bot_commands(bot: Bot):
    commands = [
        BotCommand(command="start", description="Start - Botni ishga tushirish"),
        BotCommand(command="help", description="Bot bo'yicha qo'llanma"),
        BotCommand(command="balance", description="Balansingizni tekshirish"),
        BotCommand(command="buy", description="Balansni to'ldirish / Xarid"),
        BotCommand(command="chek", description="Chek yuborish"),
        BotCommand(command="cancel", description="Operatsiyani bekor qilish"),
    ]
    await bot.set_my_commands(commands)

@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    try:
        await bot.set_chat_menu_button(chat_id=message.from_user.id, menu_button=MenuButtonDefault())
    except Exception:
        pass

    # Автоматически регистрируем пользователя в БД и даем бонус при первом входе
    get_user_balance(message.from_user.id)

    welcome_text = (
        f"Salom **{message.from_user.full_name}**, AI Taqdimot | Slayd | PPT | PDF Bot ga xush kelibsiz\n\n"
        "Men taqdimot tayyorlash, hujjatlarni vizual tarzda taqdim etish va fayllarni boshqarish bo'yicha yordam beradigan aqlli botman.\n\n"
        "/balance - sizning bot hamyonidagi balansingiz. Agar yangi a'zo bo'lgan bo'lsangiz sizda 3000 so'm bonus bor\n"
        "/buy - xarid bo'limi"
    )
    await message.answer(welcome_text, reply_markup=get_main_keyboard(), parse_mode="Markdown")

@dp.message(Command("help"))
async def cmd_help(message: types.Message):
    help_text = (
        "📖 **Botdan foydalanish bo'yicha qo'llanma:**\n\n"
        "Batafsil ma'lumot va ko'rsatmalarni quyidagi havola orqali o'qishingiz mumkin:"
    )
    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📖 Qo'llanmani ochish", url="https://telegra.ph/PDF-maker-bot--PDF-file-qollanmasi-12-05")]
        ]
    )
    await message.answer(help_text, reply_markup=keyboard, parse_mode="Markdown")

@dp.message(Command("cancel"))
async def cmd_cancel(message: types.Message):
    waiting_for_receipt.discard(message.from_user.id)
    waiting_for_pdf_photos.discard(message.from_user.id)
    user_pdf_images.pop(message.from_user.id, None)
    waiting_for_zip_files.discard(message.from_user.id)
    user_zip_files.pop(message.from_user.id, None)
    waiting_for_edit_file.discard(message.from_user.id)
    await message.answer("❌ Amaliyot bekor qilindi.", reply_markup=get_main_keyboard())

@dp.message(Command("balance"))
async def cmd_balance(message: types.Message):
    bal = get_user_balance(message.from_user.id)
    await message.answer(f"💳 Sizning balansingiz: **{bal} so'm**", parse_mode="Markdown", reply_markup=get_main_keyboard())

@dp.message(Command("buy"))
async def cmd_buy(message: types.Message):
    buy_text = (
        "💰 **Xarid Bo'limi**\n\n"
        "Taqdimotda yoziladiga ma'lumotlar ikki xil sifatga ega. Ushbu sifatlarga ko'ra bitta taqdimot yaratish narxlari quyidagicha:\n\n"
        "⭐️ **Standart**: 3000 so'm.\n"
        "⭐️⭐️ **Premium**: 5000 so'm.\n"
        "⭐️⭐️⭐️ **Yuqori sifatli Premium**: 7000 so'm."
    )
    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="💳 Bank kartasi orqali to'lov", callback_data="pay_card")]
        ]
    )
    await message.answer(buy_text, reply_markup=keyboard, parse_mode="Markdown")

@dp.callback_query(F.data == "pay_card")
async def process_card_choice(callback: types.CallbackQuery):
    card_info_text = (
        "Taqdimot uchun to'lov:\n\n"
        "1. To'lov: 9860170115658388 (Farrux Khalilov)\n"
        "2. Buyruq: Botga `/chek` buyrug'ini yuboring.\n"
        "3. Yuboring: To'lov chekining rasm yoki faylini botga tashlang."
    )
    admin_keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="👨‍💻 Administratorga yozish (@farrux204)", url="https://t.me/farrux204")]
        ]
    )
    await callback.message.answer(card_info_text, reply_markup=admin_keyboard, parse_mode="Markdown")
    await callback.answer()

@dp.message(Command("chek"))
async def cmd_chek(message: types.Message):
    waiting_for_receipt.add(message.from_user.id)
    await message.answer(
        "📷 **Iltimos, to'lov chekining rasm yoki skrinshotini yuboring**:",
        reply_markup=ReplyKeyboardMarkup(
            keyboard=[[KeyboardButton(text="🔙 Ortga qaytish")]],
            resize_keyboard=True
        ),
        parse_mode="Markdown"
    )

@dp.message(F.text == "🔙 Ortga qaytish")
async def btn_back(message: types.Message):
    waiting_for_receipt.discard(message.from_user.id)
    waiting_for_pdf_photos.discard(message.from_user.id)
    user_pdf_images.pop(message.from_user.id, None)
    waiting_for_zip_files.discard(message.from_user.id)
    user_zip_files.pop(message.from_user.id, None)
    waiting_for_edit_file.discard(message.from_user.id)
    await message.answer("🏠 Asosiy menyu:", reply_markup=get_main_keyboard())

@dp.message(F.text == "PDF yaratish 📕")
async def btn_pdf_start(message: types.Message):
    waiting_for_pdf_photos.add(message.from_user.id)
    user_pdf_images[message.from_user.id] = []
    
    keyboard = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="✅ Tayyor (PDF yaratish)")],
            [KeyboardButton(text="🔙 Ortga qaytish")]
        ],
        resize_keyboard=True
    )
    await message.answer("iltimos PDF generatsiya qilish uchun rasm yuboring. Barcha rasmlarni yuborib bo'lgach, **'✅ Tayyor (PDF yaratish)'** tugmasini bosing.", reply_markup=keyboard, parse_mode="Markdown")

@dp.message(F.text == "✅ Tayyor (PDF yaratish)")
async def generate_pdf_from_images(message: types.Message):
    user_id = message.from_user.id
    if user_id not in waiting_for_pdf_photos or not user_pdf_images.get(user_id):
        await message.answer("⚠️ Hali hech qanday rasm yubormadingiz!", reply_markup=get_main_keyboard())
        waiting_for_pdf_photos.discard(user_id)
        return

    await message.answer("⏳ Rasmlar PDF formatiga o'tkazilmoqda...", parse_mode="Markdown")

    try:
        images = user_pdf_images[user_id]
        pil_images = []

        for img_bytes in images:
            img = Image.open(io.BytesIO(img_bytes))
            if img.mode != 'RGB':
                img = img.convert('RGB')
            pil_images.append(img)

        pdf_stream = io.BytesIO()
        pil_images[0].save(
            pdf_stream, 
            format="PDF", 
            save_all=True, 
            append_images=pil_images[1:]
        )
        pdf_stream.seek(0)

        input_file = BufferedInputFile(pdf_stream.getvalue(), filename="hujjat.pdf")
        await message.answer_document(input_file, caption="✅ Sizning PDF faylingiz tayyor!", reply_markup=get_main_keyboard())

    except Exception as e:
        logging.error(f"PDF generation error: {e}")
        await message.answer("❌ PDF yaratishda xatolik yuz berdi.", reply_markup=get_main_keyboard())

    waiting_for_pdf_photos.discard(user_id)
    user_pdf_images.pop(user_id, None)

@dp.message(F.text == "Fayllarni zip qilish 🗂")
async def btn_zip_start(message: types.Message):
    waiting_for_zip_files.add(message.from_user.id)
    user_zip_files[message.from_user.id] = []
    
    keyboard = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="✅ Tayyor (ZIP yaratish)")],
            [KeyboardButton(text="🔙 Ortga qaytish")]
        ],
        resize_keyboard=True
    )
    await message.answer("📁 Iltimos, arxivga qo'shish uchun fayllar yoki rasmlarni yuboring. Hammasini yuborib bo'lgach, **'✅ Tayyor (ZIP yaratish)'** tugmasini bosing.", reply_markup=keyboard, parse_mode="Markdown")

@dp.message(F.text == "✅ Tayyor (ZIP yaratish)")
async def generate_zip_archive(message: types.Message):
    user_id = message.from_user.id
    if user_id not in waiting_for_zip_files or not user_zip_files.get(user_id):
        await message.answer("⚠️ Hali hech qanday fayl yubormadingiz!", reply_markup=get_main_keyboard())
        waiting_for_zip_files.discard(user_id)
        return

    await message.answer("⏳ Fayllar ZIP arxivga jamlanmoqda...", parse_mode="Markdown")

    try:
        zip_stream = io.BytesIO()
        with zipfile.ZipFile(zip_stream, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for idx, (fname, fbytes) in enumerate(user_zip_files[user_id]):
                filename = fname if fname else f"file_{idx+1}.bin"
                zip_file.writestr(filename, fbytes)
        
        zip_stream.seek(0)
        input_file = BufferedInputFile(zip_stream.getvalue(), filename="arxiv.zip")
        await message.answer_document(input_file, caption="✅ Sizning ZIP arxivingiz tayyor!", reply_markup=get_main_keyboard())

    except Exception as e:
        logging.error(f"ZIP generation error: {e}")
        await message.answer("❌ Arxiv yaratishda xatolik yuz berdi.", reply_markup=get_main_keyboard())

    waiting_for_zip_files.discard(user_id)
    user_zip_files.pop(user_id, None)

@dp.message(F.text == "pptx ➔ pdf")
async def btn_conv_start(message: types.Message):
    await message.answer("⚠️ Ushbu funksiya hozirda serverda vaqtincha o'chirilgan.", reply_markup=get_main_keyboard())

@dp.message(F.text == "Slaydni tahrirlash 🛠")
async def btn_edit_slide_start(message: types.Message):
    waiting_for_edit_file.add(message.from_user.id)
    keyboard = ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="🔙 Ortga qaytish")]],
        resize_keyboard=True
    )
    await message.answer("📎 Tahrirlash uchun **.pptx** faylni yuboring:", reply_markup=keyboard, parse_mode="Markdown")

@dp.message(F.photo | F.document)
async def handle_incoming_files(message: types.Message):
    user = message.from_user
    
    if user.id in waiting_for_receipt:
        waiting_for_receipt.discard(user.id)
        admin_caption = f"📥 <b>Yangi to'lov cheki keldi!</b>\n👤 Foydalanuvchi: {html.escape(user.full_name)} (ID: <code>{user.id}</code>)"
        
        admin_keyboard = InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(text="3000 so'm", callback_data=f"accept_{user.id}_3000"),
                    InlineKeyboardButton(text="5000 so'm", callback_data=f"accept_{user.id}_5000"),
                    InlineKeyboardButton(text="7000 so'm", callback_data=f"accept_{user.id}_7000")
                ]
            ]
        )
        if message.photo:
            await bot.send_photo(ADMIN_ID, photo=message.photo[-1].file_id, caption=admin_caption, reply_markup=admin_keyboard, parse_mode="HTML")
        await message.answer("✅ Chek adminga yuborildi!", reply_markup=get_main_keyboard())
        return

    if user.id in waiting_for_pdf_photos:
        if message.photo:
            photo = message.photo[-1]
            file = await bot.get_file(photo.file_id)
            file_bytes = await bot.download_file(file.file_path)
            user_pdf_images[user.id].append(file_bytes.read())
            count = len(user_pdf_images[user.id])
            await message.answer(f"📸 Rasm qabul qilindi ({count} ta). Yana rasm yuborishingiz yoki **'✅ Tayyor (PDF yaratish)'** tugmasini bosishingiz mumkin.", parse_mode="Markdown")
        else:
            await message.answer("⚠️ Iltimos, rasm formatidagi fayl yuboring.")
        return

    if user.id in waiting_for_zip_files:
        file_id = message.document.file_id if message.document else message.photo[-1].file_id if message.photo else None
        file_name = message.document.file_name if message.document else "image.jpg"
        
        if file_id:
            file = await bot.get_file(file_id)
            file_bytes = await bot.download_file(file.file_path)
            user_zip_files[user.id].append((file_name, file_bytes.read()))
            count = len(user_zip_files[user.id])
            await message.answer(f"📁 Fayl qo'shildi ({count} ta). Yana fayl yuborishingiz yoki **'✅ Tayyor (ZIP yaratish)'** tugmasini bosing.", parse_mode="Markdown")
        return

    if user.id in waiting_for_edit_file:
        if message.document:
            file_name = message.document.file_name
            if not file_name.endswith(('.pptx', '.ppt')):
                await message.answer("⚠️ Iltimos, faqat .pptx formatidagi taqdimot faylini yuboring.")
                return
            
            waiting_for_edit_file.discard(user.id)
            
            webapp_url = "https://farruh707.github.io/ai-presentation-bot/editor.html"

            response_text = (
                f"✅ **Fayl qabul qilindi!**\n\n"
                f"📄 **Fayl nomi:** {html.escape(file_name)}\n\n"
                f"Muharrirni ochib, **📁 .pptx ochish** tugmasi orqali ushbu faylni tanlang:"
            )
            
            webapp_keyboard = InlineKeyboardMarkup(
                inline_keyboard=[
                    [InlineKeyboardButton(
                        text="✏️ Taqdimotni tahrirlash", 
                        web_app=WebAppInfo(url=webapp_url)
                    )]
                ]
            )
            
            await message.answer(response_text, reply_markup=webapp_keyboard, parse_mode="Markdown")
        return

@dp.callback_query(F.data.startswith("accept_"))
async def process_admin_decision(callback: types.CallbackQuery):
    data_parts = callback.data.split("_")
    target_user_id = int(data_parts[1])
    add_amount = int(data_parts[2])
    
    current_bal = get_user_balance(target_user_id)
    new_bal = current_bal + add_amount
    update_user_balance(target_user_id, new_bal)

    await bot.send_message(target_user_id, f"🎉 To'lovingiz tasdiqlandi! Balansingizga +{add_amount} so'm qo'shildi.")
    await callback.message.edit_caption(caption=callback.message.caption + f"\n\n✅ Tasdiqlandi (+{add_amount} so'm)")
    await callback.answer("Bajarildi!")

@dp.message(F.web_app_data)
async def handle_web_app_data(message: types.Message):
    user_id = message.from_user.id
    current_balance = get_user_balance(user_id)

    raw_data = message.web_app_data.data
    try:
        data = json.loads(raw_data)
    except Exception:
        data = {"mavzu": raw_data}

    # Определение тарифа и списание средств
    tarif = data.get("tarif", "standard")
    tariff_prices = {
        "standard": 3000,
        "premium": 5000,
        "high_premium": 7000
    }
    cost = tariff_prices.get(tarif, 3000)

    if current_balance < cost:
        await message.answer(
            f"❌ **Balansingiz yetarli emas!**\n\n"
            f"Siz tanlagan tarif narxi: {cost} so'm.\n"
            f"Sizning balansingiz: {current_balance} so'm.\n\n"
            f"Balansni to'ldirish uchun /buy buyrug'idan foydalaning.",
            reply_markup=get_main_keyboard(),
            parse_mode="Markdown"
        )
        return

    topic = data.get('mavzu') or data.get('topic') or data.get('title')
    
    if topic:
        total_requested_slides = int(data.get('sahifalar', 7))
        muallif = data.get('muallif', 'Farrux')
        til = data.get('til', "O'zbek tili")
        
        # Списание средств и сохранение в БД
        update_user_balance(user_id, current_balance - cost)
        new_balance = get_user_balance(user_id)
        
        await message.answer(f"⏳ **Sun'iy intellekt '{topic}' mavzusida {total_requested_slides} ta slayddan iborat akademik taqdimot tayyorlamoqda...** (-{cost} so'm)", parse_mode="Markdown")
        
        slides_data = []
        try:
            if tarif == "standard":
                quality_instruction = "Oddiy va tushunarli akademik til, qisqaroq va aniq ma'lumotlar."
            elif tarif == "premium":
                quality_instruction = "Chuqurlashtirilgan tahliliy ma'lumotlar, ilmiy yondashuv va kengroq tushuntirishlar."
            else:  # high_premium
                quality_instruction = "Maksimal darajada batafsil, yuqori ilmiy darajadagi professional tahlil, boy va keng qamrovli kontent."

            content_slides_count = max(1, total_requested_slides - 2)

            prompt = (
                f"Siz professional akademik taqdimot yaratuvchisiz. '{topic}' mavzusida {til}da to'liq taqdimot tuzing.\n"
                f"Sifat darajasi: {quality_instruction}\n\n"
                "Siz quyidagi qat'iy tuzilishga ega JSON massivini qaytarishingiz SHART:\n"
                "1-element (2-slayd) QАТ'ИЙ равишда quyidagicha bo'lishi kerak:\n"
                "- Sarlavhasi (title): 'Reja, Xulosa va Adabiyotlar'\n"
                "- Nuqtalar (points): \n"
                "  1. Reja: 3 ta asosiy reja bandi.\n"
                "  2. Xulosa: Qisqacha ilmiy xulosa.\n"
                "  3. Foydalanilgan adabiyotlar: 2-3 ta ilmiy manba.\n\n"
                f"Qolgan {content_slides_count} ta element esa yuqoridagi rejaga muvofiq keyingi slaydlar uchun batafsil matnlar bo'lishi kerak.\n"
                f"Jami elementlar soni {content_slides_count + 1} ta bo'lishi shart.\n"
                "Javobni FAQAT quyidagi JSON formatida yuboring, boshqa hech qanday matn yozmang:\n"
                "[\n"
                "  {\"title\": \"Reja, Xulosa va Adabiyotlar\", \"points\": [\"1. Reja: ...\", \"2. Xulosa: ...\", \"3. Adabiyotlar: ...\"]},\n"
                "  {\"title\": \"1-Reja bo'limi sarlavhasi\", \"points\": [\"Batafsil ma'lumot 1\", \"Batafsil ma'lumot 2\"]}\n"
                "]"
            )
            
            response = ai_client.models.generate_content(
                model='gemini-2.5-flash',
                contents=prompt,
                config={
                    "response_mime_type": "application/json"
                }
            )
            
            clean_text = response.text.strip()
            if "```json" in clean_text:
                clean_text = clean_text.split("```json")[1].split("```")[0].strip()
            elif "```" in clean_text:
                clean_text = clean_text.split("```")[1].split("```")[0].strip()
                
            json_start = clean_text.find('[')
            json_end = clean_text.rfind(']')
            if json_start != -1 and json_end != -1:
                clean_text = clean_text[json_start:json_end+1]

            slides_data = json.loads(clean_text)

        except Exception as e:
            logging.error(f"JSON parsing or AI error: {e}")
            slides_data = []

        try:
            prs = Presentation()
            
            # Слайд 1: Титульный
            slide_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = topic
            slide.placeholders[1].text = f"Tayyorladi: {muallif}\nTarif: {tarif.upper()}"

            if not slides_data:
                slides_data = [
                    {
                        "title": "Reja, Xulosa va Adabiyotlar",
                        "points": [
                            "Reja: 1. Mavzuning dolzarbligi; 2. Asosiy tahlil; 3. Yakuniy natijalar.",
                            "Xulosa: Mavzu bo'yicha umumiy ilmiy xulosalar.",
                            "Adabiyotlar: O'zbekiston Respublikasi qonunchiligi va ilmiy manbalar."
                        ]
                    }
                ]
                for i in range(2, total_requested_slides):
                    slides_data.append({
                        "title": f"{i}-bo'lim: {topic} tahlili",
                        "points": [f"{topic} bo'yicha ilmiy yondashuvlar", f"Amaliy ahamiyati va tahlili {i}"]
                    })

            for s_info in slides_data:
                slide_layout = prs.slide_layouts[1]
                slide = prs.slides.add_slide(slide_layout)
                
                title_shape = slide.shapes.title
                title_shape.text = s_info.get('title', 'Sarlavha')
                
                body_shape = slide.placeholders[1]
                tf = body_shape.text_frame
                tf.text = ""
                
                for point in s_info.get('points', []):
                    p = tf.add_paragraph()
                    p.text = point
                    p.level = 0
                    p.space_after = Pt(8)

            file_stream = io.BytesIO()
            prs.save(file_stream)
            file_stream.seek(0)

            filename = f"Slayd_{muallif}_{topic[:10]}.pptx".replace(" ", "_")
            input_file = BufferedInputFile(file_stream.getvalue(), filename=filename)
            
            total_created = len(slides_data) + 1
            await message.answer(
                f"✅ **Akademik taqdimotingiz tayyor!**\n\n📌 Mavzu: {topic}\n🏷 Tarif: {tarif.upper()} ({cost} so'm)\n📄 Sahifalar soni: {total_created} ta\n💳 Qoldiq balans: **{new_balance} so'm**\n\nMarhamat, yuklab oling:",
                reply_markup=get_main_keyboard(),
                parse_mode="Markdown"
            )
            await message.answer_document(input_file)
            return

        except Exception as e:
            logging.error(f"Presentation generation error: {e}")
            await message.answer("❌ Taqdimotni yaratishda xatolik yuz berdi. Iltimos qaytadan urinib ko'ring.", reply_markup=get_main_keyboard())
            return

@dp.message(F.text)
async def handle_other_text(message: types.Message):
    await message.answer("❗ Iltimos, menyudan kerakli tugmani tanlang.", reply_markup=get_main_keyboard())

async def main():
    logging.basicConfig(level=logging.INFO)
    
    # Инициализируем базу данных при запуске
    init_db()
    
    # Запускаем FastAPI в отдельном потоке
    server_thread = Thread(target=run_fastapi, daemon=True)
    server_thread.start()
    
    # Сбрасываем старые веб-хуки
    await bot.delete_webhook(drop_pending_updates=True)
    
    print("Bot va FastAPI server muvaffaqiyatli ishga tushdi (Baza ulangan)!")
    await set_bot_commands(bot)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
