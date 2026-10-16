import os
import sys
import json
import html
import io
import zipfile
#import comtypes.client
from google import genai
from pptx import Presentation
from pptx.util import Pt
from docx import Document
from docx.shared import Inches
from PIL import Image

# Задаем кодировку UTF-8
os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["LANG"] = "en_US.UTF-8"

if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr.encoding.lower() != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8')

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

BOT_TOKEN = "8801689561:AAFgWYikHCNmJwrAM8RfZELM36laUyU4Prc"
GEMINI_API_KEY = "AQ.Ab8RN6IVvQz8x1XDCPWPFTy0SrcKpDEosTp9VTLKOMVpqs_qRQ"
ADMIN_ID = 762012445                  

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
ai_client = genai.Client(api_key=GEMINI_API_KEY)

user_balances = {}
waiting_for_receipt = set()
waiting_for_pdf_photos = set()
user_pdf_images = {}

waiting_for_zip_files = set()
user_zip_files = {}

waiting_for_conv_file = set()
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

    if message.from_user.id not in user_balances:
        user_balances[message.from_user.id] = 3000

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
    waiting_for_conv_file.discard(message.from_user.id)
    waiting_for_edit_file.discard(message.from_user.id)
    await message.answer("❌ Amaliyot bekor qilindi.", reply_markup=get_main_keyboard())

@dp.message(Command("balance"))
async def cmd_balance(message: types.Message):
    bal = user_balances.get(message.from_user.id, 3000)
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
    waiting_for_conv_file.discard(message.from_user.id)
    waiting_for_edit_file.discard(message.from_user.id)
    await message.answer("🏠 Asosiy menyu:", reply_markup=get_main_keyboard())

# --- PDF YARATISH ---
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

# --- FAYLLARNI ZIP QILISH ---
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

# --- PPTX ➔ PDF KONVERTATSIYA ---
@dp.message(F.text == "pptx ➔ pdf")
async def btn_conv_start(message: types.Message):
    waiting_for_conv_file.add(message.from_user.id)
    keyboard = ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="🔙 Ortga qaytish")]],
        resize_keyboard=True
    )
    await message.answer("📤 Iltimos, PDF formatiga o'tkazish uchun **.pptx** (PowerPoint) faylini yuboring:", reply_markup=keyboard, parse_mode="Markdown")

# --- SLAYDNI TAHRIRLASH 🛠 ---
@dp.message(F.text == "Slaydni tahrirlash 🛠")
async def btn_edit_slide_start(message: types.Message):
    waiting_for_edit_file.add(message.from_user.id)
    keyboard = ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="🔙 Ortga qaytish")]],
        resize_keyboard=True
    )
    await message.answer("📎 Tahrirlash uchun **.pptx** faylni yuboring:", reply_markup=keyboard, parse_mode="Markdown")


# --- FAYLLAR VA RASMLARNI QABUL QILISH ---
@dp.message(F.photo | F.document)
async def handle_incoming_files(message: types.Message):
    user = message.from_user
    
    # 1. Chek yuborish
    if user.id in waiting_for_receipt:
        waiting_for_receipt.discard(user.id)
        admin_caption = f"📥 <b>Yangi to'lov cheki keldi!</b>\n👤 Foydalanuvchi: {html.escape(user.full_name)} (ID: <code>{user.id}</code>)"
        admin_keyboard = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="✅ Tasdiqlash (+5000 so'm)", callback_data=f"accept_{user.id}_5000")]
            ]
        )
        if message.photo:
            await bot.send_photo(ADMIN_ID, photo=message.photo[-1].file_id, caption=admin_caption, reply_markup=admin_keyboard, parse_mode="HTML")
        await message.answer("✅ Chek adminga yuborildi!", reply_markup=get_main_keyboard())
        return

    # 2. PDF uchun rasmlar
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

    # 3. ZIP uchun fayllar
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

    # 4. PPTX ➔ PDF Konvertatsiya
    if user.id in waiting_for_conv_file:
        if message.document:
            file_name = message.document.file_name
            if not file_name.endswith(('.pptx', '.ppt')):
                await message.answer("⚠️ Iltimos, faqat .pptx formatidagi taqdimot faylini yuboring.")
                return
                
            file = await bot.get_file(message.document.file_id)
            f_bytes = await bot.download_file(file.file_path)
            
            await message.answer("⏳ Fayl PDF formatiga o'tkazilmoqda...", parse_mode="Markdown")
            
            input_path = f"temp_{user.id}_{file_name}"
            with open(input_path, "wb") as f:
                f.write(f_bytes.read())

            try:
                output_filename = file_name.rsplit('.', 1)[0] + ".pdf"
                output_path = os.path.abspath(f"temp_{user.id}_{output_filename}")
                abs_input_path = os.path.abspath(input_path)

                powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
                presentation = powerpoint.Presentations.Open(abs_input_path, WithWindow=False)
                presentation.SaveAs(output_path, 32)
                presentation.Close()
                powerpoint.Quit()

                with open(output_path, "rb") as f_out:
                    pdf_data = f_out.read()
                
                input_file = BufferedInputFile(pdf_data, filename=output_filename)
                await message.answer_document(input_file, caption="✅ PPTX muvaffaqiyatli PDF ga o'tkazildi!", reply_markup=get_main_keyboard())

                if os.path.exists(output_path):
                    os.remove(output_path)

            except Exception as e:
                logging.error(f"Conversion error: {e}")
                await message.answer("❌ Konvertatsiya qilishda xatolik yuz berdi.", reply_markup=get_main_keyboard())
            
            if os.path.exists(input_path):
                os.remove(input_path)
                
        waiting_for_conv_file.discard(user.id)
        return

    # 5. Slaydni tahrirlash uchun faylni tahlil qilish (Qisqa URL orqali xatolikni oldini olish)
    if user.id in waiting_for_edit_file:
        if message.document:
            file_name = message.document.file_name
            if not file_name.endswith(('.pptx', '.ppt')):
                await message.answer("⚠️ Iltimos, faqat .pptx formatidagi taqdimot faylini yuboring.")
                return
            
            file = await bot.get_file(message.document.file_id)
            f_bytes = await bot.download_file(file.file_path)
            file_content = f_bytes.read()
            
            slides_data = []
            try:
                prs = Presentation(io.BytesIO(file_content))
                for idx, slide in enumerate(prs.slides):
                    title = f"Slayd {idx + 1}"
                    content = ""
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for p in shape.text_frame.paragraphs:
                                if p.text.strip():
                                    if not title or title.startswith("Slayd"):
                                        title = p.text.strip()
                                    else:
                                        content += p.text.strip() + "\n"
                    slides_data.append({"title": title, "content": content.strip()})
            except Exception as e:
                logging.error(f"PPTX read error: {e}")
                slides_data = [{"title": "Slayd 1", "content": "Ma'lumotni o'qishda xatolik."}]

            waiting_for_edit_file.discard(user.id)
            
            # URL uzayib ketib xatolik bermasligi uchun qisqa havola va ID ishlatamiz
            webapp_url = f"https://farruh707.github.io/ai-presentation-bot/editor.html?uid={user.id}&count={len(slides_data)}"

            response_text = (
                f"✅ **Fayl qabul qilindi!**\n\n"
                f"📄 **Fayl nomi:** {html.escape(file_name)}\n"
                f"🖼 **Slaydlar soni:** {len(slides_data)}\n\n"
                f"Tahrirlash uchun quyidagi tugmani bosing:"
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
    user_balances[target_user_id] = user_balances.get(target_user_id, 0) + add_amount
    await bot.send_message(target_user_id, f"🎉 To'lovingiz tasdiqlandi! Balansingizga +{add_amount} so'm qo'shildi.")
    await callback.message.edit_caption(caption=callback.message.caption + "\n\n✅ Tasdiqlandi")
    await callback.answer("Bajarildi!")

@dp.message(F.web_app_data)
async def handle_web_app_data(message: types.Message):
    raw_data = message.web_app_data.data
    
    try:
        data = json.loads(raw_data)
    except Exception:
        data = {"mavzu": raw_data}

    if isinstance(data, dict) and data.get("action") == "edit_presentation":
        slides_data = data.get("slides", [])
        await message.answer("⏳ Tahrirlangan taqdimot tayyorlanmoqda...", parse_mode="Markdown")
        try:
            prs = Presentation()
            for s_info in slides_data:
                slide_layout = prs.slide_layouts[1]
                slide = prs.slides.add_slide(slide_layout)
                slide.shapes.title.text = s_info.get('title', 'Sarlavha')
                
                body_shape = slide.placeholders[1]
                tf = body_shape.text_frame
                tf.text = ""
                
                content = s_info.get('content', '')
                for line in content.split('\n'):
                    p = tf.add_paragraph()
                    p.text = line
                    p.level = 0
                    p.space_after = Pt(10)

            file_stream = io.BytesIO()
            prs.save(file_stream)
            file_stream.seek(0)

            input_file = BufferedInputFile(file_stream.getvalue(), filename="tahrirlangan_taqdimot.pptx")
            await message.answer(
                "✅ **Sizning tahrirlangan taqdimotingiz tayyor!** Marhamat, yuklab oling:",
                reply_markup=get_main_keyboard(),
                parse_mode="Markdown"
            )
            await message.answer_document(input_file)
            return
        except Exception as e:
            logging.error(f"Error generating edited presentation: {e}")
            await message.answer("❌ Taqdimotni saqlashda xatolik yuz berdi.", reply_markup=get_main_keyboard())
            return

    topic = data.get('mavzu') or data.get('topic') or data.get('title')
    
    if topic:
        num_slides = int(data.get('sahifalar', 7))
        muallif = data.get('muallif', 'Farrux')
        
        await message.answer(f"⏳ **Sun'iy intellekt '{topic}' mavzusida {num_slides} ta slayddan iborat taqdimot tayyorlamoqda...**", parse_mode="Markdown")
        
        try:
            prompt = (
                f"Siz professional taqdimot yaratuvchisiz. '{topic}' mavzusida o'zbek tilida {num_slides} ta slayddan iborat to'liq tarkib tuzing. "
                f"Siz aniq {num_slides} ta elementdan iborat JSON massivini qaytarishingiz SHART. "
                "Javobni FAQAT quyidagi JSON formatida yuboring, boshqa hech qanday matn yozmang:\n"
                "[\n"
                "  {\"title\": \"Slayd sarlavhasi 1\", \"points\": [\"Asosiy ma'lumot 1\", \"Asosiy ma'lumot 2\"]},\n"
                "  {\"title\": \"Slayd sarlavhasi 2\", \"points\": [\"Asosiy ma'lumot 1\", \"Asosiy ma'lumot 2\"]}\n"
                "]"
            )
            
            response = ai_client.models.generate_content(
                model='gemini-3.8-flash',
                contents=prompt,
            )
            
            clean_text = response.text.strip()
            if clean_text.startswith("```json"):
                clean_text = clean_text[7:-3].strip()
            elif clean_text.startswith("```"):
                clean_text = clean_text[3:-3].strip()
                
            slides_data = json.loads(clean_text)

            prs = Presentation()
            
            slide_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = topic
            slide.placeholders[1].text = f"Tayyorladi: {muallif}"

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
                    p.space_after = Pt(10)

            file_stream = io.BytesIO()
            prs.save(file_stream)
            file_stream.seek(0)

            filename = f"Slayd_{muallif}_{topic[:10]}.pptx".replace(" ", "_")
            input_file = BufferedInputFile(file_stream.getvalue(), filename=filename)
            
            await message.answer(
                f"✅ **Sizning taqdimotingiz tayyor!**\n\n📌 Mavzu: {topic}\n📄 Sahifalar soni: {len(slides_data) + 1}\n\nMarhamat, yuklab oling:",
                reply_markup=get_main_keyboard(),
                parse_mode="Markdown"
            )
            await message.answer_document(input_file)
            return

        except Exception as e:
            logging.error(f"Presentation generation error: {e}")
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = topic
            slide.placeholders[1].text = f"Muallif: {muallif}\n\nTaqdimot muvaffaqiyatli yaratildi."
            
            file_stream = io.BytesIO()
            prs.save(file_stream)
            file_stream.seek(0)
            await message.answer_document(BufferedInputFile(file_stream.getvalue(), filename="taqdimot.pptx"))
            return

    if 'vazirlik' in data or 'universitet' in data:
        try:
            doc = Document()
            for section in doc.sections:
                section.top_margin = Inches(1.0)
                section.bottom_margin = Inches(1.0)
                section.left_margin = Inches(1.5)
                section.right_margin = Inches(1.0)

            def add_center_p(text, size=14, bold=False):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(text)
                run.font.name = 'Times New Roman'
                run.font.size = Pt(size)
                run.bold = bold
                return p

            add_center_p(data.get('vazirlik', ''), size=12, bold=True)
            doc.add_paragraph()
            add_center_p(data.get('universitet', ''), size=14, bold=True)
            add_center_p(data.get('fakultet', ''), size=12)
            add_center_p(data.get('kafedra', ''), size=12)
            
            for _ in range(3): doc.add_paragraph()
            
            add_center_p(f"{data.get('fan', '')} fanidan", size=12)
            add_center_p(data.get('ishturi', 'DIPLOM ISHI'), size=16, bold=True)
            
            p_mavzu = doc.add_paragraph()
            p_mavzu.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_m = p_mavzu.add_run(f"Mavzu: {data.get('mavzu', '')}")
            run_m.font.name = 'Times New Roman'
            run_m.font.size = Pt(14)
            run_m.underline = True

            for _ in range(4): doc.add_paragraph()

            table = doc.add_table(rows=1, cols=2)
            table.autofit = False
            table.columns[0].width = Inches(3.0)
            table.columns[1].width = Inches(3.0)
            
            cell_left = table.cell(0, 0)
            cell_right = table.cell(0, 1)
            
            cell_left.text = f"Bajardi: {data.get('talaba', '')}\n({data.get('guruh', '')} guruh)"
            cell_right.text = f"Qabul qildi: {data.get('ustoz', '')}\n({data.get('lavozim', '')})"
            
            for p in cell_left.paragraphs:
                p.runs[0].font.name = 'Times New Roman'
                p.runs[0].font.size = Pt(12)
            for p in cell_right.paragraphs:
                p.runs[0].font.name = 'Times New Roman'
                p.runs[0].font.size = Pt(12)

            for _ in range(3): doc.add_paragraph()
            add_center_p(f"{data.get('shahar', 'Toshkent')} — {data.get('yil', '2026')}", size=14, bold=True)

            file_stream = io.BytesIO()
            doc.save(file_stream)
            file_stream.seek(0)

            await message.answer(
                "✅ **Sizning titul varag'ingiz tayyor!**",
                reply_markup=get_main_keyboard(),
                parse_mode="Markdown"
            )
            await message.answer_document(BufferedInputFile(file_stream.getvalue(), filename="titul_varagi.docx"))
        except Exception as e:
            logging.error(f"Error handling titul: {e}")
            await message.answer("❌ Xatolik yuz berdi.", reply_markup=get_main_keyword())

@dp.message(F.text)
async def handle_other_text(message: types.Message):
    await message.answer("❗ Iltimos, menyudan kerakli tugmani tanlang.", reply_markup=get_main_keyboard())

async def main():
    logging.basicConfig(level=logging.INFO)
    print("Bot muvaffaqiyatli ishga tushdi!")
    await set_bot_commands(bot)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
