import telebot
from telebot import types
import pandas as pd
import os
import re
import qrcode
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO

# ==========================
# CONFIGURATION
# ==========================
TOKEN = '7290992199:AAE_VYBPgfuht7wF2qpCuMXljdYJUSQ6Ppg'  # <-- Replace this with your actual token
bot = telebot.TeleBot(TOKEN)

excel_file_path = 'merged_results.xlsx'
df = pd.read_excel(excel_file_path)

# Folders to store QR codes
qr_diploma_folder = 'qr_diplomas'
qr_certificate_folder = 'qr_certificates'

# Fonts and sizes
font_path = 'OpenSans-SemiBold.ttf'

font_size_name = 100
font_size_subject = 70
font_size_teacher = 90
font_size_userid = 45

# Load fonts
font1 = ImageFont.truetype(font_path, font_size_name)
font2 = ImageFont.truetype(font_path, font_size_subject)
font3 = ImageFont.truetype(font_path, font_size_userid)
font4 = ImageFont.truetype(font_path, font_size_teacher)

qr_base_url = "https://urker.kz/megamozg2025-region-results/?region="

user_states = {}               # To store user's search state (name/email)
pending_teacher_update = {}    # To store pending teacher name updates by user chat_id

# ==========================
# TEMPLATES
# ==========================
diploma_templates = {
    "Алматинская_область‎": {
        "5-6": "templates/almobl_56.jpg",
        "7-8": "templates/almobl_78.jpg",
        "9-11": "templates/almobl_911.jpg"
    },
    "Алматы": {
        "5-6": "templates/almaty_56.jpg",
        "7-8": "templates/almaty_78.jpg",
        "9-11": "templates/almaty_911.jpg"
    },
    "Актюбинская_область‎": {
        "5-6": "templates/aktobe_56.jpg",
        "7-8": "templates/aktobe_78.jpg",
        "9-11": "templates/aktobe_911.jpg"
    },
    "default": {
        "5-6": "templates/default_56.jpg",
        "7-8": "templates/default_78.jpg",
        "9-11": "templates/default_911.jpg"
    }
}

certificate_templates = {
    "Алматинская_область‎": {
        "5-6": "templates/almobl_56c.jpg",
        "7-8": "templates/almobl_78c.jpg",
        "9-11": "templates/almobl_911c.jpg"
    },
    "Алматы": {
        "5-6": "templates/almaty_56c.jpg",
        "7-8": "templates/almaty_78c.jpg",
        "9-11": "templates/almaty_911c.jpg"
    },
    "Актюбинская_область‎": {
        "5-6": "templates/aktobe_56c.jpg",
        "7-8": "templates/aktobe_78c.jpg",
        "9-11": "templates/aktobe_911c.jpg"
    },
    "Астана": {
        "5-6": "templates/astana_56c.jpg",
        "7-8": "templates/astana_78c.jpg",
        "9-11": "templates/astana_911c.jpg"
    },
    "default": {
        "5-6": "templates/default_56c.jpg",
        "7-8": "templates/default_78c.jpg",
        "9-11": "templates/default_911c.jpg"
    }
}

# ==========================
# COORDINATES
# ==========================
region_coordinates = {
    "Алматинская_область‎": {
        "name_position": (900, 1265),
        "subject_position": (900, 1400),
        "teacher_position": (1250, 1530),
        "user_id_position": (815, 2167),
        "qr_position": (3250, 2200),
    },
    "Алматы": {
        "name_position": (900, 1265),
        "subject_position": (900, 1400),
        "teacher_position": (1250, 1530),
        "user_id_position": (870, 2167),
        "qr_position": (3250, 2200),
    },
    "Актюбинская_область‎": {
            "name_position": (920, 1265),
            "subject_position": (920, 1400),
            "teacher_position": (1250, 1530),
            "user_id_position": (870, 2168),
            "qr_position": (3270, 2220),
    },
    "Астана_5-6": {
        "name_position": (900, 1280),
        "subject_position": (1000, 1400),
        "teacher_position": (1130, 1510),
        "user_id_position": (925, 2200),
        "qr_position": (3300, 2250),
    },
    "Астана_7-8": {
        "name_position": (950, 1280),
        "subject_position": (1000, 1400),
        "teacher_position": (1130, 1520),
        "user_id_position": (925, 2200),
        "qr_position": (3300, 2250),
    },
    "Астана": {  # 9-11 grades default
        "name_position": (900, 1280),
        "subject_position": (1000, 1400),
        "teacher_position": (1130, 1520),
        "user_id_position": (925, 2200),
        "qr_position": (3300, 2250),
    },
    "default": {
        "name_position": (900, 1280),
        "subject_position": (900, 1400),
        "teacher_position": (1250, 1570),
        "user_id_position": (815, 2168),
        "qr_position": (3250, 2200),
    }
}
# ==========================
# HELPER FUNCTIONS
# ==========================
def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '', name.replace(' ', '_'))

def clean_region_name(region_name):
    return region_name.strip().replace(' ', '+')

def get_grade_group(class_num):
    if class_num in [5, 6]:
        return '5-6'
    elif class_num in [7, 8]:
        return '7-8'
    elif class_num in [9, 10, 11]:
        return '9-11'
    return None

def get_place(points):
    if points >= 18:
        return 1
    elif points >= 14:
        return 2
    elif points >= 10:
        return 3
    return None

def generate_qr(region, folder):
    os.makedirs(folder, exist_ok=True)
    region_url = clean_region_name(region)
    qr_link = f"{qr_base_url}{region_url}"
    qr_img = qrcode.make(qr_link)
    qr_path = os.path.join(folder, sanitize_filename(region) + ".png")
    qr_img.save(qr_path)
    return qr_path

def load_qr(region, folder):
    qr_path = os.path.join(folder, sanitize_filename(region) + ".png")
    if os.path.exists(qr_path):
        return Image.open(qr_path)
    return None

# ==========================
# GENERATION FUNCTIONS
# ==========================
def generate_diploma(record):
    region = record.get('Область', 'без региона').strip()
    class_num = int(record.get('класс'))
    points_str = str(record.get('Баллы', '0'))
    points = int(points_str.split()[0]) if points_str else 0

    grade_group = get_grade_group(class_num)
    coords = region_coordinates.get(region, region_coordinates["default"])
    qr_img = load_qr(region, qr_diploma_folder)

    if region == "Астана":
        place = get_place(points)
        if not place:
            print(f"No place for points {points}, no diploma for {record.get('ФИО')}")
            return None
        template_path = f"templates/astana_{grade_group}_{place}.jpg"
    else:
        template_path = diploma_templates.get(region, diploma_templates["default"]).get(grade_group)

    if not template_path or not os.path.exists(template_path):
        print(f"Template not found: {template_path}")
        return None

    img = Image.open(template_path)
    draw = ImageDraw.Draw(img)

    # Draw participant name
    name = str(record.get('ФИО', '')).strip()
    name_x = (img.width - draw.textbbox((0, 0), name, font=font1)[2]) // 2
    draw.text((name_x, coords['name_position'][1]), name, font=font1, fill='black')

    # Draw subject
    subject = str(record.get('предмет', '')).strip()
    subject_x = (img.width - draw.textbbox((0, 0), subject, font=font2)[2]) // 2
    draw.text((subject_x, coords['subject_position'][1]), subject, font=font2, fill='black')

    # Draw teacher name
    teacher = str(record.get('ФИО учителя', '')).strip()
    draw.text(coords['teacher_position'], teacher, font=font4, fill='black')


    user_id = str(record.get('User ID', '')).strip()
    draw.text(coords['user_id_position'], user_id, font=font3, fill='black')
    # Draw QR code if it exists
    if qr_img:
        qr_img = qr_img.resize((200, 200))
        img.paste(qr_img, coords['qr_position'])

    output = BytesIO()
    img.save(output, format='JPEG')
    output.seek(0)
    return output

# def generate_certificate(record):
#     region = record.get('Область', 'без региона').strip()
#     class_num = int(record.get('класс'))
#     grade_group = get_grade_group(class_num)

#     coords = coordinates.get(region, coordinates["default"])
#     qr_img = load_qr(region, qr_certificate_folder)

#     template_path = certificate_templates.get(region, certificate_templates["default"]).get(grade_group)

#     if not template_path or not os.path.exists(template_path):
#         print(f"Certificate template not found: {template_path}")
#         return None

#     img = Image.open(template_path)
#     draw = ImageDraw.Draw(img)

#     name = str(record.get('ФИО', '')).strip()
#     name_x = (img.width - draw.textbbox((0, 0), name, font=font1)[2]) // 2
#     draw.text((name_x, coords['name'][1]), name, font=font1, fill='black')

#     subject = str(record.get('предмет', '')).strip()
#     subject_x = (img.width - draw.textbbox((0, 0), subject, font=font2)[2]) // 2
#     draw.text((subject_x, coords['subject'][1]), subject, font=font2, fill='black')

#     teacher = str(record.get('ФИО учителя', '')).strip()
#     draw.text(coords['teacher'], teacher, font=font4, fill='black')

#     user_id = str(record.get('User ID', '')).strip()
#     draw.text(coords['user_id'], user_id, font=font3, fill='black')

#     if qr_img:
#         qr_img = qr_img.resize((200, 200))
#         img.paste(qr_img, coords['qr'])

#     output = BytesIO()
#     img.save(output, format='JPEG')
#     output.seek(0)
#     return output

# ==========================
# BOT HANDLERS
# ==========================
@bot.message_handler(commands=['start'])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add(types.KeyboardButton("Поиск по имени"), types.KeyboardButton("Поиск по email"))
    bot.send_message(message.chat.id, "Привет! Выбери способ поиска:", reply_markup=markup)

@bot.message_handler(func=lambda message: message.text in ["Поиск по имени", "Поиск по email"])
def choose_search_method(message):
    chat_id = message.chat.id
    user_states[chat_id] = "name" if message.text == "Поиск по имени" else "email"
    prompt = "Введите ФИО для поиска:" if user_states[chat_id] == "name" else "Введите email для поиска:"
    bot.send_message(chat_id, prompt)

@bot.message_handler(func=lambda message: True)
def search_and_send(message):
    chat_id = message.chat.id
    search_method = user_states.get(chat_id)

    if chat_id in pending_teacher_update:
        idx = pending_teacher_update.pop(chat_id)
        df.at[idx, 'ФИО учителя'] = message.text.strip()
        certificate = generate_certificate(df.loc[idx])
        if certificate:
            bot.send_photo(chat_id, certificate, caption=f"✅ Исправленный сертификат для {df.at[idx, 'ФИО']}")
            df.to_excel(excel_file_path, index=False)
        else:
            bot.send_message(chat_id, "❌ Не удалось создать исправленный сертификат.")
        return

    if not search_method:
        bot.send_message(chat_id, "Пожалуйста, выберите метод поиска сначала /start")
        return

    user_input = message.text.strip().lower()
    if search_method == "name":
        matches = df[df['ФИО'].fillna('').str.lower().str.contains(user_input)]
    else:
        matches = df[df['Email'].fillna('').str.lower().str.contains(user_input)]

    if matches.empty:
        bot.send_message(chat_id, "Ничего не найдено. Попробуйте снова! /start")
        return

    for idx, record in matches.iterrows():
        diploma = generate_diploma(record)
        # certificate = generate_certificate(record)

        if diploma:
            bot.send_photo(chat_id, diploma, caption=f"🎓 Диплом: {record['ФИО']}")
        else:
            bot.send_message(chat_id, f"❌ Диплом не найден для {record['ФИО']}")

        # if certificate:
        #     bot.send_photo(chat_id, certificate, caption=f"📄 Сертификат: {record['ФИО']}")
        # else:
        #     bot.send_message(chat_id, f"❌ Сертификат не найден для {record['ФИО']}")

        teacher_name = str(record.get('ФИО учителя', '')).strip()
        if teacher_name.lower() in ['', 'nan']:
            pending_teacher_update[chat_id] = idx
            bot.send_message(chat_id, f"Учитель не указан для {record['ФИО']}. Напишите имя учителя для исправления:")

    user_states.pop(chat_id, None)
    if chat_id not in pending_teacher_update:
        bot.send_message(chat_id, "✅ Поиск завершён! Напишите /start для нового поиска.")

# ==========================
# MAIN ENTRY POINT
# ==========================
def main():
    regions = df['Область'].fillna('без региона').unique()
    for region in regions:
        generate_qr(region, qr_diploma_folder)
        generate_qr(region, qr_certificate_folder)

    print("Бот запущен!")
    bot.polling(none_stop=True)

if __name__ == "__main__":
    main()