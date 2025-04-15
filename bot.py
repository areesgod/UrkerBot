import telebot
from telebot import types
import pandas as pd
import os
import re
import requests
import json
import qrcode
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import gspread
from google.oauth2 import service_account
# Define scope
dropbox_link = 'https://www.dropbox.com/scl/fi/avc052ue08x6gn20g0c6x/raribxy-bb9d464839a3.json?rlkey=av8ohjqtaxhqbp3tsg00o1wsk&st=x45ajabc&dl=1'
response = requests.get(dropbox_link)
if response.status_code != 200:
    raise Exception(f"Failed to download file: {response.status_code}")
service_account_info = json.loads(response.content)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = service_account.Credentials.from_service_account_info(service_account_info, scopes=scope)
client = gspread.authorize(creds)





# Open your spreadsheet and worksheet
GOOGLE_SHEET_NAME = 'UrkerData'  # Spreadsheet name
worksheet_name = 'Sheet1'  # Worksheet name inside spreadsheet
sheet = client.open(GOOGLE_SHEET_NAME).worksheet(worksheet_name)
# Load initial data into DataFrame
data = sheet.get_all_records()
df = pd.DataFrame(data)
# ==========================
# CONFIGURATION
# ==========================
TOKEN = '7607579283:AAG3PIj74AY7j4PBWzpy-ts16H_3sbL3Hlw'  # <-- Replace this with your actual token
bot = telebot.TeleBot(TOKEN)



# Folders to store QR codes
qr_diploma_folder = 'qr_diplomas'
qr_certificate_folder = 'qr_certificates'

# Fonts and sizes
font_path = 'OpenSans-SemiBold.ttf'

font_size_name = 100
font_size_subject = 70
font_size_teacher = 90
font_size_userid = 45
font_place = 150
# Load fonts
font1 = ImageFont.truetype(font_path, font_size_name)
font2 = ImageFont.truetype(font_path, font_size_subject)
font3 = ImageFont.truetype(font_path, font_size_userid)
font4 = ImageFont.truetype(font_path, font_size_teacher)
font5 = ImageFont.truetype(font_path,font_place)
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
coordinates = {
    "Алматинская_область‎": {
        "name": (900, 1265),
        "subject": (900, 1400),
        "teacher": (1250, 1530),
        "user_id": (815, 2167),
        "qr": (3250, 2200),
    },
    "Алматы": {
        "name": (900, 1265),
        "subject": (900, 1400),
        "teacher": (1250, 1530),
        "user_id": (870, 2167),
        "qr": (3250, 2200),
    },
    "Актюбинская_область‎": {
        "name": (920, 1265),
        "subject": (920, 1400),
        "teacher": (1250, 1530),
        "user_id": (870, 2168),
        "qr": (3270, 2220),
    },
    "Астана": {
        "name": (900, 1280),
        "subject": (1000, 1400),
        "teacher": (1130, 1520),
        "user_id": (925, 2200),
        "qr": (3300, 2250),
    },
    "default": {
        "name": (900, 1280),
        "subject": (900, 1400),
        "teacher": (1250, 1530),
        "user_id": (815, 2168),
        "qr": (3250, 2200),
    }
}

region_coordinates = {
    "Алматинская_область‎": {
        "name_position": (900, 1265),
        "subject_position": (900, 1400),
        "teacher_position": (1250, 1570),
        "user_id_position": (820, 2168),
        "qr_position": (3250, 2200),
        "place_position": (1400, 650)  # New place text position
    },
    "Алматы": {
        "name_position": (900, 1265),
        "subject_position": (900, 1400),
        "teacher_position": (1250, 1570),
        "user_id_position": (870, 2167),
        "qr_position": (3250, 2200),
        "place_position": (1400, 650)
    },
    "Актюбинская_область‎": {
        "name_position": (920, 1265),
        "subject_position": (920, 1400),
        "teacher_position": (1250, 1570),
        "user_id_position": (870, 2167),
        "qr_position": (3270, 2220),
        "place_position": (1400, 650)
    },
    "Астана_5-6": {
        "name_position": (900, 1200),
        "subject_position": (1000, 1340),
        "teacher_position": (1130, 1550),
        "user_id_position": (925, 2200),
        "qr_position": (3300, 2250)
    },
    "Астана_7-8": {
        "name_position": (950, 1200),
        "subject_position": (1000, 1340),
        "teacher_position": (1130, 1550),
        "user_id_position": (925, 2200),
        "qr_position": (3300, 2300)
    },
    "Астана": {  # 9-11 grades default
        "name_position": (900, 1150),
        "subject_position": (1000, 1290),
        "teacher_position": (1130, 1550),
        "user_id_position": (925, 2200),
        "qr_position": (3300, 2250)
    },
    "default": {
        "name_position": (900, 1265),
        "subject_position": (900, 1400),
        "teacher_position": (1250, 1570),
        "user_id_position": (820, 2168),
        "qr_position": (3250, 2200),
        "place_position": (1400, 650) 
    }
}

def refresh_sheet_data():
    global df  # Make sure it updates the global df
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    print("🔄 Данные Google Sheet обновлены!")


def rename_subject(subject_name):
    if not subject_name:
        return ''
    
    subject_renaming_rules = [
        (['кял'], 'қазақ тілі мен әдебиеті'),
        (['математика', 'math'], 'математика'),
        (['информатика'], 'информатика'),
        (['физика', 'physics'], 'физика'),
        (['английский', 'англ'], 'ағылшын тілі'),
        (['немецкий'], 'неміс тілі'),
        (['ик','история'], 'Қазақстан тарихы'),
        (['химия'], 'химия'),
        (['биология'], 'биология'),
        (['ркш', 'русяз'], 'қазақ сыныптарындағы орыс тілі'),
        (['крш'], 'орыс сыныптарындағы қазақ тілі'),
        (['музыка'], 'музыка'),
        (['право'], 'құқық негіздері'),
        (['естествознание', 'жаратылыстану'], 'жаратылыстану'),
        (['худ'], 'көркем еңбек'),
        (['физра', 'дш'], 'дене шынықтыру'),
        (['рял'], 'орыс тілі мен әдебиеті'),
        (['география'], 'география'),
    ]

    for keywords, new_name in subject_renaming_rules:
        if any(re.search(keyword, subject_name, re.IGNORECASE) for keyword in keywords):
            return new_name
    
    # Return original subject if no match
    return subject_name

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
def place_to_roman(place):
    roman_places = {
        1: 'I',
        2: 'II',
        3: 'III'
    }
    return roman_places.get(place, '')

# ==========================
# GENERATION FUNCTIONS
# ==========================
def generate_diploma(record):
    region = record.get('Область', 'без региона').strip()
    class_num = int(record.get('класс'))
    points_str = str(record.get('Баллы', '0'))
    points = int(points_str.split()[0]) if points_str else 0

    grade_group = get_grade_group(class_num)

    # Determine region key for coordinates
    if region == "Астана":
        if grade_group == "9-11":
            region_key = "Астана"
        else:
            region_key = f"Астана_{grade_group}"
    else:
        region_key = region

    coords = region_coordinates.get(region_key, region_coordinates["default"])
    qr_img = load_qr(region, qr_diploma_folder)

    # Template selection and place determination
    if region == "Астана":
        place = get_place(points)
        if not place:
            print(f"No place for points {points}, no diploma for {record.get('ФИО')}")
            return None
        template_path = f"templates/astana_{grade_group}_{place}.jpg"
    else:
        # For non-Astana, you can decide whether place is needed or not.
        place = get_place(points)
        if not place:
            print(f"No place for points {points}, no diploma for {record.get('ФИО')}")
            return None
        template_path = diploma_templates.get(region, diploma_templates["default"]).get(grade_group)

    if not template_path or not os.path.exists(template_path):
        print(f"Template not found: {template_path}")
        return None

    # Open and draw on image
    img = Image.open(template_path)
    draw = ImageDraw.Draw(img)

    # Draw participant name
    name = str(record.get('ФИО', '')).strip()
    name_x = (img.width - draw.textbbox((0, 0), name, font=font1)[2]) // 2
    draw.text((name_x, coords['name_position'][1]), name, font=font1, fill='black')

    # Draw renamed subject
    raw_subject = str(record.get('предмет', '')).strip()
    subject = rename_subject(raw_subject)
    subject_x = (img.width - draw.textbbox((0, 0), subject, font=font2)[2]) // 2
    draw.text((subject_x, coords['subject_position'][1]), subject, font=font2, fill='black')

    # Draw teacher name
    teacher = str(record.get('ФИО учителя', '')).strip()
    draw.text(coords['teacher_position'], teacher, font=font4, fill='black')

    # Draw user ID
    user_id = str(record.get('User ID', '')).strip()
    draw.text(coords['user_id_position'], user_id, font=font3, fill='black')

    # Draw QR code if it exists
    if qr_img:
        qr_img = qr_img.resize((200, 200))
        img.paste(qr_img, coords['qr_position'])

    # ✅ Draw the place (I, II, III)
    if region!="Астана":
        roman_place = place_to_roman(place)
        if roman_place and 'place_position' in coords:
            draw.text(coords['place_position'], roman_place, font=font5, fill='black')


    # Save to BytesIO and return
    output = BytesIO()
    img.save(output, format='JPEG')
    output.seek(0)
    return output



def generate_certificate(record):
    region = record.get('Область', 'без региона').strip()
    class_num = int(record.get('класс'))
    grade_group = get_grade_group(class_num)

    coords = coordinates.get(region, coordinates["default"])
    qr_img = load_qr(region, qr_certificate_folder)

    template_path = certificate_templates.get(region, certificate_templates["default"]).get(grade_group)

    if not template_path or not os.path.exists(template_path):
        print(f"Certificate template not found: {template_path}")
        return None

    img = Image.open(template_path)
    draw = ImageDraw.Draw(img)

    name = str(record.get('ФИО', '')).strip()
    name_x = (img.width - draw.textbbox((0, 0), name, font=font1)[2]) // 2
    draw.text((name_x, coords['name'][1]), name, font=font1, fill='black')

    # Rename subject before drawing it
    raw_subject = str(record.get('предмет', '')).strip()
    subject = rename_subject(raw_subject)
    subject_x = (img.width - draw.textbbox((0, 0), subject, font=font2)[2]) // 2
    draw.text((subject_x, coords['subject'][1]), subject, font=font2, fill='black')

    teacher = str(record.get('ФИО учителя', '')).strip()
    draw.text(coords['teacher'], teacher, font=font4, fill='black')

    user_id = str(record.get('User ID', '')).strip()
    draw.text(coords['user_id'], user_id, font=font3, fill='black')

    if qr_img:
        qr_img = qr_img.resize((200, 200))
        img.paste(qr_img, coords['qr'])

    output = BytesIO()
    img.save(output, format='JPEG')
    output.seek(0)
    return output


# ==========================
# BOT HANDLERS
# ==========================
# ==========================
# BOT HANDLERS
# ==========================
# === User States ===
user_states = {}
pending_teacher_update = {}
pending_name_update = {}
pending_region_update = {}
pending_region_selection = {}
pending_selection = {}

# ========== /START HANDLER ==========
@bot.message_handler(commands=['start'])
def start(message):
    chat_id = message.chat.id
    reset_user(chat_id)

    refresh_sheet_data()  # 🔄 Refresh here

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add(types.KeyboardButton("Поиск по имени"), types.KeyboardButton("Поиск по email"))

    bot.send_message(chat_id, "Привет! Выбери способ поиска:", reply_markup=markup)

    

# ========== SEARCH METHOD SELECTION ==========
@bot.message_handler(func=lambda message: message.text in ["Поиск по имени", "Поиск по email"])
def choose_search_method(message):
    chat_id = message.chat.id
    method = "name" if message.text == "Поиск по имени" else "email"

    user_states[chat_id] = {
        "method": method,
        "step": "waiting_for_query"
    }

    prompt = "Введите ФИО для поиска:" if method == "name" else "Введите email для поиска:"
    bot.send_message(chat_id, prompt)

# ========== GENERAL MESSAGE HANDLER ==========
@bot.message_handler(func=lambda message: True)
def handle_message(message):
    chat_id = message.chat.id
    text = message.text.strip()

    # --- Region Number Selection ---
    if chat_id in pending_region_selection:
        idx, region_list = pending_region_selection.pop(chat_id)

        if not text.isdigit():
            regions_str = "\n".join([f"{i + 1}. {region}" for i, region in enumerate(region_list)])
            bot.send_message(chat_id, f"❌ Введите номер региона из списка:\n{regions_str}")
            pending_region_selection[chat_id] = (idx, region_list)
            return

        region_index = int(text) - 1
        if region_index < 0 or region_index >= len(region_list):
            regions_str = "\n".join([f"{i + 1}. {region}" for i, region in enumerate(region_list)])
            bot.send_message(chat_id, f"❌ Некорректный номер региона.\nВыберите из списка:\n{regions_str}")
            pending_region_selection[chat_id] = (idx, region_list)
            return

        selected_region = region_list[region_index]
        df.at[idx, 'Область'] = selected_region
        save_to_google_sheet()

        record = df.loc[idx]

        if check_missing_fields(chat_id, idx, record):
            return

        send_diploma_and_certificate(chat_id, record)
        bot.send_message(chat_id, "✅ Готово! Введите /start для нового поиска.")
        reset_user(chat_id)
        return

    # --- Teacher Update ---
    if chat_id in pending_teacher_update:
        idx = pending_teacher_update.pop(chat_id)

        if idx not in df.index:
            bot.send_message(chat_id, "❌ Ошибка: Запись не найдена.")
            return

        df.at[idx, 'ФИО учителя'] = text
        save_to_google_sheet()

        record = df.loc[idx]

        if check_missing_fields(chat_id, idx, record):
            return

        send_diploma_and_certificate(chat_id, record)
        bot.send_message(chat_id, "✅ Готово! Введите /start для нового поиска.")
        reset_user(chat_id)
        return

    # --- Name Update ---
    if chat_id in pending_name_update:
        idx = pending_name_update.pop(chat_id)

        if idx not in df.index:
            bot.send_message(chat_id, "❌ Ошибка: Запись не найдена.")
            return

        df.at[idx, 'ФИО'] = text
        save_to_google_sheet()

        record = df.loc[idx]

        if check_missing_fields(chat_id, idx, record):
            return

        send_diploma_and_certificate(chat_id, record)
        bot.send_message(chat_id, "✅ Готово! Введите /start для нового поиска.")
        reset_user(chat_id)
        return

    # --- Regular flow ---
    state = user_states.get(chat_id)
    if not state:
        bot.send_message(chat_id, "Пожалуйста, выберите метод поиска сначала /start")
        return

    # --- Waiting for query ---
    if state["step"] == "waiting_for_query":
        user_input = text.lower()

        if state["method"] == "name":
            matches_df = df[df['ФИО'].str.lower().str.contains(user_input, na=False)]
        else:
            matches_df = df[df['Email'].fillna('').str.lower().str.contains(user_input)]

        if matches_df.empty:
            bot.send_message(chat_id, "❌ Ничего не найдено. Попробуйте снова! /start")
            return

        matches = [(idx, record) for idx, record in matches_df.iterrows()]
        pending_selection[chat_id] = {"matches": matches}

        # Build message list
        result_msg = "🔍 Найдено записей:\n\n"
        for i, (idx, record) in enumerate(matches, 1):
            subject = rename_subject(record.get('предмет', 'Без предмета'))
            grade = record.get('класс', 'no grade')
            result_msg += f"{i}. {record['ФИО']} - {subject} - {grade}\n"

        result_msg += "\nВведите номер записи, чтобы получить диплом и сертификат."

        bot.send_message(chat_id, result_msg)

        state["step"] = "waiting_for_selection"
        return

    # --- Waiting for selection ---
    if state["step"] == "waiting_for_selection":
        if not text.isdigit():
            bot.send_message(chat_id, "❌ Введите номер из списка!")
            return

        selection_index = int(text) - 1
        matches_info = pending_selection.get(chat_id)

        if not matches_info or selection_index < 0 or selection_index >= len(matches_info["matches"]):
            bot.send_message(chat_id, "❌ Некорректный номер. Попробуйте снова.")
            return

        idx, record = matches_info["matches"][selection_index]

        if check_missing_fields(chat_id, idx, record):
            return

        send_diploma_and_certificate(chat_id, record)
        bot.send_message(chat_id, "✅ Готово! Введите /start для нового поиска.")
        reset_user(chat_id)
        return

    # Fallback
    bot.send_message(chat_id, "❌ Неверная команда. Введите /start чтобы начать сначала.")

# ========== HELPER FUNCTIONS ==========

def check_missing_fields(chat_id, idx, record):
    """Check and prompt for missing fields. Returns True if waiting for input."""

    # Check name
    name = str(record.get('ФИО', '')).strip()
    if not name or name.lower() == 'nan':
        pending_name_update[chat_id] = idx
        bot.send_message(chat_id, f"❗ ФИО отсутствует. Введите ФИО участника:")
        return True

    # Check region
    region = str(record.get('Область', '')).strip()
    if not region or region.lower() == 'nan':
        region_list = df['Область'].dropna().unique().tolist()
        if not region_list:
            bot.send_message(chat_id, "❌ Нет доступных регионов в данных.")
            return True

        regions_str = "\n".join([f"{i + 1}. {region}" for i, region in enumerate(region_list)])
        bot.send_message(chat_id, f"❗ Регион отсутствует. Выберите номер из списка:\n{regions_str}")

        pending_region_selection[chat_id] = (idx, region_list)
        return True

    # Check teacher
    teacher = str(record.get('ФИО учителя', '')).strip()
    if not teacher or teacher.lower() == 'nan':
        pending_teacher_update[chat_id] = idx
        bot.send_message(chat_id, f"❗ Учитель отсутствует. Введите ФИО учителя:")
        return True

    return False

def send_diploma_and_certificate(chat_id, record):
    """Generate and send diploma and certificate"""
    diploma = generate_diploma(record)
    certificate = generate_certificate(record)

    if diploma:
        bot.send_photo(chat_id, diploma, caption=f"🎓 Диплом участника: {record['ФИО']}")
    else:
        bot.send_message(chat_id, f"❌ Диплом не найден для {record['ФИО']}")

    if certificate:
        bot.send_photo(chat_id, certificate, caption=f"📄 Сертификат участника: {record['ФИО']}")
    else:
        bot.send_message(chat_id, f"❌ Сертификат не найден для {record['ФИО']}")

def reset_user(chat_id):
    """Reset user states and pending actions"""
    user_states.pop(chat_id, None)
    pending_teacher_update.pop(chat_id, None)
    pending_name_update.pop(chat_id, None)
    pending_region_update.pop(chat_id, None)
    pending_region_selection.pop(chat_id, None)
    pending_selection.pop(chat_id, None)

def save_to_google_sheet():
    # Clean up NaN values
    clean_df = df.fillna('')

    # Convert to list of lists (header + data)
    data = [clean_df.columns.tolist()] + clean_df.values.tolist()

    # Push to Google Sheet
    sheet.update(data)

    print("Google Sheet successfully updated!")


# ========== MAIN ==========
def main():
    # Optional pre-generation of QR codes
    regions = df['Область'].fillna('без региона').unique()
    for region in regions:
        generate_qr(region, 'qr_diploma_folder')
        generate_qr(region, 'qr_certificate_folder')

    print("✅ Бот запущен и готов к работе!")
    bot.polling(none_stop=True, interval=0, timeout=20)

if __name__ == "__main__":
    main()