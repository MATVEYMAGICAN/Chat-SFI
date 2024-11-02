import time
import telebot
from telebot import types
import openpyxl
import datetime

# Токен вашего Telegram-бота
TOKEN = '7908187199:AAE2SutLItRzkKsh7ujmaIUyUA4TfP6UitQ'
bot = telebot.TeleBot(TOKEN)

# Путь к файлам с расписанием
SCHEDULE_FILE_1 = 'Новая таблица.xlsx'
SCHEDULE_FILE_2 = 'Новая таблица2.xlsx'

# Загружаем расписание из Excel файла и находим все доступные классы
def load_schedule(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    schedule_data = {}
    available_classes = []

    # Предполагаем, что первая строка - это заголовок с номерами и буквами классов
    header = [cell.value for cell in sheet[1] if cell.value]
    available_classes = header[2:]  # Пропускаем первый столбец (дни недели)

    # Собираем данные расписания по каждому классу
    for class_index, class_name in enumerate(available_classes, start=3):
        class_schedule = {}
        current_day = None

        # Считываем строки, начиная со второй (где расписание)
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0]:  # Первая колонка (название дня недели)
                current_day = row[0]
                class_schedule[current_day] = []

            # Добавляем предметы в расписание на текущий день недели
            if current_day:
                lesson_content = row[class_index - 1] if row[class_index - 1] else ""
                class_schedule[current_day].append(lesson_content)

        schedule_data[class_name] = class_schedule

    return schedule_data, available_classes

# Загружаем расписание для 1 смены по умолчанию
schedule_data, available_classes = load_schedule(SCHEDULE_FILE_1)

# Функция для получения расписания на выбранный день
def get_schedule_for_day(class_name, day):
    # Получаем расписание для указанного дня
    class_schedule = schedule_data.get(class_name, {})
    day_schedule = class_schedule.get(day, ["Расписание для этого дня не найдено"])

    # Формируем текст с нумерацией уроков
    formatted_schedule = []
    for i, lesson in enumerate(day_schedule, start=1):
        if lesson:
            formatted_schedule.append(f"{i}. {lesson}")
        else:
            formatted_schedule.append(f"{i}. ")

    return "\n".join(formatted_schedule)

# Команда /start
@bot.message_handler(commands=['start'])
def start_message(message):
    # Меню выбора смены
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton(text="1 смена", callback_data="shift_1"))
    markup.add(types.InlineKeyboardButton(text="2 смена", callback_data="shift_2"))

    bot.send_message(
        message.chat.id,
        "Привет! Я бот для получения расписания в школе. Выберите смену:",
        reply_markup=markup
    )

# Обработчик выбора смены
@bot.callback_query_handler(func=lambda call: call.data.startswith("shift_"))
def select_shift(call):
    shift = call.data.split("_")[1]
    global schedule_data, available_classes

    if shift == "2":
        # Загрузка расписания для 2 смены
        schedule_data, available_classes = load_schedule(SCHEDULE_FILE_2)
    else:
        # Загрузка расписания для 1 смены
        schedule_data, available_classes = load_schedule(SCHEDULE_FILE_1)

    # Переход к выбору класса
    bot.delete_message(call.message.chat.id, call.message.message_id)  # Удаляем сообщение с выбором смены
    markup = types.InlineKeyboardMarkup()
    for grade in range(1, 12):
        markup.add(types.InlineKeyboardButton(text=f"{grade} класс", callback_data=f"grade_{grade}"))
    markup.add(types.InlineKeyboardButton(text="Вернуться в начало", callback_data="back_to_start"))
    bot.send_message(call.message.chat.id, "Выберите номер класса:", reply_markup=markup)

# Обработчик выбора номера класса
@bot.callback_query_handler(func=lambda call: call.data.startswith("grade_"))
def select_class(call):
    bot.delete_message(call.message.chat.id, call.message.message_id)  # Удаляем предыдущее сообщение
    grade = call.data.split("_")[1]  # Извлекаем номер класса
    # Отбираем только те классы, которые начинаются с выбранного номера
    classes_for_grade = [class_name for class_name in available_classes if class_name.startswith(grade)]

    # Создаем новую разметку с классами для выбранного номера и кнопкой "Назад"
    markup = types.InlineKeyboardMarkup()
    for class_name in classes_for_grade:
        markup.add(types.InlineKeyboardButton(text=class_name, callback_data=f"class_{class_name}"))
    markup.add(types.InlineKeyboardButton(text="Вернуться в начало", callback_data="back_to_start"))

    # Отправляем сообщение с выбором букв класса
    bot.send_message(
        call.message.chat.id,
        f"Вы выбрали {grade} класс. Теперь выберите букву:",
        reply_markup=markup
    )

# Обработчик выбора конкретного класса
@bot.callback_query_handler(func=lambda call: call.data.startswith("class_"))
def select_day(call):
    bot.delete_message(call.message.chat.id, call.message.message_id)  # Удаляем предыдущее сообщение
    class_name = call.data.split("_")[1]

    # Создаем клавиатуру с выбором дня недели
    markup = types.InlineKeyboardMarkup()
    days_of_week = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
    for day in days_of_week:
        markup.add(types.InlineKeyboardButton(text=day, callback_data=f"day_{day}_{class_name}"))
    markup.add(types.InlineKeyboardButton(text="Вернуться в начало", callback_data="back_to_start"))

    bot.send_message(
        call.message.chat.id,
        "Выберите день недели:",
        reply_markup=markup
    )

# Обработчик выбора дня недели для получения расписания
@bot.callback_query_handler(func=lambda call: call.data.startswith("day_"))
def send_schedule_for_day(call):
    bot.delete_message(call.message.chat.id, call.message.message_id)  # Удаляем предыдущее сообщение
    _, day, class_name = call.data.split("_")
    schedule = get_schedule_for_day(class_name, day)

    # Добавляем кнопку "Назад" для возврата к начальному меню
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton(text="Вернуться в начало", callback_data="back_to_start"))

    # Отправляем расписание на выбранный день
    bot.send_message(call.message.chat.id, f"Расписание для класса {class_name} на {day}:\n{schedule}",
                     reply_markup=markup)

# Обработчик кнопки "Назад" для возврата к начальному меню
@bot.callback_query_handler(func=lambda call: call.data == "back_to_start")
def back_to_start(call):
    bot.delete_message(call.message.chat.id, call.message.message_id)  # Удаляем предыдущее сообщение
    start_message(call.message)

# Запуск бота
bot.polling()
