import telebot
from telebot import types
import os
import threading
from dotenv import load_dotenv
load_dotenv()

TOKEN = str(os.getenv("BOT_TOKEN"))

bot = telebot.TeleBot(TOKEN)

# Хранение ответов пользователя
user_data = {}


@bot.message_handler(commands=['start'])
def send_welcome(message):
    video_path = 'media/video.mp4'  # Локальный путь к видео

    with open(video_path, 'rb') as video_note:
        bot.send_video_note(message.chat.id, video_note)

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add(types.KeyboardButton("Пройти тест ❤️"))

    # Затем отправляем текстовое приветствие с именем
    name = message.from_user.first_name
    greeting_text = f"""Привет, {name}! 👋🏻 Очень рада видеть тебя здесь!

Здесь всё по-настоящему просто, удобно и с заботой о тебе.

<b>Я помогу тебе мягко и уверенно двигаться к твоим целям, какие бы они ни были</b>:
🌸 снять отёки и вернуть ощущение лёгкости
🌸 выглядеть на 15 лет моложе без ботокса и хирургии
🌸 подтянуть живот и все тело
🌸 укрепить мышцы тазового дна
🌸 подтянуть овал лица и убрать второй подбородок
🌸 почувствовать больше энергии

<b>Давай определим, какая тренировка нужна твоему телу прямо сейчас?</b>"""

    bot.send_message(message.chat.id, greeting_text, reply_markup=markup, parse_mode='HTML')


@bot.message_handler(func=lambda message: message.text == "Пройти тест ❤️")
def start_test(message):
    # Приветствие с именем
    ask_age(message.chat.id)

def ask_age(chat_id):
    markup = types.ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True)
    markup.add('До 30', '30-40', '40-50', '50+')
    bot.send_message(chat_id, "1/4 Сколько тебе лет?", reply_markup=markup)


@bot.message_handler(func=lambda message: message.text in ['До 30', '30-40', '40-50', '50+'])
def handle_age(message):
    user_data[message.chat.id] = {'age': message.text}
    ask_motivation(message.chat.id)


def ask_motivation(chat_id):
    markup = types.ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True)
    markup.add('Да, постоянно', 'Даже не начинаю, вечно что-то останавливает')
    bot.send_message(chat_id,
                     "2/4 Бывает ли у тебя такое, что ты обещаешь себе начать заниматься телом и внешним видом «с понедельника», но мотивации хватает ненадолго?",
                     reply_markup=markup)


@bot.message_handler(
    func=lambda message: message.text in ['Да, постоянно', 'Даже не начинаю, вечно что-то останавливает'])
def handle_motivation(message):
    user_data[message.chat.id]['motivation'] = message.text
    ask_concern(message.chat.id)


def ask_concern(chat_id):
    markup = types.ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True)
    markup.add('Сутулость, некрасивая осанка и постоянно болит спина', 'Второй подбородок, брыли')
    bot.send_message(chat_id, "3/4 Что тебя больше всего волнует?", reply_markup=markup)


@bot.message_handler(func=lambda message: message.text in ['Сутулость, некрасивая осанка и постоянно болит спина',
                                                           'Второй подбородок, брыли'])
def handle_concern(message):
    user_data[message.chat.id]['concern'] = message.text
    ask_time(message.chat.id)


def ask_time(chat_id):
    markup = types.ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True)
    markup.add('10-15 минут', '15-30 минут')
    bot.send_message(chat_id, "4/4 Сколько у тебя есть в день времени на себя?", reply_markup=markup)


@bot.message_handler(func=lambda message: message.text in ['10-15 минут', '15-30 минут'])
def handle_time(message):
    user_data[message.chat.id]['time'] = message.text

    # Отправляем благодарность
    thanks_text = """Спасибо, что поделилась.. Не поверишь, что я тоже далеко не всегда такой молодой и стройной:
💫 с кучей энергии и легкостью в теле
💫 совершенно без отёков и зажимов
💫в подтянутом теле без выпирающего живота, со стройной талией
💫 с подтянутым лицом без морщин и брылей

И прийти к такому реально любой:
— дело НЕ в возрасте или генетике,
— не нужно сидеть на диетах, ограничивать себя и убиваться на тренировках

<b>Я определилась, какая тренировка подойдет тебе, но сначала подпишись на канал https://t.me/elan_beauty </b> 👇🏻"""

    markup = types.InlineKeyboardMarkup()
    btn_subscribe = types.InlineKeyboardButton(
        text="Подписаться на канал",
        url="https://t.me/elan_beauty"
    )
    btn_check = types.InlineKeyboardButton(
        text="Проверить подписку",
        callback_data="check_subscription"
    )
    markup.add(btn_subscribe, btn_check)

    bot.send_message(message.chat.id, thanks_text, reply_markup=markup, parse_mode='HTML')


@bot.callback_query_handler(func=lambda call: call.data == "check_subscription")
def check_subscription_callback(call):
    try:
        chat_member = bot.get_chat_member(chat_id="@elan_beauty", user_id=call.from_user.id)
        if chat_member.status in ['member', 'administrator', 'creator']:
            # Пользователь подписан
            bot.answer_callback_query(call.id, "Спасибо за подписку! Сейчас подберу для тебя тренировку...")
            send_training(call.message.chat.id)
        else:
            # Пользователь не подписан
            bot.answer_callback_query(call.id, "❌ Кажется, ты еще не подписалась. Пожалуйста, подпишись и нажми снова", show_alert=True)
    except Exception as e:
        print(f"Ошибка проверки подписки: {e}")
        bot.answer_callback_query(call.id, "Произошла ошибка при проверке. Попробуй еще раз позже", show_alert=True)


def send_training(chat_id):
    user = user_data.get(chat_id, {})
    concern = user.get('concern', '')

    markup = types.InlineKeyboardMarkup()

    if concern == 'Сутулость, некрасивая осанка и постоянно болит спина':
        training_text = """<b>Лови тренировку «здоровая спина», сразу после которой ты почувствуешь</b>: 
•легкость
•осанка выпрямиться
•больше свободы в движениях и прилив энергии
•через несколько дней регулярных тренировок — минус лишние сантиметры в талии и более подтянутый живот, королевская осанка 

Всего 11 минут — и ты увидишь, как твое тело изменилось! Начнём? 🚀"""
        button = types.InlineKeyboardButton(text="Смотреть на YouTube",
                                            url="https://youtu.be/ev23H3TBDp0?si=GylBJr1rf-stdYmP")
    else:
        training_text = """<b>Лови тренировку «убираем гипертонус жевательной и височной мышцы», сразу после которой ты почувствуешь</b>: 
•расслабление мышц лица, ты почувствуешь «свободу» лица в прямом смысле 
•овал лица подтянется 
•заметно уменьшаться морщины и брыли

Всего 2 минуты — и ты увидишь, как твое лицо изменилось, овал подтянулся и вообще ты помолодела! Начнём? 🚀"""
        button = types.InlineKeyboardButton(text="Смотреть на YouTube",
                                            url="https://youtu.be/VeRg8-BpfOI?si=SLd0a-pqB5FEgWLl")

    restart_button = types.InlineKeyboardButton(text="🔄 Перезапустить тест", callback_data="restart_test")
    markup.add(restart_button)
    markup.add(button)
    bot.send_message(chat_id, training_text, reply_markup=markup, parse_mode='HTML')

    # Через некоторое время спрашиваем, сделала ли комплекс
    timer = threading.Timer(60*60, ask_if_done, args=[chat_id])  # Через 1 час
    timer.start()

@bot.callback_query_handler(func=lambda call: call.data == "restart_test")
def restart_test(call):
    bot.answer_callback_query(call.id, "Начинаем тест заново!")
    ask_age(call.message.chat.id)

def ask_if_done(chat_id):
    name = bot.get_chat(chat_id).first_name
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add(types.KeyboardButton("Конечно! 😊"))
    bot.send_message(chat_id, f"{name}, подскажи, сделала комплекс?", reply_markup=markup)


@bot.message_handler(func=lambda message: message.text == "Конечно! 😊")
def handle_done_confirmation(message):
    follow_up_text = """Уже почувствовала изменения после тренировки\?

Теперь представь, что за месяц работы на осанку и лицо ты сможешь выглядеть на 10 лет моложе\:

🌸 убираешь отёки, а вместе с ними — тяжесть и лишние объёмы в ногах, талии и животе
🌸 смотришь в зеркало и видишь, как плечи расправились как у модели или балерины
🌸 ты больше не сутулишься, а походка стала лёгкой и женственной
🌸 твое лицо стало подтянутым, морщинки исчезли и ты стала выглядеть моложе
🌸 тебе делают комплименты и не верят, когда ты называешь свой возраст

Это не сказка — это эффект системы, которая работает в клубе Elan beauty 

🔥3 онлайн в неделю тренировки по 30 минут 
🔥Занимайся LIVE с нами или в записи — когда удобно\:  
\- Утром с кофе   
\- В обеденный перерыв   
\- Вечером вместо сериала  
🔥 Ежедневная 7 минутная тренировка для шеи
🔥 Тренировки по 15 минут в день для идеальной осанки
🔥 2 раза в неделю пилатес на мяче 
🔥 Дыхательные практики для расслабления плеч
🔥 Уроки по тейпированию лица и тела, чтобы омолаживаться во сне

📌 Доступ ВСЕГО за 2599 вместо ~5599~"""

    # Убираем клавиатуру после нажатия
    markup = types.InlineKeyboardMarkup()
    btn_subscribe = types.InlineKeyboardButton(
        text="Хочу больше!",
        url="https://elviraelan.taplink.ws"
    )
    markup.add(btn_subscribe)

    bot.send_message(
        message.chat.id,
        follow_up_text,
        reply_markup=markup,  # инлайн-кнопка
        parse_mode='MarkdownV2'
    )

    threading.Timer(60*60*24, send_day_after_message, args=[message.chat.id]).start()


def send_day_after_message(chat_id):
    name = bot.get_chat(chat_id).first_name

    # Первое сообщение с фото и видео
    success_story = f"""Ира 39 лет 

Пришла в клуб:
❌Лишний вес 
❌Отечность
❌Двойной подбородок 
❌Боль в шее
❌Целлюлит 
❌Сутулость 

После: 
✅Убрали отеки 
✅Лишний вес пошел вниз 
✅Улучшилась осанка
✅Ушла отечность лица 
✅Раскрыла себя как женщина 
✅Научилась нравится себе"""

    media = [
        types.InputMediaVideo('BAACAgIAAxkDAAOLaEmZ12Y-n6di3SZCnm_hM2yH4fgAAn5yAALgMVBK6JHVUCwaVD02BA',
                              caption=success_story),
        types.InputMediaPhoto('AgACAgIAAxkDAAONaEmaWvnG0QP0CmvucYALAy0leDsAAgT5MRun7FBKsX9XfNJoAAHeAQADAgADeQADNgQ')
    ]

    # Отправляем медиагруппу
    bot.send_media_group(chat_id, media)

    # Через 15 минут отправляем второе сообщение
    threading.Timer(15*60, send_follow_up_message, args=[chat_id]).start()


def send_follow_up_message(chat_id):
    name = bot.get_chat(chat_id).first_name

    follow_up_text = f"""{name}, ты так помолодела и постройнела 🔥

Хочешь получать такие же комплименты?

Это секрет работы с осанкой и отеками, с мышцами лица
а не диет и жестких тренировок, походов к косметологу 

У девушек подтянулся овал лица, вес пошел вниз, ушел второй подбородок 

При этом они тренировалась на лайте:
- 3-4 раза в неделю дома
- делали 15-30 минутные комплексы и омолаживались даже ночью 

🏆 Итог: ушли отеки, подтянулся живот, ушел второй подбородок, осанка стала КОРОЛЕВСКОЙ

А ты готова к таким изменениям?"""

    media = [
        types.InputMediaPhoto(open('media/results1.jpg', 'rb'), caption=follow_up_text),
        types.InputMediaPhoto(open('media/results2.jpg', 'rb'))
    ]

    # Отправляем медиагруппу
    bot.send_media_group(chat_id, media)

    markup = types.InlineKeyboardMarkup()
    btn_join = types.InlineKeyboardButton(
        text="Присоединиться к клубу!",
        url="https://elviraelan.taplink.ws"
    )
    markup.add(btn_join)

    bot.send_message(chat_id, "👉🏻 Присоединяйся к клубу <b>«Elan beauty»</b> и почувствуй этот кайф на себе!", parse_mode='HTML', reply_markup=markup)
    # Через некоторое время отправляем последнее сообщение
    threading.Timer(3*60, send_final_pitch, args=[chat_id]).start()


def send_final_pitch(chat_id):
    name = bot.get_chat(chat_id).first_name

    final_text = f"""{name}, ты когда-нибудь задумывалась, куда уходят твои деньги?

<b><i>Каждый поход в магазин — 1500–3000 рублей на продукты, которые...</i></b>

🍟 добавляют лишние сантиметры и кг
🍟 создают ощущение тяжести
🍟 отражаются на твоем лице не в лучшую сторону

<b>А теперь сравни: 2599₽ — столько же и больше ты тратишь на еду за неделю.</b> Так ведь?

<b>А за эту же сумму ты получаешь:</b>
💪🏻 Тренировки для осанки, подтянутого овала лица, стройного тела и САМОЕ ГЛАВНОЕ - молодости
🤗 Уроки тейпированию
🔥 Советы по питанию

2599₽ — это вклад в стройное тело и самочувствие

👉🏻 <b>Разве это не самая выгодная инвестиция в себя и свое тело?</b>"""

    markup = types.InlineKeyboardMarkup()
    btn_join = types.InlineKeyboardButton(
        text="Убедила, иду!",
        url="https://elviraelan.taplink.ws"
    )
    markup.add(btn_join)

    bot.send_message(chat_id, final_text, reply_markup=markup, parse_mode='HTML')

if __name__ == '__main__':
    bot.polling(none_stop=True)