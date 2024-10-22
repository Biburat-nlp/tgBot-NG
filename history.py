from telegram import Update, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import Application, CallbackQueryHandler, MessageHandler, filters, CommandHandler, CallbackContext
import pandas as pd
import hashlib

# Загрузка данных из файла Excel
DATA_PATH = "C:\\Users\\gamag\\Downloads\\ОИВ Сообщения 2024-10-22 ID-5915256.xlsx"

def load_data():
    df = pd.read_excel(DATA_PATH)
    return df

df = load_data()

async def start(update: Update, context: CallbackContext):
    keyboard = [[InlineKeyboardButton("История объекта", callback_data='history')]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text('Выберите опцию:', reply_markup=reply_markup)

async def history_callback(update: Update, context: CallbackContext):
    query = update.callback_query
    await query.answer()
    await query.message.reply_text("Напишите адрес")
    context.user_data['waiting_for_address'] = True

async def address_handler(update: Update, context: CallbackContext):
    if context.user_data.get('waiting_for_address'):
        user_input = update.message.text
        context.user_data['user_address'] = user_input
        filtered_data = df[df['Адрес'].str.contains(user_input, case=False, na=False)]
        print(f"Введённый адрес: {user_input}")
        print(f"Данные после фильтрации: {filtered_data}")
        if not filtered_data.empty:
            grouped_data = (
                filtered_data.groupby('Проблемная тема')['Номер сообщения']
                .nunique()
                .reset_index()
                .sort_values(by='Номер сообщения', ascending=False)
            )

            keyboard = []
            for _, row in grouped_data.iterrows():
                problem_theme = row['Проблемная тема']
                message_count = row['Номер сообщения']
                callback_data = f'theme_{hashlib.md5(problem_theme.encode()).hexdigest()[:10]}'
                button_text = f"{problem_theme} ({message_count})"
                keyboard.append([InlineKeyboardButton(button_text, callback_data=callback_data)])
                context.user_data[callback_data] = problem_theme

            reply_markup = InlineKeyboardMarkup(keyboard)
            await update.message.reply_text(f"Результаты по адресу: {user_input}", reply_markup=reply_markup)
        else:
            await update.message.reply_text("Адрес не найден, попробуйте ещё раз.")
        context.user_data['waiting_for_address'] = False

async def button_handler(update: Update, context: CallbackContext):
    query = update.callback_query
    await query.answer()

    problem_theme = context.user_data.get(query.data)
    user_address = context.user_data.get('user_address')

    if problem_theme and user_address:
        filtered_data = df[(df['Проблемная тема'] == problem_theme) & (df['Адрес'].str.contains(user_address, case=False, na=False))]

        for _, row in filtered_data.iterrows():
            object_id = int(row['ID объекта'])
            message_id = row['Номер сообщения']

            details = (
                f"Номер сообщения: {row['Номер сообщения']}\n"
                f"Адрес: {row['Адрес']}\n"
                f"Дата публикации сообщения: {row['Дата публикации сообщения']}\n"
                f"Проблемная тема: {row['Проблемная тема']}\n"
                f"Статус подготовки ответа на сообщение: {row['Статус подготовки ответа на сообщение']}\n"
                f"Ответственный за подготовку ответа: {row['Ответственный за подготовку ответа']}\n"
                f"Ссылка на объект: [Перейти к объекту](https://gorod.mos.ru/objects/{object_id}/messages#{message_id})\n"
            )

            await query.message.reply_text(details)

    else:
        await query.message.reply_text("Произошла ошибка, попробуйте снова.")

def main():
    application = Application.builder().token('token').build()

    # Обработчики
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CallbackQueryHandler(history_callback, pattern='history'))
    application.add_handler(CallbackQueryHandler(button_handler, pattern='theme_'))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, address_handler))

    application.run_polling()

if __name__ == '__main__':
    main()
