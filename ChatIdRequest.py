from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, CallbackContext

TELEGRAM_API_TOKEN = 'token'
YOUR_CHAT_ID = '1008683615'
OTHER_BOT_LINK = 'https://t.me/AlbertHochuRabotatDoma_bot'


async def start(update: Update, context: CallbackContext):
    await update.message.reply_text('Напишите ваш номер телефона, район, должность(если портальщик то зону ответственности, например: МКД) для прохождения регистрации.')


async def handle_message(update: Update, context: CallbackContext):
    user_chat_id = update.message.chat.id
    user_message = update.message.text

    await context.bot.send_message(chat_id=YOUR_CHAT_ID, text=f"Получено сообщение от chat_id {user_chat_id}:\n{user_message}")
    await update.message.reply_text(text=f"Благодарю за регистрацию! Вы сможете воспользоваться функционалом в течении 10 минут. Вот ссылка на основного бота:https://t.me/AlbertHochuRabotatDoma_bot")


def main():
    application = Application.builder().token(TELEGRAM_API_TOKEN).build()

    application.add_handler(CommandHandler('start', start))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    application.run_polling()


if __name__ == "__main__":
    main()