import telegram
import secret_keys

# TODO: add funcs description
def send_message_for_users_by_list(msg):
    bot = telegram.Bot(token=secret_keys.bot_token, request=telegram.utils.request.Request())
    for user in secret_keys.admins_tg_id_list:
        bot.sendMessage(chat_id=user, text=msg)
