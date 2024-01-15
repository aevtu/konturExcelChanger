
from telegram.ext import Updater, MessageHandler, Filters, CommandHandler
import openpyxl

text = ""

def set_val(stv):
    global text
    print("получили ", stv)
    text = stv

def get_val():
    print("получ", text)
    return text

# Обработчик команды /start
def start(update, context):
    user = update.message.from_user
    context.bot.send_message(chat_id=update.message.chat_id, text=f"Привет, {user.first_name}!")
    context.bot.send_message(chat_id=update.effective_chat.id, text="Введи значение:")

# Обработчик текстовых сообщений
def handle_text(update, context):
    text = update.message.text
    context.bot.send_message(chat_id=update.effective_chat.id, text=f"Ты ввел: {text}")
    value = context.user_data.get('value')
    print("получили", text)
    set_val(text)

    # Сохраняем значение в переменную
    context.user_data['value'] = text


def handle_excel(update, context):
    # Get the file from the update

    file = context.bot.get_file(update.message.document.file_id)
    # Download the file
    file.download('file.xlsx')

    COLUMN_A = 'A'
    COLUMN_B = 'B'
    COLUMN_C = 'C'

    def count_row(max_row, count=0):
        print("попали сюда")
        for row in range(1, max_row + 1):
            cell_value = sheet[COLUMN_A + str(row)].value
            if cell_value is not None:
                count += 1
        count -= 1
        print("Количество ячеек в столбце A: " + str(count))
        nextFormula(count)

    def nextFormula(count_max):
        start_row = 2
        end_row = count_max + 1
        sheet['C1'] = "Link"
        for row in range(start_row, end_row + 1):
            cell = 'C{}'.format(row)
            formula = '="{}" &  A{}'.format(get_val(),row)
            sheet[cell].value = formula
        workbook.save('file_updated.xlsx')
        context.bot.send_document(chat_id=update.message.chat_id, document=open('file_updated.xlsx', 'rb'))

    # Open the Excel file
    workbook = openpyxl.load_workbook('file.xlsx')
    sheet = workbook.active
    max_row = sheet.max_row
    count_row(max_row)

updater = Updater('6675800876:AAGStcXCPcKrftN5naYNaTkK515sKSwfjL4', use_context=True)
dispatcher = updater.dispatcher
dispatcher.add_handler(MessageHandler(Filters.document, handle_excel))
start_handler = CommandHandler('start', start)
dispatcher.add_handler(start_handler)

text_handler = MessageHandler(Filters.text, handle_text)
dispatcher.add_handler(text_handler)


updater.start_polling()





