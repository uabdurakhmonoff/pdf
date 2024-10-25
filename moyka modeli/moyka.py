import os
import zipfile
import requests
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, CallbackQueryHandler
from telegram.ext.filters import Document
from fpdf import FPDF
import docx2pdf
import pandas as pd

async def hello(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(f'Hello {update.effective_user.first_name}')

async def unzip(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    file = await context.bot.get_file(update.message.document.file_id)
    file_path = file.file_path
    r = requests.get(file_path)
    with open('temp.zip', 'wb') as f:
        f.write(r.content)

    with zipfile.ZipFile('temp.zip', 'r') as zip_ref:
        zip_ref.extractall('extracted')
        file_names = zip_ref.namelist()
        for file_name in file_names:
            print(file_name)

    await update.message.reply_text("Zip fayl chiqarildi!")

    for file_name in file_names:
        file_path = os.path.join('extracted', file_name)
        await context.bot.send_document(chat_id=update.message.chat_id, document=open(file_path, 'rb'))

async def convert_to_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    file = await context.bot.get_file(update.message.document.file_id)
    file_path = file.file_path
    file_name = update.message.document.file_name
    r = requests.get(file_path)
    with open(file_name, 'wb') as f:
        f.write(r.content)

    if file_name.endswith('.docx'):
        output_file = file_name.replace('.docx', '.pdf')
        docx2pdf.convert(file_name, output_file)
        await context.bot.send_document(chat_id=update.message.chat_id, document=open(output_file, 'rb'))
    elif file_name.endswith('.xlsx'):
        output_file = file_name.replace('.xlsx', '.pdf')
        excel = pd.read_excel(file_name, engine='openpyxl')
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        for i in range(len(excel)):
            pdf.cell(200, 10, txt=str(excel.iloc[i]), ln=True)
        pdf.output(output_file)
        await context.bot.send_document(chat_id=update.message.chat_id, document=open(output_file, 'rb'))

app = ApplicationBuilder().token("7531287424:AAHO2ELX542VralDfLbyLYX-7idTZL2QL8w").build()

app.add_handler(CommandHandler("hello", hello))
app.add_handler(MessageHandler(Document.MimeType("application/zip"), unzip))
app.add_handler(MessageHandler(Document.MimeType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"), convert_to_pdf))
app.add_handler(MessageHandler(Document.MimeType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), convert_to_pdf))

app.run_polling()
