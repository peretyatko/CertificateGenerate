import os.path

import config
import telebot
import os
import zipfile
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import tempfile


bot = telebot.TeleBot(config.TOKEN)



@bot.message_handler(content_types=['document'])
def handle_file(message):
    try:
        print(message.document.file_id)
        chat_id = message.chat.id
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        src = os.path.join(message.document.file_name)
        with open(src, 'wb') as new_file:
            new_file.write(downloaded_file)

        form = pd.read_excel(message.document.file_name)

        def certificate_generate(row):
            im = Image.open(row.Curse)
            d = ImageDraw.Draw(im)
            location = (430, 1800)
            text_color = (0, 0, 0)
            font = ImageFont.truetype("calibri.ttf", 50)
            d.text(location, row.Date, fill=text_color, font=font, align="center")
            location = (430, 2500)
            d.text(location, row.Teacher, fill=text_color, font=font, align="center")
            location = (420, 1300)
            text_color = (0, 0, 0)
            font = ImageFont.truetype("calibri.ttf", 120)
            d.text(location, row.Name, fill=text_color, font=font, align="middle")
            file_name = f"certificate_{row.Name}.pdf"
            im.save(file_name, quality=100, resolution=300)

            return file_name

        def create_zip(file_list: list):
            file = tempfile.mktemp('.zip')
            with zipfile.ZipFile('students.zip', 'w') as my_zip:
                for i in file_list:
                    try:
                        my_zip.write(i)
                    except:
                        raise
                    else:
                        os.remove(i)
            return file

        file_list = [certificate_generate(x) for x in form.itertuples()]

        create_zip(file_list)
        def send_file():
            st = open('students.zip', 'rb')
            bot.send_document(message.chat.id, st)
        send_file()



    except Exception as e:
        bot.reply_to(message, e)



if __name__ == '__main__':
    bot.infinity_polling()
