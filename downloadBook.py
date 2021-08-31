import os
import requests
from bs4 import BeautifulSoup
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF
from PyPDF2 import PdfFileReader, PdfFileWriter
import time
from PIL import Image
import getpass

'''
Скачивание книг с сайта https://urait.ru/
'''

email = input('Твой логин (email): ')
passwd = input('Твой пароль: ')


user_agent = 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Safari/537.36'

url = 'https://urait.ru/login'
sess = requests.Session()
# sess.verify = False # ВКЛ если у сайта самоподписанный сертификат

resp = sess.post(url, {
    'user-agent': user_agent,
    'email': email,
    'password': passwd,
})


resp = sess.get('https://urait.ru/cabinet/profile')


ss = BeautifulSoup(resp.content, "html.parser")
try:
    ss_1 = ss.find('span', class_='content-title__name')
    print(f'Вошел как: {ss_1.text}')
except AttributeError:
    input('проверь данные и попробуй авторизоваться еще разок')
    exit()

link = input('Ввести URL изображения: ')

first_sheet = int(input('С какой страницы начинать: '))
last_sheet = int(input("До какой страницы скачивать (включительно): "))
if isinstance(first_sheet, int) and isinstance(last_sheet, int):
    paths = []
    my_link = link[0:-1]
    for i in range(first_sheet, last_sheet + 1):
        time.sleep(2)
        ufr = sess.get(f'{my_link}{i}')
        try:
            file = f'Try_{i}.svg'
            f = open(file, "wb")  # открываем файл для записи, в режиме wb
            f.write(ufr.content)  # записываем содержимое в файл; как видите - content запроса
            f.close()

            # Запись файла в pdf
            drawing = svg2rlg(file)
            renderPDF.drawToFile(drawing, f"file{i}.pdf")
            paths.append(f"file{i}.pdf")
        except:
            file = f'Try_{i}.png'
            os.remove(f'Try_{i}.svg')
            f = open(file, "wb")  # открываем файл для записи, в режиме wb
            f.write(ufr.content)  # записываем содержимое в файл; как видите - content запроса
            f.close()

            image = Image.open(file)
            im1 = image.convert('RGB')
            im1.save(f"file{i}.pdf")
            paths.append(f"file{i}.pdf")


        # Удаление файла svg
        os.remove(file)

    # объединение файлов
    pdf_writer = PdfFileWriter()

    for path in paths:
        pdf_reader = PdfFileReader(path)
        for page in range(pdf_reader.getNumPages()):
            # Add each page to the writer object
            pdf_writer.addPage(pdf_reader.getPage(page))
    output = input('Как назвать файл: ')
    # Write out the merged PDF
    with open(f"{output}.pdf", 'wb') as out:
        pdf_writer.write(out)

    for i in paths:
        os.remove(i)

else:
    print('только целые числа')

input('...конец! (нажми enter)')
