
from pdf2image import convert_from_path
import os
from PIL import Image
from utils import png_loading_imgbb
from openpyxl import load_workbook

path = r'C:\Users\Димас жренанас\Desktop\cheks\pdf'
xl_path = r'C:\Users\Димас жренанас\Desktop\cheks\payment.xlsx'

files = os.listdir(path)  # берётся список файлов из директории
list_files = [[file.split('.')[0], os.path.join(path, file)] for file in files]  # к названиям файлов прикрепляется путь

count = 1
row = 1

for file in list_files:

    path_pdf_file = file[1]  # путь файла
    file_name = file[0]
    pages = convert_from_path(path_pdf_file, 500,
                              poppler_path=r'C:\poppler-22.12.0\Library\bin')  # указываем что конвертировать и в каком качестве
    path_png_file = fr'C:\Users\Димас жренанас\Desktop\cheks\png\{count}.png'
    for page in pages:  # Проходимся по страницам файла (у нас везде по одной, но без этого никак)
        page.save(path_png_file)  # Сохраняем файл в выбранную папку в формате PNG

    if len(file_name) < 16:
        chek = Image.open(path_png_file)  # Выбираем файл
        chek_crop = chek.crop((0, 0, 2000, 3500))  # Обрезаем по точкам
        chek_crop.save(path_png_file)  # Сохраняем файл

    workbook = load_workbook(xl_path)   # Загружаем файл xmlx
    worklist = workbook['list1']    # Говорим на каком листе будем работать
    i = 1
    while not (file_name in worklist[f'A{i}'].value and worklist[f'D{i}'].value is None):
        i += 1
    worklist[f'I{i}'] = png_loading_imgbb(path_png_file)
    workbook.save(xl_path)  # Сохраняем xmlx файл
    workbook.close()    # закрываем xmlx файл
    row += 1
    print(f'Осталось: {len(list_files) - count}')
    count += 1
