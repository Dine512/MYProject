
from pdf2image import convert_from_path
import os
from PIL import Image
from utils import png_loading_imgbb
from openpyxl import load_workbook

path = r'C:\Users\Димас жренанас\Desktop\cheks\pdf'
xl_path = r'C:\Users\Димас жренанас\Desktop\cheks\payment.xlsx'

files = os.listdir(path)  # берётся список файлов из директории
list_files = [os.path.join(path, file) for file in files]  # к названиям файлов прекрепляется путь
time_list = [[x, os.path.getmtime(x)] for x in list_files]  # добавляется время создания файла
sort_time_list = sorted(time_list, key=lambda x: x[1])  # список сортируется по времени создания файлов
count = 1
row = 1

for file in sort_time_list:

    path_pdf_file = file[0]  # путь файла
    pages = convert_from_path(path_pdf_file, 500,
                              poppler_path=r'C:\poppler-22.12.0\Library\bin')  # указываем что конвертировать и в каком качестве
    path_png_file = fr'C:\Users\Димас жренанас\Desktop\cheks\png\{count}.png'
    for page in pages:  # Проходимся по страницам файла (у нас везде по одной, но без этого никак)
        page.save(path_png_file)  # Сохраняем файл в выбранную папку в формате PNG

    if 'Operation' in path_pdf_file:
        chek = Image.open(path_png_file)  # Выбираем файл
        chek_crop = chek.crop((0, 0, 2000, 3500))  # Обрезаем по точкам
        chek_crop.save(path_png_file)  # Сохраняем файл

    workbook = load_workbook(xl_path)   # Загружаем файл xmlx
    worklist = workbook['list1']    # Говорим на каком листе будем работать
    cell = worklist.cell(row=row, column=8)     # столбец комментариев о не прошедших платежах
    cell2 = worklist.cell(row=row, column=4)    # столбец комиссий
    cell3 = worklist.cell(row=row, column=9)    # столбец ссылок на чек


# Проверка что, в строке нет чека(не оплачено), что нет комиссии и нет комментария, что платёж не прошёл

    while (cell3.value is not None) or ((cell2.value is not None) and (cell2.value != 30)) or (cell.value is not None):
        row += 1    # переходим на следующую строку
        cell = worklist.cell(row=row, column=8)  # переназначаем строку
        cell2 = worklist.cell(row=row, column=4)
        cell3 = worklist.cell(row=row, column=9)
        print(50)

    worklist[f'I{row}'] = png_loading_imgbb(path_png_file)  # Вызываем функцию и записываем ссылку в ячейку
    workbook.save(xl_path)  # Сохраняем xmlx файл
    workbook.close()    # закрываем xmlx файл
    row += 1
    count += 1
