"""
Скрипт для объединения docx файлов со списками фильмов, найденных по шаблону (PATTERN) в один файл (OUTPUT).
Файлы со списками сортируются по дате в подпапке (DATE_SOURCE = 1), или по дате в имени файла (DATE_SOURCE = 2).
"""

# PATTERN = R'Z:\Фильмы по годам\Фильмы 2023\Фильмы 2023\Новая папка\*.docx'
# OUTPUT = R'Z:\Фильмы по годам\Фильмы 2023\Фильмы 2023\Новая папка\Объединенные.docx'
PATTERN = R'Z:\Фильмы по годам\Фильмы 2023\Фильмы 2023\*\*.docx'
OUTPUT = R'Z:\Фильмы по годам\Фильмы 2023\Фильмы 2023\Фильмы 2023_combined.docx'

DATE_SOURCE = 1

import os, sys
from datetime import datetime
from glob import glob
from os.path import basename, dirname, splitext

sys.path.append('lib')
from docx import Document as Document_compose
from docxcompose.composer import Composer


def combine_all_docx(files_list, output):
    filename_master = files_list.pop(0)
    master = Document_compose(filename_master)
    composer = Composer(master)
    for file in files_list:
        temp = Document_compose(file)
        composer.append(temp)
    composer.save(output)


def sort_type(date, type):
    """
    Функция - выбор даты, как критерия для сортировки
    
    type = 1: дата в подпапке
    type = 2: дата в имени файла
    """

    if type == 1:
        return basename(dirname(date))
    elif type == 2:
        return splitext(basename(date))[0]
    else:
        raise Exception("Неверно выбран тип сортировки!")


def main():
    files = glob(PATTERN)
    if not files:
        print("Файлы не найдены")
        return
    files.sort(key=lambda date: datetime.strptime(sort_type(date, DATE_SOURCE), "%d.%m.%Y"))  # сортировка файлов
    print("Файлы для объединения:", "\n".join(files), sep='\n')
    print('Всего файлов:', len(files))
    combine_all_docx(files, OUTPUT)
    print('Создан файл:', OUTPUT)
    print('Выполнено')


if __name__ == "__main__":
    main()
    os.system('pause')
