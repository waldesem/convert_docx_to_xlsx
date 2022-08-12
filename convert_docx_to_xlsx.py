###pip install pandas==1.4.3 python-docx openpyxl
import sys
import os,os.path
from pathlib import Path
from tkinter import EXCEPTION
from docx.api import Document
import pandas as pd
###Входим в каталог для перебора папок
input_dir = r'C:\Users\ubuntu\Documents\Кандидаты\\'
###Перебираем каталоги; ищем файлы с таблицами Word
for subdir, dirs, files in os.walk(input_dir):
    for file in files:
        if file.endswith(".docx"):
            in_file = os.path.join(subdir, file)
            output_file = file.split('.')[0]
            out_file = r'C:\Users\ubuntu\Documents\Кандидаты\\'+output_file+'.xlsx'
            ###Читаем документ, обрабатываем исключения файлов, которые созданы пустыми
            try:
                document = Document(in_file)
            except:
                document = Document()
            tables = document.tables
            ###Проверяем документы из которых удалено содержимое, 
            all_paras = document.paragraphs
            if len(all_paras) < 1:
                break            
            df = pd.DataFrame()
            ###Разбираем документ
            for table in document.tables:
                for row in table.rows:
                    text = [cell.text for cell in row.cells]
                    ### The frame.append method is deprecated and will be removed from pandas in a future version. Use pandas.concat instead.
                    df = df.append([text], ignore_index=True)
            ###Количество Column зависит от количество колонок в файле заключения 2 или 3, скрываем заголовки и индексы
            ###Добавляем строку путь к папке проверки
            if len(df.columns) == 2:
                df.columns = ["Column1", "Column2"]
                df.loc[len(df)] = ["url", subdir]
                df=df.T
                df.to_excel(out_file, header = False, index = False)
                df=df.T
            elif len(df.columns) == 3:
                df.columns = ["Column1", "Column2", "Column3"]
                df.loc[len(df)] = ["url", "url", subdir]
                df=df.T
                df.to_excel(out_file, columns=["Column1", "Column3"], header = False, index = False)
            else:
                continue
        else:
            continue
##Собираем информацию из файлов в одну таблицу (количество и названия столбцов должны совпадать)
path = Path(input_dir)
df = pd.concat([pd.read_excel(f) for f in path.glob("*.xlsx")], ignore_index=True)
df.to_excel(r'C:\Users\ubuntu\Documents\Кандидаты\Таблица.xlsx', header = False, index = False)
print(df)
