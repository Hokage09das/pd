import os
import pandas as pd
from openpyxl import load_workbook

# print(pd.__version__)

def parse_excel_to_dict_list(filepath: str, sheet_name='Sheet1'):
    df = pd.read_excel(filepath, sheet_name=sheet_name)

    dict_list = df.to_dict(orient='records')

    return dict_list


def create_empty_excel(columns: list, filename: str, sheet_name: str = 'Лист1'):
    df = pd.DataFrame(columns)

    if not os.path.exists('Винтажный Анализ'):
        os.makedirs('Винтажный Анализ')

    filepath = os.path.join('Винтажный Анализ', filename)
    excel_writer = pd.ExcelWriter(filepath, engine='openpyxl')
    print(excel_writer)
    df.to_excel(excel_writer, index=False, sheet_name=sheet_name)
    excel_writer._save()

    return filepath

def create_or_append_excel(columns: list, filename: str, sheet_name: str = 'Лист1'):
    df = pd.DataFrame(columns)

    if not os.path.exists('Винтажный Анализ'):
        os.makedirs('Винтажный Анализ')

    filepath = os.path.join('Винтажный Анализ', filename)

    if os.path.exists(filepath):
        with pd.ExcelWriter(filepath, engine='openpyxl', mode='a', if_sheet_exists='new') as excel_writer:
            df.to_excel(excel_writer, index=False, sheet_name=sheet_name)
    else:
        with pd.ExcelWriter(filepath, engine='openpyxl') as excel_writer:
            df.to_excel(excel_writer, index=False, sheet_name=sheet_name)

    return filepath