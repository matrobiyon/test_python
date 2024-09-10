#Для запуска файла нужны: pandas and openpyxl
# Можно было установить в docker

import pandas as pd

df = pd.read_excel('export.xlsx')

# Преобразование в числовой тип с обработкой ошибок
df['OSTTEHNO'] = pd.to_numeric(df['OSTTEHNO'], errors='coerce')
# Преобразование в int
df['OSTTEHNO'] = df['OSTTEHNO'].fillna(0).round().astype(int)
# 1. Увеличение цены на 2 у позиций с четным числом в ячейке 'OSTTEHNO'
df.loc[df['OSTTEHNO'] % 2 == 0, 'CENA'] += 2

df.loc[df["NAME"].str.contains(" 10 "), "NAME"] = df["NAME"] + " " + df["CENA"].astype(str) + " цена"

# 2. Добавление новых позиций
new_positions = [
    [427820, "Кружка PERFECTO LINEA Amber 400 мл.", 510, 0, 0, 2, 0, 0, "шт.", "026. Посуда"],
    [427821, "Набор ключей PRO STARTUL TORX 9 шт.Т10-Т50 CrV (PRO-87209)", 330, 1, 0, 2, 0, 0, "шт.", "091. Ключи, головки"],
    [427822, "Набор ключей PRO STARTUL HEX 9 шт.1,5-10мм (PRO-89409)", 150, 0, 1, 4, 0, 0, "шт.", "091. Ключи, головки"]
]

df_new = pd.DataFrame(new_positions, columns=df.columns)
df = pd.concat([df, df_new], ignore_index=True)


# 3. Удаление строки с кодом 421730
df = df[df['KOD'] != 421730]


# 5. Сортировка файла по цене в ячейке 'CENA'
df = df.sort_values(by='CENA')

df.to_excel("export_modified.xlsx", index=False)

