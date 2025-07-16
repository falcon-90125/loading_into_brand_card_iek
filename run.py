import pandas as pd
import numpy as np
import os
import io
from PyPDF2 import PdfReader
import tkinter as tk
import json

# Вытаскиваем PathSupplierLetter – путь к письму поставщика на ПК
# folder_path = 'C:\Users\sokolov\Yandex_Disk\MyData\Projects\loading_into_brand_card_iek\exchange\'
folder_path = 'C:/Users/sokolov/Yandex_Disk/MyData/Projects/loading_into_brand_card_iek/exchange/'
for file in os.listdir(folder_path):
    if '.msg' in file:
#     if file.lower().endswith('.msg'):
#         PathSupplierLetter = os.path.join(folder_path, file)
        PathSupplierLetter = folder_path + file
# print(PathSupplierLetter)

specification_xlsx = pd.read_excel('exchange\Т250715-36059.xlsx') #загружаем спецификацию

# Вытаскиваем DiscountEKS – наша скидка
discount_from_base = "Устанавливаемая скидка (скидка от базовой), %"
# Поиск индексов (строка, столбец), где встречается текст
result = np.where(specification_xlsx.apply(lambda x: x.astype(str).str.contains(discount_from_base, regex=False)))
row_idx = result[0][0]  # Индекс строки
col_idx = result[1][0]  # Индекс столбца
DiscountEKS = float(specification_xlsx.iat[row_idx+1, col_idx])
# print(DiscountEKS)

with open('exchange\P250716-38591.pdf', 'rb') as file:
    pdf_bytes = file.read()

# Создаем PdfReader из бинарных данных
pdf_reader = PdfReader(io.BytesIO(pdf_bytes))

# Извлекаем текст со всех страниц
pdf_text = ""
for page in pdf_reader.pages:
    pdf_text += page.extract_text() + "\n"

text = pdf_text.strip()
text = text.replace('\n', ' ')
text = text.replace('    ', ' ')
text = text.replace('   ', ' ')
text = text.replace('  ', ' ')

# Вытаскиваем DiscountBuyer – скидка клиента
text_discount_index = text.index('По продукции IEK - скидка до ')
text_discount = text[text_discount_index+29:text_discount_index+40]
text_discount = text_discount[:text_discount.index('%')]
text_discount = text_discount.replace(',', '.')
DiscountBuyer = float(text_discount)
# print(DiscountBuyer)

# Вытаскиваем VendorCode – код вендора
vendor_code = "Номер проекта в системе IEK"
# Поиск индексов (строка, столбец), где встречается текст
result = np.where(specification_xlsx.apply(lambda x: x.astype(str).str.contains(vendor_code, regex=False)))
row_idx = result[0][0]  # Индекс строки
col_idx = result[1][0]  # Индекс столбца
VendorCode = specification_xlsx.iat[row_idx, col_idx+1]
# print(VendorCode)

# Вытаскиваем ReceivedCondition – сумма полученных условий
text_order_amount_index = text.index('Сумма заказа ')
text_order_amount = text[text_order_amount_index+12:text_order_amount_index+40]
text_order_amount = text_order_amount[:text_order_amount.index(' рублей с учетом')]
text_order_amount = text_order_amount.replace(' ', '')
text_order_amount = text_order_amount.replace(',', '.')
ReceivedCondition = float(text_order_amount)
# print(ReceivedCondition)

# Вытаскиваем DateCondition – до какой даты действуют
date_condition = "Конец действия цен"
# Поиск индексов (строка, столбец), где встречается текст
result = np.where(specification_xlsx.apply(lambda x: x.astype(str).str.contains(date_condition, regex=False)))
row_idx = result[0][0]  # Индекс строки
col_idx = result[1][0]  # Индекс столбца
DateCondition = specification_xlsx.iat[row_idx, col_idx+1]
DateCondition = str(DateCondition)
# print(DateCondition)

# Создание главного окна
root = tk.Tk()
root.title("Преференции")

label1 = tk.Label(                      # Создание виджета метки
    text="Выберите тип преференции", # Задание отображаемого текста
    fg="#07e374",                       # Установка цвета текста
    bg="#2f457c",                       # Устанавка фона
    width=60,                           # Установка ширина виджета (в текстовых юнитах)
    height=5                           # Установка высоты виджета (в текстовых юнитах) 
)
label1.pack()  

def preference_nom_exclusive():
    global Preference
    Preference = 1
    return Preference

def preference_nom_on_request():
    global Preference
    Preference = 2
    return Preference

# Создание кнопки
button1 = tk.Button(text="Эксклюзив", command=preference_nom_exclusive)
button1.pack(pady=20)

button1 = tk.Button(text="По запросу", command=preference_nom_on_request)
button1.pack(pady=20)

label2 = tk.Label(                      # Создание виджета метки
    text="После выбора закройте это окно", # Задание отображаемого текста
    fg="#07e374",                       # Установка цвета текста
    bg="#2f457c",                       # Устанавка фона
    width=60,                           # Установка ширина виджета (в текстовых юнитах)
    height=5                           # Установка высоты виджета (в текстовых юнитах) 
)
label2.pack()  
root.mainloop()
# print(Preference)

# Данные для сохранения
data = {
    "DiscountEKS": DiscountEKS,
    "DiscountBuyer": DiscountBuyer,
    "Preference": Preference,
    "VendorCode": VendorCode,
    "ReceivedCondition": ReceivedCondition,
    "DateCondition ": DateCondition,
    "PathSupplierLetter ": PathSupplierLetter
}

# Укажите путь, куда сохранить файл
json_file_path = "exchange/data.json"  # Замените на свой путь

# Сохраняем данные в JSON-файл
with open(json_file_path, "w", encoding="utf-8") as file:
    json.dump(data, file, indent=4, ensure_ascii=False)  # indent для красивого форматирования