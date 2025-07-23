import pandas as pd
import numpy as np
import os
import json

import io
import extract_msg
from PyPDF2 import PdfReader  # Для работы с PDF
from typing import Tuple  # Для типизации

# Вытаскиваем PathSupplierLetter – путь к письму поставщика на ПК
folder_path = 'C:/Users/sokolov/Yandex_Disk/MyData/Projects/loading_into_brand_card_iek/1/'

for file in os.listdir(folder_path):
    if '.msg' in file:
        PathSupplierLetter = folder_path + file
# print(PathSupplierLetter)

def process_msg_file(PathSupplierLetter: str) -> Tuple[pd.DataFrame, str]:
    """
    Обрабатывает файл .msg, извлекает Excel и PDF вложения,
    преобразует Excel в DataFrame, PDF в текст.
    
    Args:
        msg_path (str): Путь к файлу .msg
        
    Returns:
        Tuple[pd.DataFrame, str]: DataFrame из Excel и текст из PDF
    """
    excel_df = None
    pdf_text = ""

    msg = extract_msg.Message(PathSupplierLetter)
        
    for attachment in msg.attachments:
        filename = attachment.longFilename
        
        # Обработка Excel файла
        if filename.lower().endswith('.xlsx'):
            excel_data = io.BytesIO(attachment.data)
            specification_xlsx = pd.read_excel(excel_data)
            
        # Обработка PDF файла
        elif filename.lower().endswith('.pdf'):
            pdf_data = io.BytesIO(attachment.data)
            pdf_reader = PdfReader(pdf_data)
            
            # Собираем текст со всех страниц
            for page in pdf_reader.pages:
                pdf_text += page.extract_text() + "\n"
                
        else:
            raise ValueError(f"Неподдерживаемый формат файла: {filename}")
            
    msg.close()
        
    return specification_xlsx, pdf_text

specification_xlsx, pdf_text = process_msg_file(PathSupplierLetter)
text = pdf_text.strip()
text = text.replace('\n', ' ')
text = text.replace('    ', ' ')
text = text.replace('   ', ' ')
text = text.replace('  ', ' ')
# print(text[:30])

# Вытаскиваем DiscountEKS – наша скидка
discount_from_base = "Устанавливаемая скидка (скидка от базовой), %"
# Поиск индексов (строка, столбец), где встречается текст
result = np.where(specification_xlsx.apply(lambda x: x.astype(str).str.contains(discount_from_base, regex=False)))
row_idx = result[0][0]  # Индекс строки
col_idx = result[1][0]  # Индекс столбца
DiscountEKS = float(specification_xlsx.iat[row_idx+1, col_idx])
# print(DiscountEKS)

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
text_order_amount = text[text_order_amount_index+12:text_order_amount_index+45]
text_order_amount = text_order_amount[:text_order_amount.index('рублей с учетом')]
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

# Данные для сохранения
Preference = 1 # Задаётся вручную

data = {
    "DiscountEKS": DiscountEKS,
    "DiscountBuyer": DiscountBuyer,
    "Preference": Preference,
    "VendorCode": VendorCode,
    "ReceivedCondition": ReceivedCondition,
    "DateCondition": DateCondition,
    "PathSupplierLetter": PathSupplierLetter
}

# Укажите путь, куда сохранить файл
json_file_path = "data.json"  # Замените на свой путь

# Сохраняем данные в JSON-файл
with open(json_file_path, "w", encoding="utf-8") as file:
    json.dump(data, file, indent=4, ensure_ascii=False)  # indent для красивого форматирования