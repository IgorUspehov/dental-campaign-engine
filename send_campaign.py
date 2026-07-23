import pandas as pd
import os
import re
from datetime import datetime

FILE_NAME = "patients.xlsx"
MESSAGE_TEXT = """Спасибо, что посетили нашу клинику!
Нам очень важно знать ваше мнение. Если у вас есть минутка, пожалуйста, оставьте отзыв о вашем посещении в Google Maps. Это поможет другим пациентам сделать выбор и поможет нам становиться ещё лучше.
Оставить отзыв:
https://g.page/ВАША_ССЫЛКА/review"""

SMS_API_KEY = ""
TWILIO_ACCOUNT = ""
TWILIO_TOKEN = ""
WA_API_KEY = ""

def clean_phone(num):
    num = str(num)
    num = re.sub(r'[^0-9+]', '', num)
    if num.startswith('8'): return '+' + num
    if num.startswith('7'): return '+' + num
    return num

def process_file():
    print("\n--- ЗАПУСК СКРИПТА ---")
    file_path = os.path.join("data", FILE_NAME)
    if not os.path.exists(file_path):
        print(f"Ошибка: Файл {FILE_NAME} не найден в папке data!")
        return
    df = pd.read_excel(file_path)
    df['Phone_Clean'] = df['Phone'].apply(clean_phone)
    print(f"Загружено записей: {len(df)}")
    results = []
    for index, row in df.iterrows():
        name = row.get('Name', 'Пациент')
        phone = row['Phone_Clean']
        sms_status = "Sent (Demo)"
        wa_status = "Sent (Demo)"
        with open("logs/activity.log", "a", encoding='utf-8') as log_file:
            log_file.write(f"{datetime.now()}: Обработан {phone}\n")
        results.append({
            'Name': name,
            'Phone': phone,
            'SMS Status': sms_status,
            'WhatsApp Status': wa_status
        })
    report = pd.DataFrame(results)
    filename = f"reports/report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    report.to_excel(filename, index=False)
    print(f"Готово! Отчет сохранен в папку reports: {filename}")
    print("--- ЗАВЕРШЕНО ---")

if __name__ == "__main__":
    process_file()
    input("\nНажмите Enter, чтобы закрыть окно...")
