import pandas as pd
import os
import re
from datetime import datetime

# ==========================================
# НАСТРОЙКИ (Клиент меняет эти данные перед запуском)
# ==========================================
FILE_NAME = "patients.xlsx"
MESSAGE_TEMPLATE = """Здравствуйте, {name}!
Спасибо, что посетили нашу клинику!
Нам очень важно знать ваше мнение. Если у вас есть минутка, пожалуйста, оставьте отзыв о вашем посещении в Google Maps. Это поможет другим пациентам сделать выбор и поможет нам становиться ещё лучше.
Оставить отзыв:
https://g.page/ВАША_ССЫЛКА/review"""

SMS_API_KEY = ""
TWILIO_ACCOUNT = ""
TWILIO_TOKEN = ""
WA_API_KEY = ""

NAME_COLUMNS = ["Name", "Имя", "Vorname", "Patient", "ФИО"]
SURNAME_COLUMNS = ["Nachname", "Фамилия", "Surname"]
PHONE_COLUMNS = ["Phone", "Телефон", "Mobil", "Handy", "Telefon", "Telefon mobil", "Mobile"]


def find_column(df, candidates):
    lower_map = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower().strip() in lower_map:
            return lower_map[cand.lower().strip()]
    return None


def build_full_name(row, name_col, surname_col):
    first = str(row.get(name_col, "")).strip() if name_col else ""
    last = str(row.get(surname_col, "")).strip() if surname_col else ""
    if first and first.lower() == "nan":
        first = ""
    if last and last.lower() == "nan":
        last = ""
    full = " ".join(part for part in [first, last] if part)
    return full if full else "Пациент"


def clean_phone(num):
    if pd.isna(num):
        return ""
    num = str(num)
    num = re.sub(r'[^0-9+]', '', num)
    if num.startswith('8'):
        return '+' + num
    if num.startswith('7'):
        return '+' + num
    if num.startswith('0'):
        return '+49' + num[1:]
    return num


def is_valid_phone(phone):
    digits = re.sub(r'[^0-9]', '', phone)
    return len(digits) >= 8


def load_patients_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        return pd.read_excel(file_path, dtype=str)
    elif ext == ".csv":
        return pd.read_csv(file_path, dtype=str)
    elif ext == ".txt":
        try:
            return pd.read_csv(file_path, sep="\t", dtype=str)
        except Exception:
            return pd.read_csv(file_path, sep=",", dtype=str)
    else:
        raise ValueError(f"Неподдерживаемый формат файла: {ext}")


def process_file():
    print("\n--- ЗАПУСК СКРИПТА ---")

    for folder in ["data", "reports", "logs"]:
        os.makedirs(folder, exist_ok=True)

    file_path = os.path.join("data", FILE_NAME)

    if not os.path.exists(file_path):
        print(f"ОШИБКА: Файл {FILE_NAME} не найден в папке data!")
        print("Положите файл с пациентами (xlsx, csv или txt) в папку data и запустите снова.")
        return

    try:
        df = load_patients_file(file_path)
    except Exception as e:
        print(f"ОШИБКА при чтении файла: {e}")
        return

    name_col = find_column(df, NAME_COLUMNS)
    surname_col = find_column(df, SURNAME_COLUMNS)
    phone_col = find_column(df, PHONE_COLUMNS)

    if phone_col is None:
        print("ОШИБКА: не найдена колонка с телефоном.")
        print(f"Колонки в файле: {list(df.columns)}")
        print(f"Ожидались одно из названий: {PHONE_COLUMNS}")
        return

    if name_col is None and surname_col is None:
        print("Внимание: колонки с именем не найдены, будет использовано 'Пациент'.")

    print(f"Загружено записей: {len(df)}")
    print(f"Колонка с именем: {name_col or '(не найдена)'}")
    print(f"Колонка с фамилией: {surname_col or '(не найдена)'}")
    print(f"Колонка с телефоном: {phone_col}")

    results = []
    skipped = 0
    log_path = os.path.join("logs", "activity.log")

    with open(log_path, "a", encoding="utf-8") as log_file:
        for index, row in df.iterrows():
            full_name = build_full_name(row, name_col, surname_col)
            raw_phone = row.get(phone_col, "")
            phone = clean_phone(raw_phone)

            if not is_valid_phone(phone):
                skipped += 1
                log_file.write(f"{datetime.now()}: ПРОПУЩЕН (невалидный номер) — исходное значение: {raw_phone}\n")
                continue

            personalized_message = MESSAGE_TEMPLATE.format(name=full_name)

            sms_status = "Sent (Demo)"
            wa_status = "Sent (Demo)"

            log_file.write(f"{datetime.now()}: Обработан {phone} ({full_name})\n")

            results.append({
                "Пациент": full_name,
                "Телефон": phone,
                "SMS": sms_status,
                "WhatsApp": wa_status,
                "Текст сообщения": personalized_message,
            })

    if results:
        report = pd.DataFrame(results)
        filename = os.path.join("reports", f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            report.to_excel(writer, index=False, sheet_name="Report")
            worksheet = writer.sheets["Report"]
            phone_col_idx = list(report.columns).index("Телефон") + 1
            for row_num in range(2, len(report) + 2):
                cell = worksheet.cell(row=row_num, column=phone_col_idx)
                cell.number_format = "@"
                cell.value = str(cell.value)
        print(f"ГОТОВО! Отчёт сохранён: {filename}")
    else:
        print("Нет ни одной валидной записи для обработки.")

    print(f"\nСтатистика: обработано {len(results)}, пропущено (невалидный номер) {skipped}, всего в файле {len(df)}")
    print("--- ЗАВЕРШЕНО ---")


if __name__ == "__main__":
    process_file()
    input("\nНажмите Enter, чтобы закрыть окно...")
