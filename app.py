"""
Dental Campaign Engine — веб-версия
Загрузка списка пациентов → персонализация сообщений → отчёт
"""

import os
import re
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.requests import Request

# ──────────────────────────────────────────────
# Пути
# ──────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
REPORTS_DIR = BASE_DIR / "reports"
LOGS_DIR = BASE_DIR / "logs"
UPLOADS_DIR = BASE_DIR / "uploads"

for d in (DATA_DIR, REPORTS_DIR, LOGS_DIR, UPLOADS_DIR):
    d.mkdir(exist_ok=True)

# ──────────────────────────────────────────────
# Настройки по умолчанию
# ──────────────────────────────────────────────
DEFAULT_TEMPLATE = """Здравствуйте, {name}!
Спасибо, что посетили нашу клинику!
Нам очень важно знать ваше мнение. Если у вас есть минутка, пожалуйста, оставьте отзыв о вашем посещении в Google Maps. Это поможет другим пациентам сделать выбор и поможет нам становиться ещё лучше.
Оставить отзыв:
https://g.page/ВАША_ССЫЛКА/review"""

NAME_COLUMNS = ["Name", "Имя", "Vorname", "Patient", "ФИО", "First Name", "first_name"]
SURNAME_COLUMNS = ["Nachname", "Фамилия", "Surname", "Last Name", "last_name", "Family"]
PHONE_COLUMNS = [
    "Phone", "Телефон", "Mobil", "Handy", "Telefon",
    "Telefon mobil", "Mobile", "phone", "mobile_phone", "Тел"
]

# ──────────────────────────────────────────────
# Приложение
# ──────────────────────────────────────────────
app = FastAPI(
    title="Dental Campaign Engine",
    description="Система персонализированных рассылок для стоматологических клиник",
    version="1.0.0",
)

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


# ──────────────────────────────────────────────
# Утилиты (логика из оригинального скрипта)
# ──────────────────────────────────────────────
def find_column(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    lower_map = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower().strip() in lower_map:
            return lower_map[cand.lower().strip()]
    return None


def build_full_name(row, name_col, surname_col) -> str:
    first = str(row.get(name_col, "")).strip() if name_col else ""
    last = str(row.get(surname_col, "")).strip() if surname_col else ""
    if first.lower() in ("nan", "none", ""):
        first = ""
    if last.lower() in ("nan", "none", ""):
        last = ""
    full = " ".join(part for part in [first, last] if part)
    return full if full else "Пациент"


def clean_phone(num) -> str:
    if pd.isna(num):
        return ""
    num = str(num)
    num = re.sub(r"[^0-9+]", "", num)
    if num.startswith("8") and len(num) >= 10:
        return "+" + num
    if num.startswith("7") and len(num) >= 10:
        return "+" + num
    if num.startswith("0") and len(num) >= 10:
        return "+49" + num[1:]
    return num


def is_valid_phone(phone: str) -> bool:
    digits = re.sub(r"[^0-9]", "", phone)
    return len(digits) >= 8


def load_patients_file(file_path: Path) -> pd.DataFrame:
    ext = file_path.suffix.lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(file_path, dtype=str)
    if ext == ".csv":
        return pd.read_csv(file_path, dtype=str)
    if ext == ".txt":
        try:
            return pd.read_csv(file_path, sep="\t", dtype=str)
        except Exception:
            return pd.read_csv(file_path, sep=",", dtype=str)
    raise ValueError(f"Неподдерживаемый формат файла: {ext}")


def process_patients(
    df: pd.DataFrame,
    message_template: str,
    job_id: str,
) -> dict:
    """Обрабатывает DataFrame и возвращает статистику + путь к отчёту."""
    name_col = find_column(df, NAME_COLUMNS)
    surname_col = find_column(df, SURNAME_COLUMNS)
    phone_col = find_column(df, PHONE_COLUMNS)

    if phone_col is None:
        raise ValueError(
            f"Не найдена колонка с телефоном. "
            f"Колонки в файле: {list(df.columns)}. "
            f"Ожидались: {PHONE_COLUMNS}"
        )

    results = []
    skipped = 0
    log_path = LOGS_DIR / f"activity_{job_id}.log"

    with open(log_path, "a", encoding="utf-8") as log_file:
        for _, row in df.iterrows():
            full_name = build_full_name(row, name_col, surname_col)
            raw_phone = row.get(phone_col, "")
            phone = clean_phone(raw_phone)

            if not is_valid_phone(phone):
                skipped += 1
                log_file.write(
                    f"{datetime.now()}: ПРОПУЩЕН (невалидный номер) — {raw_phone}\n"
                )
                continue

            try:
                personalized = message_template.format(name=full_name)
            except KeyError:
                personalized = message_template.replace("{name}", full_name)

            # Demo-режим (реальная отправка SMS/WA подключается отдельно)
            sms_status = "Sent (Demo)"
            wa_status = "Sent (Demo)"

            log_file.write(f"{datetime.now()}: Обработан {phone} ({full_name})\n")

            results.append({
                "Пациент": full_name,
                "Телефон": phone,
                "SMS": sms_status,
                "WhatsApp": wa_status,
                "Текст сообщения": personalized,
            })

    report_filename = f"report_{job_id}.xlsx"
    report_path = REPORTS_DIR / report_filename

    if results:
        report = pd.DataFrame(results)
        with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
            report.to_excel(writer, index=False, sheet_name="Report")
            ws = writer.sheets["Report"]
            phone_idx = list(report.columns).index("Телефон") + 1
            for row_num in range(2, len(report) + 2):
                cell = ws.cell(row=row_num, column=phone_idx)
                cell.number_format = "@"
                cell.value = str(cell.value)

    return {
        "job_id": job_id,
        "total": len(df),
        "processed": len(results),
        "skipped": skipped,
        "report_file": report_filename if results else None,
        "name_col": name_col or "(не найдена)",
        "surname_col": surname_col or "(не найдена)",
        "phone_col": phone_col,
        "preview": results[:5] if results else [],
    }


# ──────────────────────────────────────────────
# Роуты
# ──────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(
        request,
        "index.html",
        {"default_template": DEFAULT_TEMPLATE},
    )


@app.post("/api/process")
async def process_campaign(
    file: UploadFile = File(...),
    message_template: str = Form(DEFAULT_TEMPLATE),
):
    """Загрузка файла и обработка кампании."""
    if not file.filename:
        raise HTTPException(400, "Файл не выбран")

    ext = Path(file.filename).suffix.lower()
    if ext not in (".xlsx", ".xls", ".csv", ".txt"):
        raise HTTPException(
            400,
            "Поддерживаются только файлы: .xlsx, .xls, .csv, .txt",
        )

    job_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6]
    upload_path = UPLOADS_DIR / f"{job_id}{ext}"

    try:
        content = await file.read()
        with open(upload_path, "wb") as f:
            f.write(content)

        df = load_patients_file(upload_path)
        if df.empty:
            raise HTTPException(400, "Файл пустой или не содержит данных")

        result = process_patients(df, message_template.strip(), job_id)
        return result

    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(500, f"Ошибка обработки: {e}")
    finally:
        # Можно оставить файл для отладки, либо удалять
        pass


@app.get("/api/download/{filename}")
async def download_report(filename: str):
    """Скачать отчёт."""
    # Защита от path traversal
    safe_name = Path(filename).name
    if not safe_name.startswith("report_") or not safe_name.endswith(".xlsx"):
        raise HTTPException(400, "Недопустимое имя файла")

    path = REPORTS_DIR / safe_name
    if not path.exists():
        raise HTTPException(404, "Отчёт не найден")

    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=safe_name,
    )


@app.get("/og-image-de.jpg", include_in_schema=False)
async def og_image_de():
    """Картинка для превью ссылки (Open Graph)."""
    return FileResponse(BASE_DIR / "static" / "og-image-de.jpg", media_type="image/jpeg")


@app.get("/health")
async def health():
    return {"status": "ok", "service": "Dental Campaign Engine"}


# ──────────────────────────────────────────────
# Запуск локально: uvicorn app:app --reload --host 0.0.0.0 --port 8000
# ──────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("app:app", host="0.0.0.0", port=port, reload=False)
