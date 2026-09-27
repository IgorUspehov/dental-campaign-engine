# Dental Campaign Engine

Веб-приложение для стоматологических клиник: загрузка списка пациентов → персонализация сообщений с просьбой об отзыве → формирование отчёта.

Сейчас работает в **демо-режиме** (сообщения не отправляются, только генерируется Excel-отчёт). Готово к подключению SMS / WhatsApp API.

## Возможности

- Загрузка Excel / CSV / TXT с пациентами
- Автоопределение колонок (Имя, Фамилия, Телефон) на русском, английском и немецком
- Редактируемый шаблон сообщения с плейсхолдером `{name}`
- Нормализация телефонов (RU / DE)
- Скачивание отчёта в Excel
- Современный тёмный UI

## Локальный запуск

```bash
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
uvicorn app:app --reload --host 0.0.0.0 --port 8000
```

Откройте http://localhost:8000

## Деплой на Render.com

### Вариант 1 — через Blueprint (рекомендуется)

1. Залейте этот репозиторий на GitHub
2. На [Render Dashboard](https://dashboard.render.com) → **New** → **Blueprint**
3. Подключите репозиторий — Render подхватит `render.yaml`
4. Нажмите **Apply**

### Вариант 2 — вручную

1. **New** → **Web Service**
2. Подключите GitHub-репозиторий
3. Настройки:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `uvicorn app:app --host 0.0.0.0 --port $PORT`
4. Создайте сервис (Free plan подойдёт)

После деплоя приложение будет доступно по адресу вида `https://dental-campaign-engine.onrender.com`

## Структура

```
├── app.py                 # FastAPI бэкенд
├── requirements.txt
├── render.yaml            # Blueprint для Render
├── templates/index.html   # Фронтенд
├── static/
│   ├── style.css
│   └── app.js
├── send_campaign.py       # Оригинальный консольный скрипт
└── data/ reports/ logs/   # Рабочие папки
```

## Формат файла пациентов

| Имя / Name / Vorname | Фамилия / Surname | Телефон / Phone / Mobil |
|----------------------|-------------------|-------------------------|
| Иван                 | Петров            | +79161234567            |
| Anna                 | Müller            | 01761234567             |

Колонки определяются автоматически по названию.

## Дальнейшее развитие

- Подключение Twilio / WhatsApp Business API
- Сохранение истории кампаний
- Авторизация для клиник
- Планировщик рассылок
