# 📊 BudgetSheet — Django Web Spreadsheet

Веб-приложение на Django, которое отображает финансовую таблицу с автоматическим пересчётом формул при изменении значений ячеек.

## ✨ Возможности

- **Интерактивная таблица** — идентична листу Excel, все формулы работают
- **Автопересчёт** — изменение одной ячейки мгновенно пересчитывает все зависимые формулы
- **Импорт Excel** — загрузите свой .xlsx файл через веб-интерфейс
- **История изменений** — полный лог всех изменений с датой/временем
- **Граф зависимостей** — в БД хранится граф зависимостей ячеек для эффективного BFS-пересчёта
- **Поддержка формул**: SUM, AVERAGE, IF, IFERROR, ROUND, MIN, MAX, COUNT, AND, OR, MOD и др.

---

## 🚀 Быстрый старт

### 1. Установка зависимостей

```bash
cd budget_app
pip install -r requirements.txt
```

### 2. Настройка базы данных

**SQLite (по умолчанию, без настройки):**
```bash
mkdir static
python manage.py makemigrations
python manage.py migrate
```

**PostgreSQL (рекомендуется для продакшена):**

В `budget_app/settings.py` замените секцию `DATABASES`:
```python
DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.postgresql',
        'NAME': 'budget_db',
        'USER': 'your_user',
        'PASSWORD': 'your_password',
        'HOST': 'localhost',
        'PORT': '5432',
    }
}
```
Затем:
```bash
createdb budget_db
python manage.py migrate
```

### 3. Запуск с демо-данными

```bash
mkdir static                   # Создать папку для статических файлов
python manage.py makemigrations # Создать файлы миграций
python manage.py migrate        # Применить миграции к БД
python manage.py seed_demo      # Создаёт демо таблицу бюджета
python manage.py runserver
```

Откройте: **http://127.0.0.1:8000**

---

## 📥 Импорт вашего Excel файла

### Через веб-интерфейс
1. Откройте приложение в браузере
2. Нажмите **«Импорт Excel»** в верхней панели
3. Загрузите ваш `.xlsx` файл

### Через Python код
```python
from spreadsheet.importer import import_excel

sheet = import_excel(
    'path/to/your/file.xlsx',
    sheet_index=0,        # Индекс листа (0 = первый)
    sheet_name='Бюджет'   # Название для БД
)
print(f"Импортировано {sheet.cells.count()} ячеек")
```

### Через management command
```bash
python manage.py shell -c "
from spreadsheet.importer import import_excel
import_excel('your_file.xlsx', sheet_index=0, sheet_name='Мой бюджет')
"
```

---

## 🏗️ Архитектура

```
budget_app/
├── budget_app/
│   ├── settings.py          # Конфигурация Django
│   └── urls.py              # Главный роутинг
└── spreadsheet/
    ├── models.py            # Модели: Sheet, Cell, CellDependency, ChangeHistory
    ├── formula_engine.py    # Движок формул (Excel→Python транспилер + eval)
    ├── importer.py          # Импортёр .xlsx файлов
    ├── demo_data.py         # Демо данные бюджета
    ├── views.py             # Django views + REST API
    ├── urls.py              # URL patterns
    └── templates/
        └── spreadsheet/
            └── index.html   # Основной UI
```

### Модели БД

| Модель | Описание |
|--------|----------|
| `Sheet` | Лист таблицы |
| `Cell` | Ячейка: координата, значение, формула (Excel и Python), тип, форматирование |
| `CellDependency` | Граф зависимостей: source→target для BFS-пересчёта |
| `ChangeHistory` | Лог изменений ячеек |

### Как работает пересчёт

1. Пользователь редактирует ячейку → POST `/api/sheet/<id>/update/`
2. Новое значение сохраняется в БД
3. BFS по графу `CellDependency` определяет все зависимые формульные ячейки
4. Формулы вычисляются в правильном порядке через `eval()` с безопасным контекстом
5. Обновлённые значения сохраняются в БД и возвращаются JSON-ответом
6. JavaScript обновляет DOM без перезагрузки страницы + анимация изменённых ячеек

---

## 🔌 API

| Метод | URL | Описание |
|-------|-----|----------|
| GET | `/` | Главная страница |
| GET | `/sheet/<id>/` | Просмотр конкретного листа |
| GET | `/api/sheet/<id>/data/` | Данные листа в JSON |
| POST | `/api/sheet/<id>/update/` | Обновить ячейку |
| GET | `/api/sheet/<id>/cell/<row>/<col>/history/` | История ячейки |
| POST | `/api/import/` | Импорт Excel файла |
| GET | `/seed/` | Загрузить демо данные |

### Пример: обновление ячейки
```json
POST /api/sheet/1/update/
{
  "row": 3,
  "col": 1,
  "value": "1500000"
}

Response:
{
  "success": true,
  "edited_cell": {"row": 3, "col": 1, "value": "1500000", "formatted": "1 500 000.00 ₸"},
  "updates": [
    {"row": 6, "col": 1, "value": "2450000", "formatted": "2 450 000.00 ₸"},
    {"row": 15, "col": 1, "value": "950000", "formatted": "950 000.00 ₸"}
  ],
  "updates_count": 5
}
```

---

## ⚙️ Поддерживаемые Excel формулы

`SUM`, `AVERAGE`, `MIN`, `MAX`, `IF`, `IFERROR`, `ROUND`, `ABS`, `COUNT`, `COUNTA`,
`AND`, `OR`, `NOT`, `INT`, `MOD`, `CONCATENATE`, `TEXT`, `LEFT`, `RIGHT`, `MID`,
`ISBLANK`, `ISERROR`, `LEN`, `SQRT`, `POWER`, `LOG`, `EXP`, `PI`, `TODAY`, `NOW`

Диапазоны: `A1:B10`, Вложенные формулы: `=IF(SUM(A1:A5)>1000, "OK", "LOW")`

---

## 🛡️ Безопасность

- Формулы выполняются через `eval()` с пустым `__builtins__` — только наш безопасный контекст
- CSRF защита на всех POST запросах
- Валидация числовых значений перед сохранением

---

## 🗄️ Для продакшена

```bash
# .env файл
SECRET_KEY=your-secret-key-here
DEBUG=False
DATABASE_URL=postgresql://user:pass@localhost/budget_db

# Запуск
python manage.py collectstatic
gunicorn budget_app.wsgi:application --bind 0.0.0.0:8000
```