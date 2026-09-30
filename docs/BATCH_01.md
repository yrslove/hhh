# Батч 1 — технічна схема

## Потік

`Excel upload -> workbook inspection -> sheet selection -> patient import -> validation -> duplicate scan`

## Принципи

- вихідний Excel read-only;
- локальна копія в `data/imports`;
- поточний стан у SQLite;
- повторний імпорт того самого filename+sheet є snapshot replacement;
- заголовки мають пріоритет над буквами колонок;
- кожна функція повертає структурований `ActionResult`;
- кожний запуск функції має `action.start` / `action.finish` або `action.error`;
- UI-кліки мають view, елемент, функцію (коли відома) та координати;
- персональні деталі результатів не дублюються у технічний action-log.

## Модулі

- `system_check.py` — готовність локального середовища;
- `upload_workbook.py` — локальна робоча копія;
- `inspect_workbook.py` — структура книги;
- `import_patients.py` — snapshot-імпорт у SQLite;
- `validate_patients.py` — первинна перевірка;
- `find_duplicates.py` — дублікати;
- `common.py` — спільні чисті утиліти першого батчу.
