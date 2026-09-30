# Roadmap для наступного проходу Codex

## Мета

Не переписувати програму з нуля. Спочатку перевірити зв'язки модулів та фактичні селектори на робочому Doctor Eleks.

## Порядок

1. `app/core/function_registry.py`
   - усі 16 системних ID повинні мати handler.

2. `config/functions.json`
   - звірити опис функцій з бізнес-процесом.

3. `app/excel/schedule.py`
   - прогнати на копії реального Графік 1;
   - перевірити date header detection;
   - перевірити theme/indexed/rgb fill colors;
   - перевірити повторні time grids усередині ЕТ/ФТ.

4. `procedures_create`
   - перша end-to-end функція для Doctor Eleks Web;
   - safe_mode=true;
   - перевірити patient search;
   - referral row;
   - create procedure;
   - start datetime;
   - duration;
   - sign later;
   - Excel bold.

5. Після процедур:
   - `procedures_batch_sign`.

6. Потім:
   - diagnostic_interaction;
   - rehab_interaction;
   - care plan functions.

## Під час перевірки кожного Doctor Eleks кроку

Якщо selector не знайдений:
1. відкрити `logs/errors/.../screen.png`;
2. відкрити `page.html`;
3. знайти стабільний label/role/text;
4. додати його в `config/selectors.json`;
5. не змінювати workflow, якщо змінилася тільки назва елемента.

## Desktop

`app/automation/desktop.py` навмисно використовує UI Automation, а не OCR.

Якщо висновок не знаходиться:
- audit містить дерево control type / name / automation_id;
- замінити загальний пошук на точні AutomationId конкретної інсталяції.

## Після калібрування

- safe_mode=false тільки для перевірених функцій;
- окремий commit після кожного робочого workflow;
- тести schedule не ламати.
