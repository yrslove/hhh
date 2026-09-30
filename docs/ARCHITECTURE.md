# Архітектура

```text
run.bat
  ↓
FastAPI / 127.0.0.1
  ↓
Web UI
  ↓
FunctionRegistry
  ├─ workflows/procedures.py
  ├─ workflows/schedule_checks.py
  ├─ workflows/intake.py
  ├─ workflows/first_interaction.py
  └─ workflows/care_plan.py
       ↓
  ┌──────────────┬──────────────────┬────────────────┐
  │ Excel layer  │ Browser adapter  │ Desktop adapter│
  │ openpyxl     │ Playwright/CDP   │ pywinauto/UIA  │
  └──────────────┴──────────────────┴────────────────┘
       ↓
  SQLite + audit logs + backups
```

## Принцип

Workflow не повинен припускати, що користувач уже відкрив правильну сторінку.

Перед дією код:
1. підключається/запускає потрібне середовище;
2. шукає пацієнта;
3. переходить до потрібного розділу;
4. перевіряє precondition;
5. виконує крок;
6. логуює дію;
7. зберігає checkpoint.

## Стабільні ID

Назви у UI можна змінювати. Системний ID залишається незмінним:

- `procedures_create`
- `procedures_batch_sign`
- `important_000`
- `check_frm_before_etft`
- `check_etft_overlaps`
- `coordinator_to_graph`
- `desktop_neurologist_report`
- `diagnosis_prepare`
- `diagnostic_interaction`
- `rehab_referral_from_diagnostic`
- `rehab_interaction`
- `rehab_referral_after_interaction`
- `care_plan_draft`
- `care_plan_submit`
- `primary_prescriptions`
- `referrals_from_prescriptions`

## Відновлення після збою

`workflow_states` у SQLite зберігає progress.

Приклад процедур:
- completed Excel cells;
- specialist;
- total count.

Приклад OTP:
- час заявки;
- кількість запитів;
- перелік кодів.

Приклад КЕП:
- `awaiting_kep` / `safe_mode_final`.

Це зменшує ризик дублювання після перезапуску.
