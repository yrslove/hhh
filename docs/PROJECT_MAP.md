# Project map

```text
app/
  main.py                    HTTP API + static UI
  core/
    database.py              SQLite schema / state
    audit.py                 audit trail
    function_registry.py     ID → handler
    ui_layout.py             rename / notes / groups / order
    text.py                  exact-name normalization
  excel/
    schedule.py              Appointment parser + IMPORTANT_000 algorithms
    workbook.py              Graph 1 / coordinators / backups / status writes
  automation/
    browser.py               Edge/Chrome CDP + Playwright semantic actions
    desktop.py               Doctor Eleks Desktop UI Automation
  workflows/
    procedures.py
    schedule_checks.py
    intake.py
    first_interaction.py
    care_plan.py
  web/
    index.html
    static/css/app.css
    static/js/app.js

config/
  functions.json             user-facing function definitions
  function_dependencies.json dependency graph for review
  selectors.json             Doctor Eleks semantic selector candidates
  settings.json              local paths / URL / specialists
  ui_layout.json             custom grouping / rename / 300-char notes
  diagnosis_rules.txt        approved diagnostic formula slot

scripts/
  launcher.py
  check_links.py

tests/
  test_schedule_checks.py
```

## Where to change what

- Button wording/description/parameters: `config/functions.json`
- Doctor Eleks button/field renamed: `config/selectors.json`
- Business workflow sequence: `app/workflows/*.py`
- Excel schedule interpretation: `app/excel/schedule.py`
- Excel columns/status side effects: `app/excel/workbook.py`
- UI layout/appearance: `app/web/*`
- Dependencies/roadmap: `config/function_dependencies.json`, `docs/*`
