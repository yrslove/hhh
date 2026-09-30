# Doctor Eleks selector guide

Browser automation спочатку пробує semantic locators:

- role + visible name;
- label;
- placeholder;
- visible text;
- CSS тільки як fallback.

Список alias → candidates:

`config/selectors.json`

Наприклад:

```json
"new_interaction": [
  "text=Нова взаємодія",
  "text=Створити нову взаємодію"
]
```

Якщо кнопка в конкретній версії Doctor Eleks називається інакше, достатньо додати варіант.

## Помилка селектора

При невдалому обов'язковому кліку створюється:

```text
logs/errors/YYYYMMDD_HHMMSS_<function>_<step>/
  screen.png
  page.html
```

Це головний матеріал для калібрування.

## Чого не робити

Не переводити основний workflow на координатні `click(x,y)`.

Координати UI нестабільні при:
- іншому DPI;
- іншому масштабі браузера;
- resize;
- зміні DOM;
- появі повідомлення/банера.

Координати кліку користувача залишаються в audit-log тільки як діагностика.
