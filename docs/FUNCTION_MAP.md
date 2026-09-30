# Карта функцій

## 1. Процедури

`procedures_create`
- джерело: ЕТ/ФТ + R;
- Doctor Eleks: Направлення → Процедури;
- результат: процедури `Підписати пізніше`;
- side effect: клітинка графіка bold.

`procedures_batch_sign`
- Пакетне опрацювання → Процедура → пацієнт → КЕП.

## 2. IMPORTANT_000

`important_000`
- один parse ФРМ + один parse ЕТ/ФТ;
- запускає дві незалежні перевірки.

`check_frm_before_etft`
- тільки годинні ФРМ;
- помилка: `ET/FT start < FRM end`.

`check_etft_overlaps`
- `patient → date → appointments`;
- interval intersection.

## 3. Пацієнт / документи / діагнози

`coordinator_to_graph`
- координатори → C/E/F/I.

`desktop_neurologist_report`
- Doctor Eleks Desktop → Документи/Файли пацієнта → висновок невролога.

`diagnosis_prepare`
- структурує висновок;
- не затверджує клінічний діагноз автоматично;
- після ручного підтвердження пише коди у N.

## 4. Перша взаємодія

`diagnostic_interaction`
- Діагностика 30 хв;
- ТОВ 2;
- Ургентний;
- A67-02-22;
- основний діагноз;
- D;
- погашення e-referral;
- КЕП.

`rehab_referral_from_diagnostic`
- нестаціонарна реабілітація;
- Медична реабілітація;
- взяти в роботу;
- КЕП.

`rehab_interaction`
- Реабілітація 60 хв;
- Плановий;
- newest own referral;
- AR1/2 template;
- N reconciliation;
- 96 0 22 00;
- МКФ;
- R;
- КЕП;
- O = E PV Reab.

`rehab_referral_after_interaction`
- друге направлення на нестаціонарну реабілітацію.

`care_plan_draft`
- тільки на Rehab interaction;
- Зберегти;
- НЕ створити заявку;
- координатори = Можемо кидати.

## 5. План лікування

`care_plan_submit`
- очікує Кидайте;
- редагує yellow draft;
- Створити заявку;
- КЕП.

`primary_prescriptions`
- auth method;
- SMS OTP;
- максимум 2 запити;
- 10 хв;
- R codes;
- AR2 420 хв;
- 1 місяць;
- Підтримка стану;
- Q = E + призначення.

`referrals_from_prescriptions`
- Дії → Створити направлення;
- не брати в роботу;
- P = повний;
- Q = E + направлення.
