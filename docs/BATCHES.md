# 5 робочих батчів

## Батч 1 — Процедури
Перший для фактичного тестування на Doctor Eleks:
1. `procedures_create`
2. `procedures_batch_sign`

## Батч 2 — IMPORTANT_000
1. `important_000`
2. `check_frm_before_etft`
3. `check_etft_overlaps`

Цей батч не залежить від Doctor Eleks і може бути повністю протестований на копії Excel.

## Батч 3 — Пацієнт, документи, діагнози
1. `coordinator_to_graph`
2. `desktop_neurologist_report`
3. `diagnosis_prepare`

## Батч 4 — Перша взаємодія
1. `diagnostic_interaction`
2. `rehab_referral_from_diagnostic`
3. `rehab_interaction`
4. `rehab_referral_after_interaction`
5. `care_plan_draft`

## Батч 5 — План лікування
1. `care_plan_submit`
2. `primary_prescriptions`
3. `referrals_from_prescriptions`

UI-групи можна перейменовувати й переставляти функції, але системні ID і залежності не змінюються.
