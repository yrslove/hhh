# Doctor Eleks procedure form: live findings

## Observed behavior

The procedure editor is an Angular form with DevExtreme 20.1 widgets. The
patient page contains several `dxForm` instances; the procedure form is the one
containing `input[id$="_isDateOrPeriod"]`. Selecting a referral closes its picker,
then asynchronously navigates to a new `ProcedureRef` route. All subsequent
editor reads must resolve the new form and widget instances.

The old workflow set date fields with `dxDateBox.option('value', new Date(...))`
and staged `formData` directly. Its save requests serialized the dates as a
localized JavaScript `Date.toString()` value. A failed request with a selected
referral also contained the literal string `"null"` for `procedurePeriod` and
received HTTP 500. The frontend then raised `Cannot read properties of
undefined (reading 'data')`, followed by DevExtreme E1035 and stale DOM errors.

`dxForm.updateData`, replacing `formData`, and `beginUpdate/endUpdate` were
observed to fire individual field actions in this installation. The form's
date and division action XML hardcodes `<param name="procedurePeriod">null</param>`.
The period item's own action XML uses `{{procedurePeriod}}`. Direct mutation of
the model triggered saves in the wrong order and is no longer used.

## Manual successful sequence

The operator selected the division, start date and time in the calendar,
`30 хвилин`, then end date and time. The sanitized capture is in
`runtime/procedure-manual-capture.jsonl`.

| Request | Action state | Period | Start | End | HTTP |
| --- | --- | --- | --- | --- | --- |
| 3 | Referral and division selected | `"null"` | Initial timestamp | Initial timestamp | 200 |
| 4 | Start selected | `"null"` | `2026-09-02 18:00` | Initial timestamp | 200 |
| 5 | Duration selected | `"30"` | `2026-09-02T18:00:00` | Initial timestamp | 200 |
| 6 | Final `toValidate=1` | Omitted by Doctor Eleks | `2026-09-02T18:00:00` | `2026-09-02T18:30:00` | 200 |

The manual UI itself sends intermediate saves with a null period. Those saves
are accepted when the other fields retain the application's string format.
The final validation request omits `procedurePeriod`, so the last period
autosave and the final request must both be checked. The successful record
appeared in the patient procedure grid with a 30 minute duration.

## Field interaction map

| Field | Previous workflow | Current adapter | Save behavior |
| --- | --- | --- | --- |
| Electronic referral | Playwright picker click | Same picker, then wait for `ProcedureRef` route and fresh form | May save and rerender |
| Division | Direct `formData` stage or legacy widget setter | Loaded option click, verify editor, `dxForm`, Angular | Item click saves with literal null period |
| Start date/time | `dxDateBox.option('value', Date)` | Calendar day, hour, minute, `OK`; verify all layers | Value change saves with literal null period |
| Period | SelectBox option click | Stock `30 хвилин` item, verify code `"30"` and payload | Item click saves period code |
| End date/time | `dxDateBox.option('value', Date)` | Calendar day, hour, minute, `OK`; verify all layers | Value change may save with literal null period |
| Final action | Button click and popup | Require complete preflight, `toValidate=1` HTTP 2xx, then persisted grid row | Final request omits period |

The adapter stores only serialized snapshots and Playwright `Locator` objects
used within an action. It does not retain DOM nodes, `ElementHandle`,
DevExtreme widgets or `dxForm` instances across saves and rerenders.

## Diagnostic artifacts

- `runtime/procedure-manual-success.json`: successful period autosave.
- `runtime/procedure-manual-final-success.json`: successful final validation.
- `runtime/procedure-manual-diff.json`: failed automation versus manual period save.
- `runtime/procedure-save-capture.jsonl`: sanitized action, form and save logs.

All captures omit Authorization headers, Bearer tokens, KEP passwords and OTP.

## Verified automation run and false corruption signal

On 2026-09-24 the automation completed one real procedure for 2026-09-03
09:30–10:00 with code 96130 and the selected electronic referral. The
successful duration save sent `procedurePeriod="30"`, matching start/end,
and received HTTP 200. The final `toValidate=1` request received HTTP 200.
The patient grid contained the same procedure ID, start/end, 30 minute
duration, code and performer. The sanitized evidence is in
`runtime/procedure-save-capture.jsonl` and `runtime/e2e-procedure-result.json`.

The old corruption check treated every E1035 after a field change as a
broken form. The live capture showed E1035 during initialization and again
after a successful final request while Doctor Eleks closed the form and
returned to the patient list. The adapter now checks fresh editor instances,
attached nodes, loaded options and the current procedure ID before declaring
an open form corrupted. After a confirmed final request, it checks the
persisted grid row instead of expecting the closed editor to remain attached.
HTTP 401 and failed save responses still stop the workflow.

The grid has multiple pages. An existing manual procedure was on page 2;
the original duplicate check inspected only the visible page. The duplicate
check now visits the pager pages and waits for new rows to finish loading.
The manually completed 2026-09-02 18:00–18:30 procedure and the automated
2026-09-03 09:30–10:00 procedure were reconciled against server responses and
persisted grid rows before the corresponding Excel cells AJ287 and AK270
were marked bold in both the imported workbook and its source copy.

## Batch run and pager correction

The 2026-09-25 batch completed all 14 planned procedures. The original
idempotency check visibly switched through every patient procedure page
before **each** new `+` click, and the post-save row check repeated that sweep
after `Підписати пізніше` / `ОК`. This was the source of the long pause between
procedures. The create button itself was clicked once per procedure.

The workflow now records `attempted_cell` before opening a new form. A fresh
cell goes directly to `+`; a retry of an interrupted cell searches all grid
pages for a possible saved record before allowing creation. Legacy stopped
states also receive one conservative search on their first pending cell.
The final action requires the validated `spiprocedureSave` HTTP 2xx response,
the expected record ID and complete payload, and Doctor Eleks returning to
the patient's procedure list. It checks the current page for the exact new
row without turning the pager. When the row sorts onto another page, the
server-confirmed save and closed form are used as completion evidence; the
grid's page order is no longer changed after every save.

After the final planned procedure, the workflow makes one full grid pass and
requires exactly one matching row for every planned date/time/code/performer.
It then marks the workflow complete. On the finished 14-procedure patient,
this read-only batch audit found all 14 rows across six pages and restored
page 1. A retry after an interrupted individual procedure still scans the
list before opening another form. If only the final batch audit is delayed or
fails, rerunning the workflow repeats that audit without creating any more
procedures.

The referral picker stalled once because its data row appeared after the old
10 second deadline. The wait now checks for the actual row and `Вибрати`
control for up to 45 seconds. The existing draft was recovered and saved;
the remaining three records then completed. The last procedure's click trace
shows one trusted click on `Підписати пізніше`, followed by the visible `ОК`
confirmation, final HTTP 200 and a closed procedure form.
