# Troubleshooting

## Feishu send limit

Symptom:

- Feishu returns `code=11232`
- message contains `frequency limited`

Action:

- retry with backoff
- do not regenerate data; only resend the card payload

## Existing date rerun did not update workbook

Symptom:

- input JSON changed
- workbook rows or month totals stayed old

Action:

- ensure reruns rewrite workbook rows for that date
- recompute month progress after the rewrite
- sync Base afterward

## Youzan cumulative refund source incomplete

Symptom:

- process stops with a message that the cumulative refund source is incomplete

Action:

- export a refund-complete file covering month start through the report date
- rerun the target date only after that file is in place

## Workbook corruption

Symptom:

- openpyxl throws zip or CRC errors

Action:

- restore the latest healthy backup
- rerun affected days one by one, not in parallel
- keep a copy of the corrupt workbook for forensics if needed

## File-shape mismatch

Symptom:

- parser skips a platform or raises missing-column errors

Action:

- inspect the new file name and column shape
- update parsing rules and reference docs together
- do not silently remap fields without documenting the new rule
