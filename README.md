# Bushan Data Ops Skill

This package turns Bushan's daily reporting workflow into a reusable Skill and file bundle for GitHub, migration, and future reuse.

## Included workflow entrypoints

1. Local full workflow
   Download raw files, classify them into local date/platform folders, upload to Feishu, and continue into daily reporting.
2. Feishu button workflow
   Click one Feishu interactive card, scan newly uploaded files, and continue the report automatically.
3. Rerun workflow
   Reprocess one historical date after source fixes, rule updates, or workbook/Base mismatches.

## Directory guide

- `SKILL.md`: the skill entry instructions.
- `scripts/core/`: parsing, workbook, Base-sync, and send logic.
- `scripts/entrypoints/`: day runners.
- `scripts/automation/`: Feishu button listener and trigger sender.
- `references/`: business logic and operating runbooks.
- `config/`: sanitized configuration templates.
- `examples/`: sample input and output artifacts.
- `assets/`: static artifacts such as trigger cards and workbook templates.

## Before publishing

1. Replace real webhook URLs, base tokens, operator open IDs, and folder tokens with placeholders.
2. Replace absolute local paths with environment-variable-driven or relative paths.
3. Add sanitized workbook templates into `assets/templates/`.
4. Zip the skill folder after validation.

## Suggested GitHub layout

```text
repo-root/
├── skill/
│   └── bushan-data-ops/
├── docs/
├── CHANGELOG.md
└── README.md
```

Use `skill/bushan-data-ops/` as the distributable folder.
