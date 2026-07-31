# GitHub Publish Guide

## Sanitization checklist

Before publishing:

- replace real webhook URLs
- replace real Base tokens
- replace real Feishu folder tokens
- replace real operator open IDs
- replace absolute local paths
- remove personal logs and state files
- remove customer-sensitive data from examples

## Versioning

Recommended version tags:

- `v0.1.0`: first packaged skill
- `v0.2.0`: workflow-level changes
- `v0.2.1`: bug fixes only

## Recommended repository structure

```text
repo-root/
├── skill/
│   └── bushan-data-ops/
├── docs/
├── CHANGELOG.md
└── README.md
```

## Packaging

Validate and zip the skill after sanitization.

Suggested command:

```bash
python3 /Users/jixukuangdian/.agents/skills/skill-creator/scripts/package_skill.py /path/to/bushan-data-ops
```
