# Airia Skills

Agent Skills served through the Airia MCP Gateway (Skills over MCP, SEP-2640).

## Layout

```
skills/<skill-name>/
├── SKILL.md        # frontmatter + instructions
├── references/     # optional
└── scripts/        # optional
```

## Rules

- `name` in frontmatter must match the folder name.
- Frontmatter keys: `name`, `description`, `license`, `compatibility`, `metadata`, `allowed-tools` only.
- `SKILL.md` must be under 256 KB.
- Only 20 skills per server are security-scanned.
- Set `license: Apache-2.0` in each skill's frontmatter.
- This repo is **public**: no credentials, tenant IDs, or internal-only content.

Invalid skills are skipped silently by the Gateway. Always validate:

```
python3 scripts/validate.py
```

## Connecting to Airia

MCP → Custom Servers → set Transport Type to **Skills Repository**, enter this repo's HTTPS URL, run **Test repository**, then approve under Server Management.

## License

Apache License 2.0. See [LICENSE](LICENSE).
