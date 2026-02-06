# Outlook Suite Plugin

Claude Code plugin package for Outlook automation on Windows.

## Included Skills

- `outlook-suite`
- `outlook-mail`
- `outlook-calendar`
- `outlook-contacts`
- `outlook-tasks`
- `outlook-notes`

Each skill is under `skills/<skill-name>/` and includes `SKILL.md` plus scripts where applicable.

## Structure

```
plugins/outlook-suite/
├── .claude-plugin/
│   └── plugin.json
└── skills/
    ├── outlook-suite/
    ├── outlook-mail/
    ├── outlook-calendar/
    ├── outlook-contacts/
    ├── outlook-tasks/
    └── outlook-notes/
```

## Install via Marketplace

From a repo that contains `.claude-plugin/marketplace.json` at the root:

```bash
/plugin marketplace add <owner>/<repo>
/plugin install outlook-suite@outlook-suite-marketplace
```

## Manage Installed Plugin

```bash
/plugin marketplace update outlook-suite-marketplace
/plugin update outlook-suite
/plugin uninstall outlook-suite
```

## Requirements

- Windows
- Microsoft Outlook desktop installed
- PowerShell 5.1+
- Claude Code with plugin support
