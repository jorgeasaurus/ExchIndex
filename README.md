# Exch Index — Exchange & Office PowerShell Cmdlet Reference

A fast, static reference site for Exchange and Office PowerShell cmdlets — covering Exchange Online, Teams, Skype for Business, Whiteboard, Office Web Apps, and SharePoint Migration Tool.

Browse, search, and explore every cmdlet with instant filtering, syntax highlighting, parameter tables, copy-ready examples, and multiple visual themes — all in a single HTML file with no runtime dependencies.

---

## Live Site

Deploy the `public/` directory to any static host (GitHub Pages, Azure Static Web Apps, Netlify, etc.).

With the included GitHub Actions workflow, data is refreshed daily from the official [office-docs-powershell](https://github.com/MicrosoftDocs/office-docs-powershell) repository.

---

## Themes

| Theme | File | Description |
|---|---|---|
| Acrylic | `index.html` | Glassmorphism, blue accent — default |
| Cyberpunk | `cyberpunk.html` | Neon cyan/magenta, scanlines |
| CRT | `crt.html` | Green phosphor terminal |
| Synthwave | `synthwave.html` | Purple/pink retro gradient |
| Blueprint | `blueprint.html` | Technical blueprint grid |
| Solarized | `solarized.html` | Classic solarized dark/light |
| Geocities | `geocities.html` | 90s web nostalgia |
| ISE | `ise.html` | PowerShell ISE style |
| Nord | `nord.html` | Nord color palette |

---

## Using Locally

No build step required — just open in a browser via a local server (required because of `fetch()` calls):

```bash
cd public
python -m http.server 8080
# Then open http://localhost:8080
```

---

## Features

- **Instant search** with fuzzy matching across cmdlet names, modules, descriptions
- **Filter** by module, category, and verb
- **Sort** by name, module, verb, or category
- **Lazy loading** in chunks of 50 for performance
- **Expandable cards** with syntax, parameters table, examples, and related cmdlets
- **Copy button** on every example
- **Export** results as JSON or CSV
- **Presets** saved to `localStorage`
- **URL hash state** — bookmarkable/shareable filter states
- **Keyboard shortcuts**: `/` to search, `j`/`k` to navigate, `Enter` to expand, `Esc` to clear
- **Light/dark mode** toggle (persisted)
- **Alpha navigation** sidebar
- **MS Learn links** on every cmdlet name

---

## Workloads

| Workload | Module | Cmdlets |
|---|---|---|
| Exchange Online | ExchangePowerShell | ~1415 |
| Microsoft Teams | MicrosoftTeams | ~613 |
| Skype for Business | SkypeForBusiness | ~898 |
| Office Web Apps | officewebapps | ~18 |
| Whiteboard | WhiteboardAdmin | ~14 |
| SharePoint Migration | Microsoft.SharePoint.MigrationTool.PowerShell | ~8 |

Exchange cmdlets are sub-categorized into Mailbox Management, Transport Rules, Threat Protection, Client Access, Groups & Distribution, Permissions & Roles, Migration, eDiscovery, and more.

---

## Data Pipeline

```
MicrosoftDocs/office-docs-powershell (GitHub)
        |  (cloned daily via GitHub Actions)
scripts/parse_docs.py
        |
public/data/manifest.json          <- cmdlet list + metadata
public/data/descriptions.json      <- one-line descriptions
public/data/modules/{Module}.json  <- syntax + examples per module
        |
GitHub Pages  ->  index.html / *.html
```

Module versions are fetched live from the PowerShell Gallery API during each parse run.

### Generating data locally

```bash
# Clone the docs repo
git clone --depth 1 https://github.com/MicrosoftDocs/office-docs-powershell.git /tmp/office-docs

# Run the parser (Python 3, stdlib only)
python scripts/parse_docs.py /tmp/office-docs
```

---

## License

MIT
