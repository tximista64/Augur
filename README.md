# Augur — AI-Powered Threat Intelligence Briefing

![image](assets/analyst.png)

## What is this?

A curated collection of cybersecurity RSS/Atom feeds paired with `augur.py` — an AI agent that fetches today's articles, analyzes them via the Claude API, extracts IOCs, and generates a ready-to-present **PowerPoint briefing** in one command.

Feeds are grouped into sections:

| Section | Content |
|---|---|
| **Blueteam** | DFIR, malware analysis, threat intel, CVEs |
| **Offensive** | Red team, exploit research, offensive tooling |
| **Bug Bounty / Web** | Bug bounty writeups, web security research |
| **Persons of Interest** | Individual researchers & practitioners |
| **Podcasts** | Cyber & geopolitics podcasts |
| **CTF** | Capture The Flag platforms & walkthroughs |
| **Conf** | Black Hat, DEF CON, CCC and more |
| **Osint & Geopo** | OSINT, geopolitics, strategic intelligence |

Most feeds are in **English**, but you'll also find many in **French**, and some in **Spanish**, **Hebrew**, and **German**.

---

## Output

A dark-themed PowerPoint briefing (`briefing_YYYYMMDD.pptx`) with:

| Slide | Content |
|---|---|
| 1 | Title — date & stats |
| 2 | Executive Summary — top 3 threats |
| 3 | CVE & Vulnerabilities |
| 4 | APT & Threat Actors |
| 5 | Offensive IT / Red Team |
| 6 | Ransomware & Cybercrime |
| 7 | Geopolitics & Intel |
| 8 | IOCs du jour — regex + Claude extraction |
| 9 | Stats & Sources |

Each item is scored: 🔴 Critical / 🟠 High / 🟡 Medium / 🟢 Info

---

## Quick start

**Dependencies:**
```bash
pip install feedparser anthropic python-pptx
```

**API key** — add to `~/.zshrc`:
```bash
export ANTHROPIC_API_KEY=sk-ant-...
```

**Run:**
```bash
# Local feed list
python augur.py --feeds urls.md

# Always up-to-date from this repo
python augur.py --feeds-url https://raw.githubusercontent.com/tximista64/Augur/main/urls.md

# Custom output
python augur.py --feeds urls.md --output briefing.pptx
```

---

## Options

| Flag | Description | Default |
|---|---|---|
| `--feeds FILE` | Local feed list (mutually exclusive with `--feeds-url`) | — |
| `--feeds-url URL` | Remote feed list URL | — |
| `--max-age H` | Only fetch articles newer than H hours | 48 |
| `--profile` | `pentest` / `blueteam` / `geopo` / `all` | `all` |
| `--output FILE` | Output `.pptx` path | `briefing_YYYYMMDD.pptx` |

**Profile examples:**
```bash
python augur.py --feeds urls.md --profile blueteam
python augur.py --feeds urls.md --profile pentest
python augur.py --feeds urls.md --profile geopo
```

---

## Contributing

Feel free to share your favorite feeds by opening an issue or pull request — especially if you know great resources in Chinese, Farsi or Hebrew. Let's keep this list growing together!
