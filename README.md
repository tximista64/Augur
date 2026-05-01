# Cyberwatch — Automated Daily Intelligence from RSS Feeds

![image](assets/analyst.png)

## What is this?

A curated list of cybersecurity RSS/Atom feeds organized by theme, paired with `cyber_watch.py` — a script that pulls today's articles and pipes them to [Claude Code](https://claude.ai/code) to generate a structured daily threat intelligence digest, ready for Obsidian.

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

## Quick start

**Dependencies:**
```bash
pip install feedparser
# Claude Code required: https://claude.ai/code
```

**Run with local feed list:**
```bash
python cyber_watch.py --feeds urls.md | claude
```

**Run directly from this repo (always up to date):**
```bash
python cyber_watch.py --feeds-url https://raw.githubusercontent.com/tximista64/cybersecurity_rss_feed/main/urls.md | claude
```

**Export to markdown:**
```bash
python cyber_watch.py --feeds urls.md --scored | claude > veille_$(date +%Y%m%d).md
```

---

## Options

| Flag | Description | Default |
|---|---|---|
| `--feeds FILE` | Local feed list (mutually exclusive with `--feeds-url`) | — |
| `--feeds-url URL` | Remote feed list URL | — |
| `--scored` | Group by theme + assign criticality level | off |
| `--max-age H` | Only fetch articles newer than H hours | 48 |
| `--profile` | `pentest` / `blueteam` / `geopo` / `all` | `all` |

**Profile filter examples:**
```bash
# SOC / blue team focus
python cyber_watch.py --feeds urls.md --profile blueteam --scored | claude

# Red team & bug bounty
python cyber_watch.py --feeds urls.md --profile pentest | claude

# Geopolitics & strategic intel
python cyber_watch.py --feeds urls.md --profile geopo | claude
```

---

## Contributing

Feel free to share your favorite feeds by opening an issue or pull request — especially if you know great resources in Chinese, Farsi or Hebrew. Let's keep this list growing together!
