<div align="center">

# Multi-SMTP Email Automation

<img src="https://readme-typing-svg.demolab.com?font=Inter&weight=700&size=28&duration=2800&pause=700&color=EA4335&center=true&vCenter=true&width=900&lines=Multi-Account+SMTP+Automation;Parallel+Sender+Launching;Retry+%2B+Logging+%2B+Templates;Python+outreach+engine" alt="Typing SVG" />

<p>
  <img src="https://komarev.com/ghpvc/?username=mafzalkalwardev&label=Project%20Views&color=ea4335&style=for-the-badge" alt="Project views" />
  <img src="https://img.shields.io/github/stars/mafzalkalwardev/multi-smtp-email-automation?style=for-the-badge&color=22c55e" alt="GitHub stars" />
  <img src="https://img.shields.io/github/forks/mafzalkalwardev/multi-smtp-email-automation?style=for-the-badge&color=f59e0b" alt="GitHub forks" />
  <img src="https://img.shields.io/github/license/mafzalkalwardev/multi-smtp-email-automation?style=for-the-badge&color=64748b" alt="License" />
</p>

<p>
  <img src="https://skillicons.dev/icons?i=python" alt="Skill icons" />
</p>

<p>
  <strong>Python multi-account SMTP automation with parallel senders, recipient distribution, retries, logging, and randomized templates.</strong>
</p>

</div>

---

## Contribution Snake

<div align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/mafzalkalwardev/google-voice-dispatch-agent/output/github-contribution-grid-snake-dark.svg">
    <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/mafzalkalwardev/google-voice-dispatch-agent/output/github-contribution-grid-snake.svg">
    <img alt="Contribution snake" src="https://raw.githubusercontent.com/mafzalkalwardev/google-voice-dispatch-agent/output/github-contribution-grid-snake.svg">
  </picture>
</div>

---

## SEO Keywords

`SMTP automation` `multi account email` `Python outreach`

---

## Features

\- Multi-account SMTP email sending

\- Parallel sender launching

\- Recipient chunk distribution

\- Gmail App Password support

\- Excel and CSV support

\- Retry failed emails

\- Randomized subject/body templates

\- Delay handling to reduce spam detection

\- Email validation

\- Logging system

\- Automatic sender warm-up

\- Separate sender terminals

\- Bulk email automation

\## Tech Stack

\- Python

\- Pandas

\- SMTP

\- OpenPyXL

\- CSV

\- Threading

\- Subprocess Automation

\## Project Structure

```text

multi-smtp-email-automation/

│

├── email\_launcher.py

├── email\_sender.py

├── recipients.xlsx

├── senders.csv

├── run.cmd

├── README.md

└── .gitignore

```

\## File Details

\### recipients.xlsx

Contains recipient information.

Required columns:

| Email | Name | State |

|------|------|-------|

\---

\### senders.csv

Contains sender accounts.

Required columns:

| Email | AppPassword | Name |

|------|-------------|------|

\---

\## Installation

Install required packages:

```bash

pip install pandas openpyxl

```

\## How to Run

Run launcher:

```bash

python email\_launcher.py

```

OR

```bash

run.cmd

```

\## Features Overview

\### Multi Sender Distribution

Recipients are automatically divided equally among all sender accounts.

\### Randomized Templates

Subjects and email bodies are randomized to reduce spam detection risk.

\### Logging System

All sent emails are logged inside:

```text

sent\_log.csv

```

\### Retry System

Failed emails automatically retry multiple times.

\### Warm-Up Feature

Each sender account sends a warm-up email before bulk sending.

\## Security Note

Do NOT upload:

\- Gmail app passwords

\- recipient lists

\- sender credentials

Keep sensitive files private or use environment variables.

\## Author

Muhammad Afzal Kalwar

GitHub:

@mafzalkalwardev
