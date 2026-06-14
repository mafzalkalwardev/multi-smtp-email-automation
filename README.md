<div align="center">

# 🚀 Multi SMTP Email Automation

**A Python-based multi-account SMTP email automation system with parallel sender launching, recipient distribution, retries, logging, and randomized email templates.**

Documented · MIT licensed · Maintained

[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)](LICENSE)
[![PRs](https://img.shields.io/badge/PRs-welcome-brightgreen?style=for-the-badge)](CONTRIBUTING.md)

</div>

---

## 🐍 Contribution graph

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/mafzalkalwardev/multi-smtp-email-automation/output/snake-dark.svg" />
  <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/mafzalkalwardev/multi-smtp-email-automation/output/snake.svg" />
  <img alt="Contribution snake" src="https://raw.githubusercontent.com/mafzalkalwardev/multi-smtp-email-automation/output/snake.svg" />
</picture>

---

\# Multi SMTP Email Automation

A Python-based multi-account SMTP email automation system designed for bulk email dispatching using multiple sender accounts simultaneously.

The system distributes recipients evenly across sender accounts, launches parallel email sender terminals, supports retries, logging, randomized templates, delays, and Excel/CSV recipient management.

\## Screenshots

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
