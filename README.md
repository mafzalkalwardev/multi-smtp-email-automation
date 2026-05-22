\# Multi SMTP Email Automation



A Python-based multi-account SMTP email automation system designed for bulk email dispatching using multiple sender accounts simultaneously.



The system distributes recipients evenly across sender accounts, launches parallel email sender terminals, supports retries, logging, randomized templates, delays, and Excel/CSV recipient management.



\## Features



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

