PIDWS Alert Report Automation
Python-based Excel automation system that processes PIDWS alerts, classifies day/night events, groups alerts by Custom Scale, aggregates durations, applies business rules, resolves conflicts, handles operator-specific alerts, generates map links, and produces a cleaned final report.

# ALARM Alerts Automation

## Overview

PIDWS Alerts Automation is a Python-based automation tool for processing pipeline monitoring alerts from Excel files.

The project automatically:

- Separates Day Alerts and Night Alerts
- read multiply sheets and create sheets with custome names
- Merges duplicate alerts
- Calculates total alert duration
- Selects Operator Alerts
- Applies Excel formatting automatically
- Generates a professional PDF report
- Supports multiple Excel date formats

---
Raw Excel
   ↓
Night / Day Classification
   ↓
Night Alert Processing
   ↓
Day Alert Processing
   ↓
Day/Night Conflict Resolution
   ↓
Operator Alert Processing
   ↓
Map Location Links
   ↓
Custom Scale Sorting
   ↓
FINAL Excel Report

## Features

- Excel Automation using openpyxl
- Automatic Day/Night alerts detection
- Duplicate alert values merging
- Duration calculation by unique vlaues
- Operator alert (custom values )selection
- Mixed date format support example (24 hours and AM/PM)
- Automatic Excel formatting

---

## Technologies Used for this project 

- Python
- openpyxl
- pandas
- reportlab

---

## Project Structure

ALARM-ALERTS/

├── scripts/

├── sample_input/

├── sample_output/

├── requirements.txt

├── README.md

├── LICENSE

└── .gitignore

---

## Installation
Install the required libraries:
```bash
pip install -r requirements.txt
```

## Usage
Run the main script:
```bash
python scripts/daily_report.py
```
---

## Sample Input

The repository contains sample Excel files for testing.

---

## Output

The script generates:

- Processed Excel Report

---

## License

MIT License
