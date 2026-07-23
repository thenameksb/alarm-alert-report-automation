# ALARM Alerts Automation

## Overview

PIDWS Alerts Automation is a Python-based automation tool for processing pipeline monitoring alerts from Excel files.

The project automatically:

- Separates Day Alerts and Night Alerts
- eead multiply sheets and create sheets with custome names
- Merges duplicate alerts
- Calculates total alert duration
- Selects Operator Alerts
- Applies Excel formatting automatically
- Generates a professional PDF report
- Supports multiple Excel date formats

---

## Features

- Excel Automation using openpyxl
- Automatic Day/Night alert detection
- Duplicate alert merging
- Duration calculation
- Operator alert selection
- PDF report generation
- Mixed date format support
- Automatic Excel formatting

---

## Technologies Used

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

---

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
- PDF Report

---

## License

MIT License
