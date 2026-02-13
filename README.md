# Quiz Grader (Lightweight)

A minimal local grading tool for quiz scans.

## Motivation

This project exists because there are many scanned quiz PDFs to grade, and Gradescope does not support the massive upload workflow needed here.

## What it does

- Loads roster from `roster.csv` (`Role == Student` only)
- Loads submissions from `Quiz1/*.pdf` (filename stem = netid)
- Flags unmatched PDFs for manual mapping
- Auto-assigns 0 for missing submissions
- Supports rubric deductions from a default full score
- Saves grading progress locally
- Exports final grades to CSV
- Includes a basic embedded PDF viewer in the app

## Setup

```bash
conda activate env
pip install -r requirements.txt
```

## Run

```bash
python3 quiz_grader_app.py
```

## Output files

- `grading_state.json`: saved grading progress/state
- `grades_export.csv`: exported grades
