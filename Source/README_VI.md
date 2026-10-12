## Overview

This prototype implements an end-to-end Order-to-Cash automation workflow using the provided Excel dataset. It loads and transforms OTC records, enriches them with customer master data, calls two LLM workflows, validates the decision signals, and writes the results back to Excel, JSON, and CSV.

## Architecture

```text
OTC_RAW + MASTER_CUSTOMER
          |
          v
Filter, join, and derive risk flags
          |
          +--> Aggregated summary --> LLM #1 --> OTC insights
          |
          +--> High-risk/exception records --> LLM #2 --> Decision signals
                                                       |
                                                       v
                                      Validate and write output files
```

LLM #1 analyzes grouped OTC exposure. LLM #2 generates structured decision signals for records where `is_high_risk`, `has_dispute`, or `has_receipt_exception` is true.

## Setup

Requirements:

- Python 3.10+
- `pandas`
- `openpyxl`
- `requests`

Install dependencies:

```powershell
pip install pandas openpyxl requests
```

Update the configuration at the top of `otc_automation.py`:

```python
MODEL_NAME = "your-model"
API_KEY = "your-api-key"
INPUT_FOLDER = Path(r"D:\Project\INPUT")
OUTPUT_FOLDER = Path(r"D:\Project\OUTPUT")
```

## Run

```powershell
py "D:\Project\Source\otc_automation.py"
```

Optional overrides:

```powershell
python "D:\Project\Source\otc_automation.py" `
  --input "D:\data\input.xlsx" `
  --output "D:\data\output" `
  --model "your-model" `
  --decision-limit 20
```

## Transformation Logic

- Keep records where `AR_STATUS <> Closed`, `BLART IN (DR, DG)`, and `DMBTR <> 0`.
- Create `source_key` as `BELNR-BUZEI-GJAHR`.
- Derive overdue, high-risk, dispute, and receipt-exception flags.
- Join `MASTER_CUSTOMER` by `KUNNR` to obtain credit limit, collection owner, and risk tier.
- Aggregate by company code, customer segment, region, AR status, and receipt-match status.
- Send the highest-priority exception records to LLM #2; the default demo limit is 20.

## Output

The output folder contains:

- `OTC_Automation_API_Result.xlsx`
- `decision_signal.json` and `decision_signal.csv`
- `otc_summary.csv`
- `llm1_analysis.json`
- `run_audit.json`
- Raw LLM responses and provider response envelopes

The Excel workbook contains `otc_summary`, `decision_signal`, `run_audit`, and `llm1_analysis` sheets.

## Validation

Before write-back, the program checks required fields, allowed risk levels and actions, valid source keys, and confidence values between 0 and 1. Invalid JSON or schema violations stop the run.
