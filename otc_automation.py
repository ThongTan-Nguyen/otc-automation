import pandas as pd
import requests
import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

# LLM configuration.
MODEL_NAME = "deepseek-v4.1-flash"
API_KEY = ""
INPUT_FOLDER = Path(r"D:\Project\INPUT")
OUTPUT_FOLDER = Path(r"D:\Project\OUTPUT")

# Required Field, Action, Risk for LLM #2
REQUIRED = ["signal_id", "source_key", "company_code", "customer_id", "risk_level",
            "recommended_action", "reason_code", "amount_sgd", "days_overdue",
            "confidence", "owner_team", "created_timestamp", "llm_model", "json_payload"]
ACTIONS = {"No Action", "Monitor", "Collection Reminder", "Priority Collection Follow-up",
           "Request Remittance Review", "Escalate Dispute"}
RISKS = {"Low", "Medium", "High", "Critical"}

# Read Input
def read_workbook(path: Path, names: list[str]) -> dict[str, list[dict[str, str]]]:
    frames = pd.read_excel(path, sheet_name=names, dtype=str, keep_default_na=False, engine="openpyxl")
    return {
        name: frame.rename(columns=lambda column: str(column).strip()).to_dict(orient="records")
        for name, frame in frames.items()
    }

# Trim
def text(row: dict, key: str) -> str:
    return str(row.get(key, "") or "").strip()

# Convert to Float
def number(row: dict, key: str) -> float:
    try:
        return float(text(row, key))
    except ValueError:
        return 0.0

# Check value is not empty
def has_value(value: str) -> bool:
    return bool(value.strip()) and value.strip().lower() not in {"none", "null", "n/a"}

# Handle raw and master data
def prepare(raw: list[dict], master_rows: list[dict]):
    master = {text(row, "KUNNR"): row for row in master_rows if text(row, "KUNNR")}
    working = []
    # Filter
    for row in raw:
        AR_STATUS = text(row, "AR_STATUS")
        BLART = text(row, "BLART")
        DMBTR = number(row, "DMBTR")
        if AR_STATUS.lower() == "closed" or BLART not in {"DR", "DG"} or DMBTR == 0:
            continue
        KUNNR = text(row, "KUNNR")
        BELNR = text(row, "BELNR")
        BUZEI = text(row, "BUZEI")
        GJAHR = text(row, "GJAHR")
        BUKRS = text(row, "BUKRS")
        CUSTOMER_SEGMENT = text(row, "CUSTOMER_SEGMENT")
        REGION = text(row, "REGION")
        # Lookup
        customer_master = master.get(KUNNR, {})
        CREDIT_LIMIT_SGD = number(customer_master, "CREDIT_LIMIT_SGD")
        COLLECTION_OWNER = text(customer_master, "COLLECTION_OWNER")
        RISK_TIER = text(customer_master, "RISK_TIER")
        # source_key
        source_key = f"{BELNR}-{BUZEI}-{GJAHR}"
        # Normalize the risk inputs before deriving boolean workflow flags.
        DAYS_OVERDUE = int(number(row, "DAYS_OVERDUE"))
        RISK_SCORE = number(row, "RISK_SCORE")
        EXCEPTION_REASON = text(row, "EXCEPTION_REASON")
        DISPUTE_STATUS = text(row, "DISPUTE_STATUS")
        RECEIPT_MATCH_STATUS = text(row, "RECEIPT_MATCH_STATUS")
        has_exception_reason = has_value(EXCEPTION_REASON)
        has_dispute_status = has_value(DISPUTE_STATUS)
        has_receipt_exception = RECEIPT_MATCH_STATUS.lower() != "matched"
        working.append({
            "source_key": source_key, "company_code": BUKRS, "customer_id": KUNNR,
            "customer_segment": CUSTOMER_SEGMENT, "region": REGION,
            "ar_status": AR_STATUS, "receipt_match_status": RECEIPT_MATCH_STATUS,
            "amount_sgd": round(DMBTR, 2), "days_overdue": DAYS_OVERDUE, "risk_score": RISK_SCORE,
            "is_overdue": DAYS_OVERDUE > 0,
            "is_high_risk": RISK_SCORE >= 70 or DAYS_OVERDUE > 60 or has_exception_reason,
            "has_dispute": has_dispute_status, "has_receipt_exception": has_receipt_exception,
            "credit_limit_sgd": CREDIT_LIMIT_SGD,
            "owner_team": COLLECTION_OWNER or "UNASSIGNED",
            "risk_tier": RISK_TIER or "UNKNOWN",
            "exception_reason": EXCEPTION_REASON})
    return working


def aggregate(rows: list[dict]) -> list[dict]:
    if not rows:
        return []
    # Aggregation 5 Fields Required
    fields = ("company_code", "customer_segment", "region", "ar_status", "receipt_match_status")
    frame = pd.DataFrame(rows)
    # Calculate invoice_count, open_amount_sgd, avg_days_overdue, high_risk_count, dispute_count, exception_count.
    summary = frame.groupby(list(fields), as_index=False, dropna=False, sort=False).agg(
        invoice_count=("source_key", "size"),
        open_amount_sgd=("amount_sgd", "sum"),
        avg_days_overdue=("days_overdue", "mean"),
        high_risk_count=("is_high_risk", "sum"),
        dispute_count=("has_dispute", "sum"),
        exception_count=("has_receipt_exception", "sum"),
    )
    summary[["open_amount_sgd", "avg_days_overdue"]] = summary[
        ["open_amount_sgd", "avg_days_overdue"]].round(2)
    return summary.sort_values("open_amount_sgd", ascending=False, kind="stable").to_dict(orient="records")


def api_call(endpoint: str, api_key: str, model: str, prompt: str, max_tokens: int) -> tuple[str, dict]:
    body = {"model": model, "messages": [{"role": "user", "content": prompt}],
            "stream": False, "temperature": 0.1, "max_tokens": max_tokens,
            "thinking": {"type": "disabled"}}
    try:
        response = requests.post(endpoint, json=body, timeout=120, headers={
            "Authorization": f"Bearer {api_key}"})
        if not response.ok:
            raise RuntimeError(f"LLM API returned HTTP {response.status_code}: {response.text[:1000]}")
        payload = response.json()
    except requests.RequestException as exc:
        raise RuntimeError(f"LLM API request failed: {exc}") from exc
    message = payload["choices"][0]["message"]
    content = message.get("content") or ""
    return content, payload


def parse_json_array(content: str) -> list[dict]:
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", content.strip(), flags=re.I)
    value = json.loads(cleaned)
    if not isinstance(value, list):
        raise ValueError("LLM #2 output root must be a JSON array")
    return value


def parse_json_object(content: str) -> dict:
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", content.strip(), flags=re.I)
    try:
        value = json.loads(cleaned)
    except json.JSONDecodeError:
        start, end = cleaned.find("{"), cleaned.rfind("}")
        if start < 0 or end <= start:
            raise ValueError("LLM #1 response contains no JSON object")
        value = json.loads(cleaned[start:end + 1])
    if not isinstance(value, dict):
        raise ValueError("LLM #1 output root must be a JSON object")
    return value


def validate_signals(signals: list[dict], valid_keys: set[str]) -> None:
    errors = []
    for index, signal in enumerate(signals):
        missing = [field for field in REQUIRED if field not in signal]
        if missing: errors.append(f"row {index}: missing {missing}")
        if signal.get("risk_level") not in RISKS: errors.append(f"row {index}: invalid risk_level")
        if signal.get("recommended_action") not in ACTIONS: errors.append(f"row {index}: invalid action")
        if signal.get("source_key") not in valid_keys: errors.append(f"row {index}: unknown source_key")
        try:
            if not 0 <= float(signal.get("confidence")) <= 1: errors.append(f"row {index}: confidence out of range")
        except (TypeError, ValueError): errors.append(f"row {index}: invalid confidence")
    if errors:
        raise ValueError("; ".join(errors[:20]))


def llm_prompts(summary: list[dict], llm2_records: list[dict], run_at: str, model: str):
    # LLM #1
    analytic = """You are an OTC finance analyst.

                    Review the JSON summary generated from SAP-like AR data.

                    Return a valid JSON object only.
                    No markdown.
                    No explanation outside the JSON object.

                    Required schema:
                    {
                      "top_5_risk_observations": [],
                      "top_3_customer_segment_actions": [],
                      "data_quality_issues": [],
                      "executive_summary": ""
                    }

                    Rules:
                    - Use only the provided data.
                    - Do not invent missing facts.
                    - Return exactly 5 risk observations and 3 customer or segment actions.
                    - Keep executive_summary under 80 words.
                    - Focus on overdue invoices, receipt matching exceptions, disputes, and high-risk exposure. \nDATA:\n""" + json.dumps(summary[:10], ensure_ascii=False)
    # LLM #2
    required_fields = "\n".join(f"- {field}" for field in REQUIRED if field != "signal_id")
    allowed_risks = "\n".join(f"- {risk}" for risk in sorted(RISKS))
    allowed_actions = "\n".join(f"- {action}" for action in sorted(ACTIONS))
    decision = f"""You are a decision-signal generator for OTC exception handling.

                    For each input record, output one JSON object only.

                    Return a valid JSON array only.
                    No markdown.
                    No explanation.

                    Required schema:
                    {required_fields}

                    risk_level must be one of:
                    {allowed_risks}

                    Risk logic:
                    - Critical if days_overdue > 90 or risk_score >= 85
                    - High if days_overdue > 60 or exception exists
                    - Medium if days_overdue > 30
                    - Low otherwise

                    recommended_action must be one of:
                    {allowed_actions}

                    Use created_timestamp={run_at} and llm_model={model}. json_payload must contain the JSON decision rationale as a string.
                    INPUT:\n""" + json.dumps(llm2_records, ensure_ascii=False)
    return analytic, decision


def write_csv(path: Path, rows: list[dict]) -> None:
    if not rows: return
    pd.DataFrame(rows).to_csv(path, index=False, encoding="utf-8-sig")


def write_xlsx(path: Path, sheets: list[tuple[str, list[dict]]]) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, rows in sheets:
            frame = pd.DataFrame(rows)
            frame = frame.map(
                lambda value: json.dumps(value, ensure_ascii=False)
                if isinstance(value, (dict, list)) else value)
            frame.to_excel(writer, sheet_name=name, index=False)


def main() -> int:
    # Step 1: Config
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=INPUT_FOLDER / "OTC_AI_Automation_Developer_Test_Dummy_Data.xlsx")
    parser.add_argument("--output", type=Path, default=OUTPUT_FOLDER)
    parser.add_argument("--api-url", default="https://api.b.ai/v1/chat/completions")
    parser.add_argument("--model", default=MODEL_NAME)
    parser.add_argument("--decision-limit", type=int, default=20)
    args = parser.parse_args() 
    args.output.mkdir(parents=True, exist_ok=True)
    # Step 2: Read Input
    sheets = read_workbook(args.input, ["OTC_RAW", "MASTER_CUSTOMER"])
    # Step 3: Handle raw and master data
    working = prepare(sheets["OTC_RAW"], sheets["MASTER_CUSTOMER"])
    # Step 4: Aggregate data for LLM #1
    summary = aggregate(working)
    # Step 5: Filter data for LLM #2
    llm2_records = [row for row in working if row["is_high_risk"] or row["has_dispute"] or row["has_receipt_exception"]]
    llm2_records.sort(key=lambda row: (
        row["days_overdue"] > 90 or row["risk_score"] >= 85,
        row["days_overdue"], row["risk_score"], abs(row["amount_sgd"])), reverse=True)
    selected = llm2_records[:args.decision_limit]; run_at = datetime.now().astimezone().isoformat(timespec="seconds")
    # Step 6: Build independent prompts.
    analytic_prompt, decision_prompt = llm_prompts(summary, selected, run_at, args.model)
    # Step 7: Call LLM #1
    analytic_raw, analytic_envelope = api_call(args.api_url, API_KEY, args.model, analytic_prompt, 5000)
    #Output LLM #1
    (args.output / "llm1_response_envelope.json").write_text(json.dumps(analytic_envelope, ensure_ascii=False, indent=2), encoding="utf-8")
    (args.output / "llm1_raw_response.txt").write_text(analytic_raw, encoding="utf-8")
    analytic = parse_json_object(analytic_raw)
    # Step 8: Call LLM #2
    decision_raw, decision_envelope = api_call(args.api_url, API_KEY, args.model, decision_prompt, 12000)
    #Output LLM #2
    (args.output / "llm2_response_envelope.json").write_text(json.dumps(decision_envelope, ensure_ascii=False, indent=2), encoding="utf-8")
    (args.output / "llm2_raw_response.txt").write_text(decision_raw, encoding="utf-8")
    signals = parse_json_array(decision_raw)
    signal_date = datetime.fromisoformat(run_at).strftime("%Y%m%d")
    # Create unique signal_id
    for sequence, item in enumerate(signals, start=1):
        signal = {"signal_id": f"SIG-{signal_date}-{sequence:06d}", **item}
        signal["json_payload"] = json.dumps(signal, ensure_ascii=False)
        signals[sequence - 1] = signal
    # Step 9: Validate every signal before write output.
    validate_signals(signals, {row["source_key"] for row in selected})
    # Create Audit
    audit = {"raw_rows": len(sheets["OTC_RAW"]), "master_rows": len(sheets["MASTER_CUSTOMER"]),
             "working_rows": len(working), "candidate_rows": len(llm2_records), "signal_rows": len(signals),
             "api_used": True, "run_timestamp": run_at}
    # Step 10: Persist JSON, CSV, Excel, and audit artifacts for review and reruns.
    for name, value in (("decision_signal.json", signals), ("llm1_analysis.json", analytic), ("run_audit.json", audit)):
        (args.output / name).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(args.output / "decision_signal.csv", signals); write_csv(args.output / "otc_summary.csv", summary)
    audit_rows = [{"metric": key, "value": json.dumps(value, ensure_ascii=False) if isinstance(value, list) else value} for key, value in audit.items()]
    insight_rows = [{"section": key, "value": json.dumps(value, ensure_ascii=False) if isinstance(value, list) else value} for key, value in analytic.items()]
    write_xlsx(args.output / "OTC_Automation_API_Result.xlsx", [("otc_summary", summary), ("decision_signal", signals), ("run_audit", audit_rows), ("llm1_analysis", insight_rows)])
    error_file = args.output / "last_error.txt"
    if error_file.exists(): error_file.unlink()
    print(json.dumps(audit, ensure_ascii=False, indent=2)); return 0


if __name__ == "__main__":
    try: raise SystemExit(main())
    except Exception as exc:
        error_path = OUTPUT_FOLDER / "last_error.txt"
        error_path.parent.mkdir(parents=True, exist_ok=True)
        error_path.write_text(f"{type(exc).__name__}: {exc}\n", encoding="utf-8")
        print(f"ERROR: {exc}", file=sys.stderr); raise
