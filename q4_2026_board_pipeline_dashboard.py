"""Q4 2026 board pipeline dashboard. Paste into one Google Colab code cell,
or upload this .py file and run: %run q4_2026_board_pipeline_dashboard.py

Dependencies: numpy, pandas, openpyxl, plotly (normally preinstalled in Colab).
If Plotly is missing, run a separate cell: !pip install -q "plotly>=5.24,<7"
Plotly JavaScript is embedded in the HTML for offline interaction and SVG export.
No API key, chart CDN or automatic download. The workbook is read, never edited.
Local use: python q4_2026_board_pipeline_dashboard.py workbook.xlsx --output board.html
"""
from __future__ import annotations
import argparse
import base64
import hashlib
import html
import io
import json
import re
from pathlib import Path
import sys
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

# 1. EDITABLE MODELLING ASSUMPTIONS. These are choices, not observed outcomes.
CONFIG = {
    "simulations": 20000,
    "random_seed": 20261003,
    "historical_blend": 0.50,   # p(win) = (1-blend)*assigned + blend*historic
    "slippage_delay_days": 30, # calendar days; shift the delivery schedule
    "stale_after_days": 30,    # measured against the Salesforce snapshot
    "fx_shock_pct": 10,
    "score_weights": {"concentration": 0.25, "slippage": 0.25,
                      "stale": 0.20, "inflation": 0.20, "count": 0.10},
    "classification": "Course project; supplied Salesforce-style dataset",
}
Q_START = np.datetime64("2026-10-01", "D")
Q_END = np.datetime64("2027-01-01", "D")  # exclusive
REQUIRED_COLUMNS = {
    "A": "opportunity_id", "D": "opportunity_name", "G": "service_line",
    "H": "region", "K": "opportunity_owner", "L": "stage",
    "M": "probability_pct", "N": "historical_stage_win_rate_pct",
    "W": "slippage_risk_pct", "AA": "currency", "AJ": "expected_start_date",
    "AK": "delivery_duration_months", "AN": "last_updated_date", "AO": "snapshot_date",
    "AP": "report_match_check", "AQ": "pipeline_relevance", "AR": "numbered_stage",
    "AS": "relevant_contract_value_eur", "AV": "Q4-2026",
    "BI": "probability_Q4-2026", "BV": "weighted_Q4-2026", "CI": "confidence_var_Q4-2026",
}
CHECK_COLUMNS = ["AP", "BE", "BR", "CE", "CR"]


def section_heading(text):
    """Green notebook headings explain the next step when run in Colab."""
    if "google.colab" in sys.modules:
        from IPython.display import HTML, display
        display(HTML('<h3 style="color:#167044;font-family:Arial">' + html.escape(text) + '</h3>'))
    else:
        print(text)


def workday_fraction(start, months, delay=0):
    """Exactly mirrors EDATE + NETWORKDAYS: start inclusive, end exclusive.
    A delay shifts both schedule endpoints by calendar days; duration unchanged.
    Holidays are deliberately excluded, as in the source workbook.
    """
    start = pd.Timestamp(start).normalize()
    end = start + pd.DateOffset(months=int(months))
    a = np.datetime64((start + pd.Timedelta(days=int(delay))).date(), "D")
    b = np.datetime64((end + pd.Timedelta(days=int(delay))).date(), "D")
    total = int(np.busday_count(a, b))
    if total <= 0:
        raise ValueError("Delivery schedule has no Monday-Friday workdays.")
    lo, hi = max(a, Q_START), min(b, Q_END)
    return (int(np.busday_count(lo, hi)) / total) if hi > lo else 0.0


def load_pipeline(source):
    """Detect the Analysis header, skip the three summary rows, validate caches.
    openpyxl does not evaluate formulas. A formula without a saved numeric value
    causes a useful error instead of silently becoming zero.
    """
    raw = Path(source).read_bytes() if isinstance(source, (str, Path)) else source
    wb = load_workbook(io.BytesIO(raw), data_only=True)
    if "Analysis" not in wb.sheetnames:
        raise ValueError("Please upload the analysed Salesforce workbook containing an 'Analysis' worksheet.")
    sheet = wb["Analysis"]
    header_row = next((r for r in range(1, min(30, sheet.max_row) + 1)
                       if str(sheet.cell(r, 1).value).strip() == "opportunity_id"), None)
    if not header_row:
        raise ValueError("Could not find opportunity_id in column A of Analysis.")
    for letter, name in REQUIRED_COLUMNS.items():
        got = sheet.cell(header_row, column_index_from_string(letter)).value
        if got != name:
            raise ValueError(f"Analysis!{letter}{header_row}: expected '{name}', found '{got}'. Use the supplied column layout.")
    headers = [c.value for c in sheet[header_row]]
    rows, excel_rows = [], []
    for row in sheet.iter_rows(min_row=header_row + 1, values_only=False):
        if row[0].value is None:
            continue
        rows.append([c.value for c in row]); excel_rows.append(row[0].row)
    df = pd.DataFrame(rows, columns=headers)
    if df.empty:
        raise ValueError("Analysis contains no opportunities.")
    df["source_row"] = excel_rows
    df["opportunity_id"] = df["opportunity_id"].astype(str).str.strip()
    if df["opportunity_id"].eq("").any() or df["opportunity_id"].duplicated().any():
        raise ValueError("Blank or duplicate opportunity IDs found. Each ID must appear exactly once.")
    for letter in CHECK_COLUMNS:
        name = headers[column_index_from_string(letter) - 1]
        bad = ~df[name].astype(str).str.upper().eq("PASS")
        if bad.any():
            raise ValueError(f"{name} failed or has no cached result at rows {df.loc[bad, 'source_row'].head(8).tolist()}. Recalculate and save in Excel/LibreOffice first.")
    aliases = {"Q4-2026": "gross", "probability_Q4-2026": "probability",
               "weighted_Q4-2026": "standard", "confidence_var_Q4-2026": "sales_confidence"}
    df = df.rename(columns=aliases)
    for col in aliases.values():
        values = pd.to_numeric(df[col], errors="coerce")
        if values.isna().any() or not np.isfinite(values).all():
            raise ValueError(f"Missing, invalid or uncached formula results in {col}. Open the workbook, recalculate, save, and upload again.")
        df[col] = values.astype(float)
    if (df[["gross", "probability", "standard"]] < -0.01).any().any():
        raise ValueError("Negative Q4 pipeline amounts found; resolve credit/adjustment treatment before modelling.")
    if not np.allclose(df.sales_confidence, df.probability - df.standard, rtol=1e-8, atol=0.01):
        raise ValueError("CI does not reconcile to BI minus BV by opportunity.")
    open_mask = (~df.stage.astype(str).str.contains("Closed", case=False, na=False)
                 & df.pipeline_relevance.astype(str).str.lower().eq("yes"))
    if (df.loc[~open_mask, ["gross", "probability", "standard"]].abs() > 0.01).any().any():
        raise ValueError("Closed/non-relevant records have non-zero Q4 values. Resolve the workbook's scope check first.")
    q = df.loc[open_mask & (df.gross > 0)].copy().reset_index(drop=True)
    if q.empty:
        raise ValueError("No open opportunities have a positive Q4-2026 workday allocation.")
    for c in ["probability_pct", "historical_stage_win_rate_pct", "slippage_risk_pct"]:
        q[c] = pd.to_numeric(q[c], errors="coerce")
        if q[c].isna().any() or not q[c].between(0, 1).all():
            raise ValueError(f"{c} must be present as a decimal between 0 and 1 (e.g. 0.65), not 65.")
    for c in ["expected_start_date", "last_updated_date", "snapshot_date"]:
        q[c] = pd.to_datetime(q[c], errors="coerce")
        if q[c].isna().any():
            raise ValueError(f"Missing or invalid {c} in Q4 opportunities.")
    if q.snapshot_date.dt.normalize().nunique() != 1:
        raise ValueError("Q4 records contain multiple snapshot dates; use one consistent Salesforce snapshot.")
    snapshot = q.snapshot_date.iloc[0].normalize()
    q["stale_days"] = (snapshot - q.last_updated_date.dt.normalize()).dt.days
    if (q.stale_days < 0).any():
        raise ValueError("A last updated date is later than the snapshot date.")
    q["delivery_duration_months"] = pd.to_numeric(q.delivery_duration_months, errors="coerce")
    if (q.delivery_duration_months.isna().any() or (q.delivery_duration_months <= 0).any()
            or (q.delivery_duration_months % 1 != 0).any()):
        raise ValueError("Delivery durations must be positive whole months.")
    contract = pd.to_numeric(q.relevant_contract_value_eur, errors="coerce").to_numpy(float)
    f0 = np.array([workday_fraction(r.expected_start_date, r.delivery_duration_months) for r in q.itertuples()])
    if not np.allclose(q.gross, contract * f0, rtol=1e-7, atol=0.02):
        raise ValueError("AV does not match the supplied Monday-Friday delivery schedule. Resolve workbook allocation before simulating.")
    if not np.allclose(q.probability, q.gross * q.probability_pct, rtol=1e-7, atol=0.02):
        raise ValueError("BI does not reconcile to AV multiplied by the assigned probability in M.")
    if (q.standard > q.gross + 0.01).any():
        raise ValueError("Standard weighted amounts exceed gross Q4 pipeline.")
    for c in ["service_line", "region", "opportunity_owner", "currency", "stage"]:
        if q[c].isna().any() or q[c].astype(str).str.strip().eq("").any():
            raise ValueError(f"Missing {c} for a Q4 opportunity; complete the source before producing the board pack.")
    q["currency"] = q.currency.astype(str).str.upper().str.strip()
    fx = []
    if "FX" not in wb.sheetnames:
        raise ValueError("The workbook must include its FX worksheet and rate provenance.")
    fxrows = list(wb["FX"].values)
    fxh = next((i for i, r in enumerate(fxrows) if str(r[0]).strip().upper() == "CURRENCY"), None)
    if fxh is None:
        raise ValueError("Cannot locate the FX currency lookup header.")
    for r in fxrows[fxh + 1:]:
        if r[0] is not None and isinstance(r[1], (int, float)):
            if r[1] <= 0:
                raise ValueError("FX multipliers must be positive.")
            fx.append({"currency": str(r[0]).upper(), "multiplier": float(r[1]), "month": str(r[2]), "quote": float(r[3])})
    if not set(q.currency).issubset({r["currency"] for r in fx}):
        raise ValueError("A Q4 currency is missing from the FX rate lookup.")
    # Reconcile the source subtotal row, when present, to all 2,000 detail rows.
    checks = {}
    for letter, key in [("AV", "gross"), ("BI", "probability"), ("BV", "standard"), ("CI", "sales_confidence")]:
        total = float(df[key].sum()); stated = sheet.cell(header_row - 2, column_index_from_string(letter)).value if header_row >= 3 else None
        if isinstance(stated, (int, float)) and not np.isclose(total, stated, rtol=1e-8, atol=0.05):
            raise ValueError(f"{letter} summary total does not match its opportunity detail.")
        checks[key] = total
    meta = {"source_rows": len(df), "excluded_rows": len(df) - len(q), "header_row": header_row,
            "snapshot": snapshot.strftime("%d %B %Y"), "snapshot_iso": snapshot.strftime("%Y-%m-%d"),
            "sha256": hashlib.sha256(raw).hexdigest(), "fx": fx, "source_totals": checks,
            "validated_checks": CHECK_COLUMNS}
    return q, meta


def q4_month_shares(start, months, delay=0):
    """Q4 workday phasing, using the source workbook's delivery schedule.
    Shares sum to one when the schedule intersects Q4, otherwise zero.
    The existing model determines retained Q4 value; this only phases it.
    """
    start = pd.Timestamp(start).normalize()
    end = start + pd.DateOffset(months=int(months))
    shift = pd.Timedelta(days=int(delay))
    a = np.datetime64((start + shift).date(), "D")
    b = np.datetime64((end + shift).date(), "D")
    boundaries = np.array(["2026-10-01", "2026-11-01", "2026-12-01", "2027-01-01"], dtype="datetime64[D]")
    counts = np.array([max(0, int(np.busday_count(max(a, lo), min(b, hi))))
                       if min(b, hi) > max(a, lo) else 0
                       for lo, hi in zip(boundaries[:-1], boundaries[1:])], dtype=float)
    return counts / counts.sum() if counts.sum() else np.zeros(3)


def board_forecast_risks(q, base_phasing, config):
    """Describe existing model risks; do not change model outcomes or draws."""
    forecast = float(q.probability.sum())
    calibrated = q.gross.to_numpy(float) * q.model_win_probability.to_numpy(float)
    slip = q.slippage_risk_pct.to_numpy(float)
    retention = q.slip_retention.to_numpy(float)
    calibration = float(calibrated.sum() - forecast)
    # Mutually exclusive components of the EXISTING delivery-slippage loss.
    full_deferral = -float(np.sum(calibrated * slip * (1-retention) * (retention == 0)))
    partial_timing = -float(np.sum(calibrated * slip * (1-retention) * (retention > 0)))
    model = float(q.model_expected.sum())
    if not np.isclose(forecast+calibration+full_deferral+partial_timing, model, rtol=1e-12, atol=.01):
        raise ValueError("Board risk waterfall does not reconcile to the existing model.")
    order = np.argsort(-q.probability.to_numpy(float), kind="stable")
    ranks = {int(index): rank+1 for rank, index in enumerate(order)}
    records = []
    for i, row in enumerate(q.itertuples()):
        share = float(row.probability / forecast * 100) if forecast else 0
        december = float(row.probability * base_phasing[i, 2])
        december_share = float(base_phasing[i, 2]*100)
        inflation = float((row.probability_pct-row.historical_stage_win_rate_pct)*100)
        tags = []
        if ranks[i] == 1: tags.append("single")
        if ranks[i] <= 5 or share >= 10: tags.append("concentration")
        if row.slippage_risk_pct >= .4: tags.append("slippage")
        if december_share >= 50: tags.append("timing")
        if row.stale_days > config["stale_after_days"]: tags.append("stale")
        if inflation > 10: tags.append("inflation")
        high = (share >= 20 or row.slippage_risk_pct >= .6 or row.stale_days > 60
                or inflation > 20 or (december_share >= 75 and share >= 5))
        severity = 2 if high else 1 if any(t != "single" for t in tags) else 0
        records.append({
            "id": str(row.opportunity_id), "name": str(row.opportunity_name),
            "owner": str(row.opportunity_owner), "rank": ranks[i], "tags": tags,
            "exposure": float(row.probability), "share_pct": share,
            "probability": float(row.probability_pct), "historic": float(row.historical_stage_win_rate_pct),
            "slippage": float(row.slippage_risk_pct), "december": december,
            "december_share_pct": december_share, "days": int(row.stale_days),
            "inflation_pp": inflation, "model": float(row.model_expected),
            "severity": severity, "status": ["Low exposure", "Review", "High risk"][severity]
        })
    records.sort(key=lambda r: r["rank"])
    top = lambda n: sum(r["exposure"] for r in records[:n])
    exposure = sum(r["exposure"] for r in records if "slippage" in r["tags"])
    return {
        "records": records,
        "waterfall": {"weighted": forecast, "calibration": calibration,
                      "full_deferral": full_deferral, "partial_timing": partial_timing,
                      "other": 0.0, "model": model},
        "largest": top(1), "top3": top(3), "top5": top(5),
        "remaining": forecast-top(5),
        "december": sum(r["december"] for r in records),
        "slippage_exposure": exposure,
        "slippage_exposure_pct": exposure/forecast*100 if forecast else 0,
        "optimism_gap": forecast-model,
        "optimism_pct": (forecast-model)/forecast*100 if forecast else 0
    }


def build_model(q, meta, config=None):
    cfg = dict(CONFIG); cfg.update(config or {})
    if not 0 <= cfg["historical_blend"] <= 1 or cfg["slippage_delay_days"] < 0:
        raise ValueError("Use historical_blend between 0 and 1 and a non-negative delay.")
    if not 1000 <= cfg["simulations"] <= 100000:
        raise ValueError("Use 1,000–100,000 simulations.")
    g = q.gross.to_numpy(float); p = q.probability_pct.to_numpy(float)
    h = q.historical_stage_win_rate_pct.to_numpy(float); s = q.slippage_risk_pct.to_numpy(float)
    def retention(delay):
        f0 = np.array([workday_fraction(r.expected_start_date, r.delivery_duration_months) for r in q.itertuples()])
        fd = np.array([workday_fraction(r.expected_start_date, r.delivery_duration_months, delay) for r in q.itertuples()])
        # Slippage is a downside case: no artificial Q4 uplift from calendar shifts.
        return np.minimum(1, fd / f0)
    r = retention(cfg["slippage_delay_days"])
    blend = cfg["historical_blend"]
    pw = (1 - blend) * p + blend * h
    expected_retention = (1 - s) + s * r
    expected = g * pw * expected_retention
    rng = np.random.default_rng(cfg["random_seed"])
    n = int(cfg["simulations"])
    totals = np.zeros(n)
    monthly_samples = np.zeros((n, 3))
    base_phasing = np.array([q4_month_shares(row.expected_start_date, row.delivery_duration_months)
                            for row in q.itertuples()])
    delayed_phasing = np.array([q4_month_shares(row.expected_start_date, row.delivery_duration_months,
                                              cfg["slippage_delay_days"]) for row in q.itertuples()])
    if not np.allclose(base_phasing.sum(axis=1), 1):
        raise ValueError("Monthly phasing does not reconcile to Q4 opportunity allocations.")
    order = (q[["stage", "numbered_stage"]].drop_duplicates("stage")
             .sort_values("numbered_stage").stage.tolist())
    stages = {stage: np.zeros(n) for stage in order}
    # Deal-by-deal arrays avoid an n_simulations × n_opportunities memory allocation.
    for i, row in enumerate(q.itertuples()):
        won = rng.random(n) < pw[i]
        slipped = rng.random(n) < s[i]  # independent of win; used only on won deals
        revenue = won * g[i] * np.where(slipped, r[i], 1.0)
        totals += revenue; stages[row.stage] += revenue
        # Reuse the SAME wins and slips: no new random draws or model changes.
        phasing = np.where(slipped[:, None], delayed_phasing[i], base_phasing[i])
        monthly_samples += revenue[:, None] * phasing
    cumulative_samples = np.cumsum(monthly_samples, axis=1)
    if not np.allclose(cumulative_samples[:, -1], totals, rtol=1e-12, atol=1e-7):
        raise ValueError("Monthly simulated revenue does not reconcile to the existing Q4 model.")
    # Retain identical quarter-end paths, including their floating-point values.
    cumulative_samples[:, -1] = totals
    monthly_expected = np.sum(g[:, None] * pw[:, None] *
                              ((1-s[:, None])*base_phasing + (s*r)[:, None]*delayed_phasing), axis=0)
    executive_monthly = {
        "months": ["October", "November", "December"],
        "month_ends": ["31 October", "30 November", "31 December"],
        "gross": np.sum(g[:, None]*base_phasing, axis=0).tolist(),
        "weighted": np.sum(q.probability.to_numpy(float)[:, None]*base_phasing, axis=0).tolist(),
        "standard": np.sum(q.standard.to_numpy(float)[:, None]*base_phasing, axis=0).tolist(),
        "model_expected": monthly_expected.tolist(),
        "cumulative_samples": cumulative_samples.T.tolist(),
        "method": "Monthly allocation uses Monday-Friday delivery workdays with no holidays. Existing deal-level win/slippage draws are reused. Delayed retained Q4 revenue follows the shifted delivery schedule; quarter-end outcomes are identical to the reliability model."
    }
    quantiles = np.quantile(totals, [0.10, 0.25, 0.50, 0.75, 0.90])
    current = float(q.probability.sum()); gross = float(g.sum()); standard = float(q.standard.sum())
    shares = q.probability.to_numpy(float) / current if current > 0 else g / gross
    hhi = float(np.sum(shares ** 2))
    risks = {"concentration": min(hhi / 0.10, 1),
             "slippage": float(np.sum(shares * s)),
             "stale": float(np.sum(shares * (q.stale_days.to_numpy() > cfg["stale_after_days"]))),
             "inflation": min(float(np.sum(g * np.maximum(p - h, 0)) / gross) / 0.20, 1),
             "count": max(0, 1 - len(q) / 50)}
    score = float(np.clip(100 * (1 - sum(cfg["score_weights"][k] * risks[k] for k in risks)), 0, 100))
    q = q.copy()
    q["model_win_probability"] = pw; q["slip_retention"] = r; q["model_expected"] = expected
    q["shortfall_to_current"] = np.maximum(q.probability - expected, 0)
    q["non_delivery_likelihood"] = 1 - pw * expected_retention
    q["stale"] = q.stale_days > cfg["stale_after_days"]
    q["inflation_pp"] = (p - h) * 100
    def reasons(row):
        out = []
        if row.slippage_risk_pct >= 0.4: out.append("Slippage ≥40%")
        if row.stale: out.append(f"Updated {int(row.stale_days)} days ago")
        if row.inflation_pp > 10: out.append("Assigned probability > historic by 10pp")
        if row.probability / max(current, 1) >= 0.10: out.append("≥10% of forecast")
        return out or ["Validate delivery and next step"]
    actions = []
    for row in q.itertuples():
        rr = reasons(row)
        impact = float(row.shortfall_to_current)
        issue_count = sum([row.slippage_risk_pct >= .4, row.stale, row.inflation_pp > 10,
                           row.probability / max(current, 1) >= .10])
        priority = min(100, 25 * issue_count + 25 * row.non_delivery_likelihood)
        action = []
        if row.slippage_risk_pct >= .4: action.append("Confirm delivery date and unblock dependencies")
        if row.stale: action.append("Refresh the opportunity and evidence")
        if row.inflation_pp > 10: action.append("Challenge probability against stage evidence")
        if row.probability / max(current, 1) >= .10: action.append("Agree sponsor escalation and fallback")
        if not action: action = ["Confirm buyer decision and delivery readiness"]
        actions.append({"id": row.opportunity_id, "name": str(row.opportunity_name),
                        "owner": str(row.opportunity_owner), "stage": str(row.stage),
                        "current": float(row.probability), "model": float(row.model_expected),
                        "impact": impact, "priority": float(priority), "likelihood": float(row.non_delivery_likelihood),
                        "gross": float(row.gross), "reasons": rr, "action": "; ".join(action),
                        "next_step": str(getattr(row, "next_step", "") or "No next step recorded"),
                        "due": "Within 7 days of review"})
    actions.sort(key=lambda a: a["impact"] * (1 + a["priority"] / 100), reverse=True)
    def groups(column):
        t = q.groupby(column, dropna=False)[["gross", "probability", "standard"]].sum().sort_values("probability", ascending=False)
        return [{"label": str(k), **{c: float(row[c]) for c in t.columns}} for k, row in t.iterrows()]
    funnel = groups("stage"); funnel.sort(key=lambda a: order.index(a["label"]))
    # A smoothed histogram approximates each stage density; total simulations include zeros.
    violin = []
    for stage, values in stages.items():
        vmax = max(float(values.max()), 1)
        counts, edges = np.histogram(values, bins=60, range=(0, vmax))
        kernel_x = np.arange(-4, 5); kernel = np.exp(-0.5 * (kernel_x / 1.5) ** 2); kernel /= kernel.sum()
        smooth = np.convolve(counts.astype(float), kernel, mode="same")
        violin.append({"stage": stage, "y": ((edges[:-1] + edges[1:]) / 2).tolist(),
                       "density": (smooth / max(smooth.max(), 1)).tolist(),
                       "p25": float(np.quantile(values, .25)), "p50": float(np.median(values)),
                       "p75": float(np.quantile(values, .75)), "zero_pct": float(np.mean(values == 0) * 100)})
    def analytical(win=pw, slip=s, retain=r):
        return float(np.sum(g * np.clip(win, 0, 1) * ((1 - np.clip(slip, 0, 1)) + np.clip(slip, 0, 1) * retain)))
    baseline = float(expected.sum())
    tornado = [
        {"label": "Assigned probability ±10pp", "a": analytical((1-blend)*np.clip(p-.1,0,1)+blend*h), "b": analytical((1-blend)*np.clip(p+.1,0,1)+blend*h)},
        {"label": "Historic win rate ±10pp", "a": analytical((1-blend)*p+blend*np.clip(h-.1,0,1)), "b": analytical((1-blend)*p+blend*np.clip(h+.1,0,1))},
        {"label": "Slippage probability ±10pp", "a": analytical(slip=s+.1), "b": analytical(slip=s-.1)},
        {"label": "Delay 15 / 45 calendar days", "a": analytical(retain=retention(45)), "b": analytical(retain=retention(15))},
        {"label": "Historic blend 25% / 75%", "a": analytical(.25*p+.75*h), "b": analytical(.75*p+.25*h)},
    ]
    tornado.sort(key=lambda a: abs(a["b"] - a["a"]), reverse=True)
    pareto = [{"id": row.opportunity_id, "name": str(row.opportunity_name), "value": float(row.probability)}
              for row in q.sort_values("probability", ascending=False).itertuples()]
    cfg = json.loads(json.dumps(cfg))
    return {"config": cfg, "meta": meta, "executive_monthly": executive_monthly,
            "forecast_risks": board_forecast_risks(q, base_phasing, cfg),
            "kpis": {"gross": gross, "probability": current, "standard": standard,
                     "gap": float(q.sales_confidence.sum()), "count": len(q),
                     "non_eur_pct": float(q.loc[q.currency.ne("EUR"), "gross"].sum() / gross * 100),
                     "largest_pct": float(q.probability.max() / current * 100) if current > 0 else 0,
                     "score": score, "model_mean": baseline, "simulation_mean": float(totals.mean()),
                     "reach_forecast_pct": float(np.mean(totals >= current) * 100),
                     "stale_pct": risks["stale"] * 100, "slippage_pct": risks["slippage"] * 100,
                     "inflation_pp": float(np.sum(g * (p - h)) / gross * 100),
                     "hhi": hhi, "effective_count": float(1/hhi), "rating": "High" if score >= 80 else "Moderate" if score >= 60 else "Low"},
            "quantiles": dict(zip(["p10", "p25", "p50", "p75", "p90"], map(float, quantiles))),
            "risk_components": risks, "simulated_totals": totals.tolist(),
            "funnel": funnel, "regions": groups("region"), "offerings": groups("service_line"),
            "owners": groups("opportunity_owner"), "currencies": groups("currency"),
            "violin": violin, "tornado": tornado, "pareto": pareto, "actions": actions,
            "deals": [{"id": r.opportunity_id, "row": int(r.source_row), "stage": str(r.stage),
                       "gross": float(r.gross), "probability": float(r.probability), "standard": float(r.standard),
                       "sales_confidence": float(r.sales_confidence), "assigned": float(r.probability_pct),
                       "historic": float(r.historical_stage_win_rate_pct), "slippage": float(r.slippage_risk_pct),
                       "calibrated": float(r.model_win_probability), "retention": float(r.slip_retention),
                       "expected": float(r.model_expected)} for r in q.itertuples()]}

# 4. SELF-CONTAINED DASHBOARD. CSS tokens reproduce the supplied board-pack guide.
HTML_TEMPLATE = r'''<!doctype html><html lang="en-GB"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Fuchsia Works | Q4 2026 board pipeline</title><link href="https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,600;8..60,700&amp;family=IBM+Plex+Sans:wght@400;500;600&amp;display=swap" rel="stylesheet"><style>
:root{
  --fuchsia:#C2124F;      /* primary brand: rules, key figures, median series */
  --fuchsia-deep:#7D0B3F; /* secondary series, emphasis */
  --pink-600:#E0457F;       /* band edge */
  --pink-200:#F8C9DA;       /* band fill */
  --pink-100:#FCE4EC;       /* table header tint, highlight */
  --pink-50:#FFF7FA;        /* page */
  --paper:#FFFFFF;
  --ink:#2A0A1A;            /* body text */
  --ink-2:#5E3A4C;          /* secondary text */
  --ink-3:#8A6B7A;          /* tertiary: axis, footnotes */
  --rule:#E9D3DD;           /* hairline rules */
  --rule-strong:#2A0A1A;
  --favourable:#0B6E8A;--adverse:#B4460A;
  --serif:"Source Serif 4",Georgia,"Times New Roman",serif;
  --sans:"IBM Plex Sans",system-ui,-apple-system,"Segoe UI",Arial,sans-serif;
}
*{box-sizing:border-box}
html{background:var(--pink-50)}
body{margin:0;padding:40px 24px;color:var(--ink);font:400 14px/21px var(--sans);font-variant-numeric:tabular-nums lining-nums;-webkit-font-smoothing:antialiased}
.sheet{max-width:1180px;margin:0 auto;background:var(--paper);border:1px solid var(--rule);padding:0 48px 32px}
:focus-visible{outline:2px solid var(--ink);outline-offset:2px}
p{margin:0}

/* Running header */
.runhead{display:flex;justify-content:space-between;align-items:center;gap:16px;padding:18px 0 14px;border-bottom:4px solid var(--fuchsia);font-size:12px;line-height:16px;color:var(--ink-2)}
.wordmark{font:700 15px/18px var(--serif);color:var(--ink);letter-spacing:.01em}
.wordmark b{color:var(--fuchsia);font-weight:700}
.runhead .doc{text-align:right}
.runhead .doc span{color:var(--fuchsia);font-weight:600}

/* Title block */
.titleblock{padding:28px 0 22px;border-bottom:1px solid var(--rule)}
.kicker{font:600 11px/14px var(--sans);letter-spacing:.12em;text-transform:uppercase;color:var(--fuchsia)}
h1{font:700 30px/37px var(--serif);margin:8px 0 10px;max-width:980px;letter-spacing:-.005em}
h1 em{font-style:normal;color:var(--fuchsia)}
.dek{color:var(--ink-2);font-size:14px}

/* KPI row */
.kpis{display:grid;grid-template-columns:repeat(4,1fr);border-bottom:1px solid var(--rule)}
.kpi{padding:20px 24px 20px 0}
.kpi+.kpi{padding-left:24px;border-left:1px solid var(--rule)}
.kpi .l{font:600 11px/14px var(--sans);letter-spacing:.08em;text-transform:uppercase;color:var(--ink-2)}
.kpi .v{font:600 34px/40px var(--serif);margin:8px 0 6px;color:var(--ink)}
.kpi.lead .v{color:var(--fuchsia)}
.kpi .v.sm{font-size:28px}
.kpi .c{font-size:12px;line-height:17px;color:var(--ink-3)}
.kpi .c b{color:var(--ink);font-weight:600}

/* Section heads */
.sec{padding-top:26px}
.sechead{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;flex-wrap:wrap;padding-bottom:8px;border-bottom:1px solid var(--rule-strong);margin-bottom:14px}
h2{font:700 18px/24px var(--serif);margin:0}
.units{font-size:12px;color:var(--ink-3)}
.sup{font-size:9px;vertical-align:super;line-height:0;color:var(--fuchsia);font-weight:600;margin-left:1px}

.two{display:grid;grid-template-columns:minmax(0,1.75fr) minmax(0,1fr);gap:40px}

/* Scenario control: discreet, board-appropriate */
.scenario{display:flex;align-items:center;gap:10px;font-size:12px;color:var(--ink-2)}
.seg{display:inline-flex;border:1px solid var(--rule-strong);border-radius:2px;overflow:hidden}
.seg button{border:0;border-left:1px solid var(--rule-strong);background:var(--paper);color:var(--ink);font:500 12px/16px var(--sans);padding:4px 10px;cursor:pointer}
.seg button:first-child{border-left:0}
.seg button[aria-pressed="true"]{background:var(--fuchsia);color:#fff}
.seg button:hover:not([aria-pressed="true"]){background:var(--pink-100)}
.scenario input{accent-color:var(--fuchsia);width:96px}
.scenario output{min-width:30px;font-weight:600;color:var(--ink)}

/* Chart */
.legend{display:flex;gap:20px;flex-wrap:wrap;font-size:12px;color:var(--ink-2);margin-bottom:4px}
.legend span{display:inline-flex;align-items:center;gap:8px}
.k{width:20px;border-top:2.5px solid var(--fuchsia)}
.k.e{border-top:1.5px dashed var(--fuchsia-deep)}
.k.b{height:10px;border:1px solid var(--pink-600);background:var(--pink-200)}
#chart{display:block;width:100%;height:auto}
#chart .gl{stroke:var(--rule);stroke-width:1}
#chart .ax{stroke:var(--rule-strong);stroke-width:1}
#chart .yt{fill:var(--ink-3);font:400 11px var(--sans)}
#chart .xt{fill:var(--ink-2);font:500 11px var(--sans)}
#chart .band{fill:var(--pink-200);fill-opacity:.8;stroke:var(--pink-600);stroke-width:.75}
#chart .med{fill:none;stroke:var(--fuchsia);stroke-width:2.5}
#chart .exp{fill:none;stroke:var(--fuchsia-deep);stroke-width:1.5;stroke-dasharray:5 4}
#chart .dot{fill:var(--fuchsia);stroke:#fff;stroke-width:1.5}
#chart .lk{font:500 10px var(--sans);fill:var(--ink-3);letter-spacing:.04em;text-transform:uppercase}
#chart .lv{font:600 12px var(--sans);fill:var(--ink)}
#chart .lv.m{fill:var(--fuchsia)}
#chart .ld{stroke:var(--ink-3);stroke-width:.75}
#chart .ann{font:400 11px var(--sans);fill:var(--ink-2)}
.fignote{font-size:12px;color:var(--ink-3);margin-top:6px}

/* Commentary */
.summary{margin:0;padding:0;list-style:none}
.summary li{padding:0 0 12px 16px;margin-bottom:12px;border-bottom:1px solid var(--rule);position:relative;font-size:13.5px;line-height:20px}
.summary li:last-child{border-bottom:0;margin-bottom:0}
.summary li::before{content:"";position:absolute;left:0;top:8px;width:6px;height:6px;background:var(--fuchsia)}
.summary strong{font-weight:600}
.summary .n{color:var(--fuchsia);font-weight:600}

/* Phasing mini-table */
.phase{margin-top:22px}
.phase h3{font:600 11px/14px var(--sans);letter-spacing:.08em;text-transform:uppercase;color:var(--ink-2);margin:0 0 8px}
.prow{display:grid;grid-template-columns:76px 1fr 64px 36px;align-items:center;gap:10px;font-size:13px;padding:5px 0}
.pbar{height:10px;background:var(--pink-100)}
.pbar i{display:block;height:100%;background:var(--fuchsia)}
.prow .v{text-align:right;font-weight:500}
.prow .p{text-align:right;color:var(--ink-3);font-size:12px}

/* Financial table */
table{width:100%;border-collapse:collapse;font-size:13px}
thead th{font:600 11px/14px var(--sans);letter-spacing:.06em;text-transform:uppercase;color:var(--ink-2);padding:8px 10px;border-bottom:1px solid var(--rule-strong);text-align:right;white-space:nowrap;vertical-align:bottom}
thead th:first-child{text-align:left;padding-left:0}
thead tr.grp th{border-bottom:1px solid var(--rule);color:var(--ink-3);padding-bottom:4px;text-align:center}
thead tr.grp th.blank{border-bottom:0}
tbody td{padding:9px 10px;border-bottom:1px solid var(--rule);text-align:right;white-space:nowrap}
tbody td:first-child{text-align:left;padding-left:0;color:var(--ink)}
tbody td.hl{background:var(--pink-100);color:var(--fuchsia);font-weight:600}
thead th.hl{background:var(--pink-100);color:var(--fuchsia)}
tbody tr.qe td{font-weight:600;border-top:1px solid var(--rule-strong);border-bottom:3px double var(--rule-strong)}
.tablewrap{overflow-x:auto}

/* Notes */
.notes{margin-top:28px;padding-top:12px;border-top:1px solid var(--rule-strong);display:grid;grid-template-columns:1fr auto;gap:24px}
.notes ol{margin:0;padding-left:16px;font-size:11.5px;line-height:17px;color:var(--ink-3)}
.notes li{margin-bottom:3px}
.notes li::marker{color:var(--fuchsia);font-weight:600}
.actions{display:flex;gap:8px;align-items:flex-start}
.btn{border:1px solid var(--rule-strong);border-radius:2px;background:var(--paper);color:var(--ink);padding:5px 12px;font:500 12px/16px var(--sans);cursor:pointer;white-space:nowrap}
.btn:hover{background:var(--pink-100)}
.foot{display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-top:18px;padding-top:10px;border-top:4px solid var(--fuchsia);font-size:11px;color:var(--ink-3)}
.foot b{color:var(--fuchsia);font-weight:600;letter-spacing:.08em;text-transform:uppercase}

@media (max-width:900px){.sheet{padding:0 24px 24px}.two{grid-template-columns:1fr;gap:24px}.kpis{grid-template-columns:1fr 1fr}.kpi:nth-child(3){border-left:0;padding-left:0}.kpi:nth-child(n+3){border-top:1px solid var(--rule)}h1{font-size:25px;line-height:31px}.notes{grid-template-columns:1fr}}
@media (max-width:480px){body{padding:0}.sheet{border:0}.kpis{grid-template-columns:1fr}.kpi+.kpi{border-left:0;padding-left:0;border-top:1px solid var(--rule)}.runhead{flex-direction:column;align-items:flex-start}.runhead .doc{text-align:left}}
@media print{
  @page{size:A4 landscape;margin:12mm}
  html,body{background:#fff;padding:0}.sheet{border:0;max-width:none;padding:0}
  .scenario,.actions{display:none}
  *{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .sec,.kpis{break-inside:avoid}
}
/* Seven requested sections use the board-pack page inside a restrained sidebar. */
body{padding:24px}.layout{display:grid;grid-template-columns:230px minmax(0,1fr);gap:24px;max-width:1530px;margin:auto}.sheet{width:100%;max-width:1280px;padding:0 32px 28px;min-width:0}
nav{position:sticky;top:24px;height:max-content;border-top:4px solid var(--fuchsia);padding:18px 12px;background:#fff;border-bottom:1px solid var(--rule)}nav .wordmark{margin:0 6px 16px}nav button{display:block;width:100%;text-align:left;padding:13px 8px;border:0;border-bottom:1px solid var(--rule);background:#fff;font:500 13px/18px var(--sans);color:var(--ink);cursor:pointer}nav button[aria-selected=true]{background:var(--pink-100);color:var(--fuchsia);border-left:3px solid var(--fuchsia)}nav .scope{margin:20px 6px 0;font-size:11.5px;color:var(--ink-3)}
.panel{display:none}.panel.active{display:block}.kpis{grid-template-columns:repeat(4,minmax(0,1fr))}.kpi{min-width:0;padding:20px 15px!important;border-bottom:1px solid var(--rule)}.kpi:nth-child(4n+1){padding-left:0!important;border-left:0}.kpi .l{min-height:29px}.kpi .v{font-size:30px}.kpi .c{min-height:34px}.status{display:block;margin-top:8px;font:500 11px/16px var(--sans)}.favourable{color:#0B6E8A}.caution{color:#8A6B7A}.concern{color:#B4460A}.neutral{color:var(--ink-3)}
.key{display:flex;gap:18px;flex-wrap:wrap;margin:14px 0;font-size:12px}.key span{white-space:nowrap}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:24px}.wide{grid-column:1/-1}.chart{min-width:0}.chart svg{width:100%;height:auto;display:block;overflow:visible}.chart text{font-family:var(--sans);fill:var(--ink-3);font-size:11px}.chart .label{fill:var(--ink-2);font-size:11px}.chart .value{fill:var(--ink);font-size:12px}.chart .gl{stroke:var(--rule);stroke-width:1}.chart .ax{stroke:var(--ink);stroke-width:1}.chart .band{fill:var(--pink-200);opacity:.55}.chart [tabindex]:focus{outline:none;stroke:var(--ink);stroke-width:2}
.chart-heading{font:700 18px/24px var(--serif);margin:0 0 3px}.chart-unit{font-size:12px;color:var(--ink-3);margin-bottom:12px}.fignote{margin:8px 0 20px}.legend{line-height:20px}.swatch{width:14px;height:8px;display:inline-block}.controls{display:flex;flex-wrap:wrap;align-items:center;gap:12px;margin:12px 0;font-size:12px}select,input,button{font-family:var(--sans)}select{border:1px solid var(--ink);border-radius:2px;padding:4px;background:#fff;color:var(--ink)}input[type=range]{accent-color:var(--fuchsia);width:130px}.quantiles{display:grid;grid-template-columns:repeat(5,1fr);margin:16px 0;border-bottom:1px solid var(--rule)}.quantiles div{padding:10px;border-left:1px solid var(--rule);font-size:12px}.quantiles div:first-child{border-left:0}.quantiles b{display:block;font:600 22px/28px var(--serif)}.quantiles .median{background:var(--pink-100);color:var(--fuchsia)}.metric-note{padding:12px 0;border-bottom:1px solid var(--rule);font-size:12px;color:var(--ink-2)}.summary{margin-top:18px}.summary li{padding-bottom:14px;margin-bottom:14px}.summary li:last-child{border-bottom:1px solid var(--rule)}
.tablewrap{margin:15px 0}table.watch{min-width:1000px}table.watch th,table.watch td{white-space:normal;text-align:left;vertical-align:top}.watch td.number{white-space:nowrap;text-align:right}.watch td:nth-child(1){min-width:180px}.watch td:nth-child(5){min-width:230px}.watch td:nth-child(6){min-width:240px}.small{font-size:11.5px;color:var(--ink-3)}details{margin:14px 0;border-bottom:1px solid var(--rule);padding:10px 0}summary{cursor:pointer;color:var(--fuchsia);font-weight:500}.method{line-height:1.65}.method li{margin:8px 0}#fatal{display:none;color:#B4460A;padding:20px;border:1px solid #B4460A}.foot{margin-top:26px}.print-btn{margin-top:20px}.sr-only{position:absolute;left:-9999px}td{font-variant-numeric:tabular-nums}sup a{color:var(--fuchsia)}
@media(max-width:1100px){.layout{grid-template-columns:190px minmax(0,1fr);gap:12px}.sheet{padding:0 24px 24px}.kpi .v{font-size:26px}.grid{grid-template-columns:1fr}.two{grid-template-columns:1fr}.wide{grid-column:auto}}
@media(max-width:850px){body{padding:12px}.layout{display:block}nav{position:static;display:flex;flex-wrap:wrap;gap:4px;margin-bottom:16px;padding:8px}nav .wordmark{width:100%;margin:4px 8px}nav button{width:auto;font-size:12px;padding:9px}nav .scope{display:none}.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.kpi:nth-child(2n+1){border-left:0;padding-left:0!important}.runhead{gap:10px}.quantiles{grid-template-columns:repeat(5,minmax(0,1fr))}.quantiles b{font-size:18px}}
@media(max-width:480px){.sheet{padding:0 16px 20px}.kpis{grid-template-columns:1fr}.kpi+.kpi{padding-left:0!important}.quantiles{grid-template-columns:repeat(3,1fr)}h1{font-size:24px}.runhead .doc{text-align:left}}
@media print{@page{size:A4 landscape;margin:12mm}body{padding:0}nav,.controls,.print-btn,.actions,details{display:none!important}.layout{display:block}.sheet{border:0;padding:0;max-width:none}.panel{display:block!important;break-before:page}.panel:first-of-type{break-before:auto}.grid{grid-template-columns:1fr 1fr}.wide{grid-column:1/-1}.chart,.sec,.quantiles{break-inside:avoid}.runhead{padding:8px 0}.panel{padding-top:16px}.kpis{grid-template-columns:repeat(4,1fr)}h1{font-size:26px}table.watch{font-size:10px;min-width:0}.watch td:nth-child(n){min-width:0}.chart svg{max-height:95mm}.tablewrap{overflow:visible}.foot{position:static}.fignote{font-size:10px}.sheet{font-size:12px}}
/* Isolate Plotly's SVGs from the generic inline-chart CSS. */
.funnel-wrap{overflow-x:auto}.plotly-funnel{min-width:640px;min-height:680px;width:100%}.funnel-print{display:none;width:100%;height:auto}.plotly-funnel .trace{transition:opacity .35s ease}
@media(prefers-reduced-motion:reduce){.plotly-funnel .trace{transition:none}}
@media print{.funnel-wrap,#funnel-selection,#funnel-error{display:none!important}.funnel-print{display:block!important;max-height:125mm;object-fit:contain}.js-plotly-plot .plotly .modebar{display:none!important}}
</style><script>__PLOTLY_BUNDLE__</script><style>
#owners.owner-stage-panel{position:relative;background:#fff;border:1px solid var(--rule);border-radius:16px;overflow:hidden;min-width:0}
#owners .owner-stage-tools{position:absolute;top:8px;left:14px;right:14px;height:58px;display:flex;align-items:center;justify-content:space-between;gap:12px}
#owner-size-key{display:flex;align-items:center;gap:14px;font:11px var(--sans);color:var(--ink-2)}
#owner-size-key .owner-size-item{display:flex;align-items:center;gap:6px;white-space:nowrap}
#owner-size-key .owner-size-circle{display:inline-block;flex:none;border:1px solid #AA1047;border-radius:50%;background:#F8C9DA}
#owners #owner-reset{border-radius:8px;padding:6px 10px;font-size:11px;white-space:nowrap}
#owner-stage-plot{position:absolute;top:70px;left:0;right:0;bottom:26px}
#owner-stage-status{position:absolute;bottom:4px;left:14px;right:14px;height:18px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;font:11px/18px var(--sans);color:var(--ink-2)}
@media(max-width:650px){#owner-size-key{gap:6px;font-size:9px}#owner-size-key .owner-size-item{gap:3px}#owners .owner-stage-tools{left:8px;right:8px;gap:6px}}
@media print{#owners #owner-reset{display:none}#owners.owner-stage-panel{break-inside:avoid}}
</style>
<style>
#exec-forecast-extension{margin-top:24px;border-top:1px solid var(--rule)}
#exec-forecast-extension .two{grid-template-columns:minmax(0,1.75fr) minmax(0,1fr);gap:36px}
#exec-forecast-extension .exec-interval-controls{flex-wrap:wrap;margin:14px 0 12px;padding-bottom:8px;border-bottom:1px solid var(--rule-strong)}
#exec-forecast-extension .exec-fan-legend{gap:14px;margin-bottom:10px;font-size:11px}
#exec-forecast-extension .summary{margin-top:0}
#exec-forecast-extension .summary li{font-size:13px;line-height:20px}
#exec-forecast-extension .exec-table-heading{font:600 13px/18px var(--sans);margin:18px 0 8px;color:var(--ink-2)}
#exec-forecast-extension .chart{min-width:0}
#exec-forecast-extension .chart text.exec-end-label{fill:var(--fuchsia);font-size:11px;font-weight:600}
#exec-forecast-extension .chart text.exec-end-title{fill:var(--ink-3);font-size:9px}
#exec-forecast-extension .exec-table-note{color:var(--ink-3)}
@media(max-width:1050px){#exec-forecast-extension .two{grid-template-columns:1fr;gap:24px}}
@media print{#exec-forecast-extension .exec-interval-controls{display:none}#exec-forecast-extension .two{grid-template-columns:minmax(0,1.75fr) minmax(0,1fr);gap:24px}#exec-forecast-extension{break-before:page}#exec-forecast-extension .tablewrap{break-inside:avoid}}
</style>
<style>
#risk .board-risk-cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:24px 0 0}
#risk .board-risk-card{display:flex;flex-direction:column;align-items:flex-start;text-align:left;background:#fff;border:1px solid var(--rule);border-radius:12px;padding:17px 15px;color:var(--ink);cursor:pointer;min-width:0}
#risk .board-risk-card:hover{background:var(--pink-50);border-color:var(--fuchsia)}
#risk .board-risk-card[aria-pressed="true"]{outline:2px solid var(--fuchsia);outline-offset:1px;background:var(--pink-50)}
#risk .board-risk-card .risk-name{font:600 12px/17px var(--sans);min-height:34px;color:var(--ink-2)}
#risk .board-risk-card .risk-number{font:600 31px/38px var(--serif);margin:9px 0 5px}
#risk .board-risk-card .risk-amount{font:600 12px/17px var(--sans);margin-bottom:6px}
#risk .board-risk-card .risk-detail{font:400 11px/16px var(--sans);color:var(--ink-3);min-height:32px}
#risk .board-risk-card .risk-rag{margin-top:10px;font:600 11px/16px var(--sans)}
#risk .risk-high{color:#AD2145}#risk .risk-review{color:#A66511}#risk .risk-low{color:#167044}
#risk .board-risk-visuals{display:grid;grid-template-columns:minmax(0,1.4fr) minmax(0,1fr);gap:20px}
#risk .board-risk-chart-card{min-width:0;border:1px solid var(--rule);border-radius:12px;background:#fff;padding:18px 16px 8px}
#risk .board-risk-plot{height:375px;width:100%;min-width:0}
#risk .board-risk-wheel-key{display:grid;grid-template-columns:1fr 1fr;gap:8px 14px;font:11px/16px var(--sans)}
#risk .board-risk-wheel-key b{font-size:15px;color:var(--ink)}
#risk .board-risk-wheel-key .risk-wheel-swatch{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:5px}
#risk .board-risk-selection{font:12px/18px var(--sans);color:var(--ink-2);margin-bottom:10px}
#risk table.board-risk-table{min-width:880px;font-size:12px;table-layout:fixed}
#risk .board-risk-table th{padding:8px 7px;text-align:left;font-size:10px;white-space:normal;letter-spacing:.02em}
#risk .board-risk-table th button{border:0;background:none;color:inherit;font:inherit;text-align:inherit;padding:0;cursor:pointer;width:100%;line-height:16px}
#risk .board-risk-table th button:hover{color:var(--fuchsia)}
#risk .board-risk-table td{padding:10px 7px;text-align:left;white-space:normal;vertical-align:top;font-size:11px;line-height:16px}
#risk .board-risk-table td.number{text-align:right;white-space:nowrap}
#risk .board-risk-table td small{display:block;color:var(--ink-3);font-size:10px}
#risk .board-risk-table tr.risk-selected{background:#FCE4EC;box-shadow:inset 3px 0 0 #C2124F}
#risk .board-risk-table tr.risk-muted{opacity:.45}
#risk .risk-status{display:inline-block;border-radius:5px;padding:3px 6px;font-weight:600;white-space:nowrap}
#risk .risk-status.risk-high{background:#FBE8EC}#risk .risk-status.risk-review{background:#FFF3DE}#risk .risk-status.risk-low{background:#EAF5EB}
#risk .board-risk-definitions{font-size:11px;line-height:17px;color:var(--ink-3);margin:12px 0}
#risk .board-risk-definitions p{margin:8px 0}
@media(max-width:1050px){#risk .board-risk-cards{grid-template-columns:repeat(2,minmax(0,1fr))}#risk .board-risk-visuals{grid-template-columns:1fr}}
@media(max-width:480px){#risk .board-risk-cards{grid-template-columns:1fr}}
@media print{#risk .board-risk-cards{grid-template-columns:repeat(4,minmax(0,1fr))}#risk .board-risk-visuals{grid-template-columns:minmax(0,1.4fr) minmax(0,1fr)}#risk #board-risk-reset{display:none}#risk table.board-risk-table{min-width:0;font-size:9px}#risk .board-risk-table td{font-size:9px;padding:5px}#risk .board-risk-chart-card{break-inside:avoid}}
</style></head><body><div class="layout"><nav aria-label="Dashboard sections"><div class="wordmark"><b>Fuchsia Works</b><br>Digital Consultancy</div>
<button data-tab="executive" aria-selected="true">01 · Executive summary</button><button data-tab="composition" aria-selected="false">02 · Forecast composition</button><button data-tab="reliability" aria-selected="false">03 · Forecast reliability</button><button data-tab="risk" aria-selected="false">04 · Forecast risks</button><button data-tab="fx" aria-selected="false">05 · FX sensitivity</button><button data-tab="management" aria-selected="false">06 · Management action</button><button data-tab="assumptions" aria-selected="false">07 · Assumptions</button><div class="scope">Q4 2026 · EUR reporting<br><span id="nav-snapshot"></span><br><br>Confidential · Prepared for the Board</div></nav>
<article class="sheet"><header class="runhead"><div class="wordmark"><b>Fuchsia Works</b> Digital Consultancy</div><div class="doc">FP&amp;A Working Prototype · Terri Jackson · Beyond the Weighted Pipeline · <span>Q4 2026 pipeline</span></div></header><div id="fatal" role="alert"></div>
<section id="executive" class="panel active" aria-labelledby="exec-title"><div class="titleblock"><div class="kicker">Q4 2026 · Sales pipeline delivery outlook</div><h1 id="exec-title"></h1><p class="dek" id="exec-dek"></p></div><div class="kpis" id="kpis"></div><div class="key" aria-label="Status key"><span class="favourable">▲ Favourable · stronger reliability / lower risk</span><span class="caution">◆ Caution · review required</span><span class="concern">▼ Concern · material risk</span><span class="neutral">— Context · no target supplied</span></div><p class="small">Colours assess forecast reliability and exposure. A larger pipeline or a positive forecast gap alone does not establish favourable performance.</p>
<section class="sec"><div class="sechead"><h2>Executive commentary</h2><span class="units">Generated from this dashboard's measured and modelled KPIs</span></div><ul class="summary" id="commentary"></ul></section>
<section id="exec-forecast-extension" class="sec">
  <div class="two">
    <section>
      <h2 class="chart-heading">Cumulative Q4 delivery: forecast range</h2>
      <p class="chart-unit">EUR millions · cumulative at month end</p>
      <div class="scenario exec-interval-controls">
        <span>Interval</span>
        <div class="seg" role="group" aria-label="Central simulated outcome interval">
          <button type="button" data-exec-interval="50" aria-pressed="false">50%</button>
          <button type="button" data-exec-interval="80" aria-pressed="true">80%</button>
          <button type="button" data-exec-interval="90" aria-pressed="false">90%</button>
          <button type="button" data-exec-interval="95" aria-pressed="false">95%</button>
        </div>
        <input id="exec-interval" type="range" min="50" max="95" step="1" value="80" aria-label="Central outcome interval percentage">
        <output id="exec-interval-value" for="exec-interval">80%</output>
      </div>
      <div class="legend exec-fan-legend">
        <span><i class="k b"></i><span id="exec-band-label">Central 80% range (P10–P90)</span></span>
        <span><i class="k"></i>Simulated median</span>
        <span><i class="k e"></i>Probability weighted</span>
      </div>
      <div id="exec-cumulative-fan" class="chart"></div>
      <p id="exec-fan-note" class="fignote"></p>
    </section>
    <aside>
      <div class="sechead"><h2>Executive summary</h2></div>
      <ul id="exec-dynamic-summary" class="summary" aria-live="polite"></ul>
      <div class="phase"><h3>Phasing of weighted forecast</h3><div id="exec-month-phasing"></div></div>
    </aside>
  </div>
<section class="sec">
    <div class="sechead"><div><h2>Forecast range tables</h2><p id="exec-table-unit" class="units"></p></div></div>
    <h3 class="exec-table-heading">Cumulative from 1 October 2026</h3>
    <div id="exec-cumulative-table" class="tablewrap"></div>
    <h3 class="exec-table-heading">Revenue in each month</h3>
    <div id="exec-monthly-table" class="tablewrap"></div>
    <p class="fignote" id="exec-month-method"></p>
    <p id="exec-extension-error" class="concern" role="alert" hidden></p>
  </section>
</section>
</section>
<section id="composition" class="panel" aria-labelledby="composition-title"><div class="titleblock"><div class="kicker">Forecast composition</div><h1 id="composition-title">Where is the forecast coming from?</h1><p class="dek">Open opportunities with a positive Q4 workday allocation. Group totals reconcile to the executive summary.</p></div><div class="grid sec"><section class="wide"><h2 class="chart-heading">Pipeline funnel by sales stage</h2><p class="chart-unit">Nested curved funnel · stage width proportional to EUR value · stage composition, not historical conversion</p>
<div class="controls"><label for="funnel-stage">Highlight stage</label><select id="funnel-stage"><option value="-1">All stages</option></select><button class="btn" id="funnel-reset">Reset highlight</button><button class="btn" id="export-funnel-svg">Export funnel SVG</button></div>
<div class="funnel-wrap"><div id="funnel" class="plotly-funnel" role="region" aria-label="Interactive nested Q4 sales stage funnel"></div></div><img id="funnel-print" class="funnel-print" alt="Q4 2026 nested sales pipeline funnel for board-pack printing"><p id="funnel-selection" class="metric-note" aria-live="polite">All stages shown.</p><p id="funnel-error" class="concern" role="alert" hidden></p>
<p class="fignote">Gross, probability weighted and standard weighted layers use a common EUR scale. Both weighted boundaries remain visible where their values cross. The silhouette may widen at later stages because it follows the actual pipeline composition. Hover for values; click a stage to highlight it. SVG export contains vector artwork.</p></section><section><h2 class="chart-heading">Forecast by region</h2><p class="chart-unit">Probability weighted · EUR millions · descending</p><div id="regions" class="chart"></div></section><section><h2 class="chart-heading">Forecast by business offering</h2><p class="chart-unit">Probability weighted · EUR millions · descending</p><div id="offerings" class="chart"></div><p class="fignote">Business offering uses the workbook's service_line field; a separate business unit field is not supplied.</p></section><section class="wide"><h2 class="chart-heading">Owner pipeline by sales stage</h2><p class="chart-unit">Bubble area: probability weighted EUR · colour: win probability · owners ranked by forecast</p>
<div id="owners" class="owner-stage-panel">
  <div class="owner-stage-tools"><div id="owner-size-key" aria-label="Bubble size legend"></div><button type="button" id="owner-reset" class="btn">Reset owner</button></div>
  <div id="owner-stage-plot" role="region" aria-label="Q4 forecast by opportunity owner and sales stage"></div>
  <p id="owner-stage-status" class="owner-stage-status" aria-live="polite">Click a bubble to highlight its owner.</p>
</div><p class="fignote">All owners shown. Win probability is gross-pipeline-weighted; lines show composition. Tiny bubbles have a 4px visibility floor.</p></section></div></section>
<section id="reliability" class="panel" aria-labelledby="reliability-title"><div class="titleblock"><div class="kicker">Forecast reliability</div><h1 id="reliability-title">How believable is the current forecast?</h1><p class="dek">Simulated Q4 delivery uses assigned probabilities, supplied historic win rates and an explicit delivery-delay scenario.</p></div><section class="sec"><div class="sechead"><div><h2>Simulated Q4 revenue distribution</h2><p class="units">Revenue: EUR millions · density per EUR million</p></div></div><div class="controls"><label>Central outcome interval <input id="confidence" type="range" min="50" max="95" step="5" value="80"> <output id="confidence-value">80%</output></label><label>Histogram bins <select id="bins"><option>15</option><option selected>30</option><option>60</option></select></label><button class="btn" id="reset-hist">Reset</button></div><div class="legend"><span>Dashed ink: current probability forecast</span><span>Fuchsia: P50</span><span>Blue: P90</span></div><div id="histogram" class="chart"></div><p class="fignote" id="interval-note"></p><div id="quantiles" class="quantiles"></div><p class="small">P10 is the lower 10th percentile; P90 is the upper 90th percentile. Only about 10% of simulated outcomes exceed P90. These are modelled outcomes, not a validated statistical guarantee.</p></section><div class="grid sec"><section><h2 class="chart-heading">Monte Carlo confidence gauge</h2><p class="chart-unit">Forecast confidence score · heuristic reliability index, 0–100</p><div id="gauge" class="chart"></div><div id="score-components"></div><p class="fignote">The gauge combines five forecast risk inputs. It is separate from the simulated probability of reaching the forecast.</p></section><section><h2 class="chart-heading">Forecast distribution by sales stage</h2><p class="chart-unit">Simulated Q4 revenue contribution · EUR millions</p><div id="violin" class="chart"></div><p class="fignote">Each violin shows the stage's total simulated contribution, with median and interquartile range. Width is normalised within each stage; smoothing includes zero outcomes. Full zero-outcome shares appear on focus or hover.</p></section></div></section>
<section id="risk" class="panel" aria-labelledby="risk-title">
<div class="titleblock"><div class="kicker">Forecast risks</div><h1 id="risk-title">What could prevent delivery of the forecast?</h1><p class="dek">Revenue exposure, delivery dependencies and the opportunities requiring attention.</p></div>
<div id="board-risk-cards" class="board-risk-cards" aria-label="Executive risk cards"></div>
<div class="board-risk-visuals sec">
  <section class="board-risk-chart-card"><h2 class="chart-heading">Forecast risk waterfall</h2><p class="chart-unit">EUR millions · current weighted forecast to model expected outcome</p><div id="board-risk-waterfall" class="board-risk-plot" role="region" aria-label="Forecast risk waterfall"></div><p id="board-risk-waterfall-note" class="fignote"></p></section>
  <section class="board-risk-chart-card"><h2 class="chart-heading">Forecast dependency wheel</h2><p class="chart-unit">Share of probability weighted forecast · hover a deal for its owner and value</p><div id="board-risk-wheel" class="board-risk-plot" role="region" aria-label="Forecast dependency wheel"></div><div id="board-risk-wheel-key" class="board-risk-wheel-key"></div><p class="fignote">Top 3 and Top 5 are cumulative shares. Wheel segments represent individual opportunities. Click a segment to locate that deal in the watchlist.</p></section>
</div>
<section class="sec">
  <div class="sechead"><div><h2>Forecast watchlist</h2><p class="units">Probability weighted Q4 revenue exposed · click a column heading to sort</p></div><button id="board-risk-reset" class="btn" type="button">Clear highlight</button></div>
  <p id="board-risk-selection" class="board-risk-selection" aria-live="polite"></p>
  <div id="board-risk-watchlist" class="tablewrap"></div>
  <p class="fignote">One row per opportunity; exposure is its full weighted Q4 forecast, not an estimated loss. Risk categories overlap. Status denotes modelled severity, not CRM action progress. Update age is measured at the Salesforce snapshot.</p>
  <details class="board-risk-definitions"><summary>Risk definitions and severity thresholds</summary><div id="board-risk-definitions"></div></details>
  <p id="board-risk-error" class="concern" role="alert" hidden></p>
</section>
</section>
<section id="fx" class="panel" aria-labelledby="fx-title"><div class="titleblock"><div class="kicker">FX sensitivity</div><h1 id="fx-title">How exposed is the forecast to currency movements?</h1><p class="dek" id="fx-dek"></p></div><div class="controls"><label>FX rate shock <input id="fx-shock" type="range" min="0" max="20" step="1" value="10"> <output id="fx-shock-value">10%</output></label></div><section class="sec"><h2 class="chart-heading">FX tornado</h2><p class="chart-unit">Change in source probability weighted forecast · EUR millions</p><div id="fx-tornado" class="chart"></div><p class="fignote">Each currency is shocked independently. For an ECB quote of currency units per €1, EUR strengthening raises the quote and reduces EUR revenue: base ÷ (1 + shock). EUR weakening uses base ÷ (1 − shock). EUR deals are unchanged; no hedge data is supplied.</p></section><div class="tablewrap" id="currency-table"></div></section>
<section id="management" class="panel" aria-labelledby="management-title"><div class="titleblock"><div class="kicker">Management action</div><h1 id="management-title">Which deals need attention first?</h1><p class="dek">Priorities combine modelled shortfall and the source risk flags. They support management review, with named owners and next steps.</p></div><section class="sec"><div class="sechead"><div><h2>Executive watchlist</h2><p class="units">EUR thousands · top 10 deals by shortfall × (1 + priority / 100)</p></div><button id="export-watchlist" class="btn">Export watchlist CSV</button></div><div id="watchlist" class="tablewrap"></div><p class="fignote">Impact is max(current probability forecast − model expectation, 0). Recommendations are rules generated from recorded risks; interventions require owner judgement.</p></section><section class="sec"><h2 class="chart-heading">Action priority matrix</h2><p class="chart-unit">X: action priority index, 0–100 · Y: modelled shortfall to current forecast, EUR millions</p><div id="action-matrix" class="chart"></div><p class="fignote">Priority = 25 points for each risk flag, plus 25 × non-delivery likelihood, capped at 100. Flags: ≥40% slippage; record older than 30 days; assigned probability > historic by 10pp; deal ≥10% of current forecast. Lines: priority 60 and median shortfall.</p></section></section>
<section id="assumptions" class="panel" aria-labelledby="assumptions-title"><div class="titleblock"><div class="kicker">Assumptions and source</div><h1 id="assumptions-title">How the forecast is calculated</h1><p class="dek">Source measures, scenario choices and quality checks remain visible for board review.</p></div><section class="sec method" id="method"></section><section class="sec"><div class="sechead"><h2>ECB FX lookup supplied in the workbook</h2><p class="units">Rates per €1 and EUR multiplier</p></div><div id="fx-rates" class="tablewrap"></div></section><details><summary>Inspect Q4 opportunity inputs and calculations</summary><p class="small">Monetary columns: EUR thousands. Probabilities shown as percentages. Source row points to Analysis.</p><button class="btn" id="export-inputs">Export model inputs CSV</button><div id="input-table" class="tablewrap"></div></details></section>
<div class="actions print-btn"><button class="btn" id="print">Print / PDF · all sections</button></div><footer class="foot"><span><b>Portfolio demonstration</b> · AIAC Course Project</span><span id="footer-source"></span></footer></article></div>
<script id="forecastData" type="application/json">__DATA__</script>
<script>
'use strict';
const D=JSON.parse(document.getElementById('forecastData').textContent), K=D.kpis, C=D.config;
const $=id=>document.getElementById(id), E=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const M=v=>'€'+(v/1e6).toLocaleString('en-GB',{minimumFractionDigits:2,maximumFractionDigits:2})+'m';
const m1=v=>'€'+(v/1e6).toFixed(1)+'m', P=v=>v.toFixed(1)+'%', N=v=>v.toLocaleString('en-GB',{maximumFractionDigits:0}), S=v=>(v<0?'−':'+')+M(Math.abs(v));
const colors={gross:'#F8C9DA',probability:'#C2124F',standard:'#7D0B3F'}, ink='#2A0A1A',blue='#0B6E8A',orange='#B4460A';
const txt=(x,y,s,anchor='start',cls='label')=>`<text x="${x}" y="${y}" text-anchor="${anchor}" class="${cls}">${E(s)}</text>`;
const line=(x1,y1,x2,y2,color='#E9D3DD',extra='')=>`<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${color}" ${extra}/>`;
const mark=(body,tip)=>`<g tabindex="0" aria-label="${E(tip)}"><title>${E(tip)}</title>${body}</g>`;
function svg(id,title,desc,inner,w=820,h=380){$(id).innerHTML=`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" role="img" aria-label="${E(title)}"><title>${E(title)}</title><desc>${E(desc)}</desc>${inner}</svg>`;}
function nice(v){if(v<=0)return 1;const p=10**Math.floor(Math.log10(v)),r=v/p;return (r<=1?1:r<=2?2:r<=5?5:10)*p;}
function wrapLabel(s,x,y,max=17){let words=String(s).split(' '),a='',b='';for(const word of words){if((a+' '+word).trim().length<=max&&!b)a=(a+' '+word).trim();else b+=(b?' ':'')+word;}return txt(x,y,a,'middle')+txt(x,y+15,b,'middle');}
function axes(max,x0,y0,w,h,xAxis=false){let o='';for(let i=0;i<=4;i++){const v=max*i/4;if(xAxis){const x=x0+w*i/4;o+=line(x,y0,x,y0+h)+txt(x,y0+h+24,(v/1e6).toFixed(2),'middle');}else{const y=y0+h-h*i/4;o+=line(x0,y,x0+w,y)+txt(x0-9,y+4,(v/1e6).toFixed(2),'end');}}return o;}
function hbar(id,rows,title,color='#C2124F'){const w=820,left=230,right=90,top=12,rh=31,h=top+rows.length*rh+38,max=nice(Math.max(...rows.map(r=>r.probability))*1.05),pw=w-left-right;let b=axes(max,left,top,pw,rows.length*rh,true);rows.forEach((r,i)=>{const y=top+i*rh,label=r.label.length>32?r.label.slice(0,31)+'…':r.label;b+=mark(txt(left-12,y+19,label,'end')+`<rect x="${left}" y="${y+4}" width="${pw*r.probability/max}" height="20" fill="${color}"/>`+txt(left+pw*r.probability/max+7,y+19,M(r.probability),'start','value'),r.label+': '+M(r.probability));});svg(id,title,'Descending probability weighted EUR forecast by category.',b,w,h);}
// Smooth, value-preserving Plotly funnel. Stage width = its EUR amount, not an
// assumed conversion rate. Smoothstep interpolation avoids spline overshoot.
let funnelSelection=-1, funnelPromise=null, funnelHasAnimated=false, funnelBusy=false;
const funnelLayerNames=['Gross pipeline','Probability weighted','Standard weighted'];
const funnelKeys=['gross','probability','standard'];
const funnelFills=['#F8C9DA','rgba(194,18,79,0.40)','rgba(125,11,63,0.38)'];
const funnelLines=['#E0457F','#C2124F','#7D0B3F'];
const reducedMotion=()=>window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
function funnelWidth(y,key){
  const rows=D.funnel,n=rows.length;
  if(y<=0)return rows[0][key]/1e6;
  if(y>=n-1)return rows[n-1][key]/1e6;
  const i=Math.floor(y),t=y-i,u=t*t*(3-2*t);
  return (rows[i][key]*(1-u)+rows[i+1][key]*u)/1e6;
}
function funnelPolygon(key,stage=null,scale=1){
  const n=D.funnel.length,low=stage===null?-.38:Math.max(-.38,stage-.43),
    high=stage===null?n-1+.38:Math.min(n-1+.38,stage+.43),steps=stage===null?Math.max(100,n*36):60;
  const y=Array.from({length:steps+1},(_,i)=>low+(high-low)*i/steps);
  const widths=y.map(v=>funnelWidth(v,key)*scale/2);
  return {x:widths.map(v=>-v).concat(widths.slice().reverse(),[-widths[0]]),
          y:y.concat(y.slice().reverse(),[y[0]])};
}
function funnelBoundary(key,sign=1,scale=1){
  const n=D.funnel.length,steps=Math.max(100,n*36),y=Array.from({length:steps+1},(_,i)=>-.38+(n-1+.76)*i/steps);
  return {x:y.map(v=>sign*funnelWidth(v,key)*scale/2),y};
}
function funnelStageDetails(){
  const i=funnelSelection;
  $('funnel-stage').value=String(i);
  if(i<0){$('funnel-selection').textContent='All stages shown. Click inside a stage or choose one from the menu to highlight it.';return;}
  const r=D.funnel[i],count=D.deals.filter(d=>d.stage===r.label).length;
  $('funnel-selection').textContent=`${r.label} · ${count} opportunities · gross ${M(r.gross)} · probability ${M(r.probability)} · standard ${M(r.standard)} · gap ${S(r.probability-r.standard)}`;
}
function funnelError(error){
  $('funnel-error').hidden=false;
  $('funnel-error').textContent='The interactive funnel could not render: '+error.message;
  console.error(error);
}
async function snapshotFunnel(){
  const url=await Plotly.toImage($('funnel'),{format:'svg',width:1100,height:680});
  $('funnel-print').src=url;
  return url;
}
function renderFunnel(){
  if(funnelPromise)return funnelPromise;
  funnelPromise=(async()=>{
    if(!window.Plotly)throw new Error('The embedded Plotly library is unavailable. Regenerate the HTML with the updated script.');
    const n=D.funnel.length,maxWidth=Math.max(...D.funnel.map(r=>r.gross))/1e6;
    const data=[];
    // Three filled silhouettes. Both weighted boundaries are redrawn above the
    // fills so neither series vanishes where the two weighted values cross.
    funnelKeys.forEach((key,j)=>{data.push({type:'scatter',mode:'lines',...funnelPolygon(key),
      name:funnelLayerNames[j],showlegend:true,fill:'toself',fillcolor:funnelFills[j],
      line:{color:funnelLines[j],width:1.6,shape:'linear',simplify:false},
      hoverinfo:'skip',legendgroup:key,opacity:1});});
    funnelKeys.forEach((key,j)=>{data.push({type:'scatter',mode:'lines',...funnelPolygon(key,0),
      name:funnelLayerNames[j]+' selected',showlegend:false,fill:'toself',fillcolor:funnelFills[j],
      line:{color:funnelLines[j],width:2.4,shape:'linear',simplify:false},hoverinfo:'skip',opacity:0});});
    ['probability','standard'].forEach((key,j)=>{[-1,1].forEach(sign=>data.push({
      type:'scatter',mode:'lines',...funnelBoundary(key,sign),showlegend:false,
      line:{color:funnelLines[j+1],width:1.6,dash:key==='standard'?'dot':'solid',simplify:false},hoverinfo:'skip'}));});
    // Transparent, densely spaced hit targets provide one truthful hover card
    // and a reliable click target throughout each stage band.
    const hx=[],hy=[],hc=[];
    D.funnel.forEach((r,i)=>{
      const card=[E(r.label),r.gross,r.probability,r.standard,r.probability-r.standard,
                  D.deals.filter(d=>d.stage===r.label).length,i];
      for(let row=0;row<=8;row++){
        const y=Math.max(-.38,Math.min(n-1+.38,i-.40+row*.10)),half=funnelWidth(y,'gross')/2;
        for(let col=0;col<=26;col++){hx.push(-half+2*half*col/26);hy.push(y);hc.push(card);}
      }
    });
    data.push({type:'scatter',mode:'markers',x:hx,y:hy,customdata:hc,showlegend:false,
      marker:{size:15,color:'rgba(0,0,0,0)'},
      hovertemplate:'<b>%{customdata[0]}</b><br>Gross: €%{customdata[1]:,.2f}<br>Probability: €%{customdata[2]:,.2f}<br>Standard: €%{customdata[3]:,.2f}<br>Forecast gap: €%{customdata[4]:+,.2f}<br>Opportunities: %{customdata[5]}<extra></extra>'});
    const annotations=D.funnel.map((r,i)=>({xref:'paper',yref:'y',x:1.02,y:i,
      text:M(r.probability),showarrow:false,xanchor:'left',font:{size:13,color:'#C2124F'}}));
    annotations.push({xref:'paper',yref:'paper',x:1.02,y:1.055,text:'Probability',showarrow:false,
      xanchor:'left',font:{size:11,color:'#5E3A4C'}});
    const layout={width:1000,height:680,autosize:true,paper_bgcolor:'#FFFFFF',plot_bgcolor:'#FFFFFF',
      title:{text:'Q4 2026 pipeline by sales stage',x:0,xanchor:'left',y:.985,yanchor:'top',font:{family:'Source Serif 4, Georgia, serif',size:22,color:'#2A0A1A'}},
      margin:{l:180,r:140,t:110,b:32},font:{family:'IBM Plex Sans, Arial, sans-serif',size:13,color:'#2A0A1A'},
      xaxis:{range:[-maxWidth*.57,maxWidth*.57],visible:false,fixedrange:true},
      yaxis:{range:[n-1+.6,-.6],tickvals:D.funnel.map((_,i)=>i),
        ticktext:D.funnel.map(r=>E(r.label)),fixedrange:true,showgrid:false,zeroline:false,
        showline:false,tickfont:{size:12,color:'#5E3A4C'},ticks:''},
      hovermode:'closest',hoverdistance:35,clickmode:'event',dragmode:false,
      hoverlabel:{bgcolor:'#FFFFFF',bordercolor:'#C2124F',font:{family:'IBM Plex Sans, Arial, sans-serif',size:13,color:'#2A0A1A'},align:'left'},
      legend:{orientation:'h',x:0,y:1.045,xanchor:'left',yanchor:'bottom',font:{size:12},
              itemclick:false,itemdoubleclick:false},annotations};
    await Plotly.newPlot($('funnel'),data,layout,{responsive:true,displaylogo:false,
      modeBarButtonsToRemove:['zoom2d','pan2d','select2d','lasso2d','zoomIn2d','zoomOut2d','autoScale2d','resetScale2d'],
      toImageButtonOptions:{format:'svg',filename:'fuchsia_works_q4_2026_nested_funnel',width:1100,height:680}});
    $('funnel').on('plotly_click',event=>{
      const point=(event.points||[]).find(p=>p.customdata);
      if(point)selectFunnelStage(Number(point.customdata[6]));
    });
    $('funnel').on('plotly_legendclick',()=>false);
    funnelStageDetails();
    await snapshotFunnel();
    return $('funnel');
  })();
  funnelPromise.catch(funnelError);
  return funnelPromise;
}
async function selectFunnelStage(stage){
  await renderFunnel();
  if(funnelBusy)return;
  if(!Number.isInteger(stage)||stage < -1||stage>=D.funnel.length)return;
  funnelBusy=true;
  ['funnel-stage','funnel-reset','export-funnel-svg'].forEach(id=>$(id).disabled=true);
  funnelSelection=(stage===funnelSelection)?-1:stage;
  funnelStageDetails();
  const overlayStage=Math.max(0,funnelSelection),selected=funnelSelection>=0;
  try{
    const updates=funnelKeys.map(()=>({opacity:selected?.30:1}));
    funnelKeys.forEach(key=>updates.push({...funnelPolygon(key,overlayStage),opacity:selected?1:0}));
    const ms=reducedMotion()?0:450;
    await Plotly.animate($('funnel'),{data:updates,traces:[0,1,2,3,4,5]},
      {mode:'immediate',transition:{duration:ms,easing:'cubic-in-out'},frame:{duration:ms,redraw:false}});
    await snapshotFunnel();
  }catch(error){funnelError(error);}
  finally{funnelBusy=false;['funnel-stage','funnel-reset','export-funnel-svg'].forEach(id=>$(id).disabled=false);}
}
async function revealFunnel(){
  try{
    await renderFunnel();
    await Plotly.Plots.resize($('funnel'));
    if(!funnelHasAnimated&&!reducedMotion()){
      funnelHasAnimated=true;
      // Animate only X coordinates; the real stage amounts are restored before
      // any export. Highlight transitions subsequently move the selected band.
      await Plotly.restyle($('funnel'),{x:funnelKeys.map(key=>funnelPolygon(key,null,.02).x)},[0,1,2]);
      await Plotly.animate($('funnel'),{data:funnelKeys.map(key=>({x:funnelPolygon(key).x})),traces:[0,1,2]},
        {mode:'immediate',transition:{duration:650,easing:'cubic-in-out'},frame:{duration:650,redraw:false}});
      await snapshotFunnel();
    }
  }catch(error){funnelError(error);}
}
function setupFunnelControls(){
  $('funnel-stage').innerHTML='<option value="-1">All stages</option>'+D.funnel.map((r,i)=>`<option value="${i}">${E(r.label)}</option>`).join('');
  $('funnel-stage').addEventListener('change',()=>{const index=Number($('funnel-stage').value);if(index!==funnelSelection)selectFunnelStage(index);});
  $('funnel-reset').addEventListener('click',()=>{if(funnelSelection!==-1)selectFunnelStage(-1);});
  $('export-funnel-svg').addEventListener('click',async()=>{
    const button=$('export-funnel-svg');button.disabled=true;
    try{
      await renderFunnel();
      if(funnelBusy)throw new Error('Wait for the stage transition to finish, then export again.');
      await Plotly.downloadImage($('funnel'),{format:'svg',filename:'fuchsia_works_q4_2026_nested_funnel',width:1100,height:680});
    }catch(error){funnelError(error);}finally{button.disabled=false;}
  });
}

function riskStatus(v,good,bad,reverse=false){let t=reverse?(v<=good?'favourable':v<=bad?'caution':'concern'):(v>=good?'favourable':v>=bad?'caution':'concern');return [t,t==='favourable'?'▲ Favourable':t==='caution'?'◆ Caution':'▼ Concern'];}
function renderExecutive(){const q=D.quantiles; $('exec-title').innerHTML=`Q4 probability forecast of <em>${M(K.probability)}</em>; modelled 80% outcome range ${m1(q.p10)}–${m1(q.p90)}`;$('exec-dek').textContent=`EUR group reporting · snapshot ${D.meta.snapshot} · ${N(K.count)} contributing opportunities · independent deal simulation`;
const gapPct=K.standard?K.gap/K.standard*100:0,scoreState=riskStatus(K.score,80,60),largestState=riskStatus(K.largest_pct,10,20,true),fxState=riskStatus(K.non_eur_pct,25,50,true);
const cards=[['Q4-2026 Gross Pipeline',M(K.gross),'Open pipeline allocated to Q4 workdays',['neutral','— Context']],['Q4-2026 Probability Weighted Forecast',M(K.probability),P(K.probability/K.gross*100)+' of gross pipeline',['neutral','— Context']],['Q4-2026 Standard Weighted Forecast',M(K.standard),'Workbook stage-weight benchmark',['neutral','— Context']],['Q4-2026 Forecast Gap (Probability − Standard)',S(K.gap),(gapPct>=0?'+':'')+P(gapPct)+' against standard; sales_confidence in CI',Math.abs(gapPct)>10?['concern','▼ Concern · wide gap']:Math.abs(gapPct)>5?['caution','◆ Caution · review gap']:['favourable','▲ Favourable · measures aligned']],['Q4-2026 Opportunity Count',N(K.count),'Count of distinct contributing IDs',K.count>=50?['favourable','▲ Favourable · broad count']:K.count>=20?['caution','◆ Caution · limited count']:['concern','▼ Concern · small count']],['Non-EUR % of Q4-2026',P(K.non_eur_pct),'Share of gross Q4 EUR value in non-EUR deals',fxState],['Largest Deal Contribution %',P(K.largest_pct),'Largest deal ÷ probability weighted forecast',largestState],['Q4-2026 Forecast Confidence Score',Math.round(K.score)+'/100',K.rating+' · heuristic reliability score',scoreState]];
$('kpis').innerHTML=cards.map((c,i)=>`<div class="kpi ${i===1?'lead':''}"><div class="l">${E(c[0])}</div><div class="v">${E(c[1])}</div><div class="c">${E(c[2])}</div><div class="status ${c[3][0]}">${E(c[3][1])}</div></div>`).join('');
const top5=D.pareto.slice(0,5).reduce((s,d)=>s+d.value,0)/K.probability*100,topRegion=D.regions[0],topOffering=D.offerings[0],topOwner=D.owners[0];
$('commentary').innerHTML=[
`<strong>Forecast assessment.</strong> ${N(K.count)} opportunities support ${M(K.probability)}, ${S(K.gap)} (${gapPct>=0?'+':''}${P(gapPct)}) above the stage benchmark of ${M(K.standard)}. Gross Q4 delivery opportunity is ${M(K.gross)}. ${E(topRegion.label)} supplies ${P(topRegion.probability/K.probability*100)}; ${E(topOffering.label)} is the largest offering (${P(topOffering.probability/K.probability*100)}).`,
`<strong>Reliability assessment.</strong> The calibrated, slippage-adjusted model expects ${M(K.model_mean)}, with P50 ${M(q.p50)} and P10–P90 ${m1(q.p10)}–${m1(q.p90)}. ${P(K.reach_forecast_pct)} of simulations reach the current forecast. Gross-weighted assigned probabilities exceed supplied historic rates by ${K.inflation_pp.toFixed(1)} percentage points.`,
`<strong>Risk assessment.</strong> The largest deal supplies ${P(K.largest_pct)} of the weighted forecast; the top five supply ${P(top5)}. Weighted slippage risk is ${P(K.slippage_pct)} and ${P(K.stale_pct)} is in records older than ${C.stale_after_days} days. ${P(K.non_eur_pct)} of gross Q4 value is exposed to non-EUR currencies. ${E(topOwner.label)} owns the largest forecast share (${P(topOwner.probability/K.probability*100)}).`,
`<strong>Recommended actions.</strong> Review the watchlist, beginning with ${E(D.actions[0].id)} (${E(D.actions[0].owner)}), with modelled shortfall ${M(D.actions[0].impact)}. Validate delivery dates, update stale opportunities and challenge unsupported win probabilities. Review the FX scenarios alongside any hedges before deciding on mitigation.`,
`<strong>Overall confidence.</strong> ${K.rating}, ${Math.round(K.score)}/100. This is a transparent risk index. Supplied historic rates have no supporting closed-deal sample or backtest in this workbook; independence may understate shared customer, market or delivery shocks.`].map(t=>'<li>'+t+'</li>').join('');
}
const sorted=D.simulated_totals.slice().sort((a,b)=>a-b);function quantile(p){const v=(sorted.length-1)*p,i=Math.floor(v);return sorted[i]+(sorted[Math.min(i+1,sorted.length-1)]-sorted[i])*(v-i);}
function renderHistogram(){const bins=+ $('bins').value,level=+$('confidence').value,lo=quantile((1-level/100)/2),hi=quantile(1-(1-level/100)/2);$('confidence-value').textContent=level+'%';
const w=820,h=400,left=65,top=75,pw=730,ph=265,xmax=nice(Math.max(...sorted,K.probability,D.quantiles.p90)*1.02),dx=xmax/bins,counts=Array(bins).fill(0);for(const v of sorted)counts[Math.min(bins-1,Math.floor(v/dx))]++;
const density=counts.map(n=>n/sorted.length/(dx/1e6)),max=nice(Math.max(...density)*1.1),X=v=>left+pw*v/xmax,Y=v=>top+ph-ph*v/max;let b=`<rect x="${X(lo)}" y="${top}" width="${X(hi)-X(lo)}" height="${ph}" class="band" fill="#F8C9DA" fill-opacity=".55"/>`;
for(let i=0;i<=4;i++){let y=Y(max*i/4);b+=line(left,y,left+pw,y)+txt(left-8,y+4,(max*i/4).toFixed(2),'end');b+=txt(X(xmax*i/4),top+ph+25,(xmax*i/4/1e6).toFixed(1),'middle');}
density.forEach((d,i)=>{b+=mark(`<rect x="${X(i*dx)+1}" y="${Y(d)}" width="${pw/bins-2}" height="${top+ph-Y(d)}" fill="#C2124F" fill-opacity=".70"/>`,`${M(i*dx)} to ${M((i+1)*dx)}: ${P(counts[i]/sorted.length*100)} of simulations; density ${d.toFixed(3)} per €m`);});
[[K.probability,ink,'Current forecast',18,'stroke-dasharray="5 4"'],[D.quantiles.p50,'#C2124F','P50',38,''],[D.quantiles.p90,blue,'P90',58,'']].forEach(([v,c,l,y,extra])=>{const x=X(v);b+=line(x,top,x,top+ph,c,`stroke-width="2" ${extra}`)+`<text x="${Math.max(left+2,Math.min(left+pw-2,x))}" y="${y}" text-anchor="${x>left+pw*.75?'end':x<left+pw*.25?'start':'middle'}" style="fill:${c}">${E(l+' '+M(v))}</text>`;});svg('histogram','Simulated Q4 revenue probability density',`${level}% central interval ${M(lo)} to ${M(hi)}; P50 ${M(D.quantiles.p50)}.`,b,w,h);
$('interval-note').textContent=`Shading: central ${level}% modelled outcome interval, ${M(lo)}–${M(hi)}. ${N(C.simulations)} simulations, seed ${C.random_seed}. Changing this interval does not change the simulation or reliability score.`;
$('quantiles').innerHTML=['p10','p25','p50','p75','p90'].map(k=>`<div class="${k==='p50'?'median':''}">${k.toUpperCase()}<b>${M(D.quantiles[k])}</b></div>`).join('');}
function renderGauge(){const w=460,h=245,cx=230,cy=173,r=135,point=v=>[cx+r*Math.cos(Math.PI*(1-v/100)),cy-r*Math.sin(Math.PI*(1-v/100))];let b='';[[0,60,orange],[60,80,'#8A6B7A'],[80,100,blue]].forEach(([a,z,c])=>{let s=point(a),e=point(z);b+=`<path d="M ${s} A ${r} ${r} 0 0 1 ${e}" fill="none" stroke="${c}" stroke-width="16"/>`;});let v=point(K.score);b+=line(cx,cy,v[0],v[1],ink,'stroke-width="3"')+`<circle cx="${cx}" cy="${cy}" r="5" fill="${ink}"/>`+txt(cx,cy+38,`${Math.round(K.score)}/100 · ${K.rating}`,'middle','value')+txt(80,200,'0','middle')+txt(385,200,'100','middle');svg('gauge','Forecast confidence score',`${K.score.toFixed(1)} out of 100, ${K.rating}; this is not a win probability.`,b,w,h);
$('score-components').innerHTML=Object.entries(D.risk_components).map(([k,v])=>`<div class="metric-note">${E(k)} risk input: ${P(v*100)} · penalty ${(100*C.score_weights[k]*v).toFixed(1)} points</div>`).join('');}
function renderViolin(){let w=700,h=410,left=60,top=18,ph=295,pw=620,max=nice(Math.max(...D.violin.flatMap(v=>v.y))*1.02),Y=v=>top+ph-ph*v/max,b=axes(max,left,top,pw,ph),bw=pw/D.violin.length;D.violin.forEach((v,i)=>{const cx=left+(i+.5)*bw,half=bw*.37;let pts=v.y.map((y,j)=>[cx+half*v.density[j],Y(y)]).concat(v.y.map((y,j)=>[cx-half*v.density[j],Y(y)]).reverse());b+=mark(`<path d="M ${pts.map(a=>a.join(',')).join(' L ')} Z" fill="#F8C9DA" stroke="#E0457F" stroke-width="1"/>`+line(cx,Y(v.p25),cx,Y(v.p75),'#7D0B3F','stroke-width="5"')+`<circle cx="${cx}" cy="${Y(v.p50)}" r="4" fill="#C2124F"/>`,`${v.stage}: P25 ${M(v.p25)}, P50 ${M(v.p50)}, P75 ${M(v.p75)}; zero outcome ${P(v.zero_pct)}`);b+=wrapLabel(v.stage,cx,top+ph+25,16);});svg('violin','Simulated revenue distribution by sales stage','Stage total contribution; normalised density widths with medians and interquartile ranges.',b,w,h);}
function tornado(id,rows,baseline,desc){const w=820,left=270,top=18,rh=45,plotw=460,max=nice(Math.max(...rows.flatMap(r=>[Math.abs(r.a-baseline),Math.abs(r.b-baseline)]))*1.1),half=plotw/2,zero=left+half,h=top+rows.length*rh+40,X=v=>zero+v/max*half;let b='';for(let i=-2;i<=2;i++){const v=max*i/2,x=X(v);b+=line(x,top,x,top+rows.length*rh,i===0?ink:'#E9D3DD')+txt(x,h-10,(v<0?'−':v>0?'+':'')+(Math.abs(v)/1e6).toFixed(2),'middle');}rows.forEach((r,i)=>{let a=Math.min(r.a,r.b)-baseline,z=Math.max(r.a,r.b)-baseline,y=top+i*rh;b+=txt(left-12,y+24,r.label,'end');[[a,orange,y+4],[z,blue,y+21]].forEach(([v,c,y1])=>{b+=mark(`<rect x="${Math.min(zero,X(v))}" y="${y1}" width="${Math.max(0.8,Math.abs(X(v)-zero))}" height="14" fill="${c}"/>`,`${r.label}: ${S(v)} versus baseline ${M(baseline)}; scenario total ${M(baseline+v)}`);});});svg(id,'Tornado sensitivity chart',desc,b,w,h);}
function scatter(id,action=false){const w=820,h=390,left=75,top=35,pw=715,ph=290,rows=D.actions,annotations=[],max=nice(Math.max(...rows.map(d=>action?d.impact:d.gross))*1.06),cut=quantileArray(rows.map(d=>action?d.impact:d.gross),.5),X=v=>left+pw*v/100,Y=v=>top+ph-ph*v/max;let b=axes(max,left,top,pw,ph),threshold=action?60:50;b+=`<rect x="${X(threshold)}" y="${top}" width="${X(100)-X(threshold)}" height="${Y(cut)-top}" fill="#FCE4EC"/>`+line(X(threshold),top,X(threshold),top+ph,ink,'stroke-dasharray="4 4"')+line(left,Y(cut),left+pw,Y(cut),ink,'stroke-dasharray="4 4"')+txt(left+pw-6,top+15,action?'Act first':'High impact / high likelihood','end');
for(let i=0;i<=4;i++)b+=txt(X(i*25),top+ph+25,i*25+(action?'':'%'),'middle');rows.forEach((d,i)=>{const x=X(action?d.priority:d.likelihood*100),y=Y(action?d.impact:d.gross),r=3+9*Math.sqrt(d.current/Math.max(K.probability,1)),c=(action?d.priority>=60:d.likelihood>=.5)?orange:blue;b+=mark(`<circle cx="${x}" cy="${y}" r="${r}" fill="${c}" fill-opacity=".65" stroke="#fff"/>`,`${d.id}, ${d.owner}: ${action?'priority '+d.priority.toFixed(0)+', shortfall '+M(d.impact):'non-delivery '+P(d.likelihood*100)+', gross '+M(d.gross)}; ${d.reasons.join('; ')}`);if(i<3)annotations.push({x,y,label:d.id,lx:Math.min(x+12,left+pw-115),ly:Math.max(top+28,y-10)});});annotations.sort((a,b)=>a.ly-b.ly);for(let i=0;i<annotations.length;i++){const a=annotations[i];if(i)a.ly=Math.max(a.ly,annotations[i-1].ly+26);a.ly=Math.min(a.ly,top+ph-12);b+=line(a.x,a.y,a.lx,a.ly-4,'#8A6B7A','stroke-width=".7"')+txt(a.lx,a.ly,a.label);}svg(id,action?'Action priority matrix':'Opportunity risk matrix','Each bubble is an opportunity. Focus or hover for its identity and risk.',b,w,h);}
function quantileArray(a,p){a=a.slice().sort((x,y)=>x-y);const v=(a.length-1)*p,i=Math.floor(v);return a[i]+(a[Math.min(i+1,a.length-1)]-a[i])*(v-i);}
function table(headers,rows,cls=''){return `<table class="${cls}"><thead><tr>${headers.map(h=>'<th>'+E(h)+'</th>').join('')}</tr></thead><tbody>${rows.map(r=>'<tr>'+r.map(v=>'<td>'+E(v)+'</td>').join('')+'</tr>').join('')}</tbody></table>`;}
function renderFX(){let shock=+$('fx-shock').value/100;$('fx-shock-value').textContent=P(shock*100);let rows=D.currencies.filter(r=>r.label!=='EUR').map(r=>({label:r.label,a:r.probability/(1+shock)-r.probability,b:r.probability/(1-shock)-r.probability})).sort((a,b)=>Math.abs(b.b)-Math.abs(a.b));if(rows.length)tornado('fx-tornado',rows,0,'Independent currency shocks: EUR strengthens (negative orange) or weakens (positive blue).');else $('fx-tornado').textContent='All contributing deals are denominated in EUR; no translation exposure.';
$('fx-dek').textContent=`${P(K.non_eur_pct)} of gross Q4 delivery is denominated outside EUR. The existing workbook FX lookup remains the reporting baseline.`;
$('currency-table').innerHTML='<p class="units">EUR thousands · probability weighted currency exposure and scenario change</p>'+table(['Currency','Base forecast','Share','EUR stronger: change','EUR weaker: change'],D.currencies.map(r=>[r.label,N(r.probability/1000),P(r.probability/K.probability*100),r.label==='EUR'?'0':N((r.probability/(1+shock)-r.probability)/1000),r.label==='EUR'?'0':N((r.probability/(1-shock)-r.probability)/1000)]));}
function renderWatchlist(){$('watchlist').innerHTML='<table class="watch"><thead><tr><th>Opportunity / owner</th><th>Current forecast</th><th>Model expectation</th><th>Shortfall</th><th>Risk / priority</th><th>Recommended action</th></tr></thead><tbody>'+D.actions.slice(0,10).map(d=>`<tr><td><b>${E(d.id)}</b><br>${E(d.name)}<br><span class="small">${E(d.owner)} · ${E(d.stage)}</span></td><td class="number">${N(d.current/1000)}</td><td class="number">${N(d.model/1000)}</td><td class="number">${N(d.impact/1000)}</td><td>${E(d.reasons.join('; '))}<br><b>Priority ${Math.round(d.priority)}/100</b></td><td>${E(d.action)}<br><span class="small">${E(d.due)}. Recorded next step: ${E(d.next_step)}</span></td></tr>`).join('')+'</tbody></table>';}
function renderMethods(){const cfg=C,fxMonths=[...new Set(D.meta.fx.map(r=>r.month))].join(', ');$('method').innerHTML=`<ol>
<li><b>Scope and currency.</b> Q4 2026 means 1 October to 31 December. Includes open, relevant opportunities with AV &gt; 0. ${N(D.meta.source_rows)} source records; ${N(K.count)} contributing IDs; ${N(D.meta.excluded_rows)} excluded records. This is opportunity-based delivery pipeline, excluding closed deals and contracted backlog, not total company revenue. Values are EUR group reporting currency using the supplied European Central Bank monthly average lookup (${E(fxMonths)}). The script reads those rates; it does not fetch or refresh them.</li>
<li><b>Workday allocation.</b> Monday–Friday workdays only. No local public or corporate holidays. Start date is inclusive; EDATE(start, delivery months) is exclusive. AV is checked against AS × Q4 workdays / total delivery workdays.</li>
<li><b>Source forecasts.</b> Gross = AV; probability weighted = BI; standard weighted = BV; sales_confidence = CI = BI − BV. Standard stage weights are read from the workbook. Business offering = service_line. Owners are pre-validated. No budget, prior forecast or conversion history is supplied.</li>
<li><b>Monte Carlo model.</b> Each opportunity contributes its gross Q4 EUR allocation when won. Win probability = ${(100*(1-cfg.historical_blend)).toFixed(0)}% × assigned probability + ${(100*cfg.historical_blend).toFixed(0)}% × supplied historic stage win rate. The two rates are blended, not multiplied. Conditional on a win, slippage occurs with the supplied slippage probability; a slipped deal shifts both delivery endpoints by ${cfg.slippage_delay_days} calendar days. Retained Q4 allocation = min(1, delayed Q4 workday fraction / original fraction). Loss gives zero. The model excludes potential upside from slippage into Q4 and from opportunities with no original Q4 allocation.</li>
<li><b>Simulation and limits.</b> ${N(cfg.simulations)} draws, seed ${cfg.random_seed}. Deals are independent; wins and slip events are independent. No systemic correlation, uncertain contract values, new opportunities or actual revenue is modelled. Supplied historic rates have no underlying historical sample or backtest attached. The central 80% outcome interval is P10–P90; P90 is an upside percentile. Model mean ${M(K.model_mean)}; sampled mean ${M(K.simulation_mean)}. The interval slider changes the displayed range only.</li>
<li><b>Reliability score.</b> 100 − 100 × weighted risk penalties. Concentration: 25% × min(HHI / 0.10, 1), where HHI uses weighted-forecast deal shares; slippage: 25% × weighted slippage probability; stale: 20% × weighted forecast share in records older than ${cfg.stale_after_days} days; inflation: 20% × min(gross-weighted positive assigned-minus-historic probability / 20pp, 1); count: 10% × max(0, 1 − deal count / 50). Review thresholds and weights are policy assumptions. Score bands: 80–100 favourable; 60–79 caution; below 60 concern. This is a heuristic index, not the probability of achieving the forecast.</li>
<li><b>Other status thresholds.</b> Forecast gap: absolute gap ≤5% of standard is aligned; &gt;5–10% caution; &gt;10% concern. Count: ≥50 favourable, 20–49 caution, &lt;20 concern. Largest deal share: ≤10% favourable, &gt;10–20% caution, &gt;20% concern. Non-EUR gross share: ≤25% favourable, &gt;25–50% caution, &gt;50% concern. Headline amounts are contextual without a target; a positive gap is not evidence of outperformance.</li>
<li><b>Risk drivers and actions.</b> Tornado scenarios change one input at a time and report analytical expected revenue changes, not a causal decomposition or statistical attribution of variance. Pareto uses probability weighted amounts. Risk matrix Y is gross Q4 EUR at stake; X is 1 − model win probability × expected retention. Watchlist impact is max(BI − model expected revenue, 0); priority is the disclosed risk-flag index. Review and approve actions with the opportunity owner.</li>
<li><b>FX sensitivity.</b> Shocks change one currency's ECB quote independently, holding local revenue and probability fixed. Other currencies and EUR remain unchanged. No hedges are included. Scenario values are separate from the Monte Carlo distribution.</li>
<li><b>Source and validation.</b> ${E(D.meta.filename)} · Analysis, header row ${D.meta.header_row} · snapshot ${D.meta.snapshot}. SHA-256: <span style="overflow-wrap:anywhere">${E(D.meta.sha256)}</span>. Unique IDs, complete saved numeric results, non-negative source values, BI probability multiplication, CI reconciliation, workday allocation and the supplied checks (${D.meta.validated_checks.join(', ')}) passed. Workbook subtotal/detail totals reconcile. ${E(cfg.classification)}.</li></ol>`;
$('fx-rates').innerHTML=table(['Currency','EUR multiplier','Rate month','Currency units per €1'],D.meta.fx.map(r=>[r.currency,r.multiplier.toFixed(8),r.month,r.quote.toFixed(8)]));
$('input-table').innerHTML=table(['ID','Source row','Stage','Gross','Probability','Standard','CI gap','Assigned %','Historic %','Slippage %','Model win %','Slip retention %','Model expected'],D.deals.map(r=>[r.id,r.row,r.stage,N(r.gross/1000),N(r.probability/1000),N(r.standard/1000),N(r.sales_confidence/1000),P(r.assigned*100),P(r.historic*100),P(r.slippage*100),P(r.calibrated*100),P(r.retention*100),N(r.expected/1000)]));}
function exportCSV(rows,name){if(!rows.length)return;const cols=Object.keys(rows[0]),cell=v=>{let s=Array.isArray(v)?v.join('; '):String(v??'');if(typeof v==='string'&&/^[=+\-@\t\r]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';},csv='\ufeff'+[cols.map(cell).join(','),...rows.map(r=>cols.map(k=>cell(r[k])).join(','))].join('\r\n'),url=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'})),a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);}
function activate(id){document.querySelectorAll('.panel').forEach(p=>p.classList.toggle('active',p.id===id));document.querySelectorAll('nav button').forEach(b=>b.setAttribute('aria-selected',String(b.dataset.tab===id)));if(id==='composition')requestAnimationFrame(revealFunnel);}
try{
$('nav-snapshot').textContent=D.meta.snapshot;$('footer-source').textContent=`Snapshot ${D.meta.snapshot} · Analysis · EUR`;
renderExecutive();setupFunnelControls();renderFunnel();hbar('regions',D.regions,'Forecast by region');hbar('offerings',D.offerings,'Forecast by business offering');renderHistogram();renderGauge();renderViolin();renderFX();renderWatchlist();scatter('action-matrix',true);renderMethods();
document.querySelectorAll('nav button').forEach(b=>b.addEventListener('click',()=>{activate(b.dataset.tab);window.scrollTo({top:0,behavior:'smooth'});}));
$('confidence').addEventListener('input',renderHistogram);$('bins').addEventListener('change',renderHistogram);$('reset-hist').addEventListener('click',()=>{$('confidence').value=80;$('bins').value=30;renderHistogram();});$('fx-shock').value=C.fx_shock_pct;$('fx-shock').addEventListener('input',renderFX);renderFX();$('export-watchlist').addEventListener('click',()=>exportCSV(D.actions.slice(0,10),'q4_2026_executive_watchlist.csv'));$('export-inputs').addEventListener('click',()=>exportCSV(D.deals,'q4_2026_model_inputs_eur.csv'));$('print').addEventListener('click',async()=>{try{await renderFunnel();await snapshotFunnel();window.print();}catch(error){funnelError(error);}});
}catch(error){$('fatal').style.display='block';$('fatal').textContent='The dashboard could not render: '+error.message;console.error(error);}

// Uses existing D.deals and D.actions; no new model or forecast calculations.
let ownerMatrixPromise=null, ownerMatrixSelected=null, ownerMatrixResize=null;
function ownerMatrixError(error){
  $('owner-stage-status').textContent='Owner chart unavailable: '+error.message;
  $('owner-stage-status').style.color='#B4460A';
  console.error(error);
}
function ownerMatrixRows(){
  const ownerById=new Map(D.actions.map(r=>[String(r.id),r.owner]));
  const stageIndex=new Map(D.funnel.map((r,i)=>[r.label,i]));
  const groups=new Map();
  for(const deal of D.deals){
    const owner=ownerById.get(String(deal.id)), stage=stageIndex.get(deal.stage);
    if(owner===undefined||stage===undefined)throw new Error('Missing owner or stage for '+deal.id);
    const key=JSON.stringify([owner,deal.stage]);
    if(!groups.has(key))groups.set(key,{owner,stage:deal.stage,x:stage,gross:0,weighted:0,standard:0,winNumerator:0,ids:new Set()});
    const row=groups.get(key);
    row.gross+=deal.gross;row.weighted+=deal.probability;row.standard+=deal.standard;
    row.winNumerator+=deal.gross*deal.assigned;row.ids.add(String(deal.id));
  }
  return [...groups.values()].map(r=>({...r,win:r.gross>0?100*r.winNumerator/r.gross:null,count:r.ids.size}));
}
function ownerMatrixColour(p,darker=false){
  const stops=[[0,[252,228,236]],[.35,[248,201,218]],[.65,[224,69,127]],[1,[194,18,79]]];
  const t=Math.max(0,Math.min(1,p/100));let i=1;
  while(i<stops.length-1&&t>stops[i][0])i++;
  const a=stops[i-1],b=stops[i],f=(t-a[0])/(b[0]-a[0]);
  const rgb=a[1].map((v,j)=>Math.round((v+(b[1][j]-v)*f)*(darker?.78:1)));
  return 'rgb('+rgb.join(',')+')';
}
async function renderOwnerStageMatrix(){
  if(ownerMatrixPromise)return ownerMatrixPromise;
  ownerMatrixPromise=(async()=>{
    if(!window.Plotly)throw new Error('Embedded Plotly library is missing.');
    const rows=ownerMatrixRows(),owners=D.owners.map(r=>r.label),stages=D.funnel.map(r=>r.label);
    if(!rows.length)throw new Error('No Q4 owner-stage combinations.');
    const plot=$('owner-stage-plot'),shell=$('owners'),traces=[];
    const maximum=Math.max(...rows.map(r=>r.weighted),1);
    const top5=rows.slice().sort((a,b)=>b.weighted-a.weighted).slice(0,5);
    const fullEUR=v=>'€'+Number(v).toLocaleString('en-GB',{maximumFractionDigits:0});
    const compactEUR=v=>'€'+(v>=1e6?(v/1e6).toFixed(2)+'m':v>=1000?(v/1000).toFixed(0)+'k':Math.round(v));
    // Preserve the height of the original 820-wide, 31px-per-owner SVG.
    shell.style.aspectRatio='820 / '+(50+31*owners.length);
    const geometry=()=>{
      const width=shell.clientWidth||Math.max(300,document.querySelector('.sheet').clientWidth-64);
      const h=Math.max(100,plot.clientHeight||width*(50+31*owners.length)/820-96);
      const rowHeight=Math.max(1,(h-80)/owners.length);
      return {diameter:Math.min(58,rowHeight*.82),height:h};
    };
    let geo=geometry();
    const sizeRef=()=>2*maximum/(geo.diameter**2);
    const hover='<b>%{customdata[0]}</b><br>%{customdata[1]}<br><br>'+
      'Gross pipeline: %{customdata[2]}<br>Weighted Forecast: %{customdata[3]}<br>'+
      'Standard Forecast: %{customdata[4]}<br>Win Probability: %{customdata[5]}<br>'+
      'Opportunity Count: %{customdata[6]}<extra></extra>';
    owners.forEach((owner,y)=>{
      const points=rows.filter(r=>r.owner===owner).sort((a,b)=>a.x-b.x);
      if(!points.length)return;
      const xs=[],ys=[];
      for(let j=0;j<points.length-1;j++){
        const a=points[j].x,b=points[j+1].x;
        // Gentle Bezier-like arcs stay inside the owner's own row.
        for(let k=0;k<=20;k++){const t=k/20;xs.push(a+(b-a)*t);ys.push(y-.13*4*t*(1-t));}
      }
      if(xs.length){
        traces.push({type:'scatter',mode:'lines',x:xs,y:ys,showlegend:false,hoverinfo:'skip',
          line:{color:'rgba(194,18,79,.23)',width:1.1,shape:'spline',smoothing:.8},meta:{owner,kind:'flow'}});
      }
      traces.push({type:'scatter',mode:'markers',name:owner,showlegend:false,
        x:points.map(r=>r.x),y:points.map(()=>y),meta:{owner,kind:'bubble'},
        customdata:points.map(r=>[E(owner),E(r.stage),fullEUR(r.gross),fullEUR(r.weighted),fullEUR(r.standard),r.win===null?'Unavailable':r.win.toFixed(1)+'%',r.count]),
        marker:{size:points.map(r=>r.weighted),sizemode:'area',sizeref:sizeRef(),sizemin:2,
          color:points.map(r=>r.win??0),cmin:0,cmax:100,opacity:.95,
          colorscale:[[0,'#FCE4EC'],[.35,'#F8C9DA'],[.65,'#E0457F'],[1,'#C2124F']],
          line:{width:1,color:points.map(r=>ownerMatrixColour(r.win??0,true))}},hovertemplate:hover});
    });
    // A non-data trace supplies a colour scale fixed at 0–100%.
    traces.push({type:'scatter',mode:'markers',x:[null],y:[null],showlegend:false,hoverinfo:'skip',
      marker:{color:[0],cmin:0,cmax:100,colorscale:[[0,'#FCE4EC'],[.35,'#F8C9DA'],[.65,'#E0457F'],[1,'#C2124F']],showscale:true,
        colorbar:{title:{text:'Win<br>probability',font:{size:10}},tickvals:[0,25,50,75,100],ticksuffix:'%',
          tickfont:{size:10},thickness:9,len:.45,outlinewidth:0,x:1.015,y:.74}}});
    const annotations=()=>top5.map((r,i)=>({x:r.x,y:owners.indexOf(r.owner),text:compactEUR(r.weighted),
      showarrow:false,yshift:geo.diameter/2+9,font:{size:10,color:'#7D0B3F',family:'IBM Plex Sans, Arial, sans-serif'},
      bgcolor:'rgba(255,255,255,.92)',borderpad:2}));
    const margin={l:132,r:88,t:18,b:62};
    const layout={autosize:true,height:geo.height,paper_bgcolor:'#fff',plot_bgcolor:'#fff',margin,
      font:{family:'IBM Plex Sans, Arial, sans-serif',size:11,color:'#5E3A4C'},
      hoverlabel:{bgcolor:'#fff',bordercolor:'#E9D3DD',font:{family:'IBM Plex Sans, Arial, sans-serif',size:12,color:'#2A0A1A'}},
      hovermode:'closest',clickmode:'event',dragmode:false,showlegend:false,annotations:annotations(),
      xaxis:{range:[-.6,stages.length-.4],tickmode:'array',tickvals:stages.map((_,i)=>i),
        ticktext:stages.map(s=>E(s).replace(/ /,'<br>')),fixedrange:true,showgrid:true,gridcolor:'#F6EDF1',
        zeroline:false,showline:false,ticks:'',tickfont:{size:10},title:{text:'Sales stage',font:{size:11},standoff:8}},
      yaxis:{range:[owners.length-.45,-.55],tickmode:'array',tickvals:owners.map((_,i)=>i),
        ticktext:owners.map(s=>E(s.length>20?s.slice(0,19)+'…':s)),fixedrange:true,showgrid:false,
        zeroline:false,showline:false,ticks:'',tickfont:{size:11},title:{text:'Opportunity owner',font:{size:11},standoff:8}}};
    await Plotly.newPlot(plot,traces,layout,{responsive:true,displayModeBar:false,displaylogo:false});
    function drawSizeKey(){
      const values=[maximum/16,maximum/4,maximum],names=['Small','Medium','Large'];
      $('owner-size-key').innerHTML=values.map((value,i)=>{
        const diameter=geo.diameter*Math.sqrt(value/maximum);
        return '<span class="owner-size-item"><span class="owner-size-circle" style="width:'+diameter+'px;height:'+diameter+'px"></span><span>'+names[i]+'<br>'+compactEUR(value)+'</span></span>';
      }).join('');
    }
    function selectOwner(owner){
      ownerMatrixSelected=owner===ownerMatrixSelected?null:owner;
      const indices=traces.map((t,i)=>t.meta?i:null).filter(i=>i!==null);
      const opacities=indices.map(i=>!ownerMatrixSelected||traces[i].meta.owner===ownerMatrixSelected?1:.16);
      Plotly.restyle(plot,{opacity:opacities},indices).catch(ownerMatrixError);
      Plotly.relayout(plot,{annotations:annotations().map((a,i)=>({...a,opacity:!ownerMatrixSelected||top5[i].owner===ownerMatrixSelected?1:.16}))}).catch(ownerMatrixError);
      $('owner-stage-status').textContent=ownerMatrixSelected?'Highlighted: '+ownerMatrixSelected+' · click again or reset to show all owners.':'Click a bubble to highlight its owner.';
    }
    plot.on('plotly_click',event=>{
      const trace=event.points?.[0]?.data;
      if(trace?.meta?.kind==='bubble')selectOwner(trace.meta.owner);
    });
    $('owner-reset').addEventListener('click',()=>{ownerMatrixSelected=null;selectOwner(null);});
    let resizeFrame=0;
    ownerMatrixResize=()=>{
      if(!shell.clientWidth)return;
      cancelAnimationFrame(resizeFrame);
      resizeFrame=requestAnimationFrame(async()=>{
        try{
          geo=geometry();
          const indices=traces.map((t,i)=>t.meta?.kind==='bubble'?i:null).filter(i=>i!==null);
          await Plotly.restyle(plot,{'marker.sizeref':sizeRef()},indices);
          await Plotly.relayout(plot,{height:geo.height,annotations:annotations().map((a,i)=>({...a,opacity:!ownerMatrixSelected||top5[i].owner===ownerMatrixSelected?1:.16}))});
          await Plotly.Plots.resize(plot);drawSizeKey();
        }catch(error){ownerMatrixError(error);}
      });
    };
    if(window.ResizeObserver)new ResizeObserver(ownerMatrixResize).observe(shell);
    else window.addEventListener('resize',ownerMatrixResize);
    drawSizeKey();ownerMatrixResize();return plot;
  })();
  ownerMatrixPromise.catch(ownerMatrixError);return ownerMatrixPromise;
}
// Initialise after the existing dashboard code, and resize when its tab opens.
const previousActivate=activate;
activate=function(id){
  previousActivate(id);
  if(id==='composition')renderOwnerStageMatrix().then(()=>ownerMatrixResize?.()).catch(()=>{});
};
renderOwnerStageMatrix().catch(()=>{});


// Executive-only controls: existing summaries and reliability controls are unchanged.
const execMonthly=D.executive_monthly;
const execCumSorted=execMonthly.cumulative_samples.map(a=>a.slice().sort((x,y)=>x-y));
const execMonthSorted=execMonthly.cumulative_samples.map((a,i)=>
  a.map((v,j)=>Math.max(0,v-(i?execMonthly.cumulative_samples[i-1][j]:0))).sort((x,y)=>x-y));
function execQuantile(a,p){
  const v=(a.length-1)*p,i=Math.floor(v);
  return a[i]+(a[Math.min(i+1,a.length-1)]-a[i])*(v-i);
}
function execIntervalLabel(p){return 'P'+(p*100).toFixed(1).replace(/\.0$/,'');}
function execRange(a,level){const tail=(1-level/100)/2;return {low:execQuantile(a,tail),median:execQuantile(a,.5),high:execQuantile(a,1-tail)};}
function execCumulative(a){let total=0;return a.map(v=>total+=v);}
function execPath(points){return points.map(([x,y],i)=>(i?'L':'M')+x.toFixed(2)+','+y.toFixed(2)).join(' ');}

function renderExecCumulative(ranges,weighted,level,lowLabel,highLabel){
  const w=820,h=420,left=57,top=22,pw=620,ph=310;
  // Fixed vertical scale keeps interval changes visually comparable.
  const upper=execCumSorted.map(a=>execQuantile(a,.975));
  const max=nice(Math.max(...upper,K.probability,1)*1.08);
  const X=i=>left+pw*i/3,Y=v=>top+ph-ph*v/max;
  const low=[0,...ranges.map(r=>r.low)],high=[0,...ranges.map(r=>r.high)];
  const med=[0,...ranges.map(r=>r.median)],base=[0,...weighted];
  let b='';
  for(let i=0;i<=4;i++){
    const v=max*i/4,y=Y(v);
    b+=line(left,y,left+pw,y)+txt(left-9,y+4,v?'€'+(v/1e6).toLocaleString('en-GB',{maximumFractionDigits:1})+'m':'0','end');
  }
  const polygon=low.map((v,i)=>[X(i),Y(v)]).concat(high.map((v,i)=>[X(i),Y(v)]).reverse());
  b+=`<path d="${execPath(polygon)} Z" fill="#F8C9DA" fill-opacity=".8" stroke="#E0457F" stroke-width="1"/>`;
  b+=`<path d="${execPath(base.map((v,i)=>[X(i),Y(v)]))}" fill="none" stroke="#7D0B3F" stroke-width="1.6" stroke-dasharray="5 4"/>`;
  b+=`<path d="${execPath(med.map((v,i)=>[X(i),Y(v)]))}" fill="none" stroke="#C2124F" stroke-width="2.6"/>`;
  ['1 Oct','31 Oct','30 Nov','31 Dec'].forEach((label,i)=>{
    b+=line(X(i),top+ph,X(i),top+ph+5,ink)+txt(X(i),top+ph+24,label,'middle');
    if(i){
      const r=ranges[i-1],tip=`${execMonthly.month_ends[i-1]}: ${level}% cumulative range ${M(r.low)}–${M(r.high)}; median ${M(r.median)}; probability weighted ${M(weighted[i-1])}`;
      b+=mark(`<rect x="${X(i)-13}" y="${Y(r.high)-7}" width="26" height="${Math.max(16,Y(r.low)-Y(r.high)+14)}" fill="#fff" fill-opacity="0" pointer-events="all"/>`+`<circle cx="${X(i)}" cy="${Y(r.median)}" r="4" fill="#C2124F" stroke="#fff" stroke-width="1.5"/>`,tip);
    }
  });
  const last=ranges[2],labels=[
    {v:last.high,title:'HIGH · '+highLabel,value:M(last.high),color:'#8A6B7A'},
    {v:last.median,title:'MEDIAN',value:M(last.median),color:'#C2124F'},
    {v:weighted[2],title:'WEIGHTED',value:M(weighted[2]),color:'#7D0B3F'},
    {v:last.low,title:'LOW · '+lowLabel,value:M(last.low),color:'#8A6B7A'}
  ].sort((a,b)=>b.v-a.v);
  labels.forEach((r,i)=>r.y=Math.max(top+10,Y(r.v),i?labels[i-1].y+34:top+10));
  const overflow=Math.max(0,labels[labels.length-1].y-(top+ph-8));
  labels.forEach(r=>{
    r.y-=overflow;
    b+=line(X(3)+4,Y(r.v),X(3)+18,r.y,r.color,'stroke-width=".7"');
    b+=`<text x="${X(3)+22}" y="${r.y-3}" class="exec-end-title">${E(r.title)}</text>`;
    b+=`<text x="${X(3)+22}" y="${r.y+11}" class="exec-end-label" style="fill:${r.color}">${E(r.value)}</text>`;
  });
  svg('exec-cumulative-fan','Cumulative Q4 delivery forecast range',`${level}% central simulated range; hover month-end markers for values.`,b,w,h);
}

function execRangeTable(ranges,weighted,cumulative,lowLabel,highLabel){
  const th=s=>'<th scope="col">'+E(s)+'</th>';
  return '<table><thead><tr class="grp"><th class="blank"></th><th colspan="3">Simulated outcome</th><th colspan="2">Probability weighted</th><th class="blank"></th></tr><tr>'+th(cumulative?'Month end':'Month')+
    th('Low ('+lowLabel+')')+'<th scope="col" class="hl">Median</th>'+th('High ('+highLabel+')')+th(cumulative?'Cumulative':'In month')+th(cumulative?'In month':'Share of Q4')+th('Range width')+'</tr></thead><tbody>'+
    ranges.map((r,i)=>`<tr class="${cumulative&&i===2?'qe':''}"><td>${E(cumulative?execMonthly.month_ends[i]+(i===2?' · quarter end':''):execMonthly.months[i])}</td><td>${N(r.low/1000)}</td><td class="hl">${N(r.median/1000)}</td><td>${N(r.high/1000)}</td><td>${N(weighted[i]/1000)}</td><td>${cumulative?N(execMonthly.weighted[i]/1000):P(execMonthly.weighted[i]/K.probability*100)}</td><td>${N((r.high-r.low)/1000)}</td></tr>`).join('')+'</tbody></table>';
}

function renderExecForecastExtension(){
  try{
    const level=Number($('exec-interval').value),tail=(1-level/100)/2;
    const lowLabel=execIntervalLabel(tail),highLabel=execIntervalLabel(1-tail);
    const cumulative=execCumSorted.map(a=>execRange(a,level));
    const monthly=execMonthSorted.map(a=>execRange(a,level));
    const weighted=execCumulative(execMonthly.weighted),last=cumulative[2];
    const width=last.high-last.low,delta=last.median-K.probability;
    const december=execMonthly.weighted[2]/K.probability*100;
    $('exec-interval-value').textContent=level+'%';
    document.querySelectorAll('[data-exec-interval]').forEach(b=>b.setAttribute('aria-pressed',String(Number(b.dataset.execInterval)===level)));
    $('exec-band-label').textContent=`Central ${level}% range (${lowLabel}–${highLabel})`;
    renderExecCumulative(cumulative,weighted,level,lowLabel,highLabel);
    $('exec-fan-note').textContent=`Shaded area contains the central ${level}% of ${N(C.simulations)} modelled outcomes at each month end. It is a pointwise outcome range, not a guarantee that an entire delivery path stays inside the band. Hover or focus a month-end marker for values.`;
    $('exec-dynamic-summary').innerHTML=[
      `<strong>Base forecast ${M(K.probability)}.</strong> The probability weighted forecast represents <span class="n">${P(K.probability/K.gross*100)}</span> of the ${M(K.gross)} gross Q4 pipeline.`,
      `<strong>Quarter-end range.</strong> The central ${level}% of modelled outcomes falls between <span class="n">${M(last.low)}</span> and <span class="n">${M(last.high)}</span>. Range width is ${M(width)}${last.median>0?' ('+P(width/last.median*100)+' of the median)':''}.`,
      `<strong>December delivery.</strong> <span class="n">${P(december)}</span> of the weighted forecast (${M(execMonthly.weighted[2])}) is scheduled for December. ${december>=50?'Confirm the delivery dates of the largest December opportunities.':'Review monthly delivery dates alongside the largest opportunities.'}`,
      `<strong>Model versus forecast.</strong> The simulated median is <span class="n">${M(last.median)}</span>, ${M(Math.abs(delta))} ${delta>=0?'above':'below'} the current forecast. ${P(K.reach_forecast_pct)} of simulations reach that forecast. The model includes the existing historical calibration and slippage assumptions.`
    ].map(t=>'<li>'+t+'</li>').join('');
    $('exec-month-phasing').innerHTML=execMonthly.months.map((month,i)=>{
      const share=execMonthly.weighted[i]/K.probability*100;
      return `<div class="prow"><span>${E(month)}</span><div class="pbar"><i style="width:${share}%"></i></div><span class="v">${M(execMonthly.weighted[i])}</span><span class="p">${Math.round(share)}%</span></div>`;
    }).join('');
    $('exec-table-unit').textContent=`EUR thousands · central ${level}% outcome range (${lowLabel}–${highLabel})`;
    $('exec-cumulative-table').innerHTML=execRangeTable(cumulative,weighted,true,lowLabel,highLabel);
    $('exec-monthly-table').innerHTML=execRangeTable(monthly,execMonthly.weighted,false,lowLabel,highLabel);
    $('exec-month-method').textContent=execMonthly.method;
  }catch(error){
    $('exec-extension-error').hidden=false;
    $('exec-extension-error').textContent='The monthly forecast section could not render: '+error.message;
    console.error(error);
  }
}
$('exec-interval').addEventListener('input',renderExecForecastExtension);
document.querySelectorAll('[data-exec-interval]').forEach(button=>button.addEventListener('click',()=>{
  $('exec-interval').value=button.dataset.execInterval;renderExecForecastExtension();
}));
renderExecForecastExtension();


// Section 04 uses the existing model; no scenario assumptions are changed.
const BR=D.forecast_risks;
const boardRiskLabels={concentration:'Concentration risk',slippage:'Slippage risk',timing:'Timing risk',stale:'Stale opportunity',inflation:'Probability inflation'};
let boardRiskFocus=null,boardRiskDeal=null,boardRiskSort='exposure',boardRiskDescending=true,boardRiskPlotsPromise=null;
const boardRiskEUR=v=>'€'+Number(v).toLocaleString('en-GB',{maximumFractionDigits:0});
function boardRiskError(error){$('board-risk-error').hidden=false;$('board-risk-error').textContent='Forecast risks could not render: '+error.message;console.error(error);}
function boardRiskSeverity(value,amber,red){return value>=red?2:value>=amber?1:0;}
function boardRiskSelected(r){return boardRiskDeal?r.id===boardRiskDeal:boardRiskFocus?r.tags.includes(boardRiskFocus):false;}

function renderBoardRiskCards(){
  const cards=[
    {key:'single',name:'Single Deal Dependency',value:P(BR.largest/K.probability*100),amount:M(BR.largest)+' of forecast',detail:BR.records[0].name,severity:boardRiskSeverity(BR.largest/K.probability*100,10,20)},
    {key:'timing',name:'December Concentration',value:P(BR.december/K.probability*100),amount:M(BR.december)+' scheduled for December',detail:'Click to highlight deals with ≥50% of Q4 delivery in December',severity:boardRiskSeverity(BR.december/K.probability*100,40,60)},
    {key:'inflation',name:'Forecast Optimism',value:M(BR.optimism_gap),amount:P(BR.optimism_pct)+' of weighted forecast',detail:'Weighted '+M(K.probability)+' versus model '+M(K.model_mean),severity:boardRiskSeverity(BR.optimism_pct,10,20)},
    {key:'slippage',name:'Slippage Exposure',value:P(BR.slippage_exposure_pct),amount:M(BR.slippage_exposure)+' of forecast',detail:'Deals with recorded delay likelihood ≥40%',severity:boardRiskSeverity(BR.slippage_exposure_pct,25,50)}
  ];
  const status=['Low exposure','Review','High risk'],classes=['risk-low','risk-review','risk-high'],symbols=['●','◆','▲'];
  $('board-risk-cards').innerHTML=cards.map(c=>`<button type="button" class="board-risk-card" data-board-risk="${c.key}" aria-pressed="${boardRiskFocus===c.key}" title="Highlight related opportunities in the watchlist"><span class="risk-name">${E(c.name)}</span><span class="risk-number">${E(c.value)}</span><span class="risk-amount">${E(c.amount)}</span><span class="risk-detail">${E(c.detail)}</span><span class="risk-rag ${classes[c.severity]}">${symbols[c.severity]} ${status[c.severity]}</span></button>`).join('');
  document.querySelectorAll('[data-board-risk]').forEach(button=>button.addEventListener('click',()=>{
    boardRiskDeal=null;boardRiskFocus=boardRiskFocus===button.dataset.boardRisk?null:button.dataset.boardRisk;
    renderBoardRiskCards();renderBoardRiskWatchlist();
  }));
}

function renderBoardRiskWatchlist(){
  const active=Boolean(boardRiskFocus||boardRiskDeal);
  let rows=BR.records.filter(r=>r.tags.some(t=>t!=='single')||boardRiskSelected(r));
  const allCount=rows.length;
  const riskText=r=>r.tags.filter(t=>boardRiskLabels[t]).map(t=>boardRiskLabels[t]).join(' · ');
  const sortValue=r=>boardRiskSort==='risk_type'?riskText(r):boardRiskSort==='status'?r.severity:r[boardRiskSort];
  rows.sort((a,b)=>{
    if(active){const priority=Number(boardRiskSelected(b))-Number(boardRiskSelected(a));if(priority)return priority;}
    const av=sortValue(a),bv=sortValue(b);
    const comparison=typeof av==='number'?av-bv:String(av).localeCompare(String(bv));
    return (boardRiskDescending?-comparison:comparison)||a.rank-b.rank;
  });
  rows=rows.slice(0,15);
  const columns=[['name','Opportunity','25%'],['owner','Owner','14%'],['risk_type','Risk type','20%'],['exposure','Exposure EUR','12%'],['probability','Probability','9%'],['days','Days since update','10%'],['status','Status','10%']];
  const classes=['risk-low','risk-review','risk-high'];
  $('board-risk-watchlist').innerHTML='<table class="board-risk-table"><colgroup>'+columns.map(c=>`<col style="width:${c[2]}">`).join('')+'</colgroup><thead><tr>'+columns.map(c=>`<th scope="col" aria-sort="${boardRiskSort===c[0]?(boardRiskDescending?'descending':'ascending'):'none'}"><button type="button" data-board-risk-sort="${c[0]}">${E(c[1])}${boardRiskSort===c[0]?(boardRiskDescending?' ↓':' ↑'):''}</button></th>`).join('')+'</tr></thead><tbody>'+rows.map(r=>`<tr class="${active?(boardRiskSelected(r)?'risk-selected':'risk-muted'):''}"><td><b>${E(r.name)}</b><small>${E(r.id)}</small></td><td>${E(r.owner)}</td><td>${E(riskText(r)||'Portfolio dependency')}</td><td class="number" title="${E(boardRiskEUR(r.exposure))}">${boardRiskEUR(r.exposure)}</td><td class="number" title="Assigned ${P(r.probability*100)}; historic ${P(r.historic*100)}; delay likelihood ${P(r.slippage*100)}">${P(r.probability*100)}</td><td class="number">${N(r.days)}</td><td><span class="risk-status ${classes[r.severity]}">${E(r.status)}</span></td></tr>`).join('')+'</tbody></table>';
  document.querySelectorAll('[data-board-risk-sort]').forEach(button=>button.addEventListener('click',()=>{
    const field=button.dataset.boardRiskSort;
    boardRiskDescending=field===boardRiskSort?!boardRiskDescending:['exposure','probability','days','status'].includes(field);
    boardRiskSort=field;renderBoardRiskWatchlist();
  }));
  const focusLabel=boardRiskDeal?BR.records.find(r=>r.id===boardRiskDeal)?.name:boardRiskFocus?({single:'Largest opportunity',timing:'December timing risk',inflation:'Probability inflation',slippage:'Slippage exposure'}[boardRiskFocus]):null;
  const selected=rows.filter(boardRiskSelected).length;
  $('board-risk-selection').textContent=`Showing ${rows.length} of ${allCount} opportunities with risk flags${focusLabel?' · '+focusLabel+': '+selected+' highlighted; related rows listed first.':'. Select a risk card or deal to highlight related rows.'}`;
}

function renderBoardRiskDefinitions(){
  $('board-risk-definitions').innerHTML=`<p><b>Measures.</b> Largest deal, Top 3 and Top 5 use probability weighted Q4 revenue. December exposure follows the validated Monday–Friday delivery allocation. Slippage exposure is the full weighted value of deals with a recorded delay probability of at least 40%; it is distinct from the existing weighted average delay probability. Forecast optimism = weighted forecast minus model expectation; its percentage uses weighted forecast as denominator.</p><p><b>Watchlist flags.</b> Concentration: Top 5 deals or individual share ≥10%. Timing: ≥50% of the deal's Q4 delivery falls in December. Stale: older than ${C.stale_after_days} days at the snapshot. Probability inflation: assigned probability exceeds historic stage win rate by more than 10 percentage points.</p><p><b>Card RAG thresholds (amber / red).</b> Single deal: 10% / 20%. December concentration: 40% / 60%. Forecast optimism: 10% / 20%. Slippage exposure: 25% / 50%. These are review thresholds, not calibrated statistical limits.</p><p><b>Row status.</b> High risk: individual forecast share ≥20%; delay probability ≥60%; update age >60 days; probability inflation >20pp; or ≥75% of delivery in December with ≥5% forecast share. Other flagged deals are Review; unflagged deals are Low exposure.</p>`;
}

function renderBoardRiskPlots(){
  if(boardRiskPlotsPromise)return boardRiskPlotsPromise;
  boardRiskPlotsPromise=(async()=>{
    if(!window.Plotly)throw new Error('Embedded Plotly library is missing.');
    const w=BR.waterfall;
    const names=['Weighted<br>forecast','Probability<br>calibration','Slippage<br>risk','Timing<br>risk','Other<br>adjustments','Model expected<br>outcome'];
    const values=[w.weighted,w.calibration,w.full_deferral,w.partial_timing,w.other,w.model];
    const compact=v=>(v<0?'−':v>0?'+':'')+'€'+(Math.abs(v)/1e6).toFixed(2)+'m';
    const hover=[
      'Source probability weighted Q4 forecast',
      'Net change from blending assigned and historical probabilities',
      'Expected loss on slipped deals with no retained Q4 delivery',
      'Expected loss on slipped deals retaining only part of Q4 delivery',
      'No additional model adjustments',
      'Existing calibrated, slippage-adjusted model expectation'
    ];
    const font={family:'IBM Plex Sans, Arial, sans-serif',size:11,color:'#5E3A4C'};
    await Plotly.newPlot($('board-risk-waterfall'),[{
      type:'waterfall',orientation:'v',measure:['absolute','relative','relative','relative','relative','total'],
      x:names,y:values.map((v,i)=>i===5?0:v/1e6),customdata:hover,
      text:values.map((v,i)=>i===0||i===5?M(v):compact(v)),textposition:'outside',cliponaxis:false,
      connector:{line:{color:'#C7A9B8',width:1,dash:'dot'}},
      decreasing:{marker:{color:'#E0457F',line:{color:'#C2124F',width:.5}}},
      increasing:{marker:{color:'#0B6E8A'}},totals:{marker:{color:'#7D0B3F'}},
      hovertemplate:'%{customdata}<br>%{text}<extra></extra>'
    }],{
      height:375,autosize:true,paper_bgcolor:'#fff',plot_bgcolor:'#fff',font,
      margin:{l:51,r:14,t:32,b:75},showlegend:false,waterfallgap:.35,
      xaxis:{fixedrange:true,tickfont:{size:10},showgrid:false},
      yaxis:{fixedrange:true,tickprefix:'€',ticksuffix:'m',rangemode:'tozero',gridcolor:'#E9D3DD',zerolinecolor:'#E9D3DD'},
      hoverlabel:{bgcolor:'#fff',bordercolor:'#E9D3DD',font},dragmode:false
    },{responsive:true,displayModeBar:false,displaylogo:false});
    $('board-risk-waterfall-note').textContent=`${M(w.weighted)} weighted forecast reconciles to ${M(w.model)} model expectation. Slippage risk = full Q4 deferrals; timing risk = partial Q4 delivery losses. These mutually exclusive components split the existing ${M(-(w.full_deferral+w.partial_timing))} delivery adjustment. Other adjustments are zero.`;
    const records=BR.records;
    const palette=records.map(r=>r.rank===1?'#C2124F':r.rank<=3?'#E0457F':r.rank<=5?'#F8C9DA':'#FCE4EC');
    await Plotly.newPlot($('board-risk-wheel'),[{
      type:'pie',labels:records.map(r=>E(r.name)),values:records.map(r=>r.exposure),ids:records.map(r=>r.id),
      hole:.72,sort:false,direction:'clockwise',rotation:-90,textinfo:'none',showlegend:false,
      marker:{colors:palette,line:{color:'#fff',width:.7}},
      customdata:records.map(r=>[E(r.owner),boardRiskEUR(r.exposure),P(r.share_pct),r.id]),
      hovertemplate:'<b>%{label}</b><br>Owner: %{customdata[0]}<br>Forecast: %{customdata[1]}<br>Contribution: %{customdata[2]}<extra></extra>'
    }],{
      height:375,autosize:true,paper_bgcolor:'#fff',plot_bgcolor:'#fff',font,
      margin:{l:8,r:8,t:12,b:12},showlegend:false,
      annotations:[{xref:'paper',yref:'paper',x:.5,y:.53,text:P(BR.largest/K.probability*100),showarrow:false,font:{...font,size:31,color:'#C2124F'}},
                   {xref:'paper',yref:'paper',x:.5,y:.41,text:'Largest deal',showarrow:false,font:{...font,size:12}}],
      hoverlabel:{bgcolor:'#fff',bordercolor:'#E9D3DD',font}
    },{responsive:true,displayModeBar:false,displaylogo:false});
    $('board-risk-wheel-key').innerHTML=[['Largest deal',BR.largest,'#C2124F'],['Top 3 deals',BR.top3,'#E0457F'],['Top 5 deals',BR.top5,'#F8C9DA'],['Remaining portfolio',BR.remaining,'#FCE4EC']].map(([label,value,color])=>`<div><span class="risk-wheel-swatch" style="background:${color}"></span>${E(label)}<br><b>${P(value/K.probability*100)}</b> · ${M(value)}</div>`).join('');
    $('board-risk-wheel').on('plotly_click',event=>{
      const id=event.points?.[0]?.customdata?.[3];if(!id)return;
      boardRiskFocus=null;boardRiskDeal=boardRiskDeal===id?null:id;
      renderBoardRiskCards();renderBoardRiskWatchlist();
    });
    return ['board-risk-waterfall','board-risk-wheel'];
  })();
  boardRiskPlotsPromise.catch(boardRiskError);return boardRiskPlotsPromise;
}

const activateBeforeBoardRisks=activate;
activate=function(id){
  activateBeforeBoardRisks(id);
  if(id==='risk')requestAnimationFrame(()=>renderBoardRiskPlots().then(ids=>Promise.all(ids.map(id=>Plotly.Plots.resize($(id))))).catch(()=>{}));
};
$('board-risk-reset').addEventListener('click',()=>{
  boardRiskFocus=null;boardRiskDeal=null;renderBoardRiskCards();renderBoardRiskWatchlist();
});
try{renderBoardRiskCards();renderBoardRiskWatchlist();renderBoardRiskDefinitions();renderBoardRiskPlots().catch(()=>{});}catch(error){boardRiskError(error);}

</script>
<script data-goatcounter="https://terrijackson.goatcounter.com/count" async src="https://gc.zgo.at/count.js"></script>
</body></html>
'''


def write_dashboard(data, output):
    # Escaping prevents workbook labels containing </script> from ending the data block.
    payload = json.dumps(data, ensure_ascii=False, allow_nan=False).replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    try:
        from plotly.offline import get_plotlyjs
    except ImportError as exc:
        raise ValueError('Plotly is missing. In Colab run: !pip install -q "plotly>=5.24,<7" and rerun this script.') from exc
    # Include the installed Plotly JS bundle; no external chart CDN is required.
    bundle = re.sub(r"</script", r"<\\/script", get_plotlyjs(), flags=re.IGNORECASE)
    text = HTML_TEMPLATE.replace("__DATA__", payload).replace("__PLOTLY_BUNDLE__", bundle)
    Path(output).write_text(text, encoding="utf-8")
    return Path(output)


def main():
    try:
        from google.colab import files
        in_colab = True
    except ImportError:
        in_colab = False
    if in_colab:
        section_heading("1. Upload the analysed Salesforce sales pipeline report")
        print("Upload the Excel report containing Analysis and FX. Its saved formula results will be checked before modelling.")
        uploaded = files.upload()
        if len(uploaded) != 1:
            raise ValueError("Please upload exactly one .xlsx Salesforce workbook and run this cell again.")
        filename, content = next(iter(uploaded.items()))
        if not filename.lower().endswith(".xlsx"):
            raise ValueError("Please upload an .xlsx workbook.")
        output = Path("fuchsia_works_q4_2026_board_dashboard.html")
    else:
        parser = argparse.ArgumentParser(description=__doc__)
        parser.add_argument("workbook", help="Analysed Salesforce .xlsx report")
        parser.add_argument("--output", default="fuchsia_works_q4_2026_board_dashboard.html")
        args = parser.parse_args()
        filename = Path(args.workbook).name; content = Path(args.workbook).read_bytes(); output = Path(args.output)
    section_heading("2. Validate opportunity IDs, formula results, Q4 scope and EUR allocations")
    q, meta = load_pipeline(content); meta["filename"] = filename
    print(f"Validated {meta['source_rows']:,} records; {len(q):,} open opportunities contribute to Q4. Snapshot: {meta['snapshot']}.")
    section_heading("3. Simulate Q4 delivery and calculate forecast reliability and risk drivers")
    data = build_model(q, meta)
    print(f"Running {CONFIG['simulations']:,} simulations, seed {CONFIG['random_seed']}. Model choices are documented in Assumptions.")
    section_heading("4. Create the seven-section Fuchsia Works board dashboard")
    write_dashboard(data, output)
    section_heading("5. Dashboard ready — choose whether to download")
    if in_colab:
        from IPython.display import display, HTML
        # A user-activated link is deliberately used instead of files.download().
        # Embed the file bytes so the link remains usable after the runtime stops.
        encoded = base64.b64encode(output.read_bytes()).decode("ascii")
        display(HTML('<p>Open the downloaded HTML file in your browser. Charts work offline.</p>'
                     '<a download="' + html.escape(output.name, quote=True) + '" href="data:text/html;base64,' + encoded + '" '
                     'style="display:inline-block;padding:10px 16px;border:1px solid #2A0A1A;color:#C2124F;text-decoration:none;font-family:Arial">'
                     'Download Q4-2026 board dashboard (.html)</a>'))
        print("Your browser downloads the file only when you click the link above. It is also available in Colab's Files panel.")
    else:
        print(f"Created {output.resolve()}")


if __name__ == "__main__":
    try:
        main()
    except (ValueError, OSError, KeyError) as exc:
        message = f"Dashboard not created: {exc}"
        if "google.colab" in sys.modules:
            from IPython.display import display, HTML
            display(HTML('<p style="color:#B4460A;font-family:Arial">' + html.escape(message) + '</p>'))
        else:
            print(message, file=sys.stderr)
        raise
