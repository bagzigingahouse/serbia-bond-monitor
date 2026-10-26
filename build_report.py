#!/usr/bin/env python3
"""Build a public Serbian government bond secondary-trading report from UJD XLSX.

Uses only the Python standard library. Runs on GitHub Actions; no local setup.
"""
import argparse
import collections
from datetime import date, datetime, timedelta, timezone
import io
import json
from pathlib import Path
import re
import urllib.request
import zipfile
from xml.etree import ElementTree as ET

SOURCE = "https://javnidug.gov.rs/static/uploads/Sekundarno%20trgovanje.xlsx"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
BASE = date(1899, 12, 30)


def excel_date(raw):
    if not raw:
        return None
    try:
        return BASE + timedelta(days=int(float(raw)))
    except (ValueError, OverflowError):
        for pattern in ("%d/%m/%Y", "%Y-%m-%d", "%d.%m.%Y"):
            try:
                return datetime.strptime(raw.strip(), pattern).date()
            except ValueError:
                continue
    return None


def number(value):
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def percent(value):
    val = number(value)
    if val is None:
        return None
    return val * 100 if abs(val) <= 1 else val


def read_xlsx(binary):
    """Read Excel 2007+ cells without external dependencies or Excel installation."""
    with zipfile.ZipFile(io.BytesIO(binary)) as archive:
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            strings = ["".join(t.text or "" for t in item.iter(NS + "t"))
                       for item in root.findall(NS + "si")]
        update_label = None
        trades = []
        with archive.open("xl/worksheets/sheet1.xml") as handle:
            for _, element in ET.iterparse(handle, events=("end",)):
                if element.tag != NS + "row":
                    continue
                row_no = int(element.get("r", "0"))
                cells = {}
                for cell in element.findall(NS + "c"):
                    match = re.match(r"([A-Z]+)", cell.get("r", ""))
                    if not match:
                        continue
                    col = match.group(1)
                    raw_el = cell.find(NS + "v")
                    raw = raw_el.text if raw_el is not None else None
                    if cell.get("t") == "s" and raw is not None:
                        raw = strings[int(raw)]
                    elif cell.get("t") == "inlineStr":
                        raw = "".join(t.text or "" for t in cell.iter(NS + "t"))
                    cells[col] = raw
                if row_no == 8:
                    update_label = cells.get("B", "")
                if row_no <= 10:
                    element.clear()
                    continue
                trade_date = excel_date(cells.get("C"))
                maturity = excel_date(cells.get("H"))
                currency = (cells.get("J") or "").strip().upper()
                ticker = (cells.get("L") or "").strip().upper()
                isin = (cells.get("K") or "").strip().upper()
                bond_type = (cells.get("E") or "").strip()
                amount = number(cells.get("P"))
                ytm = percent(cells.get("Q"))
                price = number(cells.get("O"))
                if (not trade_date or not maturity or not isin or not ticker
                        or currency not in ("RSD", "EUR") or not amount or amount <= 0
                        or ytm is None or not (-10 <= ytm <= 100)):
                    element.clear()
                    continue
                # Government coupon bonds; restitution and T-bills are excluded.
                if not (ticker.startswith("RSOB") or "TBOND" in bond_type.upper().replace(" ", "")):
                    element.clear()
                    continue
                trades.append({
                    "date": trade_date.isoformat(),
                    "isin": isin, "ticker": ticker, "currency": currency,
                    "market": (cells.get("B") or "").strip(),
                    "maturity": maturity.isoformat(),
                    "remaining_years": round((maturity - trade_date).days / 365.25, 3),
                    "coupon_pct": percent(cells.get("I")),
                    "ytm_pct": ytm, "price": price,
                    "quantity": number(cells.get("M")) or 0,
                    "nominal_value": number(cells.get("N")),
                    "turnover": amount,
                })
                element.clear()
        if not trades:
            raise ValueError("No valid government coupon-bond trades found in Excel")
        return trades, update_label


def weighted(rows, field):
    relevant = [r for r in rows if r.get(field) is not None]
    den = sum(r["turnover"] for r in relevant)
    return round(sum(r["turnover"] * r[field] for r in relevant) / den, 5) if den else None


def maturity_bucket(years):
    if years < 1: return "<1y"
    if years < 3: return "1–3y"
    if years < 5: return "3–5y"
    if years < 7: return "5–7y"
    if years <= 10: return "7–10y"
    return ">10y"


def aggregate(trades):
    groups = collections.defaultdict(list)
    for t in trades:
        groups[(t["date"], t["isin"], t["currency"])].append(t)
    result = []
    for (day, isin, curr), rows in groups.items():
        first = rows[0]
        result.append({
            "date": day, "ticker": first["ticker"], "isin": isin,
            "currency": curr, "maturity": first["maturity"],
            "remaining_years": first["remaining_years"],
            "bucket": maturity_bucket(first["remaining_years"]),
            "coupon_pct": first["coupon_pct"],
            "trades": len(rows), "quantity": sum(r["quantity"] for r in rows),
            "turnover": round(sum(r["turnover"] for r in rows), 2),
            "weighted_ytm_pct": weighted(rows, "ytm_pct"),
            "weighted_price": weighted(rows, "price"),
        })
    result.sort(key=lambda r: (r["date"], r["turnover"]), reverse=True)
    previous = {}
    # Find immediately preceding day with execution of the *same ISIN*.
    for row in sorted(result, key=lambda r: (r["date"], r["isin"])):
        key = (row["isin"], row["currency"])
        old = previous.get(key)
        row["previous_trade_date"] = old["date"] if old else None
        row["yield_change_bp"] = (round((row["weighted_ytm_pct"] - old["weighted_ytm_pct"]) * 100, 2)
                                  if old else None)
        previous[key] = row
    return result


def format_money(value, currency):
    if currency == "RSD":
        return f"{value / 1_000_000:,.2f} mil. RSD"
    return f"{value / 1_000:,.1f} hilj. EUR"


def summary_rows(rows):
    grouped = collections.defaultdict(list)
    for r in rows:
        grouped[(r["currency"], r["bucket"])].append(r)
    return [{"currency": cur, "bucket": bucket,
             "total_turnover": round(sum(r["turnover"] for r in vals), 2),
             "transaction_count": sum(r["trades"] for r in vals),
             "weighted_ytm_pct": weighted([{"turnover": r["turnover"], "ytm_pct": r["weighted_ytm_pct"]} for r in vals], "ytm_pct"),
             "distinct_isins": len({r["isin"] for r in vals})}
            for (cur, bucket), vals in sorted(grouped.items())]


def render(trades, update_label):
    aggregates = aggregate(trades)
    dates = sorted({r["date"] for r in aggregates}, reverse=True)
    latest = dates[0]
    # Last five observed trading sessions, not assumed weekdays.
    last_five = set(dates[:5])
    last_ten = set(dates[:10])
    latest_rows = [r for r in aggregates if r["date"] in last_five]
    ten_rows = [r for r in aggregates if r["date"] in last_ten]
    latest_date_rows = [r for r in aggregates if r["date"] == latest]
    last45 = date.fromisoformat(latest) - timedelta(days=45)
    history = [r for r in aggregates if date.fromisoformat(r["date"]) >= last45]
    totals = {cur: {"turnover": round(sum(r["turnover"] for r in latest_rows if r["currency"] == cur), 2),
                    "transactions": sum(r["trades"] for r in latest_rows if r["currency"] == cur)}
              for cur in ("RSD", "EUR")}
    report = {
        "source_url": SOURCE, "source_label": update_label,
        "latest_trade_date": latest,
        "last_five_trading_dates": sorted(last_five, reverse=True),
        "source_transaction_count_coupon_bonds": len(trades),
        "last_five_trading_dates_totals": totals,
        "last_five_trading_dates_series": latest_rows,
        "last_ten_trading_dates_maturity_buckets": summary_rows(ten_rows),
        "last_45_calendar_days_series": history,
    }
    lines = [
        "# Sekundarno trgovanje državnim kuponskim obveznicama Srbije",
        "",
        f"**Poslednji evidentirani datum trgovanja:** {latest}  ",
        f"**Izvorna oznaka ažuriranja:** {update_label or 'nije dostupna'}  ",
        f"**Izvor:** [Uprava za javni dug — zbirni Excel]({SOURCE})  ",
        f"**Obuhvat:** poslednjih 5 datuma sa trgovanjem kuponskim hartijama: {', '.join(sorted(last_five, reverse=True))}.",
        "",
        "Podaci su **realizovana trgovanja**, ne tekuće kupovne/prodajne kotacije. Ponderisani YTM i cena izračunati su prema ostvarenoj vrednosti trgovanja. RSD i EUR se ne sabiraju.",
        "",
        "## Promet poslednjih pet datuma trgovanja",
        "",
    ]
    for cur in ("RSD", "EUR"):
        s = totals[cur]
        lines.append(f"- **{cur}:** {format_money(s['turnover'], cur)} u {s['transactions']} transakcija.")
    lines += ["", "## Serije — poslednjih pet datuma trgovanja", "",
              "| Datum | Valuta | Tiker / ISIN | Dospeće | Preostalo | YTM (%) | Cena | Promet | Poslova | Δ YTM prema prethodnom trgovanju iste serije |",
              "|---|---|---|---|---:|---:|---:|---:|---:|---:|"]
    for row in latest_rows:
        diff = "—" if row["yield_change_bp"] is None else f"{row['yield_change_bp']:+.1f} bp ({row['previous_trade_date']})"
        lines.append("| {date} | {currency} | {ticker} / `{isin}` | {maturity} | {remaining_years:.1f} g. | {weighted_ytm_pct:.3f} | {weighted_price:.2f} | {money} | {trades} | {diff} |".format(
            **row, money=format_money(row["turnover"], row["currency"]), diff=diff))
    lines += ["", "## Kretanje po ročnosti — poslednjih deset datuma trgovanja", "",
              "| Valuta | Preostala ročnost | Promet | Poslova | Različitih ISIN | Ponderisani YTM (%) |",
              "|---|---|---:|---:|---:|---:|"]
    for s in summary_rows(ten_rows):
        lines.append(f"| {s['currency']} | {s['bucket']} | {format_money(s['total_turnover'], s['currency'])} | {s['transaction_count']} | {s['distinct_isins']} | {s['weighted_ytm_pct']:.3f} |")
    lines += ["", "---", "", "Automatski obračun iz zvaničnog XLSX. Ako je promet mali ili postoji samo jedno trgovanje, prikazani prinos nije reprezentativan za celu ročnost. Detaljniji podaci: [`latest.json`](latest.json).", ""]
    return report, "\n".join(lines)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", help="Local XLSX file for offline testing; omit for official download")
    parser.add_argument("--output-dir", default="docs")
    args = parser.parse_args()
    if args.input:
        binary = Path(args.input).read_bytes()
    else:
        request = urllib.request.Request(SOURCE, headers={"User-Agent": "Mozilla/5.0 (compatible; SerbiaBondMonitor/1.0)", "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*"})
        with urllib.request.urlopen(request, timeout=60) as response:
            binary = response.read()
    if len(binary) < 10000 or binary[:2] != b"PK":
        raise RuntimeError(f"Download did not return a valid XLSX ZIP archive ({len(binary)} bytes)")
    trades, label = read_xlsx(binary)
    report, md = render(trades, label)
    output = Path(args.output_dir)
    output.mkdir(parents=True, exist_ok=True)
    (output / "latest.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (output / "latest.md").write_text(md, encoding="utf-8")
    print(f"SUCCESS: {len(trades)} coupon bond transactions, last date {report['latest_trade_date']}, last 5 sessions = {len(report['last_five_trading_dates_series'])} ISIN-days.")


if __name__ == "__main__":
    main()
