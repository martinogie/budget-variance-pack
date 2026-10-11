"""Build the variance pack: data, pandas analysis, commentary, Excel workbook, chart and web page.

    python -m budget_pack.run            # writes reports/ and index.html
    python -m budget_pack.run --recalc   # also recalculates the workbook with LibreOffice so cached values are stored
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

from . import analysis as A
from .data import CLOSE_MONTH, DEPTS, MONTHS, OPEX, PRODUCTS, build
from .workbook import Pack

ROOT = Path(__file__).resolve().parent.parent
REPORTS = ROOT / "reports"
XLSX = REPORTS / "FY2026_Budget_vs_Actual_Variance_Pack.xlsx"
RECALC = Path("/mnt/skills/public/xlsx/scripts/recalc.py")


def k(v):
    return f"${v / 1e3:,.0f}K"


def m(v):
    return f"${v / 1e6:,.2f}M"


def signed(v, fmt=m):
    return ("+" if v >= 0 else "-") + fmt(abs(v))


def commentary(p, bridge, hcb, dept):
    E, R, GP, OPX = (p.loc[x] for x in ("EBITDA", "Total revenue", "Gross profit", "Total operating expenses"))
    gm_a = p.loc["Gross profit", "ytd_act"] / R["ytd_act"]
    gm_b = p.loc["Gross profit", "ytd_bud"] / R["ytd_bud"]
    mkt = p.loc["Marketing programs"]
    ent = bridge.loc["Card processing - Enterprise"]
    eng = hcb.set_index("department").loc["Engineering"]
    fraud = p.loc["Fraud & chargeback losses"]
    s_ = bridge.loc["Card processing - SMB"]
    smb = s_["actual_tpv"] / s_["budget_tpv"] - 1
    return [
        ("Headline", [
            f"September EBITDA was {m(E['m_act'])}, {signed(E['m_var'])} vs budget. The beat is timing: the deferred Q3 "
            f"marketing campaign ({signed(mkt['m_var'])} favourable this month) covers a {m(-GP['m_var'])} gross-profit shortfall.",
            f"Year to date EBITDA is {m(E['ytd_act'])}, {signed(E['ytd_var'])} ({E['ytd_var_pct']:+.1%}) vs budget. "
            f"Revenue is ahead by {m(R['ytd_var'])} ({R['ytd_var_pct']:+.1%}) but gross margin is {gm_a:.1%} vs {gm_b:.1%} budget.",
            f"Full-year EBITDA outlook is {m(E['fy_out'])} vs budget {m(E['fy_bud'])} ({signed(E['fy_var'])}).",
        ]),
        ("Revenue", [
            f"Product revenue YTD {signed(bridge['total'].sum())}: volume {signed(bridge['volume'].sum())}, "
            f"mix {signed(bridge['mix'].sum())}, rate {signed(bridge['rate'].sum())}.",
            f"Enterprise card volume is running ahead of plan after Q2 wins, but at a negotiated take rate of "
            f"{ent['a_rate']:.2%} vs {ent['b_rate']:.2%} budget: rate effect {signed(ent['rate'])} YTD.",
            f"SMB card volume is {smb:.1%} vs plan YTD; shifting volume from SMB to lower-priced enterprise is the "
            f"{signed(bridge['mix'].sum())} mix effect.",
            f"Instant payouts {signed(p.loc['Instant payouts (RTP)', 'ytd_var'])} and interest on client float "
            f"{signed(p.loc['Interest on client float', 'ytd_var'])} (rates stayed higher than planned) are ahead.",
        ]),
        ("Cost of revenue", [
            f"Interchange and scheme fees {signed(p.loc['Interchange & scheme fees', 'ytd_var'])} YTD; volume driven, "
            f"so they move with the revenue beat.",
            f"Fraud and chargeback losses {signed(fraud['ytd_var'])} YTD, mostly the August card-testing attack ($850K). "
            f"Velocity rules went live in early September; September was still {signed(fraud['m_var'], k)} vs budget.",
        ]),
        ("Operating expenses", [
            f"Engineering is {int(-dept_gap(hcb))} heads behind plan: salaries {signed(eng['ytd_var'])} YTD "
            f"(headcount effect {signed(eng['volume_effect'])}, cost-per-head effect {signed(eng['rate_effect'])}).",
            f"Contractors {signed(p.loc['Contractors', 'ytd_var'])} YTD as engineering backfills open roles.",
            f"Marketing programs {signed(mkt['ytd_var'])} YTD is timing: the Q3 campaign moved to Q4 and is in the outlook.",
            f"Professional fees {signed(p.loc['Professional fees', 'ytd_var'])} YTD for SOC 2 and money-transmitter licensing.",
            f"{int(dept['flag'].sum())} department lines breach the ${A.THRESHOLD_USD / 1e3:.0f}K and "
            f"{A.THRESHOLD_PCT:.0%} thresholds this month (see Dept Opex).",
        ]),
        ("Actions", [
            "Set a 2.20% floor on enterprise renewals and review pricing on the three largest new accounts.",
            "Close the engineering hiring gap by year end and roll off contractors as roles are filled.",
            "Confirm the Q4 campaign budget ($1.08M) with Marketing before 15 October.",
            "Keep the velocity fraud rules in place; review losses weekly through peak season.",
        ]),
    ]


def dept_gap(hcb):
    e = hcb.set_index("department").loc["Engineering"]
    return e["actual_hc"] - e["budget_hc"]


def recalc(path):
    if not RECALC.exists():
        print("LibreOffice recalc script not found; open the workbook in Excel to calculate.")
        return None
    out = subprocess.run([sys.executable, str(RECALC), str(path), "90"], capture_output=True, text=True)
    res = json.loads(out.stdout[out.stdout.index("{"):])
    if res.get("status") != "success":
        raise SystemExit(f"Workbook formulas have errors: {res}")
    return res


def plot(gl, p, bridge, path):
    fig, ax = plt.subplots(1, 3, figsize=(17, 4.6))
    months = range(1, 13)
    for scen, color, ls in (("Budget", "#898781", "--"), ("Actual", "#2a78d6", "-")):
        g = gl[gl["scenario"] == scen]
        e = [g[(g.month_num == mm) & (g.type == "Revenue")]["amount"].sum() - g[(g.month_num == mm) & (g.type == "Cost")]["amount"].sum()
             for mm in months if scen == "Budget" or mm <= CLOSE_MONTH]
        ax[0].plot(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][:len(e)], np.array(e) / 1e6, color=color, ls=ls, lw=2, marker="o", ms=4, label=scen)
    ax[0].set_title("Monthly EBITDA, $M")
    ax[0].legend(frameon=False)
    steps = [("Volume", bridge["volume"].sum()), ("Mix", bridge["mix"].sum()), ("Rate", bridge["rate"].sum())]
    run = 0
    for i, (name, v) in enumerate(steps):
        ax[1].bar(i, v / 1e6, bottom=run / 1e6, color="#2a78d6" if v >= 0 else "#d03b3b", width=0.6)
        ax[1].text(i, max(run, run + v) / 1e6 + 0.12, signed(v), ha="center", fontsize=9)
        run += v
    ax[1].bar(3, run / 1e6, color="#52514e", width=0.6)
    ax[1].text(3, run / 1e6 + 0.15, signed(run), ha="center", fontsize=9)
    ax[1].set_xticks(range(4), ["Volume", "Mix", "Rate", "Total"])
    ax[1].axhline(0, color="#c3c2b7", lw=0.8)
    ax[1].set_ylim(0, max(run, steps[0][1]) / 1e6 * 1.18)
    ax[1].set_title("YTD product revenue variance, $M")
    lines = [l for l in p.index if l in OPEX or l in PRODUCTS or l in ("Interest on client float", "Interchange & scheme fees", "Fraud & chargeback losses")]
    v = p.loc[lines, "ytd_var"].sort_values()
    ax[2].barh(range(len(v)), v / 1e6, color=["#d03b3b" if x < 0 else "#2a78d6" for x in v])
    ax[2].set_yticks(range(len(v)), v.index, fontsize=8)
    ax[2].axvline(0, color="#c3c2b7", lw=0.8)
    ax[2].set_title("YTD variance by line, fav/(unfav) $M")
    for a in ax:
        a.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig(path, dpi=130)
    plt.close(fig)


def page_data(gl, drv, hc, notes):
    """Monthly actual and budget by line and department, so the page can rebuild any month or YTD view."""
    series = {}
    for (scen, line, dept), g in gl.groupby(["scenario", "line", "department"]):
        arr = [0.0] * 12
        for mm, amt in zip(g["month_num"], g["amount"]):
            arr[mm - 1] = round(float(amt))
        series.setdefault(line, {}).setdefault(dept, {})[scen] = arr
    fc = A.forecast(gl)
    return {
        "close_month": CLOSE_MONTH, "months": MONTHS,
        "products": list(PRODUCTS), "revenue": A.REVENUE, "cogs": A.COGS, "opex": OPEX, "departments": list(DEPTS),
        "series": series,
        "drivers": {p: {"bt": g["budget_tpv"].tolist(), "br": g["budget_rate"].tolist(),
                        "at": [None if x != x else x for x in g["actual_tpv"]],
                        "ar": [None if x != x else x for x in g["actual_rate"]]}
                    for p, g in drv.groupby("product")},
        "headcount": {d: {"b": g["budget_hc"].tolist(), "a": [None if x != x else int(x) for x in g["actual_hc"]]}
                      for d, g in hc.groupby("department")},
        "outlook": {l: round(float(v)) for l, v in fc["fy_outlook"].items()},
        "thresholds": {"usd": A.THRESHOLD_USD, "pct": A.THRESHOLD_PCT},
        "commentary": notes,
    }


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--recalc", action="store_true")
    args = ap.parse_args(argv)
    REPORTS.mkdir(exist_ok=True)
    gl, drv, hc = build()
    p = A.pnl(gl)
    bridge = A.price_volume_mix(drv, range(1, CLOSE_MONTH + 1))
    hcb = A.headcount_bridge(gl, hc)
    dept = A.dept_opex(gl)
    notes = commentary(p, bridge, hcb, dept)
    Pack(gl, drv, hc, notes).build(XLSX)
    if args.recalc:
        print("recalc:", recalc(XLSX))
    plot(gl, p, bridge, REPORTS / "results.png")
    gl.to_csv(REPORTS / "gl_extract.csv", index=False)
    summary = p[["m_act", "m_bud", "m_var", "ytd_act", "ytd_bud", "ytd_var", "fy_bud", "fy_out", "fy_var"]].round(0)
    summary.to_csv(REPORTS / "pnl_summary_sep2026.csv")
    md = ["# September 2026 close: commentary", ""]
    for head, bullets in notes:
        md += [f"## {head}", ""] + [f"- {b}" for b in bullets] + [""]
    (REPORTS / "COMMENTARY.md").write_text("\n".join(md))
    html = (ROOT / "web" / "template.html").read_text().replace("/*DATA*/null", json.dumps(page_data(gl, drv, hc, notes), separators=(",", ":")))
    (ROOT / "index.html").write_text(html)
    e = p.loc["EBITDA"]
    print(f"EBITDA month {m(e['m_act'])} ({signed(e['m_var'])}), YTD {m(e['ytd_act'])} ({signed(e['ytd_var'])}), "
          f"FY outlook {m(e['fy_out'])} vs {m(e['fy_bud'])}")


if __name__ == "__main__":
    main()
