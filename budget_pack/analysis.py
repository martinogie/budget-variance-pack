"""Variance analysis in pandas. The Excel pack computes the same numbers with formulas; the tests
check that both agree."""
import numpy as np
import pandas as pd

from .data import COGS, CLOSE_MONTH, DEPTS, OPEX, OTHER_REVENUE, PRODUCTS

REVENUE = list(PRODUCTS) + OTHER_REVENUE
THRESHOLD_USD, THRESHOLD_PCT = 50_000, 0.05

# Q4 forecast method per P&L line: run-rate (Q4 budget x Q3 actual/budget) or budget, plus a $ adjustment
FORECAST = {line: ("Run-rate", 0.0) for line in REVENUE + COGS + OPEX}
FORECAST.update({
    "Fraud & chargeback losses": ("Budget", 150_000.0),   # August attack treated as one-off; extra monitoring
    "Marketing programs": ("Budget", 1_080_000.0),        # deferred Q3 campaign lands in Q4
    "Professional fees": ("Budget", 150_000.0),           # licensing work continues
    "Platform subscriptions": ("Budget", 0.0),
})


def amounts(gl, scenario, months, by="line"):
    g = gl[(gl["scenario"] == scenario) & gl["month_num"].isin(months)]
    return g.groupby(by)["amount"].sum()


def fav(line_type, actual, budget):
    """Favourable (+) / unfavourable (-) variance: revenue above budget, costs below budget."""
    return actual - budget if line_type == "Revenue" else budget - actual


def forecast(gl):
    ytd = range(1, CLOSE_MONTH + 1)
    q3, q4 = range(7, 10), range(10, 13)
    out = {}
    for line, (method, adj) in FORECAST.items():
        a = lambda s, m: amounts(gl, s, m).get(line, 0.0)
        factor = a("Actual", q3) / a("Budget", q3) if method == "Run-rate" else 1.0
        q4_fc = a("Budget", q4) * factor + adj
        out[line] = {"ytd_actual": a("Actual", ytd), "q4_forecast": q4_fc,
                     "fy_outlook": a("Actual", ytd) + q4_fc, "fy_budget": a("Budget", range(1, 13))}
    return pd.DataFrame(out).T


def pnl(gl, month=CLOSE_MONTH):
    """P&L lines and subtotals for the month, year to date and full year."""
    fc = forecast(gl)
    cols = {
        "m_act": amounts(gl, "Actual", [month]), "m_bud": amounts(gl, "Budget", [month]),
        "ytd_act": amounts(gl, "Actual", range(1, month + 1)), "ytd_bud": amounts(gl, "Budget", range(1, month + 1)),
        "fy_bud": fc["fy_budget"], "fy_out": fc["fy_outlook"],
    }
    lines = pd.DataFrame(cols).reindex(REVENUE + COGS + OPEX).fillna(0.0)
    tot = lambda names: lines.loc[names].sum()
    rows = {}
    for line in REVENUE:
        rows[line] = ("Revenue", lines.loc[line])
    rows["Total revenue"] = ("Revenue", tot(REVENUE))
    for line in COGS:
        rows[line] = ("Cost", lines.loc[line])
    rows["Total cost of revenue"] = ("Cost", tot(COGS))
    rows["Gross profit"] = ("Revenue", tot(REVENUE) - tot(COGS))
    for line in OPEX:
        rows[line] = ("Cost", lines.loc[line])
    rows["Total operating expenses"] = ("Cost", tot(OPEX))
    rows["EBITDA"] = ("Revenue", tot(REVENUE) - tot(COGS) - tot(OPEX))
    out = []
    for name, (t, v) in rows.items():
        r = {"line": name, "type": t, **v.to_dict()}
        for p, (a, b) in {"m": ("m_act", "m_bud"), "ytd": ("ytd_act", "ytd_bud"), "fy": ("fy_out", "fy_bud")}.items():
            r[f"{p}_var"] = fav(t, v[a], v[b])
            r[f"{p}_var_pct"] = r[f"{p}_var"] / abs(v[b]) if v[b] else np.nan
        out.append(r)
    return pd.DataFrame(out).set_index("line")


def dept_opex(gl, month=CLOSE_MONTH):
    opex = gl[gl["section"] == "Opex"]
    rows = []
    for dept in DEPTS:
        d = opex[opex["department"] == dept]
        for line in [l for l in OPEX if l in set(d["line"])]:
            x = d[d["line"] == line]
            ma = x[(x.scenario == "Actual") & (x.month_num == month)]["amount"].sum()
            mb = x[(x.scenario == "Budget") & (x.month_num == month)]["amount"].sum()
            ya = x[(x.scenario == "Actual") & (x.month_num <= month)]["amount"].sum()
            yb = x[(x.scenario == "Budget") & (x.month_num <= month)]["amount"].sum()
            flag = abs(mb - ma) >= THRESHOLD_USD and mb and abs(mb - ma) / mb >= THRESHOLD_PCT
            rows.append({"department": dept, "line": line, "m_act": ma, "m_bud": mb, "m_var": mb - ma,
                         "ytd_act": ya, "ytd_bud": yb, "ytd_var": yb - ya, "flag": bool(flag)})
    return pd.DataFrame(rows)


def price_volume_mix(drv, months):
    """Split product revenue variance into volume, mix and rate effects.

    volume: change in total TPV at budget mix and budget rates
    mix:    TPV shifted between products, valued at each product's budget rate
    rate:   actual TPV x (actual rate - budget rate)
    For every product, volume + mix + rate = actual revenue - budget revenue.
    """
    d = drv[drv["month_num"].isin(months)].copy()
    d["b_rev"] = d["budget_tpv"] * d["budget_rate"]
    d["a_rev"] = d["actual_tpv"] * d["actual_rate"]
    g = d.groupby("product")[["budget_tpv", "actual_tpv", "b_rev", "a_rev"]].sum().reindex(list(PRODUCTS))
    g["b_rate"] = g["b_rev"] / g["budget_tpv"]
    g["a_rate"] = g["a_rev"] / g["actual_tpv"]
    B, A = g["budget_tpv"].sum(), g["actual_tpv"].sum()
    share = g["budget_tpv"] / B
    g["volume"] = (A - B) * share * g["b_rate"]
    g["mix"] = (g["actual_tpv"] - A * share) * g["b_rate"]
    g["rate"] = g["actual_tpv"] * (g["a_rate"] - g["b_rate"])
    g["total"] = g["a_rev"] - g["b_rev"]
    return g


def headcount_bridge(gl, hc, month=CLOSE_MONTH):
    """Salary variance by department split into headcount (volume) and cost-per-head (rate)."""
    rows = []
    for dept, (_, _, cph) in DEPTS.items():
        h = hc[(hc["department"] == dept) & (hc["month_num"] <= month)]
        s = gl[(gl["department"] == dept) & (gl["line"] == "Salaries & benefits") & (gl["month_num"] <= month)]
        bud, act = s[s.scenario == "Budget"]["amount"].sum(), s[s.scenario == "Actual"]["amount"].sum()
        b_hm, a_hm = h["budget_hc"].sum(), h["actual_hc"].sum()
        m = hc[(hc["department"] == dept) & (hc["month_num"] == month)].iloc[0]
        vol = (b_hm - a_hm) * cph                    # favourable when fewer heads than budget
        rows.append({"department": dept, "budget_hc": m["budget_hc"], "actual_hc": m["actual_hc"],
                     "ytd_budget": bud, "ytd_actual": act, "ytd_var": bud - act,
                     "volume_effect": vol, "rate_effect": (bud - act) - vol})
    return pd.DataFrame(rows)
