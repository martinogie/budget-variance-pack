"""Synthetic FY2026 budget and Jan-Sep actuals for a card-payments company.

Budget and actuals are built from the same drivers (payment volume, take rates, headcount, cost
per head), and the actuals carry the stories a September close would explain:

- Enterprise card volume ahead of plan from Q2, at a lower negotiated rate; SMB volume behind.
- Instant payouts (RTP) adopted faster than planned.
- Interest on client float above plan as rates stayed higher.
- A card-testing fraud attack in August.
- Engineering hiring behind plan, partly backfilled with contractors.
- The Q3 marketing campaign deferred to Q4.
- Compliance work (SOC 2, money-transmitter licences) driving professional fees.
"""
import numpy as np
import pandas as pd

YEAR = 2026
CLOSE_MONTH = 9
MONTHS = [f"{YEAR}-{m:02d}" for m in range(1, 13)]

PRODUCTS = {  # name: (budget TPV $ in Jan, monthly growth, budget take rate, interchange & scheme cost rate)
    "Card processing - SMB": (260e6, 0.012, 0.0290, 0.0155),
    "Card processing - Enterprise": (300e6, 0.015, 0.0225, 0.0155),
    "ACH": (150e6, 0.010, 0.0045, 0.0005),
    "Instant payouts (RTP)": (40e6, 0.020, 0.0100, 0.0025),
}
OTHER_REVENUE = ["Platform subscriptions", "Interest on client float"]
COGS = ["Interchange & scheme fees", "Processing & bank fees", "Fraud & chargeback losses"]
OPEX = ["Salaries & benefits", "Contractors", "Software & cloud", "Marketing programs",
        "Travel & entertainment", "Professional fees", "Rent & facilities"]
DEPTS = {  # name: (headcount Jan, headcount Dec, monthly cost per head)
    "Engineering": (95, 112, 17500),
    "Product": (24, 28, 16500),
    "Sales & Marketing": (55, 63, 13000),
    "Customer Support": (50, 54, 7500),
    "Risk & Compliance": (30, 35, 12500),
    "G&A": (32, 34, 13500),
}
OTHER_OPEX_BUDGET = {  # monthly budget by dept and account
    "Engineering": {"Contractors": 90e3, "Software & cloud": 420e3, "Travel & entertainment": 15e3},
    "Product": {"Software & cloud": 40e3, "Travel & entertainment": 8e3},
    "Sales & Marketing": {"Marketing programs": 350e3, "Travel & entertainment": 60e3, "Software & cloud": 35e3},
    "Customer Support": {"Software & cloud": 55e3, "Contractors": 40e3},
    "Risk & Compliance": {"Software & cloud": 70e3, "Professional fees": 45e3},
    "G&A": {"Professional fees": 120e3, "Rent & facilities": 260e3, "Software & cloud": 30e3, "Travel & entertainment": 10e3},
}
FLOAT_BUDGET, RATE_BUDGET = 450e6, 0.0425


def line_type(line):
    return "Revenue" if line in PRODUCTS or line in OTHER_REVENUE else "Cost"


def section(line):
    if line in PRODUCTS or line in OTHER_REVENUE:
        return "Revenue"
    return "COGS" if line in COGS else "Opex"


def drivers(rng):
    """Budget and actual payment volume (TPV) and take rate by product and month."""
    rows = []
    for name, (tpv0, g, rate, _) in PRODUCTS.items():
        for m in range(1, 13):
            season = 1 + 0.06 * np.sin((m - 3) / 12 * 2 * np.pi) + (0.08 if m in (11, 12) else 0)
            b_tpv = tpv0 * (1 + g) ** (m - 1) * season
            a_tpv = a_rate = None
            if m <= CLOSE_MONTH:
                vol = 1 + rng.normal(0, 0.015)
                r = rate * (1 + rng.normal(0, 0.006))
                if name == "Card processing - Enterprise" and m >= 4:
                    vol *= 1 + min(0.04 * (m - 3), 0.16)          # enterprise wins ramp in
                    r = 0.0216 * (1 + rng.normal(0, 0.004))       # negotiated pricing
                if name == "Card processing - SMB" and m >= 3:
                    vol *= 0.955
                if name == "Instant payouts (RTP)":
                    vol *= 1 + 0.035 * m
                a_tpv, a_rate = b_tpv * vol, r
            rows.append({"month": MONTHS[m - 1], "month_num": m, "product": name,
                         "budget_tpv": round(b_tpv, -3), "budget_rate": rate,
                         "actual_tpv": None if a_tpv is None else round(a_tpv, -3),
                         "actual_rate": None if a_rate is None else round(a_rate, 6)})
    return pd.DataFrame(rows)


def headcount(rng):
    rows = []
    for dept, (h0, h1, cph) in DEPTS.items():
        for m in range(1, 13):
            b = round(h0 + (h1 - h0) * (m - 1) / 11)
            a = None
            if m <= CLOSE_MONTH:
                gap = 0
                if dept == "Engineering" and m >= 3:
                    gap = -min(2 * (m - 2), 9)
                elif dept in ("Sales & Marketing", "Risk & Compliance"):
                    gap = int(rng.integers(-2, 2))
                a = b + gap
            rows.append({"month": MONTHS[m - 1], "month_num": m, "department": dept,
                         "budget_hc": b, "actual_hc": a, "budget_cost_per_head": cph})
    return pd.DataFrame(rows)


def build(seed=26):
    rng = np.random.default_rng(seed)
    drv = drivers(rng)
    hc = headcount(rng)
    rows = []

    def add(scenario, m, dept, line, amount):
        rows.append({"scenario": scenario, "month": MONTHS[m - 1], "month_num": m, "department": dept,
                     "section": section(line), "line": line, "type": line_type(line), "amount": round(amount, 2)})

    for m in range(1, 13):
        d = drv[drv["month_num"] == m]
        for scen in ("Budget", "Actual"):
            if scen == "Actual" and m > CLOSE_MONTH:
                continue
            tpv = dict(zip(d["product"], d[f"{scen.lower()}_tpv"]))
            rate = dict(zip(d["product"], d[f"{scen.lower()}_rate"]))
            noise = (lambda s=0.03: 1 + rng.normal(0, s)) if scen == "Actual" else (lambda s=0.0: 1.0)
            for p in PRODUCTS:
                add(scen, m, "Corporate", p, tpv[p] * rate[p])
            add(scen, m, "Corporate", "Platform subscriptions", 1.2e6 * (1.01 ** (m - 1)) * noise(0.02))
            if scen == "Budget":
                interest = FLOAT_BUDGET * RATE_BUDGET / 12
            else:
                interest = FLOAT_BUDGET * (1 + 0.004 * m) * 0.0445 / 12
            add(scen, m, "Corporate", "Interest on client float", interest)
            card = tpv["Card processing - SMB"] + tpv["Card processing - Enterprise"]
            add(scen, m, "Corporate", "Interchange & scheme fees",
                sum(tpv[p] * PRODUCTS[p][3] for p in PRODUCTS) * noise(0.004))
            add(scen, m, "Corporate", "Processing & bank fees", sum(tpv.values()) * 0.0008 * noise(0.02))
            fraud = card * 0.0006 * noise(0.08)
            if scen == "Actual" and m == 8:
                fraud += 850e3
            if scen == "Actual" and m == 9:
                fraud += 120e3
            add(scen, m, "Corporate", "Fraud & chargeback losses", fraud)

            h = hc[hc["month_num"] == m].set_index("department")
            for dept, (_, _, cph) in DEPTS.items():
                heads = h.at[dept, "budget_hc" if scen == "Budget" else "actual_hc"]
                cost_ph = cph if scen == "Budget" else cph * (1.02 if dept == "Engineering" else 1.0) * noise(0.01)
                add(scen, m, dept, "Salaries & benefits", heads * cost_ph)
                for acct, amt in OTHER_OPEX_BUDGET[dept].items():
                    a = amt
                    if acct == "Marketing programs" and scen == "Budget" and m in (8, 9):
                        a += 600e3                                   # Q3 campaign in the budget
                    if acct == "Marketing programs" and scen == "Actual" and m in (8, 9):
                        a += 60e3                                    # campaign deferred; only prep spend
                    if scen == "Actual":
                        if acct == "Contractors" and dept == "Engineering" and m >= 4:
                            a += 110e3                               # backfill for open roles
                        if acct == "Software & cloud" and dept == "Engineering":
                            a *= 1.08                                # cloud scales with transactions
                        if acct == "Professional fees" and dept == "G&A" and m in (7, 9):
                            a += {7: 90e3, 9: 180e3}[m]              # SOC 2 and licensing
                        a *= noise(0.04)
                    add(scen, m, dept, acct, a)
    gl = pd.DataFrame(rows)
    return gl, drv, hc
