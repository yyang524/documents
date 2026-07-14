#!/usr/bin/env python3
"""Build the Saudi Arabia Blanchard-Leigh workbook from IMF Article IV vintages.

Projections: IMF staff projections from the Article IV staff report published
closest before each budget year's approval (budgets are approved in Nov/Dec of
the preceding year). Actuals: latest IMF historical data (most recent staff
report containing the year), plus first-outturn values as a robustness variant.
All values transcribed from Table 1 (Selected Economic Indicators) of each
report; PDFs archived in data/imf_reports/.
"""
import statsmodels.api as sm
import numpy as np
import pandas as pd

# ---------------------------------------------------------------- source data
# vintage projections: budget year t -> dict
V = {
    2017: dict(report="CR 16/326 (2016 Art. IV)", pub="Oct 2016", page="Table 1, p.41 (PDF p.46)",
               g_proj=2.7, nopb_t=-36.2, nopb_tm1=-39.3),
    2018: dict(report="CR 17/316 (2017 Art. IV)", pub="Oct 2017", page="Table 1, p.41 (PDF p.46)",
               g_proj=1.3, nopb_t=-34.0, nopb_tm1=-39.6),
    2019: dict(report="CR 18/263 (2018 Art. IV)", pub="Aug 2018", page="Table 1, p.41 (PDF p.46)",
               g_proj=2.1, nopb_t=-36.9, nopb_tm1=-41.7),
    2020: dict(report="CR 19/290 (2019 Art. IV)", pub="Sep 2019", page="Table 1, p.44 (PDF p.49)",
               g_proj=2.7, nopb_t=-38.1, nopb_tm1=-41.1),
    2021: dict(report="none (no 2020 Art. IV; COVID)", pub="-", page="-",
               g_proj=None, nopb_t=None, nopb_tm1=None),
    2022: dict(report="CR 21/149 (2021 Art. IV)", pub="Jul 2021", page="Table 1, p.40 (PDF p.45)",
               g_proj=3.6, nopb_t=-25.7, nopb_tm1=-28.2),
    2023: dict(report="CR 22/274 (2022 Art. IV)", pub="Aug 2022", page="Table 1, p.45 (PDF p.51)",
               g_proj=3.8, nopb_t=-26.1, nopb_tm1=-27.9),
    2024: dict(report="CR 23/323 (2023 Art. IV)", pub="Sep 2023", page="Table 1, p.49 (PDF p.55)",
               g_proj=4.4, nopb_t=-25.8, nopb_tm1=-27.4),
    2025: dict(report="CR 24/280 (2024 Art. IV)", pub="Sep 2024", page="Table 1, p.52 (PDF p.58)",
               g_proj=4.4, nopb_t=-30.4, nopb_tm1=-32.4),
    2026: dict(report="CR 25/223 (2025 Art. IV)", pub="Aug 2025", page="Table 1, p.43 (PDF p.49)",
               g_proj=3.5, nopb_t=-20.3, nopb_tm1=-21.1),
}

# latest IMF historical values for year t (source = most recent report containing t)
A_LATEST = {
    2017: dict(g=1.3,  nopb=-39.6, src="CR 19/290"),
    2018: dict(g=2.1,  nopb=-41.6, src="CR 19/290"),
    2019: dict(g=3.5,  nopb=-34.7, src="CR 23/323"),
    2020: dict(g=-2.3, nopb=-37.2, src="CR 24/280"),
    2021: dict(g=5.6,  nopb=-29.1, src="CR 24/280"),
    2022: dict(g=10.9, nopb=-25.3, src="CR 25/223 (rebased non-oil GDP)"),
    2023: dict(g=5.8,  nopb=-25.4, src="CR 25/223"),
    2024: dict(g=4.5,  nopb=-24.7, src="CR 25/223"),
    2025: dict(g=None, nopb=None,  src="n.a. - 2026 Art. IV report due ~Aug 2026; PR 26/181: real GDP +4.5% in 2025"),
    2026: dict(g=None, nopb=None,  src="n.a. (year in progress)"),
}

# first post-outturn estimates for year t (earliest report showing t as actual/est.)
A_FIRST = {
    2017: dict(g=1.1,  nopb=-39.7, nopb_prev=-45.7, src="CR 18/263"),
    2018: dict(g=2.1,  nopb=-41.6, nopb_prev=-39.6, src="CR 19/290"),
    2019: dict(g=3.3,  nopb=-35.9, nopb_prev=None,  src="CR 21/149 (no 2018 column; prior year n.a.)"),
    2020: dict(g=-2.3, nopb=-38.1, nopb_prev=-35.9, src="CR 21/149"),
    2021: dict(g=4.9,  nopb=-31.4, nopb_prev=-39.7, src="CR 22/274"),
    2022: dict(g=4.8,  nopb=-32.4, nopb_prev=-29.5, src="CR 23/323"),
    2023: dict(g=3.8,  nopb=-33.0, nopb_prev=-32.2, src="CR 24/280"),
    2024: dict(g=4.5,  nopb=-24.7, nopb_prev=-25.4, src="CR 25/223 (rebased non-oil GDP)"),
    2025: dict(g=None, nopb=None,  nopb_prev=None,  src="n.a."),
    2026: dict(g=None, nopb=None,  nopb_prev=None,  src="n.a."),
}

YEARS = list(range(2017, 2027))

rows = []
for t in YEARS:
    v, al, af = V[t], A_LATEST[t], A_FIRST[t]
    plan = (v["nopb_t"] - v["nopb_tm1"]) if v["g_proj"] is not None else None
    fe_l = (al["g"] - v["g_proj"]) if (al["g"] is not None and v["g_proj"] is not None) else None
    fe_f = (af["g"] - v["g_proj"]) if (af["g"] is not None and v["g_proj"] is not None) else None
    rows.append(dict(year=t, plan=plan, fe_latest=fe_l, fe_first=fe_f))
df = pd.DataFrame(rows).set_index("year")

def ols(dfsub, ycol):
    d = dfsub.dropna(subset=["plan", ycol])
    X = sm.add_constant(d["plan"])
    m = sm.OLS(d[ycol], X).fit()
    return m, d

print("=" * 70)
for ycol, label in [("fe_latest", "FE vs LATEST actuals"), ("fe_first", "FE vs FIRST-OUTTURN actuals")]:
    for excl, elabel in [([], "full sample"), ([2020], "excl. 2020 (COVID)"),
                         ([2020, 2022], "excl. 2020 & 2022 (COVID, rebasing)")]:
        sub = df[~df.index.isin(excl)]
        m, d = ols(sub, ycol)
        print(f"{label:28s} | {elabel:34s} | n={int(m.nobs)} "
              f"| beta={m.params['plan']:+.3f} (se {m.bse['plan']:.3f}, p={m.pvalues['plan']:.3f}) "
              f"| alpha={m.params['const']:+.2f} | R2={m.rsquared:.3f}")
print("=" * 70)
df.to_csv('/tmp/claude-0/-home-user-documents/2056b728-c5cf-5ac0-a7f4-7459adce6669/scratchpad/bl_data.csv')
print(df)
