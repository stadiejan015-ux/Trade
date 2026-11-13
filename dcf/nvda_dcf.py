"""NVIDIA (NVDA) 10-year FCFF DCF. Fiscal years end late January (FY27 = Feb-2026..Jan-2027).
All $ in billions. Inputs as of 2026-11-13; see README notes for sources and assumptions."""

PRICE = 166.19            # last close (approx., 2026-11-12)
SHARES = 24.15            # diluted shares, bn
NET_CASH = 50.0           # ASSUMPTION: cash+securities less debt; verify vs latest 10-Q
TAX = 0.16

# FY27E revenue: Q1 ~81.5 (implied by Q2 +18% q/q) + Q2 96.2 + Q3 guide 108 + Q4 ~117
FY27_REV = 81.5 + 96.2 + 108.0 + 117.0

SCENARIOS = {
    #            growth FY28..FY36                                        EBIT margin FY27..FY36                                    reinvest % rev   WACC   g
    "Bear": ([0.15, 0.05, -0.05, 0.03, 0.04, 0.04, 0.04, 0.03, 0.03], [0.62, 0.55, 0.50, 0.46, 0.44, 0.42, 0.40, 0.40, 0.40, 0.40], 0.06, 0.115, 0.025),
    "Base": ([0.35, 0.20, 0.12, 0.08, 0.07, 0.06, 0.05, 0.04, 0.04], [0.64, 0.63, 0.61, 0.59, 0.57, 0.55, 0.53, 0.52, 0.50, 0.50], 0.06, 0.105, 0.030),
    "Bull": ([0.50, 0.30, 0.20, 0.15, 0.12, 0.10, 0.08, 0.06, 0.05], [0.65, 0.65, 0.64, 0.63, 0.62, 0.61, 0.60, 0.58, 0.57, 0.56], 0.06, 0.095, 0.035),
}

def dcf(growth, margins, reinvest, wacc, g, verbose=False):
    rev, revs = FY27_REV, [FY27_REV]
    for gr in growth:
        rev *= 1 + gr; revs.append(rev)
    pv = 0.0
    # FY27 is ~75% elapsed; value only the remaining Q4 cash flow + FY28 onward, discount from mid-year points
    rows = []
    for i, (r, m) in enumerate(zip(revs, margins)):
        fcff = r * m * (1 - TAX) - r * reinvest
        frac = 0.25 if i == 0 else 1.0
        t = 0.125 if i == 0 else 0.25 + (i - 1) + 0.5
        pv_i = fcff * frac / (1 + wacc) ** t
        pv += pv_i
        rows.append((f"FY{27+i}", r, m, fcff, pv_i))
    tv = rows[-1][3] * (1 + g) / (wacc - g)
    pv_tv = tv / (1 + wacc) ** (0.25 + len(revs) - 1)
    ev = pv + pv_tv
    eq = ev + NET_CASH
    px = eq / SHARES
    if verbose:
        print(f"{'FY':<6}{'Revenue':>10}{'EBIT%':>8}{'FCFF':>9}{'PV':>9}")
        for fy, r, m, f, p in rows:
            print(f"{fy:<6}{r:>10.0f}{m:>8.0%}{f:>9.0f}{p:>9.0f}")
        print(f"PV explicit {pv:,.0f} | PV terminal {pv_tv:,.0f} ({pv_tv/ev:.0%} of EV) | EV {ev:,.0f} | Equity {eq:,.0f}")
        print(f"Exit FY36 EV/FCFF implied: {tv/rows[-1][3]:.1f}x")
    return px, ev

if __name__ == "__main__":
    print(f"FY27E revenue: ${FY27_REV:.0f}bn\n")
    for name, (gr, mg, ri, w, g) in SCENARIOS.items():
        print(f"== {name}: WACC {w:.1%}, g {g:.1%}")
        px, ev = dcf(gr, mg, ri, w, g, verbose=(name == "Base"))
        print(f"   Value/share ${px:,.0f}  vs ${PRICE}  ({px/PRICE-1:+.0%})\n")
    gr, mg, ri, _, _ = SCENARIOS["Base"]
    waccs = [0.090, 0.095, 0.100, 0.105, 0.110, 0.115, 0.120]
    gs = [0.020, 0.025, 0.030, 0.035, 0.040]
    print("Base-case sensitivity ($/share), rows WACC, cols terminal g")
    print("WACC  " + "".join(f"{g:>8.1%}" for g in gs))
    for w in waccs:
        print(f"{w:<6.1%}" + "".join(f"{dcf(gr, mg, ri, w, g)[0]:>8.0f}" for g in gs))
