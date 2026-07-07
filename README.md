# M&A Deal Screener

A screening tool for mergers & acquisitions, built with Streamlit. It works for
**public *and* private targets** — the realistic case for a small / boutique
advisory firm, where most targets have no ticker.

## What it does
- **Three interchangeable data sources**, all producing the same company snapshot:
  1. **Manual entry** — type figures straight from a 10-K, CIM, or tax return,
     with an **add-backs** section that builds **Adjusted EBITDA** (the number
     small-firm deals actually transact on).
  2. **SEC EDGAR** — free, audited XBRL data pulled from the latest 10-K
     (US public filers). No API key required.
  3. **Yahoo Finance** — quick public-market pull for a fast first look.
- **Flexible offer basis**: premium to share price (public targets), EV/EBITDA
  multiple, or an absolute enterprise value (private targets).
- **Valuation & deal analytics**:
  - Unlevered free-cash-flow **DCF** with a proper enterprise → equity bridge
  - **EPS accretion / dilution** (incl. foregone interest on cash used)
  - Pro-forma **leverage** and capital structure
  - **Synergy** run-rate + NPV (with a ramp profile)
  - **Trading comps** off a peer set
  - Bull / Base / Bear **scenarios**
- **Save / load** deals as JSON, and a formatted **Excel** export.

## How to run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Assumptions & limitations (read this)
This is a **screening tool, not a fairness opinion**. It deliberately simplifies
a full merger model:
- Add-backs / normalization are the **user's judgment** — the tool doesn't verify them.
- DCF uses single-stage growth and a *maintenance-capex ≈ D&A* proxy when those
  inputs are blank.
- EDGAR covers **US public filers only** and does not contain a share price
  (enter it manually).
- **Purchase accounting** (goodwill, intangible step-up amortization) and
  transaction fees are **not** modeled.
- The verdict is an **illustrative flag**, never an investment recommendation.

*The previous single-source (Yahoo-only) version is preserved as `app_legacy_backup.py`.*
