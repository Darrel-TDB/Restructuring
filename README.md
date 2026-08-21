# Loan Restructuring Calculator

An Excel workbook that encodes an unsecured cash loan's terms, lets you enter a
borrower's **last payment date** and **days past due (DPD)**, and automatically
computes the restructured loan — outstanding principal, accrued interest,
penalty, and the new amortization schedule — with the basis for every figure
shown next to it.

No client or borrower information is stored in this repository. The shipped
workbook is pre-filled with an illustrative sample scenario only
(₱20,000 principal, 95% APR, 12 installments → restructured into 47.5% APR,
16 installments).

## Contents

| File | Purpose |
|---|---|
| `tonik-loan-restructuring-calculator.xlsx` | The calculator. Open this in Excel or Google Sheets. |
| `build/build_workbook.py` | Generates the workbook from scratch (Python + openpyxl). Source of truth for the file's structure and formulas. |
| `requirements.txt` | Python dependency for the build script. |

## Using the calculator

Open `tonik-loan-restructuring-calculator.xlsx`. It has 6 tabs:

1. **Instructions** — legend and sheet-by-sheet guide.
2. **1. Original Loan Setup** — encode the existing loan's terms (principal,
   date granted, APR, term, fees). Yellow cells are inputs.
3. **2. Original Amort Schedule** — auto-generated amortization schedule for
   the original loan (annuity installment, actual/365 daily interest accrual).
4. **3. Restructuring Workup** — enter the **Last Payment Date** and **DPD**.
   Everything else (outstanding principal, accrued interest, penalty, and the
   new restructured principal/terms) is computed, with its basis stated next
   to each figure.
5. **4. Restructured Amort Sched** — auto-generated schedule for the new,
   restructured loan.
6. **5. Summary** — side-by-side comparison of original vs. restructured loan.

Color legend used throughout: **blue bold text on yellow** = input cell,
**black** = formula, **green** = link to another sheet.

### Key assumptions (documented in-sheet)

- Interest accrues on an actual/365 daily basis over each installment's
  billing cycle (28/30/31 days), with a fixed annuity installment computed
  via `PMT()`. This reproduces disclosed bank amortization schedules closely,
  though the final installment of a real bank schedule is sometimes "trued
  up" to hit a pre-disclosed total finance charge, so the model's last row
  can differ slightly.
- Accrued interest during delinquency = Outstanding Principal × (Annual
  Rate / 365) × DPD.
- Penalty = Late Payment Fee × number of missed 30-day installment cycles
  (`ROUNDUP(DPD/30)`).
- Whether accrued interest is capitalized into the new restructured principal
  follows a fixed DPD policy (not manually editable): **DPD > 90 days** →
  restructured principal = Outstanding Principal only (interest waived).
  **DPD ≤ 90 days** → restructured principal = Outstanding Principal +
  Accrued Interest (interest capitalized). Penalty capitalization remains a
  separate manual Yes/No toggle (defaults to No).

Adjust any of these in the yellow input cells or the formulas themselves if
your institution's policy differs.

## Regenerating the workbook

The workbook is built entirely from `build/build_workbook.py` — no manual
formatting is stored anywhere else. To regenerate it:

```bash
pip install -r requirements.txt
python3 build/build_workbook.py
```

This writes `tonik-loan-restructuring-calculator.xlsx` to the repo root with
formulas but no cached values. Open it in Excel (or LibreOffice) once to
calculate, or run it through LibreOffice headless:

```bash
soffice --headless --convert-to xlsx --outdir . tonik-loan-restructuring-calculator.xlsx
```

## Branding

Sheet headers and tab colors use Tonik's brand palette:

| Color | Hex |
|---|---|
| Primary purple | `#785AFF` |
| Secondary purple ("Finn") | `#682763` |
| White | `#FFFFFF` |

## License

See [LICENSE](./LICENSE). The included license is a placeholder for internal
use — swap in an open-source license before making this repository public if
that's the intent.
