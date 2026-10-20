# Exemption Ownership Matcher

**Flags which parcel-tax exemptions should stay and which should go when the property's owner changes, using fuzzy name matching instead of eyeballing thousands of rows.**

Built for a California public-finance consulting firm that administers senior and hardship parcel-tax exemptions. Each year, analysts had to compare last period's owner of record to the current owner for every exempt parcel and decide: did the original applicant still own it (keep), or did it sell or pass to heirs (remove)? This tool does the first pass and leaves only the ambiguous cases for a human.

> All data in this repo is **synthetic**. No real owners, parcels, or districts appear here.

## The problem

Owner names are messy and change legitimately without the owner changing:

| Prior period | Current period | Right answer |
|---|---|---|
| `SMITH JOHN A & MARY B` | `SMITH JOHN A & MARY B` | Keep |
| `GARCIA MARIA L` | `GARCIA MARIA L TRUST` | Keep (moved into a trust) |
| `NGUYEN TRAN & LINH` | `NGUYEN LINH` | Keep (surviving spouse) |
| `OKAFOR DAVID` | `EST OF OKAFOR DAVID` | Remove (owner deceased) |
| `PATEL ANIL K` | `WALSH PETER J & DIANA` | Remove (sold) |
| `KOWALSKI ANNA` | `KOVALSKI ANNA` | Keep (spelling variant) |

An exact string match gets most of these wrong in both directions.

## Approach

1. **Normalize** names: uppercase, strip punctuation, drop noise tokens (`TRUST`, `TTEE`, `ET AL`, `JR`, ...).
2. **Score similarity** with Jaro-Winkler (`jellyfish`) and token-set ratio (`rapidfuzz`), falling back to `difflib` if neither is installed.
3. **Check the applicant**: does the original applicant's name appear in the current owner?
4. **Detect surviving spouse**: shared surname plus name overlap when the applicant is missing.
5. **Flag estates/heirs**: `EST OF`, `HEIRS OF`, `DECEASED`, `DECD`, etc. force a Remove.
6. **Decide and explain**: each row gets Keep/Remove, a confidence level, and a plain-English reason.

## Results on the sample data

12 synthetic parcels covering the scenarios above:

| Result | Count |
|---|---|
| Recommend Keep | 7 |
| Recommend Remove | 5 |
| Flagged as heirs/estate | 3 |

The 12 sample rows land on the expected answer. See `sample_data/sample_output.xlsx` for the color-coded workbook (tabs: All Records, Recommend Keep, Recommend Remove, Heirs-Estate Flagged, Summary).

Notable behavior: the spelling variant (`KOWALSKI` / `KOVALSKI`) scores 0.92 and is kept at *medium* confidence, while the sale to an unrelated buyer scores 0.52 and is removed. Medium-confidence rows are what an analyst should look at first.

## How to run

```bash
pip install -r requirements.txt
python make_sample_data.py          # regenerate the synthetic input
python exemption_applicant_comparison.py sample_data/sample_input.xlsx sample_data/sample_output.xlsx
```

Optional third argument: similarity threshold (default `0.85`). On Windows, `gui_launcher.py` provides a file-picker interface.

**Input columns** (auto-detected, case-insensitive): `Parcel`, `Name`, `CurrentPeriodName`, `ApplicantName`. Optional: `StatusCheck`, `PriorPeriodAPN`, situs/mail address.

## Limitations

- Rule-based, not a trained model; thresholds were tuned on a small set of confirmed matches.
- A name match is evidence of continuity, not proof of ownership. Output is a recommendation for analyst review.
- English-style `LAST FIRST` name order is assumed.

## Files

- `exemption_applicant_comparison.py` - matching and Excel report
- `gui_launcher.py` - optional Tkinter GUI
- `make_sample_data.py` - builds the synthetic input
- `sample_data/` - example input and output

## License

MIT
