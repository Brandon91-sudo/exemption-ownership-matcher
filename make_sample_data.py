"""Generate a small SYNTHETIC input file (no real owners, parcels, or districts)."""
import pandas as pd

rows = [
    # parcel, prior owner, current owner, applicant, scenario
    ("000-000-001", "SMITH JOHN A & MARY B", "SMITH JOHN A & MARY B", "JOHN SMITH", "unchanged"),
    ("000-000-002", "GARCIA MARIA L", "GARCIA MARIA L TRUST", "MARIA GARCIA", "moved into trust"),
    ("000-000-003", "NGUYEN TRAN & LINH", "NGUYEN LINH", "TRAN NGUYEN", "surviving spouse"),
    ("000-000-004", "OKAFOR DAVID", "EST OF OKAFOR DAVID", "DAVID OKAFOR", "owner deceased"),
    ("000-000-005", "JOHNSON ROBERT T", "HEIRS OF JOHNSON ROBERT T", "ROBERT JOHNSON", "heirs"),
    ("000-000-006", "PATEL ANIL K", "WALSH PETER J & DIANA", "ANIL PATEL", "sold to new owner"),
    ("000-000-007", "MARTINEZ ELENA R TR", "MARTINEZ ELENA R", "ELENA MARTINEZ", "trust to individual"),
    ("000-000-008", "LEE CHARLES W & SUSAN", "LEE SUSAN M", "CHARLES LEE", "surviving spouse"),
    ("000-000-009", "BROWN FAMILY TRUST", "BROWN FAMILY REV LIVING TRUST", "HAROLD BROWN", "trust renamed"),
    ("000-000-010", "KOWALSKI ANNA", "KOVALSKI ANNA", "ANNA KOWALSKI", "typo/spelling variant"),
    ("000-000-011", "DAVIS MICHAEL J", "ACME HOLDINGS LLC", "MICHAEL DAVIS", "sold to LLC"),
    ("000-000-012", "ROSSI GIUSEPPE & CARLA", "ROSSI CARLA DECEASED GIUSEPPE", "GIUSEPPE ROSSI", "mixed"),
]
df = pd.DataFrame(rows, columns=["Parcel", "Name", "CurrentPeriodName", "ApplicantName", "Scenario"])
df.to_excel("sample_data/sample_input.xlsx", index=False)
print(f"wrote {len(df)} synthetic rows")
