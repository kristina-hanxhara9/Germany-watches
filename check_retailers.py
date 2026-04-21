"""
Check a list of German watch retailers against the Handelsregister DB.

For each retailer (name + zip code), returns:
  - active status (based on dissolutionDate)
  - business description (Unternehmensgegenstand / objective)
  - address, register number, court, founded date

USAGE
-----
  python check_retailers.py input.xlsx output.xlsx
  python check_retailers.py input.csv  output.xlsx

INPUT FORMAT (xlsx or csv)
--------------------------
  Must have columns: name, zip
  (case-insensitive; "retailer", "company", "plz", "postcode" also accepted)

REQUIREMENTS
------------
  pip install pandas openpyxl
  handelsregister.db must be in the same folder as this script
"""

import os
import sys
import sqlite3
import pandas as pd

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "handelsregister.db")

NAME_COLS = {"name", "retailer", "company", "firm", "firma", "händler", "haendler"}
ZIP_COLS  = {"zip", "zipcode", "plz", "postcode", "postleitzahl"}


def load_input(path):
    if path.lower().endswith(".csv"):
        df = pd.read_csv(path, dtype=str)
    else:
        df = pd.read_excel(path, dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]
    name_col = next((c for c in df.columns if c in NAME_COLS), None)
    zip_col  = next((c for c in df.columns if c in ZIP_COLS),  None)
    if not name_col or not zip_col:
        raise SystemExit(f"Input must contain a name column ({NAME_COLS}) and zip column ({ZIP_COLS}). Got: {list(df.columns)}")
    df = df.rename(columns={name_col: "name", zip_col: "zip"})
    df["name"] = df["name"].fillna("").astype(str).str.strip()
    df["zip"]  = df["zip"].fillna("").astype(str).str.strip().str.zfill(5)
    return df[["name", "zip"]]


def lookup(conn, name, zip_code):
    """Find companies matching name (LIKE) and zip (in zipAndPlace)."""
    if not name:
        return []

    # Try exact name + zip match first, then progressively looser
    sql = """
        SELECT DISTINCT n.companyId
        FROM Names n
        JOIN Addresses a ON a.companyId = n.companyId
        WHERE n.name LIKE ?
          AND a.zipAndPlace LIKE ?
          AND n.isCurrent = 'True'
          AND a.isCurrent = 'True'
        LIMIT 10
    """
    rows = conn.execute(sql, [f"%{name}%", f"{zip_code}%"]).fetchall()
    if not rows:
        # Fallback: name only
        rows = conn.execute(
            "SELECT DISTINCT companyId FROM Names WHERE name LIKE ? AND isCurrent='True' LIMIT 10",
            [f"%{name}%"]
        ).fetchall()
    return [r[0] for r in rows]


def fetch_details(conn, cid):
    name = conn.execute("SELECT name FROM Names WHERE companyId=? AND isCurrent='True' LIMIT 1", (cid,)).fetchone()
    obj  = conn.execute("SELECT objective FROM Objectives WHERE companyId=? LIMIT 1", (cid,)).fetchone()
    addr = conn.execute("SELECT fullAddress, zipAndPlace FROM Addresses WHERE companyId=? AND isCurrent='True' LIMIT 1", (cid,)).fetchone()
    ref  = conn.execute("SELECT nativeReferenceNumber, courtName FROM ReferenceNumbers WHERE companyId=? LIMIT 1", (cid,)).fetchone()
    comp = conn.execute("SELECT foundedDate, dissolutionDate FROM Companies WHERE companyId=? LIMIT 1", (cid,)).fetchone()
    dissolved = comp[1] if comp else ""
    return {
        "matched_name":   name[0] if name else "",
        "objective":      obj[0]  if obj  else "",
        "address":        addr[0] if addr else "",
        "zipAndPlace":    addr[1] if addr else "",
        "register":       ref[0]  if ref  else "",
        "court":          ref[1]  if ref  else "",
        "founded":        comp[0] if comp else "",
        "dissolved":      dissolved,
        "active":         "NO" if dissolved else "YES",
    }


def main():
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)

    in_path, out_path = sys.argv[1], sys.argv[2]

    if not os.path.exists(DB_PATH):
        raise SystemExit(f"Database not found: {DB_PATH}\nCopy handelsregister.db next to this script.")
    if not os.path.exists(in_path):
        raise SystemExit(f"Input file not found: {in_path}")

    df = load_input(in_path)
    print(f"Loaded {len(df)} retailers from {in_path}")

    conn = sqlite3.connect(DB_PATH)
    results = []
    for i, row in df.iterrows():
        ids = lookup(conn, row["name"], row["zip"])
        if not ids:
            results.append({
                "input_name": row["name"], "input_zip": row["zip"],
                "active": "NOT FOUND", "matched_name": "", "objective": "",
                "address": "", "zipAndPlace": "", "register": "", "court": "",
                "founded": "", "dissolved": "",
            })
            print(f"[{i+1}/{len(df)}] {row['name']} ({row['zip']}) -> NOT FOUND")
            continue
        for cid in ids:
            d = fetch_details(conn, cid)
            results.append({"input_name": row["name"], "input_zip": row["zip"], **d})
        print(f"[{i+1}/{len(df)}] {row['name']} ({row['zip']}) -> {len(ids)} match(es)")

    conn.close()

    out = pd.DataFrame(results, columns=[
        "input_name", "input_zip",
        "active", "matched_name", "objective",
        "address", "zipAndPlace",
        "register", "court", "founded", "dissolved",
    ])
    out.to_excel(out_path, index=False)
    print(f"\nDone. Wrote {len(out)} rows to {out_path}")


if __name__ == "__main__":
    main()
