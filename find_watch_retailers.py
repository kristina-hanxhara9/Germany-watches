"""
Find all watch retailers in Germany from handelsregister.db.

Definition (applied heuristically via keyword density in Unternehmensgegenstand):
  INCLUDE: company objective mentions watches (mandatory) AND is not dominated
           by fashion-accessory brands (Pandora, Thomas Sabo, Modeschmuck) AND
           is not a pure repair shop.
  EXCLUDE: dominated by fashion-accessory signals, OR repair-only with no retail.
  REVIEW:  watch keyword present but fashion-accessory signals also present.

Run:  python find_watch_retailers.py
Output: watch_retailers_germany.xlsx
"""

import os
import re
import sqlite3
import pandas as pd
from datetime import datetime

DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "handelsregister.db")
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "watch_retailers_germany.xlsx")

# -----------------------------------------------------------------------------
# Keyword rules
# -----------------------------------------------------------------------------

# FTS5 query — narrow set so "Uhrzeit" (= clock time in phrases like "24 Uhr")
# does not pollute results. We use specific compound terms.
WATCH_FTS = (
    '"uhren" OR "uhrmacher" OR "uhrmacherei" OR "uhrenhandel" OR '
    '"uhrenfachgeschäft" OR "uhrenfachhandel" OR "armbanduhr" OR '
    '"armbanduhren" OR "taschenuhr" OR "taschenuhren" OR "chronograph" OR '
    '"chronometer" OR "standuhr" OR "standuhren" OR "wanduhr" OR "wanduhren"'
)

WATCH_RE = re.compile(
    r"\b(uhren|uhrmacher\w*|uhrenhandel|uhrenfachgesch\w*|uhrenfachhandel|"
    r"armbanduhr\w*|taschenuhr\w*|chronograph\w*|chronometer\w*|"
    r"standuhr\w*|wanduhr\w*)\b", re.IGNORECASE,
)
JEWELRY_RE = re.compile(
    r"\b(juwelier\w*|schmuck|goldschmied\w*|feinschmuck|edelmetall\w*|"
    r"silberschmiede\w*|trauring\w*)\b", re.IGNORECASE,
)
REPAIR_RE = re.compile(r"\b(reparatur\w*|instandsetzung\w*|wartung)\b", re.IGNORECASE)
RETAIL_RE = re.compile(
    r"\b(handel|verkauf|einzelhandel|vertrieb|gro\wHandel|gro\whandel|"
    r"versandhandel|onlinehandel|e-commerce|vertrieb)\b", re.IGNORECASE,
)
FASHION_RE = re.compile(
    r"\b(pandora|thomas\s*sabo|modeschmuck|modeaccessoire\w*|modeartikel|"
    r"fashion\s*accessoire\w*|bijouterie|kostümschmuck)\b", re.IGNORECASE,
)


def classify(objective: str):
    """Return (decision, reasoning)."""
    t = (objective or "")
    watch_hits   = len(WATCH_RE.findall(t))
    jewel_hits   = len(JEWELRY_RE.findall(t))
    repair_hits  = len(REPAIR_RE.findall(t))
    retail_hits  = len(RETAIL_RE.findall(t))
    fashion_hits = len(FASHION_RE.findall(t))

    if watch_hits == 0:
        return "NO_WATCH", "no watch keyword"

    reasons = []

    # Rule: fashion-accessory dominance => exclude
    if fashion_hits > 0 and fashion_hits >= watch_hits:
        return "EXCLUDE", f"fashion-accessory signals ({fashion_hits}) >= watch signals ({watch_hits})"

    # Rule: repair-only shop (repair mentioned, no retail, more repair than watch signals)
    if repair_hits >= 1 and retail_hits == 0 and repair_hits > watch_hits:
        return "EXCLUDE", f"repair-dominant ({repair_hits} repair vs {watch_hits} watch, no retail word)"

    # Rule: any fashion signal at all => REVIEW (even if watch dominant)
    if fashion_hits > 0:
        reasons.append(f"watch({watch_hits}) + fashion({fashion_hits}) — review")
        return "REVIEW", "; ".join(reasons)

    reasons.append(f"watch×{watch_hits}")
    if jewel_hits:  reasons.append(f"jewelry×{jewel_hits}")
    if retail_hits: reasons.append(f"retail×{retail_hits}")
    if repair_hits: reasons.append(f"repair×{repair_hits} (allowed)")
    return "INCLUDE", ", ".join(reasons)


# -----------------------------------------------------------------------------
# DB access
# -----------------------------------------------------------------------------

def fetch_company(conn, cid):
    name  = conn.execute("SELECT name FROM Names WHERE companyId=? AND isCurrent='True' LIMIT 1", (cid,)).fetchone()
    obj   = conn.execute("SELECT objective FROM Objectives WHERE companyId=? ORDER BY validFrom DESC LIMIT 1", (cid,)).fetchone()
    addr  = conn.execute("SELECT fullAddress, zipAndPlace FROM Addresses WHERE companyId=? AND isCurrent='True' LIMIT 1", (cid,)).fetchone()
    ref   = conn.execute("SELECT nativeReferenceNumber, courtName FROM ReferenceNumbers WHERE companyId=? LIMIT 1", (cid,)).fetchone()
    comp  = conn.execute("SELECT foundedDate, dissolutionDate FROM Companies WHERE companyId=? LIMIT 1", (cid,)).fetchone()
    cap   = conn.execute("SELECT capitalAmount, capitalCurrency FROM Capital WHERE companyId=? AND isCurrent='True' LIMIT 1", (cid,)).fetchone()
    pos   = conn.execute("SELECT firstName, lastName, foundPosition FROM Positions WHERE companyId=? ORDER BY startDate DESC", (cid,)).fetchall()

    officers = "; ".join(
        " ".join(x for x in (p["firstName"], p["lastName"], f"({p['foundPosition']})" if p["foundPosition"] else "") if x).strip()
        for p in pos
    ) if pos else ""

    dissolved = comp["dissolutionDate"] if comp else ""
    zp = (addr["zipAndPlace"] if addr else "") or ""
    zipc = zp.split()[0] if zp else ""
    city = " ".join(zp.split()[1:]) if zp else ""

    return {
        "companyId":    cid,
        "name":         name["name"]  if name  else "",
        "active":       "NO" if dissolved else "YES",
        "objective":    obj["objective"] if obj else "",
        "address":      addr["fullAddress"] if addr else "",
        "zip":          zipc,
        "city":         city,
        "register":     ref["nativeReferenceNumber"] if ref else "",
        "court":        ref["courtName"] if ref else "",
        "founded":      comp["foundedDate"] if comp else "",
        "dissolved":    dissolved or "",
        "capital":      f"{cap['capitalAmount']} {cap['capitalCurrency']}" if cap and cap["capitalAmount"] else "",
        "officers":     officers,
    }


def main():
    if not os.path.exists(DB):
        raise SystemExit(f"DB not found: {DB}")

    print(f"[{datetime.now():%H:%M:%S}] Connecting to {DB}")
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row

    print(f"[{datetime.now():%H:%M:%S}] FTS search for watch keywords…")
    ids = [r[0] for r in conn.execute(
        "SELECT DISTINCT companyId FROM ObjectivesFts WHERE ObjectivesFts MATCH ?",
        [WATCH_FTS],
    ).fetchall()]
    print(f"[{datetime.now():%H:%M:%S}] FTS returned {len(ids)} candidate companies")

    rows = []
    for i, cid in enumerate(ids, 1):
        d = fetch_company(conn, cid)
        decision, reasoning = classify(d["objective"])
        d["decision"]  = decision
        d["reasoning"] = reasoning
        rows.append(d)
        if i % 500 == 0:
            print(f"[{datetime.now():%H:%M:%S}]   processed {i}/{len(ids)}")

    conn.close()

    df = pd.DataFrame(rows, columns=[
        "decision", "reasoning",
        "name", "active",
        "zip", "city", "address",
        "objective",
        "register", "court",
        "founded", "dissolved",
        "capital", "officers",
        "companyId",
    ])

    # Sort: INCLUDE first, then REVIEW, EXCLUDE, NO_WATCH; active first; name asc
    order = {"INCLUDE": 0, "REVIEW": 1, "EXCLUDE": 2, "NO_WATCH": 3}
    df["_o"] = df["decision"].map(order).fillna(9)
    df["_a"] = (df["active"] != "YES").astype(int)
    df = df.sort_values(by=["_o", "_a", "name"]).drop(columns=["_o", "_a"])

    # Write one sheet per decision + summary
    with pd.ExcelWriter(OUT, engine="openpyxl") as xl:
        df.to_excel(xl, index=False, sheet_name="All")
        for dec in ("INCLUDE", "REVIEW", "EXCLUDE", "NO_WATCH"):
            sub = df[df["decision"] == dec]
            if len(sub):
                sub.to_excel(xl, index=False, sheet_name=dec)

        summary = (df.groupby(["decision", "active"]).size()
                     .unstack(fill_value=0).reset_index())
        summary.to_excel(xl, index=False, sheet_name="Summary")

    print(f"\n[{datetime.now():%H:%M:%S}] Done. {len(df)} rows -> {OUT}")
    print(df["decision"].value_counts().to_string())


if __name__ == "__main__":
    main()
