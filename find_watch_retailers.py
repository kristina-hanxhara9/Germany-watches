"""
Find all watch retailers in Germany from handelsregister.db.

Definition (strict):
  INCLUDE: Unternehmensgegenstand contains a clear watch keyword AND a clear
           RETAIL/SALES intent. Optionally jewellery + repair as supporting.
  EXCLUDE: holding/Verwaltungsgesellschaft, wholesale-only, pure repair shop,
           fashion-accessory dominated, or no watch keyword at all.
  REVIEW:  watch + retail but with significant fashion signal — manual check.

Run:  python find_watch_retailers.py
Output: watch_retailers_germany.xlsx
"""

import os
import re
import sqlite3
from datetime import datetime

import pandas as pd

DB  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "handelsregister.db")
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "watch_retailers_germany.xlsx")

# ── FTS keyword set (DB-level pre-filter) ────────────────────────────────────

WATCH_FTS = (
    '"uhren" OR "uhrmacher" OR "uhrmacherei" OR "uhrenhandel" OR '
    '"uhrenfachgeschäft" OR "uhrenfachhandel" OR "armbanduhr" OR '
    '"armbanduhren" OR "taschenuhr" OR "taschenuhren" OR "chronograph" OR '
    '"chronometer" OR "standuhr" OR "standuhren" OR "wanduhr" OR "wanduhren"'
)

# ── Strict regex classifiers (post-filter, applied to the actual objective) ──

WATCH_RE = re.compile(
    r"\b(uhren|uhrmacher\w*|uhrenhandel|uhrenfachgesch\w*|uhrenfachhandel|"
    r"armbanduhr\w*|taschenuhr\w*|chronograph\w*|chronometer\w*|"
    r"standuhr\w*|wanduhr\w*)\b", re.IGNORECASE,
)
JEWELRY_RE = re.compile(
    r"\b(juwelier\w*|schmuck|goldschmied\w*|feinschmuck|edelmetall\w*|"
    r"silberschmiede\w*|trauring\w*)\b", re.IGNORECASE,
)
# Retail / sales intent — must be present for INCLUDE
RETAIL_RE = re.compile(
    r"\b(einzelhandel|einzel-handel|fachgesch\wft\w*|ladengesch\wft\w*|"
    r"verkauf|vertrieb|versandhandel|onlinehandel|online-handel|e-?commerce|"
    r"shop\w*|laden\w*|boutique\w*|handelt? mit|handel mit|"
    r"verkauf\w* von|vertrieb\w* von)\b", re.IGNORECASE,
)
# Wholesale signals (we EXCLUDE wholesale-only since user wants retailers)
WHOLESALE_RE = re.compile(
    r"\b(gro[ßs]handel|wholesale|import\b|export\b|handelsvertretung)\b",
    re.IGNORECASE,
)
# Holding / management — EXCLUDE if dominant and no retail
HOLDING_RE = re.compile(
    r"\b(beteiligung\w*|holding|verwaltung\w*|verwaltungsgesellschaft|"
    r"vermögensverwaltung|management|geschäftsführung\w*)\b",
    re.IGNORECASE,
)
ACTIVE_BIZ_RE = re.compile(
    r"\b(handel\w*|verkauf\w*|vertrieb\w*|herstellung\w*|fertigung\w*|"
    r"reparatur\w*|dienstleistung\w*|produktion\w*)\b", re.IGNORECASE,
)
REPAIR_RE = re.compile(r"\b(reparatur\w*|instandsetzung\w*|wartung)\b", re.IGNORECASE)
FASHION_RE = re.compile(
    r"\b(pandora|thomas\s*sabo|modeschmuck|modeaccessoire\w*|modeartikel|"
    r"fashion\s*accessoire\w*|bijouterie|kost[üu]mschmuck|trendschmuck|"
    r"silberschmuck\s*-?\s*mode)\b", re.IGNORECASE,
)

# Unrelated product categories — if many appear alongside "Uhren", the company
# is a GENERIC retailer / online-shop, not a watch specialist.
UNRELATED_CATEGORIES = {
    "clothing":   re.compile(r"\b(bekleidung|kleidung|textilien|mode\b|schuhe|stiefel|sportbekleidung|kindermode|herrenmode|damenmode)\b", re.IGNORECASE),
    "food":       re.compile(r"\b(lebensmittel\w*|nahrungsmittel\w*|nahrungsergänz\w*|getränke|wein\b|spirituosen|tee\b|kaffee\b|süßwaren|reformwaren|diät|biolebensmittel)\b", re.IGNORECASE),
    "electronics":re.compile(r"\b(elektronik\w*|computer|laptop|hardware\b|software\b|smartphone\w*|handy\w*|tablet\w*|fernseher|kameras?|unterhaltungselektronik)\b", re.IGNORECASE),
    "household":  re.compile(r"\b(haushaltswaren|haushaltsartikel|hausrat|möbel|matratzen|heimtextilien|porzellan|glaswaren|küchenwaren|geschirr)\b", re.IGNORECASE),
    "toys":       re.compile(r"\b(spielzeug|spielwaren|spiele\b|baby\w*|kinderartikel|kinderwagen)\b", re.IGNORECASE),
    "diy":        re.compile(r"\b(werkzeug\w*|baumarkt\w*|heimwerker\w*|baustoff\w*|garten\w*|gartenbedarf)\b", re.IGNORECASE),
    "drugstore":  re.compile(r"\b(drogerie\w*|kosmetik\w*|körperpflege|hygiene|parfüm\w*|pflegeprodukt\w*|wellness)\b", re.IGNORECASE),
    "automotive": re.compile(r"\b(auto\w*|kfz|kraftfahrzeug\w*|motorrad\w*|fahrrad\w*|fahrzeug\w*|reifen|autozubehör|gebrauchtwagen|neuwagen|pkw|lkw)\b", re.IGNORECASE),
    "sports":     re.compile(r"\b(sportartikel|sportgerät\w*|sportbedarf|fitness\w*|outdoor|camping)\b", re.IGNORECASE),
    "petfood":    re.compile(r"\b(tiernahrung|tierbedarf|tierfutter|haustier\w*)\b", re.IGNORECASE),
    "office":     re.compile(r"\b(bürobedarf|büromaterial|schreibwaren|papierwaren)\b", re.IGNORECASE),
    "industrial": re.compile(r"\b(maschinen\b|stahlprodukt\w*|metallwaren|baumaterial\w*|industrielle\s*produkt\w*|chemikalien|agrarprodukt\w*)\b", re.IGNORECASE),
}


def classify(objective: str):
    """Return (decision, reasoning) — strict watch-retailer rules."""
    t = (objective or "")
    if not t.strip():
        return "EXCLUDE", "empty objective"

    watch_hits     = len(WATCH_RE.findall(t))
    jewel_hits     = len(JEWELRY_RE.findall(t))
    retail_hits    = len(RETAIL_RE.findall(t))
    wholesale_hits = len(WHOLESALE_RE.findall(t))
    holding_hits   = len(HOLDING_RE.findall(t))
    repair_hits    = len(REPAIR_RE.findall(t))
    fashion_hits   = len(FASHION_RE.findall(t))
    active_biz     = bool(ACTIVE_BIZ_RE.search(t))

    # How many UNRELATED product categories does the objective list?
    unrelated_cats = [name for name, rx in UNRELATED_CATEGORIES.items() if rx.search(t)]
    n_unrelated = len(unrelated_cats)

    # Rule 0 — must have a watch keyword (FTS got us here, but be strict)
    if watch_hits == 0:
        return "EXCLUDE", "no explicit watch keyword in objective text"

    # Rule 1 — pure holding/Verwaltungsgesellschaft (no retail, no wholesale, no active biz)
    if holding_hits >= 1 and retail_hits == 0 and wholesale_hits == 0 and not active_biz:
        return "EXCLUDE", f"holding/Verwaltungsgesellschaft (no retail/wholesale/active biz)"

    # Rule 2 — fashion-accessory dominance
    if fashion_hits > 0 and fashion_hits >= watch_hits:
        return "EXCLUDE", f"fashion-accessory signals ({fashion_hits}) ≥ watch signals ({watch_hits})"

    # Rule 3 — repair-only (no retail, no wholesale, repair ≥ watch)
    if repair_hits > 0 and retail_hits == 0 and wholesale_hits == 0 and repair_hits >= watch_hits:
        return "EXCLUDE", f"repair-dominant (repair×{repair_hits} ≥ watch×{watch_hits}, no retail)"

    # Rule 4 — wholesale-only (no retail) — user wants RETAILERS
    if wholesale_hits >= 1 and retail_hits == 0:
        return "EXCLUDE", f"wholesale-only (Großhandel) — not a retailer"

    # Rule 5 — must have explicit retail intent
    if retail_hits == 0:
        return "REVIEW", f"watch×{watch_hits} but no retail/sales keyword in objective"

    # Rule 6 — generic retailer (3+ unrelated categories like food/electronics/clothing)
    if n_unrelated >= 3:
        return "EXCLUDE", f"generic retailer — {n_unrelated} unrelated product categories: {', '.join(unrelated_cats[:6])}"

    # Rule 7 — borderline generic (1–2 unrelated cats AND watch only mentioned 1x) → REVIEW
    # Even with jewelry, watch×1 + any unrelated category means watches aren't the focus.
    if n_unrelated >= 1 and watch_hits == 1:
        return "REVIEW", f"watch×1 alongside {n_unrelated} unrelated categor(y/ies): {', '.join(unrelated_cats)}"

    # Rule 8 — fashion present but not dominant → REVIEW
    if fashion_hits > 0:
        return "REVIEW", f"watch×{watch_hits} + retail×{retail_hits} + fashion×{fashion_hits}"

    # Rule 9 — INCLUDE: watch + retail confirmed, focused on watches/jewellery
    parts = [f"watch×{watch_hits}", f"retail×{retail_hits}"]
    if jewel_hits:    parts.append(f"jewelry×{jewel_hits}")
    if repair_hits:   parts.append(f"repair×{repair_hits} (allowed)")
    if n_unrelated:   parts.append(f"+{n_unrelated} other cat({','.join(unrelated_cats)})")
    return "INCLUDE", ", ".join(parts)


# ── DB access ────────────────────────────────────────────────────────────────

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
    m = re.match(r"\d{5}", zp)
    zipc = m.group(0) if m else ""
    city = zp[len(zipc):].strip() if zipc else zp

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
        "founded":      (comp["foundedDate"] if comp and comp["foundedDate"] else "") or "",
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

    print(f"[{datetime.now():%H:%M:%S}] FTS pre-filter: any watch keyword in objective…")
    ids = [r[0] for r in conn.execute(
        "SELECT DISTINCT companyId FROM ObjectivesFts WHERE ObjectivesFts MATCH ?",
        [WATCH_FTS],
    ).fetchall()]
    print(f"[{datetime.now():%H:%M:%S}] FTS returned {len(ids)} candidates")

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

    order = {"INCLUDE": 0, "REVIEW": 1, "EXCLUDE": 2}
    df["_o"] = df["decision"].map(order).fillna(9)
    df["_a"] = (df["active"] != "YES").astype(int)
    df = df.sort_values(by=["_o", "_a", "name"]).drop(columns=["_o", "_a"])

    with pd.ExcelWriter(OUT, engine="openpyxl") as xl:
        df.to_excel(xl, index=False, sheet_name="All")
        for dec in ("INCLUDE", "REVIEW", "EXCLUDE"):
            sub = df[df["decision"] == dec]
            if len(sub):
                sub.to_excel(xl, index=False, sheet_name=dec)
        summary = (df.groupby(["decision", "active"]).size()
                     .unstack(fill_value=0).reset_index())
        summary.to_excel(xl, index=False, sheet_name="Summary")

    print(f"\n[{datetime.now():%H:%M:%S}] Done. {len(df)} rows -> {OUT}")
    print(df["decision"].value_counts().to_string())
    print()
    print("INCLUDE active vs dissolved:")
    print(df[df["decision"] == "INCLUDE"]["active"].value_counts().to_string())


if __name__ == "__main__":
    main()
