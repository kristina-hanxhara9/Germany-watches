"""
Check a list of German watch retailers against the Handelsregister DB.

v2 — much more accurate matching:
  * normalises German company names (GmbH / KG / & Co. / e.K. / umlauts / punctuation)
  * uses NamesFts (full-text index) with token AND + zip-constrained match
  * falls back progressively: all-tokens+zip -> all-tokens -> core-token+zip -> LIKE
  * scores every candidate (token overlap + zip + active) and keeps the best
  * returns FULL business description: all objectives (current + historical),
    all addresses, all historical names, officers, capital, register, court

USAGE
-----
  python check_retailers.py input.xlsx  output.xlsx
  python check_retailers.py input.csv   output.xlsx

INPUT columns (case-insensitive; first match wins):
  name    — "name", "retailer", "company", "firm", "firma", "händler"
  zip     — "zip", "plz", "postcode", "postleitzahl"
  city    — "city", "stadt", "ort"         (optional, improves scoring)

REQUIREMENTS
------------
  pip install pandas openpyxl
  handelsregister.db must be in the same folder
"""

import os
import re
import sys
import sqlite3
import unicodedata

import pandas as pd

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "handelsregister.db")

NAME_COLS = {"name", "retailer", "company", "firm", "firma", "händler", "haendler"}
ZIP_COLS  = {"zip", "zipcode", "plz", "postcode", "postleitzahl"}
CITY_COLS = {"city", "stadt", "ort"}

# ── German name normalisation ────────────────────────────────────────────────

# Legal suffixes / forms stripped before matching.
# All anchored with word boundaries so we don't eat letters out of real names.
LEGAL_SUFFIXES = [
    r"\bgmbh\s*&\s*co\.?\s*kg\b",
    r"\bgmbh\s*&\s*co\.?\b",
    r"\bgmbh\s*u\.?\s*co\.?\b",
    r"\bag\s*&\s*co\.?\s*kg\b",
    r"\bug\s*\(haftungsbeschr[aä]nkt\)",
    r"\bgmbh\b", r"\bmbh\b",
    r"\bag\b", r"\bkg\b", r"\bohg\b", r"\bgbr\b", r"\bug\b",
    r"\be\.\s*k\.?",  r"\be\.\s*kfr\.?", r"\be\.\s*kfm\.?",
    r"\be\.\s*v\.?",  r"\be\.\s*g\.?",
    r"\bgesellschaft\b", r"\bunternehmen\b",
    r"\binhaberin?\b", r"\binh\.?\b",
]
LEGAL_RE = re.compile("|".join(LEGAL_SUFFIXES), re.IGNORECASE)

# German stopwords not useful for matching
STOPWORDS = {
    "der", "die", "das", "den", "dem", "des",
    "und", "oder", "von", "vom", "zu", "zum", "zur",
    "fuer", "für", "mit", "auf", "an", "am", "im", "in",
    "&", "+",
}

FOLD = str.maketrans({"ä": "ae", "ö": "oe", "ü": "ue", "Ä": "ae", "Ö": "oe", "Ü": "ue", "ß": "ss"})


def normalise(s: str) -> str:
    """Lowercase, fold umlauts, strip legal form, remove punctuation, collapse spaces."""
    if not s:
        return ""
    s = s.translate(FOLD).lower()
    # Remove legal suffixes anywhere in the string
    s = LEGAL_RE.sub(" ", s)
    # Remove punctuation except spaces
    s = re.sub(r"[^\w\s]", " ", s, flags=re.UNICODE)
    # Collapse whitespace
    s = re.sub(r"\s+", " ", s).strip()
    # Drop accidental stopwords
    toks = [t for t in s.split() if t not in STOPWORDS and len(t) > 1]
    return " ".join(toks)


def tokens(s: str) -> list:
    return [t for t in normalise(s).split() if t]


def fts_escape(token: str) -> str:
    """FTS5-safe: wrap in double quotes, escape internal quotes."""
    return '"' + token.replace('"', '""') + '"'


# ── DB helpers ───────────────────────────────────────────────────────────────

def load_input(path: str) -> pd.DataFrame:
    if path.lower().endswith(".csv"):
        df = pd.read_csv(path, dtype=str)
    else:
        df = pd.read_excel(path, dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]
    name_col = next((c for c in df.columns if c in NAME_COLS), None)
    zip_col  = next((c for c in df.columns if c in ZIP_COLS),  None)
    city_col = next((c for c in df.columns if c in CITY_COLS), None)
    if not name_col or not zip_col:
        raise SystemExit(
            f"Input must have a name column ({sorted(NAME_COLS)}) and zip column "
            f"({sorted(ZIP_COLS)}). Got: {list(df.columns)}"
        )
    df = df.rename(columns={name_col: "name", zip_col: "zip"})
    if city_col:
        df = df.rename(columns={city_col: "city"})
    else:
        df["city"] = ""
    df["name"] = df["name"].fillna("").astype(str).str.strip()
    df["zip"]  = df["zip"].fillna("").astype(str).str.strip().str.extract(r"(\d{3,5})").fillna("")[0].str.zfill(5)
    df["city"] = df["city"].fillna("").astype(str).str.strip()
    return df[["name", "zip", "city"]]


def candidate_ids(conn, name: str, zip_code: str, limit: int = 200) -> list:
    """Progressive search: tightest match first, widen on misses."""
    toks = tokens(name)
    if not toks:
        return []

    # Tier 1 & 2 — NamesFts with AND of all significant tokens
    fts_and = " AND ".join(fts_escape(t) for t in toks)
    fts_or  = " OR ".join(fts_escape(t) for t in toks)

    # Helper to run FTS + optional zip filter
    def fts_query(fts_expr, zip_filter):
        sql = """
            SELECT DISTINCT n.companyId
            FROM NamesFts n
        """
        params = [fts_expr]
        if zip_filter:
            sql += """
                JOIN Addresses a ON a.companyId = n.companyId
                WHERE NamesFts MATCH ? AND a.zipAndPlace LIKE ? AND a.isCurrent='True'
            """
            params.append(zip_filter + "%")
        else:
            sql += " WHERE NamesFts MATCH ?"
        sql += f" LIMIT {limit}"
        return [r[0] for r in conn.execute(sql, params).fetchall()]

    # Tier 1: all tokens AND + zip
    if zip_code:
        ids = fts_query(fts_and, zip_code)
        if ids: return ids

    # Tier 2: all tokens AND (no zip)
    ids = fts_query(fts_and, None)
    if ids: return ids

    # Tier 3: any-token OR + zip
    if zip_code:
        ids = fts_query(fts_or, zip_code)
        if ids: return ids

    # Tier 4: most distinctive token (longest) alone
    if toks:
        core = max(toks, key=len)
        ids = fts_query(fts_escape(core), zip_code or None)
        if ids: return ids

    # Tier 5: LIKE fallback on unfolded text
    like = f"%{name.strip()}%"
    sql = "SELECT DISTINCT companyId FROM Names WHERE name LIKE ? LIMIT ?"
    return [r[0] for r in conn.execute(sql, [like, limit]).fetchall()]


def fetch_full(conn, cid: str) -> dict:
    """Return EVERY useful row for a company: all names, all objectives,
    all addresses, officers, capital, register, court, founded, dissolved."""
    names = conn.execute(
        "SELECT name, isCurrent, validFrom, validTill FROM Names WHERE companyId=? ORDER BY isCurrent DESC, validFrom DESC",
        (cid,),
    ).fetchall()
    current_name = next((n[0] for n in names if n[1] in ("True", 1, True)), names[0][0] if names else "")
    historical_names = [n[0] for n in names if n[1] not in ("True", 1, True)]

    objectives = [r[0] for r in conn.execute(
        "SELECT objective FROM Objectives WHERE companyId=? ORDER BY validFrom DESC", (cid,)
    ).fetchall() if r[0]]
    current_objective = objectives[0] if objectives else ""
    full_objective   = "\n\n---\n\n".join(dict.fromkeys(objectives))  # dedup, keep order

    addrs = conn.execute(
        "SELECT fullAddress, zipAndPlace, isCurrent, validFrom FROM Addresses WHERE companyId=? ORDER BY isCurrent DESC, validFrom DESC",
        (cid,),
    ).fetchall()
    cur_addr = next((a for a in addrs if a[2] in ("True", 1, True)), addrs[0] if addrs else None)
    all_addrs = "; ".join(dict.fromkeys(a[0] for a in addrs if a[0]))

    ref = conn.execute(
        "SELECT nativeReferenceNumber, courtName FROM ReferenceNumbers WHERE companyId=? LIMIT 1", (cid,)
    ).fetchone()

    comp = conn.execute(
        "SELECT foundedDate, dissolutionDate FROM Companies WHERE companyId=? LIMIT 1", (cid,)
    ).fetchone()

    cap = conn.execute(
        "SELECT capitalAmount, capitalCurrency FROM Capital WHERE companyId=? AND isCurrent='True' LIMIT 1", (cid,)
    ).fetchone()

    pos = conn.execute(
        "SELECT firstName, lastName, foundPosition, startDate, endDate FROM Positions WHERE companyId=? ORDER BY startDate DESC",
        (cid,),
    ).fetchall()
    officers = "; ".join(
        " ".join(x for x in (p[0], p[1], f"({p[2]})" if p[2] else "") if x).strip() for p in pos
    )

    zp   = (cur_addr[1] if cur_addr else "") or ""
    zipc = (re.match(r"\d{5}", zp).group(0) if re.match(r"\d{5}", zp) else "")
    city = zp[len(zipc):].strip() if zipc else zp

    dissolved = comp[1] if comp else ""
    return {
        "companyId":        cid,
        "matched_name":     current_name,
        "historical_names": "; ".join(dict.fromkeys(historical_names)),
        "active":           "NO" if dissolved else "YES",
        "zip":              zipc,
        "city":             city,
        "address":          cur_addr[0] if cur_addr else "",
        "all_addresses":    all_addrs,
        "current_objective": current_objective,
        "full_objective":    full_objective,
        "n_objectives":     len(objectives),
        "register":         ref[0] if ref else "",
        "court":            ref[1] if ref else "",
        "founded":          (comp[0] if comp and comp[0] else "") or "",
        "dissolved":        dissolved or "",
        "capital":          f"{cap[0]} {cap[1]}" if cap and cap[0] else "",
        "officers":         officers,
    }


# ── Scoring ──────────────────────────────────────────────────────────────────

def score_candidate(input_name: str, input_zip: str, input_city: str,
                    cand_name: str, cand_zip: str, cand_city: str, cand_active: str) -> float:
    """0..100 score.  Higher = better."""
    in_t  = set(tokens(input_name))
    c_t   = set(tokens(cand_name))
    if not in_t or not c_t:
        return 0.0

    overlap = len(in_t & c_t) / max(1, len(in_t))         # recall of input tokens
    precision = len(in_t & c_t) / max(1, len(c_t))
    jacc = len(in_t & c_t) / max(1, len(in_t | c_t))

    # Base: weighted mean
    score = 55 * overlap + 25 * jacc + 20 * precision      # 0..100

    # Bonuses
    if input_zip and cand_zip and input_zip == cand_zip:
        score += 15
    elif input_zip and cand_zip and input_zip[:2] == cand_zip[:2]:
        score += 4

    if input_city and cand_city:
        ic = normalise(input_city); cc = normalise(cand_city)
        if ic and (ic in cc or cc in ic):
            score += 5

    if cand_active == "YES":
        score += 2

    return round(min(score, 100.0), 1)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)
    in_path, out_path = sys.argv[1], sys.argv[2]

    if not os.path.exists(DB_PATH):
        raise SystemExit(f"Database not found: {DB_PATH}")
    if not os.path.exists(in_path):
        raise SystemExit(f"Input not found: {in_path}")

    df = load_input(in_path)
    print(f"Loaded {len(df)} retailers from {in_path}")

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = None  # plain tuples for speed

    results = []
    for i, row in df.iterrows():
        cids = candidate_ids(conn, row["name"], row["zip"])
        scored = []
        for cid in cids:
            d = fetch_full(conn, cid)
            s = score_candidate(row["name"], row["zip"], row["city"],
                                d["matched_name"], d["zip"], d["city"], d["active"])
            scored.append((s, d))
        scored.sort(key=lambda x: -x[0])

        if not scored:
            results.append({
                "input_name": row["name"], "input_zip": row["zip"], "input_city": row["city"],
                "match_score": 0, "match_quality": "NOT FOUND",
                **{k: "" for k in ["matched_name","historical_names","active","zip","city","address",
                                    "all_addresses","current_objective","full_objective","n_objectives",
                                    "register","court","founded","dissolved","capital","officers","companyId"]},
            })
            print(f"[{i+1}/{len(df)}] {row['name']} ({row['zip']}) -> NOT FOUND")
            continue

        top_score, top = scored[0]
        quality = "EXACT" if top_score >= 85 else ("GOOD" if top_score >= 65 else ("WEAK" if top_score >= 40 else "POOR"))
        alts = " | ".join(
            f"{s}:{d['matched_name']}({d['zip']})" for s, d in scored[1:4]
        )
        results.append({
            "input_name":    row["name"],
            "input_zip":     row["zip"],
            "input_city":    row["city"],
            "match_score":   top_score,
            "match_quality": quality,
            "alt_candidates": alts,
            **top,
        })
        print(f"[{i+1}/{len(df)}] {row['name']} ({row['zip']}) -> {top['matched_name']} "
              f"[{quality} {top_score}] ({len(scored)} cand.)")

    conn.close()

    cols = [
        "input_name", "input_zip", "input_city",
        "match_score", "match_quality",
        "matched_name", "active", "zip", "city", "address",
        "current_objective",
        "register", "court", "founded", "dissolved", "capital",
        "officers",
        "historical_names", "all_addresses", "n_objectives", "full_objective",
        "alt_candidates", "companyId",
    ]
    out = pd.DataFrame(results)
    for c in cols:
        if c not in out.columns:
            out[c] = ""
    out = out[cols]
    out.to_excel(out_path, index=False)

    print(f"\nDone. Wrote {len(out)} rows to {out_path}")
    print(out["match_quality"].value_counts().to_string())


if __name__ == "__main__":
    main()
