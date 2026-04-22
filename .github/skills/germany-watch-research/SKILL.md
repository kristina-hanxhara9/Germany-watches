---
name: germany-watch-research
description: Research a German watch / watches-&-jewellery retailer from Company Name, ZIP, City, HR Number and the Handelsregister objective (Unternehmensgegenstand). Classifies the retailer against the watch-retailer definition and returns structured JSON. Never estimates values.
---

# Germany Watch Retailer Research Skill

## Purpose
Actively browse the web to find real, published information about a German watch or watches-&-jewellery retailer and classify it against the watch-retailer definition below.
NEVER estimate, guess, or infer any value. Report only what you actually read.

## Watch retailer definition (MANDATORY — apply to every company)

A company qualifies as a **watch retailer** only if ALL of the following are true:

1. **Sale of watches is mandatory.** The retailer must actually sell watches (Uhren, Armbanduhren, Taschenuhren, Standuhren, Wanduhren, Chronographen, Chronometer).
2. **At least ~80% of turnover** comes from one or more of:
   - Watches (Uhren)
   - Fine jewellery (Feinschmuck, Juwelier-Qualität, Trauringe, Edelmetalle)
   - Watch/jewellery repair services (Uhrmacher-Reparatur)

A company must be **excluded** if any of the following is true:

- **More than 50% of turnover from fashion accessories** — e.g. Pandora, Thomas Sabo, Modeschmuck (costume jewellery), Modeaccessoires, Kostümschmuck, bijouterie, mass-market fashion brands.
- **More than 50% of turnover from repair services** — pure Uhrmacher / repair workshops that don't sell watches at retail.
- No watch sales at all (e.g. jewellery-only shop with no watches).

When it's genuinely ambiguous, set `classification = "REVIEW"` and explain in `classification_reason`.

## Tools to use

Use the Copilot CLI's built-in tools — no extensions, no API keys.
- `web_search` for every search step
- `web_fetch` to read a specific URL (retailer's own site, Gelbe Seiten page, Northdata entry, Handelsregister listing, LinkedIn page)
- `view` for any local files if needed

For every search: call `web_search` with the query string, then `web_fetch` on the most relevant result URLs to read full content.

## Search sequence — follow this order every time

1. `web_search`: `"{company_name}" {zip} {city}` → discover website + directory listings → `web_fetch` each
2. `web_search`: `"{company_name}" Gelbe Seiten {city}` → then `web_fetch` the gelbeseiten.de page → phone, address, hours, categories
3. `web_search`: `"{company_name}" {city} Google Maps` → rating, review count, opening hours, category
4. `web_fetch` the company's own website → products, brands carried, repair services offered, about / history, online shop, contact
5. `web_search`: `"{company_name}" site:northdata.com` → `web_fetch` → legal entity, turnover, employees, WZ 2008 / NACE code
6. `web_search`: `"{company_name}" site:handelsregister.de` OR `site:unternehmensregister.de` → legal filings, Unternehmensgegenstand
7. `web_search`: `"{company_name}" Umsatz` OR `Jahresumsatz` → only record if a real published figure is found
8. `web_search`: `"{company_name}" Marken Uhren` → watch brands carried (Rolex, Omega, Breitling, Tissot, Seiko, etc.)
9. `web_search`: `"{company_name}" Schmuck Pandora "Thomas Sabo"` → check exclusion signals (fashion-accessory dominance)
10. `web_search`: `"{company_name}" Reparatur Uhrmacher` → check repair-only signals
11. `web_search`: `"{company_name}" 2024 2025 Nachrichten` → recent news
12. `web_search`: `"{company_name}" LinkedIn` → `web_fetch` the LinkedIn page → employee count if shown

## Strict rules

- `annual_turnover`: ONLY fill if you see a stated figure like `"Umsatz: 2,3 Mio. € (2023)"`. NULL otherwise.
- `employee_count`: ONLY fill if you see a stated number. NULL otherwise.
- `google_maps_rating`: exact star rating shown, e.g. `"4.6"`. NULL if not found.
- `classification`: one of `"INCLUDE"`, `"EXCLUDE"`, `"REVIEW"`. Apply the definition above. Required.
- `classification_reason`: short factual reason citing what you actually saw on the web (brands carried, % of shop floor, words on website, categories on Gelbe Seiten/Google Maps).
- `watch_brands_carried`: list of specific watch brands you saw listed on the retailer's site or directory entries. Empty list if you only saw generic "Uhren".
- `jewellery_brands_carried`: list of specific jewellery brands (fine & fashion). Used to detect Pandora / Thomas Sabo dominance.
- `offers_repair_services`: true / false / null. Only true if the retailer advertises repair.
- `data_confidence`: `"high"` = 3+ reliable sources with detail. `"medium"` = partial. `"low"` = very little found.
- `sources_checked`: list every URL you visited, even dead ends.
- Never invent, estimate, or extrapolate any field.
- Never change the `classification` based on guessing; if the evidence is thin, use `"REVIEW"`.

## Output format

Return ONLY valid JSON — no text before, no text after, no markdown fences:

{
  "website": "https://... or null",
  "phone_number": "+49 ... or null",
  "address": "full address or null",
  "google_maps_url": "https://maps.google.com/... or null",
  "google_maps_rating": "4.6 or null",
  "google_maps_review_count": "127 or null",
  "opening_hours": "Mo-Sa 9-19 Uhr or null",
  "about": "2-3 factual sentences from what you actually read",
  "products_sold": ["product1", "product2"],
  "watch_brands_carried": ["Rolex", "Omega"],
  "jewellery_brands_carried": ["brand1", "brand2"],
  "own_brands": ["own brand or empty list"],
  "offers_repair_services": true,
  "chain_or_group": "chain name or null",
  "parent_company": "parent group or null",
  "number_of_locations": "exact number if stated or null",
  "annual_turnover": "2,3 Mio. € (2023) or null",
  "employee_count": "47 or null",
  "target_customers": "who they sell to based on what you found",
  "price_positioning": "budget | mid-range | premium | luxury | unknown",
  "online_shop_url": "https://... or null",
  "social_media": "Instagram: @handle — or null",
  "recent_news": "headline + source + date or null",
  "northdata_url": "https://www.northdata.com/... or null",
  "handelsregister_url": "https://www.handelsregister.de/... or null",
  "gelbeseiten_url": "https://www.gelbeseiten.de/... or null",
  "classification": "INCLUDE | EXCLUDE | REVIEW",
  "classification_reason": "short factual reason citing what you saw",
  "data_confidence": "high | medium | low",
  "sources_checked": ["url1", "url2", "url3"],
  "research_error": null
}
