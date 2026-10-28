# Data cleaning and conforming: what was done, and what better options exist

Scope: `implementations/elt_ingest_excel` (Great Plains customer and supplier master data into Workday, plus HCM contingent workers). Evidence is the code and reference data on disk as of 2026-10-28; paths are relative to this folder. A companion diagram is in [data-conforming-inventory.html](data-conforming-inventory.html).

**Read this first.** The data here is finance master data about *organisations* (suppliers, customers), not consumer or child personal data. Names, addresses, phone, email and bank details transfer to a personal-info design pattern. Nothing in this repo covers parent/guardian, date of birth, consent or age handling.

## 1. Summary

| Domain | What we do today | Strength | Better option exists? |
|---|---|---|---|
| Postcode | Syntax clean by Workday's per-country rule text; UK repair of one mistyped character; outward-code lookup | Good syntax handling | **Yes**: full-postcode existence check (PAF / OS AddressBase / ONSPD) |
| Address lines | Strip junk, compact to 3-4 lines, re-pack to what Workday accepts per country | Rule-driven | **Yes**: address validation APIs return structured, standardised addresses |
| Town and region | Outward code to post town and county, then 4-5 redirects to a Workday region instance | Works, but fragile | **Yes**: ONSPD / OS data to administrative areas with codes; map once to Workday ids |
| Phone | `phonenumbers` (libphonenumber) plus SQL pre-clean | Solid | Partly: add line-type and liveness lookup only if needed |
| Email | Lowercase, split, simple regex, drop failures | Basic | **Yes**: MX and deliverability check |
| Sort code to bank | First 2 digits of the sort code to a hand-built 88-row table | Weak | **Yes**: full 6-digit Pay.UK sort code directory |
| Account number | Pad to 8 digits; length check | Weak | **Yes**: modulus check (free weights file) and Confirmation of Payee |
| IBAN | Regex plus length; no checksum | Weak | **Yes**: mod-97 checksum is free and trivial |
| SWIFT / BIC | Upper-case; no format check, no lookup | Weakest | **Yes**: SWIFTRef or library registry |
| Tax / VAT id | Format guess for type; no check | Weak | **Yes**: HMRC and EU VIES APIs (free) |
| Names / dedupe | Deterministic rules; no fuzzy match at transform time | Transparent | Partly: Companies House / D&B identity match |

## 2. Cleaning activity, by domain

### Names (`src_fin_*_raw.sql`)
- Unaccent, upper-case, then ordered rules: remove `.`; `FOOTBALL CLUB` to `FC`; drop `(...)`; `UNITED` to `UTD`; `AND` to `&`; strip trailing `LIMITED/LTD/PLC/LLP/INC/COMPANY`.
- Address and city fields are title-cased with a naive `fn_initcap` ("PO BOX" becomes "Po Box").
- Dedupe is deterministic SQL: score every source row (payment terms, bank, tax id, region, postcode...), the top score is canonical, all addresses, banks, phones and emails are aggregated onto it. Suppliers also split on bank account when one name has several; customers merge on name only (risk: two different customers with one normalised name merge).
- Fuzzy matching (Jaro-Winkler) exists only in a post-run reconcile review.

### Address
- Per line: remove `"<>|;{}`, collapse whitespace, null the placeholders `[Not Known]`, empty and `-`.
- Country resolve: name map, then code map, then raw code; a UK postcode-district hit overrides to GB, JE, GG or IM; default is GB.
- Re-pack to Workday's rules (section 4): where a country marks line 3 or 4 "Not Accepted", overflow is merged into the previous line (supplier). Customer simply nulls them, which **loses data**.
- Hungary special case: split a line over 35 characters at a word boundary.
- Postcode moved into line 2 when the country marks the postal code "Not Accepted".
- City: postcode-district `post_town`, else a town-to-county table, else raw city.

### Postcode
- `udf_normalize_postal_code` (Python) is driven by Workday's free-text "Postal Code Formatting" rule for the country (regex, ZIP+4, "N digits", Eircode, NL, CA, UK).
- UK: strict regex, then repair of **one** mistyped character (`WC1H`, `SO15`); the inward letters are never guessed; garbage returns NULL.
- Fallback chain if the postcode column is empty: search the address lines, then a UK regex, then the first comma part.
- Purely syntactic: it checks the code looks right, not that it exists.

### Phone
- SQL pre-clean: split on `;`, digits and leading `+` only, `00` to `+`, `+440` to `+44`, strip trunk zeros and a repeated country code.
- `udf_parse_phone` (Python, `phonenumbers`): builds candidate parses (as given, `+`, `00`/`011` prefixes, domicile country, GB if it starts 0 or 7), keeps the first valid, else trims up to 10 trailing digits to recover over-long values (`75239943970000` becomes `7523994397`).
- Output device type is only Landline, Mobile or NULL; Workday's VoIP, TollFree and Premium are never produced. No extension handling. Fax excluded.
- Rejected if fewer than 7 digits.

### Email
- Lowercase, split on `;`, simple regex; failures dropped silently. One owner per address across all entities (lowest id wins). No MX or deliverability check.

### Bank (supplier only; customer bank tables are switched off)
- Sort code: strip separators, pad 5 digits to 6. Account: pad 7 digits to 8.
- IBAN: kept if `^[A-Z]{2}\d{2}[A-Z0-9]+$`; GB IBAN yields sort code and account number. **No mod-97 checksum anywhere.**
- SWIFT/BIC: upper-case, nulled if it equals the sort code. **No format regex, no lookup.**
- Bank name has five tiers: US routing number, UK sort-code **2-digit prefix**, the source's own text, IBAN-embedded prefix then 16-country IBAN bank-code table, and finally a "Various..." placeholder.
- Account kept only when there is an account or IBAN and a bank name.
- Validation checks UK sort code is 6 digits and account 8 digits, and the Workday per-country "IBAN required / length" flags. Per-country regexes in the Workday bank-format sheet are loaded but **not used**.

### Tax id
- Customer: keep alphanumerics, then **drop it unless all digits**, so `GB123456789` is lost. Supplier: trim and upper-case only.
- GB type guessed from shape (9 digits VAT, 8 chars company number, 10 digits UTR); other countries use a 19-row default-type table. No checksum.

### Other
Dates by a three-format cascade; payment terms, categories, business unit, country attributes by lookup. Unmapped business units and unmapped supplier payment terms are **silently dropped**.

## 3. Reference datasets built for conforming

Loaded by `config/transform/sql/ref/*.sql` from `config/data/`. Row counts verified from the files.

| Dataset | Rows | Built from | Used for | Weakness |
|---|---|---|---|---|
| `ref_post_code_district` | 3,126 | UK postcode-district gazetteer, source not documented | outward code to post town, region, UK region | outward code only; 160 blank post towns |
| `ref_post_code_workday_region` | 352 | hand-built | gazetteer region to Workday region name | 60 gazetteer regions have no override |
| `ref_workday_country_state_region` | 4,349 | Workday tenant export | valid region instances per country | 346 rows lack a Workday id; joined on country *name* |
| `ref_workday_county_obsolete` | 160 | hand-built | "(obsolete)" instance to current one | GB only |
| `ref_country_county_state_town_mapping` | 200 | hand-built | town to county fallback | all GB |
| `ref_country` | 109 | hand-built | ISO2 to name, currency, dial code, language, tax type | none noted |
| `ref_source_country_name_mapping` / `_code_mapping` | 115 / 60 | hand-built | free-text country to ISO2 | exact match only |
| `ref_country_iso3_mapping` | 105 | ISO 3166 | **not used by any transform** | would fix the name-join bug below |
| `ref_country_tax_id_type_mapping` | 19 | hand-built | default tax id type | 7 countries |
| `ref_bank_sort_code_prefix_mapping` | 88 | hand-compiled | 2-digit sort-code prefix to bank | shared prefixes (04, 23, 70-72, 93) resolve to "Various" |
| `ref_bank_routing_number_mapping` | 19,221 | Workday tenant named range | US ABA to bank name | US only |
| `ref_iban_bank_code_mapping` | 103 | national registries, source cited per row | non-UK IBAN to bank | 16 countries, partial lists |
| `ref_iban_structure` | 4 | SWIFT IBAN registry | bank code position in IBAN | only GE, KZ, UA, CR |
| `ref_bank_name_normalisation` | 52 | hand-built | **not wired in** | unused |
| `ref_supplier_category`, `ref_customer_category`, `ref_source_payment_terms`, `ref_source_business_unit_mapping`, `ref_workday_group_additional` | 25, 26, 11, 3, 7 | hand-built | business mappings | domain-specific |
| `ref_location_mapping`, `ref_worker_type_mapping` | 10, 34 | hand-built | contingent workers | domain-specific |
| `ref_workday_validate_{supplier,customer}_*` (address, region, postal_code, phone_format, bank_format, currency, tax_id_type, address_format) | n/a | Workday instruction workbooks, ingested | Workday's own rules (section 4) | joined on country name |

Dead or stale: `ref_post_code_district.json` (200-row truncated copy), `ref_workday_customer_mapping.csv` and `generate_mappings.py` (column-lineage artefacts, not lookups), `ref_post_code_county` (referenced by `validation.py`, created nowhere), and `udf_normalize_address` (calls postcodes.io, Nominatim and libpostal; wired to nothing).

## 4. The region chain (supplier; customer identical)

1. Resolve country (name map, code map, `ref_country`).
2. Postcode to district: `nrm_postal_code LIKE district || ' %'` (outward code).
3. District region to Workday region name via the override table.
4. If no district: try raw `county`, then raw `city`, against Workday instances; then the town-to-county table by county, then by city.
5. First non-null candidate wins; validate it against the Workday instance list for that country.
6. Retry with `(obsolete)` appended; remap obsolete to current.
7. Blank the region if Workday marks it "Not Accepted" for that country.
8. The same district join also decides GB / JE / GG / IM and the post town.

Two behaviours to know: when a district matches but the result is not a Workday instance, there is no further fallback (region becomes NULL); and the whole chain keys on country **name** text. A known failure: Workday's rules workbook says "Macedonia, the Former Yugoslav Republic of", the tenant says "North Macedonia", so the join fails silently.

## 5. Workday's own address rules

Held in the instruction workbooks, ingested into `ref_workday_validate_*_address`: 245 countries, each attribute Required, Optional or Not Accepted.

| Attribute | Required | Optional | Not accepted |
|---|---|---|---|
| Address line 1 | 239 | 6 | 0 |
| Address line 2 | 1 | 244 | 0 |
| Address line 3 | 0 | 44 | 201 |
| Address line 4 | 0 | 21 | 224 |
| City (municipality) | 208 | 28 | 9 |
| Region / state | 29 | 90 | 126 |
| County (subregion) | 6 | 13 | 226 |
| City subdivision | 5 | 12 | 228 |
| Postal code | 102 | 80 | 63 |

UK: line 1 Required, lines 2 and 3 Optional, line 4 Not Accepted, city Required, region Optional, county Not Accepted, postcode Required. USA: region and city Required, lines 3 and 4 Not Accepted.

Postal code rules are 245 free-text descriptions (124 are "No validation"). The only length limit encoded in the repo is 35 characters for line 1 (customer payload schema), applied just for Hungary and the postcode relocation. There is no length check or truncation for city, region or other lines.

## 6. Gaps and risks

1. Provenance of the postcode gazetteer, the Workday region override table and the tenant region export is undocumented.
2. Joins on country name rather than ISO code (Macedonia). The ISO3 table is loaded and unused.
3. Bank data quality rests on hand-built 2-digit and 16-country tables, with no checksum, no BIC validation and no account verification.
4. Customer address overflow lines are discarded, supplier ones are merged.
5. Silent drops: unmapped business unit, unmapped supplier payment terms, emails failing the regex, banks with no resolved name.
6. Address-line order depends on DuckDB 1.5.5's `list_distinct` order; an upgrade could reverse lines.
7. Validation has no severity levels; every issue row weighs the same, and the "recommended" message can never fire.
8. Docs disagree with code in places (postcode relocation condition, bank-name tiers, email regex described as RFC, 35-character limit described as enforced).

## 7. Better data and paid services to evaluate

Pricing is deliberately omitted: it changes, and I have not verified any this session. Treat every entry as a shortlist to confirm with the vendor. "Free" means no licence fee as I understand it; check each licence.

### Addresses and postcodes
| Option | Type | What it gives | Replaces |
|---|---|---|---|
| Royal Mail PAF (via licensed resellers) | Paid licence | Authoritative UK deliverable addresses | outward-code guess, `post_town` |
| Ordnance Survey AddressBase / OS Places API | Paid licence | UK addresses with UPRN and coordinates | same, plus a stable property id |
| ONS Postcode Directory (ONSPD) | Free, Open Government Licence | Every full UK postcode to county, district, region, with ONS codes | the 3,126-row district table |
| Loqate (GBG), Experian Address Validation, Ideal Postcodes, Fetchify, getAddress.io | Paid APIs | Typeahead, validation, standardised output, international | `udf_normalize_address`, line parsing |
| Google Address Validation API; Smarty (US-strong); Melissa | Paid APIs | International validation and components | same |
| Google libaddressinput (i18n address data) | Free | Per-country field names, order, required fields, postcode regex | partly overlaps the Workday rules workbook; good for **UI design** |
| postcodes.io (built on ONSPD), GeoNames, libpostal | Free | lookup, gazetteer, parsing | free alternatives |

### Phone
| Option | Type | What it gives |
|---|---|---|
| libphonenumber (in use) | Free | Parse, format, validity, type |
| Twilio Lookup, Vonage Number Insight, Telesign | Paid APIs | Carrier, line type, reachability |
| Loqate phone validation | Paid API | Validity and type |

### Bank accounts
| Option | Type | What it gives | Replaces |
|---|---|---|---|
| Pay.UK Industry Sort Code Directory (ISCD) | Licensed | Every UK sort code to bank and branch | the 2-digit prefix table |
| Modulus-checking weights file (Vocalink) | Free download | Offline UK sort code and account check | length-only check |
| Experian Bank Wizard, Loqate Bank Account Validation, similar | Paid APIs | Sort code and account validation with bank details | the whole UK bank path |
| Confirmation of Payee (via banks or Open Banking providers) | Bank service | Name matches the account | nothing today |
| SWIFTRef (SWIFT) | Paid | BIC, IBAN structure and bank directory worldwide | SWIFT handling and the 16-country IBAN table |
| `schwifty`, `python-stdnum` | Free libraries | IBAN mod-97 checksum, BIC format, bank lookups for some countries | regex-only IBAN |
| openiban.com, iban.com APIs | Free / paid | IBAN validation and bank name | same |
| US FedACH directory | Free | ABA routing numbers | tenant range (acceptable as is) |

### Tax id, organisations, names, email
| Option | Type | What it gives |
|---|---|---|
| HMRC VAT number check API; EU VIES | Free | Is this VAT number live, and whose |
| Companies House API; GLEIF LEI | Free | Authoritative UK company name and number; legal entity id |
| Dun & Bradstreet, Creditsafe, Bureau van Dijk, OpenCorporates | Paid | Entity resolution and matching across suppliers |
| `python-stdnum` | Free | Checksums for many tax id types |
| ZeroBounce, NeverBounce, Kickbox, Mailgun validation | Paid APIs | Deliverability, MX, disposable detection |
| `nameparser` (free); Melissa, Experian name tools | Free / paid | Person-name splitting and title handling |

### Where I would look first
1. **Free, high value, low effort:** IBAN mod-97 and BIC format via a library; Vocalink modulus check; HMRC and VIES tax checks; ONSPD in place of the district table.
2. **Paid, replaces the weakest hand-built part:** a UK bank validation service (ISCD-backed) and an address validation API.
3. **Not worth buying for this data:** phone lookup (the current approach is sound for batch migration).

## 8. Framework starting point: rules to adopt

These are proposals for discussion, derived from the failures above.

1. **Keep raw and conformed side by side.** Today recoveries are fill-blank only and sources are read-only; keep that.
2. **One country-rules source, keyed on ISO code,** never on country name. Workday, libaddressinput and the product UI should read from the same table.
3. **Validate at capture, conform in batch.** A service or library at the point of entry (typeahead, checksum) prevents what the ELT currently repairs.
4. **Layer the checks:** syntax (cheap, offline) then existence (reference data) then verification (paid, per transaction only where risk warrants).
5. **Never drop silently.** Every rejection writes a reason; give issues a severity.
6. **Select by risk, not uniformly.** Bank details justify paid verification; a phone number does not.
7. **Document provenance and refresh date for every reference dataset.**
8. **Decide the overflow rule once** for address lines, and test it for every target.

For the personal-info design work specifically: names, addresses, phone and email map directly onto sections 2 to 7. Parent/guardian, date of birth, consent and under-18 handling are **not explored anywhere in this repo** and need their own input (including data-protection requirements) rather than inheriting these rules.
