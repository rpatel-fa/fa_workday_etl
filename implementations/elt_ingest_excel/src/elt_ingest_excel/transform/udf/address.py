from typing import TYPE_CHECKING
import re
import duckdb
try:
    import requests
except Exception:
    requests = None
from duckdb.typing import VARCHAR

if TYPE_CHECKING:
    import duckdb as duckdb_types

_DISALLOWED = re.compile(r'["`<>|;{}]')
_WS = re.compile(r"\s+")

_PAT_GB = re.compile(r"\b[A-PR-UWYZ][A-HK-Y]?\d[A-Z\d]?\s?\d[ABD-HJLNP-UW-Z]{2}\b", re.IGNORECASE)

# Canonical single-valued OCR confusion pairs (letter<->digit) for UK postcode repair.
_GB_L2D = {"O": "0", "I": "1", "L": "1", "S": "5", "Z": "2", "B": "8", "G": "6"}
_GB_D2L = {"0": "O", "1": "I", "5": "S", "2": "Z", "8": "B", "6": "G"}
# Valid outward-code templates by length ('A' = letter slot, 'N' = digit slot).
# The inward code is always the trailing 3 chars in the fixed shape digit-letter-letter.
_GB_OUTWARD = {2: ("AN",), 3: ("ANN", "AAN", "ANA"), 4: ("AANN", "AANA")}


def _gb_ocr_repair(text: str | None) -> str | None:
    """Repair a UK postcode whose only defect is a single letter/digit OCR
    confusion (e.g. 'WCIH 9LG'->'WC1H 9LG', 'S015 2JU'->'SO15 2JU',
    'BRG 0HX'->'BR6 0HX', 'AL7 IEW'->'AL7 1EW').

    Deliberately conservative to avoid fabricating a postcode that merely looks
    valid (e.g. the county word 'ESSEX' must NOT become 'E5 5EX'):
      - at most ONE character may be substituted;
      - the two trailing inward letters are never guessed from digits (only the
        inward leading digit and the structured outward code may be coerced);
      - the result must be unambiguous and match the authoritative UK matcher.
    Returns the spaced postcode, or None if no safe single-fix repair exists.
    """
    a = re.sub(r"[^0-9A-Z]", "", (text or "").upper())
    if not (5 <= len(a) <= 7):
        return None
    outward, inward = a[:-3], a[-3:]
    # Inward leading position must be a digit; only coerce a letter there. The two
    # trailing positions must already be letters - a digit there is unverifiable.
    i0 = inward[0] if inward[0].isdigit() else _GB_L2D.get(inward[0])
    if i0 is None or not (inward[1].isalpha() and inward[2].isalpha()):
        return None
    inward_fixed = i0 + inward[1] + inward[2]
    subs_inward = 0 if inward[0].isdigit() else 1
    templates = _GB_OUTWARD.get(len(outward))
    if not templates:
        return None
    # A 2-char outward (letter+digit) carries no redundancy: coercing it merely
    # forces an incomplete code to look valid (e.g. 'IG' -> 'I6', when the real
    # value is a truncated 'IG1'). Only allow outward coercion when there is
    # enough structure (>= 3 chars) for a swap to be a genuine correction.
    allow_outward_subs = len(outward) >= 3
    candidates: set[str] = set()
    for tmpl in templates:
        fixed: list[str] = []
        subs = subs_inward
        ok = True
        for ch, slot in zip(outward, tmpl):
            if slot == "A":  # letter slot
                if ch.isalpha():
                    fixed.append(ch)
                elif ch in _GB_D2L and allow_outward_subs:
                    fixed.append(_GB_D2L[ch]); subs += 1
                else:
                    ok = False; break
            else:  # digit slot
                if ch.isdigit():
                    fixed.append(ch)
                elif ch in _GB_L2D and allow_outward_subs:
                    fixed.append(_GB_L2D[ch]); subs += 1
                else:
                    ok = False; break
        if not ok or subs > 1:
            continue
        cand = "".join(fixed) + inward_fixed
        if re.fullmatch(_PAT_GB, cand):
            candidates.add(cand)
    if len(candidates) != 1:
        return None
    g = candidates.pop()
    return g[:-3] + " " + g[-3:]


_PAT_US = re.compile(r"\b\d{5}(?:-\d{4})?\b")
_PAT_DE = re.compile(r"\b\d{5}\b")
_PAT_FR = re.compile(r"\b\d{5}\b")
_PAT_NL = re.compile(r"\b\d{4}\s?[A-Z]{2}\b", re.IGNORECASE)
_PAT_IE = re.compile(r"\b[A-Z][0-9][A-Z0-9]\s?[A-Z0-9]{4}\b", re.IGNORECASE)

_CACHE_POSTCODES: dict[str, dict] = {}
_CACHE_ZIP: dict[str, dict] = {}
_CACHE_GEOCODE: dict[str, dict] = {}


def _clean(s: str | None) -> str | None:
    if s is None:
        return None
    t = _DISALLOWED.sub("", s)
    t = _WS.sub(" ", t).strip()
    return t if t else None


def _first(*vals: str | None) -> str | None:
    for v in vals:
        if v is not None and str(v).strip():
            return v
    return None


def _find_postcode(text: str, cc_hint: str | None) -> tuple[str | None, str | None]:
    m = _PAT_GB.search(text)
    if m:
        return _WS.sub("", m.group(0)).upper(), "GB"
    m = _PAT_NL.search(text)
    if m:
        g = m.group(0).upper().replace(" ", "")
        return g, "NL"
    m = _PAT_US.search(text)
    if m:
        return m.group(0), "US"
    m = _PAT_IE.search(text)
    if m:
        return m.group(0).upper(), "IE"
    m = _PAT_DE.search(text)
    if m:
        return m.group(0), "DE"
    m = _PAT_FR.search(text)
    if m:
        return m.group(0), "FR"
    return None, cc_hint


def _libpostal_parse(address: str) -> dict | None:
    try:
        from postal.parser import parse_address
    except Exception:
        return None
    try:
        parts = parse_address(address)
        out: dict[str, str] = {}
        for value, label in parts:
            out[label] = value
        return out
    except Exception:
        return None


def _lookup_gb(postcode: str) -> dict | None:
    if requests is None:
        return None
    key = postcode.upper()
    if key in _CACHE_POSTCODES:
        return _CACHE_POSTCODES[key]
    try:
        r = requests.get(f"https://api.postcodes.io/postcodes/{key}", timeout=3)
        if r.status_code == 200:
            data = r.json().get("result") or {}
            _CACHE_POSTCODES[key] = data
            return data
    except Exception:
        pass
    return None


def _lookup_us(zipcode: str) -> dict | None:
    if requests is None:
        return None
    key = zipcode
    if key in _CACHE_ZIP:
        return _CACHE_ZIP[key]
    try:
        r = requests.get(f"https://api.zippopotam.us/us/{key}", timeout=3)
        if r.status_code == 200:
            data = r.json()
            _CACHE_ZIP[key] = data
            return data
    except Exception:
        pass
    return None


def _geocode(address: str, cc_hint: str | None) -> dict | None:
    if requests is None:
        return None
    key = (address + "|" + (cc_hint or "")).lower()
    if key in _CACHE_GEOCODE:
        return _CACHE_GEOCODE[key]
    params = {
        "q": address,
        "format": "jsonv2",
        "addressdetails": 1,
        "limit": 1,
    }
    if cc_hint:
        params["countrycodes"] = cc_hint.lower()
    try:
        r = requests.get("https://nominatim.openstreetmap.org/search", params=params, timeout=4, headers={"User-Agent": "elt-ingest-excel"})
        if r.status_code == 200:
            arr = r.json()
            if isinstance(arr, list) and arr:
                res = arr[0]
                _CACHE_GEOCODE[key] = res
                return res
    except Exception:
        pass
    return None


def register(conn: "duckdb.DuckDBPyConnection") -> None:
    def udf_normalize_address(
        address_1: str | None,
        address_2: str | None,
        address_3: str | None,
        address_4: str | None,
        city: str | None,
        region: str | None,
        postal_code: str | None,
        country_code: str | None,
    ) -> dict[str, str | None]:
        a1 = _clean(address_1)
        a2 = _clean(address_2)
        a3 = _clean(address_3)
        a4 = _clean(address_4)
        c0 = _clean(city)
        r0 = _clean(region)
        p0 = _clean(postal_code)
        cc0 = (country_code or "").upper() or None
        text = " ".join([x for x in [a1, a2, a3, a4, c0, r0, p0] if x])
        p1, cc1 = _find_postcode(text, cc0)
        p = p0 or p1
        cc = cc1 or cc0
        norm_city = c0
        norm_region = r0
        norm_cc = cc
        norm_postcode = p
        if (cc or "").upper() == "GB":
            gb = None
            if p:
                gb = _lookup_gb(p)
            if gb is None and p1 and p1 != p:
                gb = _lookup_gb(p1)
            if gb:
                norm_postcode = gb.get("postcode") or p or p1
                post_town = gb.get("post_town")
                admin_county = gb.get("admin_county")
                region_name = gb.get("region")
                country_name = gb.get("country")
                norm_city = _clean(_first(post_town, norm_city))
                norm_region = _clean(_first(admin_county, region_name, norm_region))
                norm_cc = "GB"
        elif (cc or "").upper() == "US":
            us = None
            if p:
                us = _lookup_us(p)
            if us is None and p1 and p1 != p:
                us = _lookup_us(p1)
            if us and "places" in us and us["places"]:
                place = us["places"][0]
                norm_city = _clean(_first(place.get("place name"), norm_city))
                norm_region = _clean(_first(place.get("state"), norm_region))
                norm_cc = "US"
                norm_postcode = p or p1
        if not norm_city or not norm_region or not norm_postcode or not norm_cc:
            q = " ".join([x for x in [a1, a2, a3, a4, c0, r0] if x])
            lp = _libpostal_parse(q)
            if lp:
                norm_postcode = _clean(_first(lp.get("postcode"), norm_postcode))
                norm_city = _clean(_first(lp.get("city"), lp.get("suburb"), norm_city))
                norm_region = _clean(_first(lp.get("state"), lp.get("county"), lp.get("region"), norm_region))
                cc_guess = lp.get("country_code") or None
                if not cc_guess:
                    country_name = lp.get("country") or ""
                    cc_guess = country_name.strip().upper() if len(country_name.strip()) == 2 else None
                norm_cc = _first(cc_guess, norm_cc)
        if not norm_city or not norm_region or not norm_postcode or not norm_cc:
            q = " ".join([x for x in [a1, a2, a3, a4, c0, r0] if x])
            g = _geocode(q, cc)
            if g and isinstance(g, dict):
                addr = g.get("address") or {}
                norm_postcode = _clean(_first(addr.get("postcode"), norm_postcode))
                norm_cc = (addr.get("country_code") or norm_cc or "").upper() or None
                norm_city = _clean(
                    _first(
                        addr.get("city"),
                        addr.get("town"),
                        addr.get("village"),
                        addr.get("hamlet"),
                        norm_city,
                    )
                )
                norm_region = _clean(
                    _first(
                        addr.get("state"),
                        addr.get("county"),
                        addr.get("region"),
                        norm_region,
                    )
                )
        toks = {x.lower() for x in [norm_city or "", norm_region or ""] if x}
        b1 = None if a1 and a1.lower() in toks else a1
        b2 = None if a2 and a2.lower() in toks else a2
        b3 = None if a3 and a3.lower() in toks else a3
        b4 = None if a4 and a4.lower() in toks else a4
        return {
            "address_line_1": b1,
            "address_line_2": b2,
            "address_line_3": b3,
            "address_line_4": b4,
            "city": norm_city,
            "region": norm_region,
            "postal_code": norm_postcode,
            "country_code": norm_cc,
        }

    conn.create_function(
        "udf_normalize_address",
        udf_normalize_address,
        [VARCHAR, VARCHAR, VARCHAR, VARCHAR, VARCHAR, VARCHAR, VARCHAR, VARCHAR],
        duckdb.struct_type(
            {
                "address_line_1": "VARCHAR",
                "address_line_2": "VARCHAR",
                "address_line_3": "VARCHAR",
                "address_line_4": "VARCHAR",
                "city": "VARCHAR",
                "region": "VARCHAR",
                "postal_code": "VARCHAR",
                "country_code": "VARCHAR",
            }
        ),
        null_handling="special",
    )

    def udf_normalize_postal_code(
        postal_code_raw: str | None,
        validation_rule: str | None,
    ) -> str | None:
        raw = _clean(postal_code_raw)
        if not raw:
            return None
        rule = (validation_rule or "").strip()
        rl = rule.casefold()
        s = _WS.sub(" ", raw).strip().upper()
        digits = re.sub(r"[^0-9]", "", s)
        alnum = re.sub(r"[^0-9A-Z]", "", s)
        if rl.startswith("regex:"):
            spec = rule.split(":", 1)[1].strip()
            pat, repl = (spec.split("=>", 1) + [None])[:2]
            pat = pat.strip()
            repl = repl.strip() if repl is not None else None
            flags = re.IGNORECASE
            for candidate in (alnum, s):
                try:
                    m = re.fullmatch(pat, candidate, flags=flags)
                except re.error:
                    return None
                if not m:
                    continue
                if repl is None or repl == "":
                    return candidate or None
                try:
                    out = re.sub(pat, repl, candidate, flags=flags)
                except re.error:
                    return None
                out = _WS.sub(" ", out).strip().upper()
                return out or None
            return None
        if not rl or rl == "no validation":
            return s or None
        if "zip+4" in rl or "#####-####" in rl:
            if len(digits) >= 9:
                return digits[:5] + "-" + digits[5:9]
            if len(digits) >= 5:
                return digits[:5]
            if len(digits) == 4:
                return digits.zfill(5)
            return None
        if "3 or 5 digits" in rl:
            if len(digits) >= 5:
                return digits[:5]
            if len(digits) >= 3:
                return digits[:3]
            return None
        if "5 or 7 digits" in rl:
            if len(digits) >= 7:
                return digits[:7]
            if len(digits) >= 5:
                return digits[:5]
            return None
        if "5 or 8 digits" in rl:
            if len(digits) >= 8:
                return digits[:8]
            if len(digits) >= 5:
                return digits[:5]
            return None
        if "5 digits, or 6 digits" in rl:
            if len(digits) >= 6:
                return digits[:6]
            if len(digits) >= 5:
                return digits[:5]
            return None
        if "5 digits, first three digits must be '980'" in rl:
            if len(digits) >= 5:
                return digits[:5]
            return None
        if "4 digits, may be proceeded by 'b-'" in rl or "4 digits, may be proceeded by 's-'" in rl:
            if len(digits) >= 4:
                return digits[:4]
            return None
        if "four digits, may be preceeded by 'cc' and a digit" in rl:
            if len(digits) >= 4:
                return digits[-4:]
            return None
        if "may be preceeded by" in rl or "may be preceded by" in rl:
            pm = re.search(r"'([A-Z]{1,3})'", rule.upper())
            nm = re.search(r"postal code must be\s*(\d+)\s*digits", rl)
            if pm and nm:
                prefix = pm.group(1)
                n = int(nm.group(1))
                if alnum.startswith(prefix) and len(alnum) >= len(prefix) + n:
                    tail = re.sub(r"[^0-9]", "", alnum[len(prefix) :])
                    if len(tail) >= n:
                        return prefix + tail[:n]
                if len(digits) >= n:
                    return digits[:n]
                return None
        if "4 digits followed by 2 letters" in rl:
            m = re.search(r"(\d{4})([A-Z]{2})", alnum)
            if m:
                return m.group(1) + " " + m.group(2)
            return None
        if "format: a#a #a#" in rl:
            m = re.search(r"([A-Z]\d[A-Z])(\d[A-Z]\d)", alnum)
            if m:
                return m.group(1) + " " + m.group(2)
            return None
        if "eircode" in rl:
            m = re.search(r"([A-Z]\d[A-Z0-9])([A-Z0-9]{4})", alnum)
            if m:
                return m.group(1) + " " + m.group(2)
            return None
        if "postal code is made up of two parts" in rl:
            m = _PAT_GB.search(s)
            if m:
                g = _WS.sub("", m.group(0)).upper()
                return (g[:-3] + " " + g[-3:]) if len(g) > 3 else g
            m2 = re.search(r"([A-Z]{1,2}\d[A-Z\d]?\d[A-Z]{2})", alnum)
            if m2:
                g = m2.group(1)
                return (g[:-3] + " " + g[-3:]) if len(g) > 3 else g
            # Common data-entry / OCR confusion: a single letter<->digit swap
            # (e.g. 'HA9 OWS'->'HA9 0WS', 'WCIH 9LG'->'WC1H 9LG', 'S015 2JU'->
            # 'SO15 2JU', 'BRG 0HX'->'BR6 0HX'). Attempt one safe, self-validating
            # repair before giving up. See _gb_ocr_repair for the guardrails that
            # stop it fabricating a plausible-but-wrong postcode.
            repaired = _gb_ocr_repair(s)
            if repaired is not None:
                return repaired
            # No fragment matches a genuine UK postcode shape - returning the raw
            # alnum text here (e.g. an address line with no real postcode, like
            # '6THIEPVALDRIVE') produced a garbage postal_code value that looked
            # plausible but wasn't a postcode, masking a genuinely missing source
            # value instead of surfacing it.
            return None
        if "fiqq 1zz" in rl:
            if "FIQQ1ZZ" in alnum:
                return "FIQQ 1ZZ"
            return None
        if "format: bb#d 1zz" in rl:
            m = re.search(r"(BB\d[A-Z])(\d[A-Z]{2})", alnum)
            if m:
                return m.group(1) + " " + m.group(2)
            return None
        if "1 letter followed by 4 digits and 3 letters" in rl:
            m = re.search(r"([A-Z]\d{4}[A-Z]{3})", alnum)
            if m:
                return m.group(1)
            m2 = re.search(r"([A-Z]\d{4})", alnum)
            return m2.group(1) if m2 else None
        if "4 characters, first 2 characters must be letters" in rl:
            m = re.search(r"([A-Z]{2}[A-Z0-9]{2})", alnum)
            return m.group(1) if m else None
        if "4 characters: 1 letter followed by 3 digits" in rl:
            m = re.search(r"([A-Z]\d{3})", alnum)
            return m.group(1) if m else None
        if "8 characters: 1 digit, 3 letters, then 4 digits" in rl:
            m = re.search(r"(\d[A-Z]{3}\d{4})", alnum)
            return m.group(1) if m else None
        if "format: 'je'##aa" in rl or "format: 'gy'##aa" in rl or "format: 'im'##aa" in rl:
            m = re.search(r"((?:JE|GY|IM)\d{2}[A-Z]{2})", alnum)
            return m.group(1) if m else None
        if "format: 'jm'aa##" in rl:
            m = re.search(r"(JM[A-Z]{2}\d{2})", alnum)
            return m.group(1) if m else None
        m = re.search(r"postal code must be (\d+) digits", rl)
        if m:
            n = int(m.group(1))
            if len(digits) >= n:
                return digits[:n]
            if len(digits) == n - 1:
                return digits.zfill(n)
            return None
        return s or None

    conn.create_function(
        "udf_normalize_postal_code",
        udf_normalize_postal_code,
        [VARCHAR, VARCHAR],
        VARCHAR,
        null_handling="special",
    )
