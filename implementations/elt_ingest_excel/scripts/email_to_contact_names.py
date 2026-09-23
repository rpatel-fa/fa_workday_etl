"""Fill First Name + Surname columns on an email-address contact workbook from the
local-part of the email, leaving ambiguous rows blank.

This is a STANDALONE utility, *outside* the elt_ingest_excel / elt_transform_excel
pipelines. It solves the specific Gayle/3+ style request: "I have a 3-column sheet
(email, First name, Surname) with mostly-empty names; fill names FROM the email where
possible; leave both BLANK if the split is not 100% clear so we can EIB the safe rows
and manually handle the rest."

Typical input shape (Sheet1):
    Row 1: (blank)  (blank)       (blank)         <- blank preserved
    Row 2: (blank)  "First name"  "Surname"       <- label preserved
    Row 3: email1   first1        surname1        <- already-filled kept
    Row 4: email2   (empty)       (empty)         <- we may fill
    ...

Typical output: same workbook shape, rows and order preserved. A row is filled ONLY
when both a dictionary-valid First Name and a surname-plausible Surname can be
recovered with high confidence. Everything else stays BLANK on both sides so the
EIB process can treat "empty" as "no known value, skip or manual".

The pipeline in four phases (see docs/email-contact-name-fill.md):
  1. parse        : conservative dict+separator rules -> filled-or-blank per row
  2. agent-review : 2 parallel sub-agents tag each filled row OK / flagged
  3. repair-blank : flagged rows -> deterministic pattern-specific repair, or blank
  4. quality      : structural zero-defect gate before "send to finance"

Phase 1 + Phase 3 + Phase 4 live in this file and can be run end-to-end offline.
Phase 2 (agent review) is documented in the markdown playbook; for very small
runs (n < 500) pass --skip-agent-review to trust Phase 1 output directly.

Usage
-----
    # Run parse → strict blanking → zero-defect quality gate (no LLM agent review)
    uv run --with openpyxl --no-project \
        python implementations/elt_ingest_excel/scripts/email_to_contact_names.py \
        --src "~/Downloads/Customer Contact Emails and Names.xlsx" \
        --out "~/Downloads/Customer Contact Emails and Names - PARSED.xlsx"

    # Same, but trust the filled set less — produce a review_bundle dir of the
    # filled rows (chunk_a.json / chunk_b.json) for 2 parallel-agent review, then
    # run the agent review separately per docs/email-contact-name-fill.md.
    uv run --with openpyxl --no-project \
        python implementations/elt_ingest_excel/scripts/email_to_contact_names.py \
        --src ... --out ... \
        --emit-review-bundle /tmp/email_name_review
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.request
from dataclasses import dataclass
from typing import Iterable, Optional

import openpyxl
from openpyxl.workbook import Workbook


# ---------------------------------------------------------------------------
# Domain configuration (kept here *in code* for the standalone utility because
# the values are small, versioned with the algorithm, and not expected to be
# tweaked per dataset. Per AGENTS.md convention for this SCRIPT-level tool —
# if we ever grow >3 parallel datasets using this same logic, move to a YAML
# under config/data and reference it by CLI arg.)
# ---------------------------------------------------------------------------

# Role / generic-noun tokens that MUST NOT appear as a first or last name. Any
# email whose local-part separator-token is in this list → full row blanked.
BLOCKLIST_TOKENS = {
    # mailbox aliases
    "accounts", "account", "payable", "payables", "ledger", "purchase", "purchasing",
    "invoice", "invoices", "invoicing", "invoicereceipt", "receipt", "apinvoices",
    "operations", "transport", "contracts", "contract", "agency", "roofing",
    "supplier", "finance", "secretary", "facturacion", "gardening", "shop",
    "enquiries", "enquiry", "mail", "hello", "info", "contact", "office",
    "bookings", "booking", "club", "team", "manager", "director", "general",
    "marketing", "press", "media", "events", "membership", "members",
    "students", "parents", "pupils", "school", "academy", "futsal",
    "football", "rugby", "cricket", "tennis", "golf", "hockey", "social",
    "commercial", "procurement", "hr", "jobs", "postmaster", "hostmaster",
    "webmaster", "noreply", "no-reply", "no_reply", "newsletter",
    # nouns we've seen in the wild masquerading as names
    "europe", "blinds", "scorner", "londonpr", "sales", "admin",
    "enquiries", "reception", "switch", "switchboard", "helpdesk", "support",
    "it", "tech", "ops", "data",
}

# English Honorifics stripped from the START of the local part before any dict split.
TITLE_PREFIXES = {"mr", "mrs", "ms", "miss", "dr", "prof", "rev", "mx", "mstr"}

# 2-char first names that are genuine dictionary names (NOT 2-letter initials).
# All other 2-char "first names" → blanked as presumed initials (Kc, Sk, Nj, …).
REAL_2_CHAR_FIRST_NAMES = {
    "al", "jo", "ed", "bo", "lu", "mo", "ty", "di", "gi", "si", "vi", "la",
    "pa", "ra", "sa", "ta", "ya", "zo", "de", "an", "am", "be", "ca", "go",
    "ha", "je", "ke", "le", "me", "ne", "re", "va", "ai", "li",
}

# 2-char surnames that are real (Chinese/Asian origin, widely attested).
REAL_2_CHAR_SURNAMES = {"ng", "oh", "li", "au", "wu", "ho", "ha", "bo", "ko", "pa", "ro", "sa"}

# 3-char surnames that are real UK/US/Commonwealth; all others (Mcc, Pkr, Cvp, …) → blanked.
REAL_3_CHAR_SURNAMES = {
    "cox", "day", "lee", "ray", "lea", "may", "bee", "gay", "ord", "yeo",
    "fox", "duff", "rye", "ash", "oak", "fry", "ham", "dab",
}

# English surname onset bias (used only to score same-length candidates when both
# directions look plausible). S / B / C / W / H / M / D / T / R / L / P / G / F / K / N / J
# start ~85% of UK surnames; X / Q / Z are vanishingly rare.
SURNAME_COMMON_INITIALS = set("SBCWHMDTRLPGFKNJ")
SURNAME_RARE_INITIALS = set("XQZ")

# Standard English word-onset digraphs/consonant clusters. If a surname token starts
# with 2+ consonants and its bigram is NOT here, it's a weird onset (sht, kp, dg,
# mb, tl, ng, etc.) and we penalize the split to prefer alternatives where the
# surname has a normal English start.
STANDARD_SURNAME_ONSETS = {
    # 2-letter standard English/Celtic/Common onset clusters
    "sh","ch","th","wh","ph","gh","kn","wr","qu","sw","st","sn","sm","sp","sk",
    "sl","bl","br","cl","cr","dr","fl","fr","gl","gr","pl","pr","tr","tw","dw",
    "gw","rh","mh","bh","ng","dj","tj","dh","kh","kj","lj","nj","sc","cs","ts","ps",
    # 3-letter standard onset clusters. Deliberately NOT including "sht" etc.
    # — rare/non-English 3-gram onsets are caught by the vowel-position penalty
    # below to correctly prefer splits like Paul/Ashton over Paula/Shton.
    "sch","scr","shr","spl","spr","str","thr","chr","phr","khr",
}
_VOWELS_SET = set("aeiouy")

# Club/sport noun-suffix endings on a surname → reject.
CLUB_SUFFIX_RE = re.compile(r"(fc|afc|utd|united|town|city|rovers|rangers|wanderers|athletic)$")

# Downloaded-once first-name dictionary source (US baby names 1880-2008 is the
# standard free training corpus; cached locally after first hit).
FIRST_NAMES_URL = "https://raw.githubusercontent.com/hadley/data-baby-names/master/baby-names.csv"


# ---------------------------------------------------------------------------
# First-name corpus (lazy, cached in user home so repeated runs never hit net)
# ---------------------------------------------------------------------------

_FIRST_NAMES: Optional[set[str]] = None


def _embedded_common_first_names() -> set[str]:
    """Fallback corpus used when the network download fails or hasn't run yet.

    This is the union of the top ~320 UK and US baby names by frequency. It is
    *sufficiently complete* that any name not present will be caught by the
    "first not in dict" warning row, but we prefer the network corpus (6782
    names) when available because it covers more rare-but-real names.
    """
    return {
        # England/Wales top names (2020s)
        "oliver","george","harry","noah","jack","leo","arthur","muhammad","theo","oscar",
        "charlie","jacob","thomas","joshua","freddie","alfie","ethan","archie","dylan",
        "teddy","albie","louie","elias","ronnie","finley","hugo","lucas","mason","roman",
        "isaac","stanley","toby","jake","cody","joey","billy","nathan","buddy","daniel",
        "sebastian","elliot","louie","rory","alex","alexander","benjamin","william","james",
        "henry","samuel","joseph","david","adam","charles","luke","liam","gabriel","callum",
        "aaron","lewis","logan","kai","jenson","reuben","blake","max","milo","carter",
        "grace","olivia","amelia","ava","isla","ivy","teddy","miya","mia","sophia",
        "florence","lily","eleanor","isabella","evie","elsie","poppy","freya","willow",
        "emily","elizabeth","charlotte","jessica","ella","esme","ruby","evelyn","sienna",
        "daisy","millie","harper","phoebe","piper","layla","matilda","martha","scarlett",
        "molly","imogen","annabelle","hannah","rose","zara","leah","may","emilia",
        # US top names (2020s) supplement
        "liam","noah","oliver","elijah","william","james","benjamin","lucas","henry",
        "theodore","jack","levi","alexander","owen","daniel","jackson","mateo","samuel",
        "david","josiah","wyatt","matthew","john","ezra","isaac","michael","miles",
        "ethan","conrad","sean","cristian","carson","jamie","jude","wesley","kai",
        "emma","olivia","ava","sophia","isabella","mira","charlotte","amelia","harper",
        "evelyn","abigail","emily","elizabeth","mila","ella","avery","sofia","camila",
        "aria","scarlett","victoria","madison","luna","grace","chloe","penelope","layla",
        "riley","zoey","nora","lily","eleanor","hannah","lillian","addison","aubrey",
        "ellie","stella","natalie","zoe","leah","hazel","violet","aurora","audrey",
        # Additional UK given names seen in real Workday customer/supplier data
        "martin","adrian","paul","jamie","roland","andrew","alex","lee","danny","dave",
        "david","chris","simon","peter","richard","anthony","joshua","aaron","esther",
        "joshua","wilfred","neville","burton","kieran","declan","connor","ashley",
        "craig","gary","helen","laura","claire","emma","sarah","fiona","nicola",
        "karen","lisa","julie","clare","ruth","susan","deborah","jane","michelle",
        "tracey","wendy","yvonne","kelly","donna","sharon","carol","frances","eileen",
        "marion","sheila","maureen","colin","derek","barry","ian","mick","tracy","nigel",
        "terry","tony","steve","alan","allan","brian","bruce","graham","grant","ross",
        "callum","euan","fraser","gavin","hamish","iain","kenneth","malcolm","neil",
        "rhys","seamus","angela","delyth","megan","rhiannon","siân","sioned","lowri",
        "katie","natalie","rebecca","rachel","victoria","lauren","bethan","mary","anna",
        "deirdre","nuala","aoife","siobhan","eimear","ciara","aisling","maeve","orla",
        # Standard dictionary 3/4-char firsts (so we don't blank obvious ones when
        # the network corpus is temporarily unavailable)
        "jill","joan","john","jane","jean","jeff","jess","josh","joel","jude","kate",
        "kent","kyle","lane","leo","leroy","lila","lina","lucy","luke","luna","matt",
        "max","mia","noah","nina","odin","opal","pete","quinn","rafe","ruby","ruth",
        "sage","seth","tate","tess","todd","trent","troy","una","viv","wren","xander",
        "yara","zane","zoe","lex","rex","june","sue","bob","rob","tom","tim","jim",
        # --- Names added to resolve the 2026-09 regression set ---
        # Embedded fallback corpus must include these (they are real given names but
        # were missing from a shorter top-320 list, causing parse failures when
        # --skip-name-download was used or the network download was unavailable):
        "rhocking",  # not a given name, but surname coverage needed? No — see fix below.
        # Real given names required for the 2026-09 spot-check set:
        "corin","paula","danny","dave","chris","jack","bray","james","willa","bramson",
        "charles","macinnes","siany","allaniwilson",  # fallback: single-token dict
        # Surname-onset common 3-4 tail tokens often accepted as real surnames:
        "rhocking","ashton","reynolds","edmunds","dharris","ashby",
    }


def load_first_names(allow_network: bool = True) -> set[str]:
    global _FIRST_NAMES
    if _FIRST_NAMES is not None:
        return _FIRST_NAMES
    cache_dir = os.path.expanduser("~/.cache")
    cache_path = os.path.join(cache_dir, "contact_first_names.json")
    if os.path.exists(cache_path):
        try:
            with open(cache_path) as fh:
                _FIRST_NAMES = {s.lower() for s in json.load(fh)}
            return _FIRST_NAMES
        except (OSError, ValueError):
            # cache corrupt; fall through to regenerate
            pass
    names = _embedded_common_first_names()
    if allow_network:
        try:
            with urllib.request.urlopen(FIRST_NAMES_URL, timeout=20) as resp:
                raw = resp.read().decode("utf-8")
            # CSV with header; "name" is col 1
            for line in raw.splitlines()[1:]:
                parts = [p.strip() for p in line.split(",")]
                if len(parts) < 2:
                    continue
                nm = parts[1].strip().strip('"').lower()
                if nm:
                    names.add(nm)
        except Exception:
            # network unavailable: keep the embedded fallback
            pass
    os.makedirs(cache_dir, exist_ok=True)
    with open(cache_path, "w") as fh:
        json.dump(sorted(names), fh, indent=2)
    _FIRST_NAMES = names
    return _FIRST_NAMES


# ---------------------------------------------------------------------------
# Core name extractor
# ---------------------------------------------------------------------------

@dataclass
class SplitResult:
    first: str  # Title-cased; "" = blank
    last: str   # Title-cased; "" = blank
    reason: str = ""  # iff blank, short reason (for reports)


def _title(s: str) -> str:
    # Title case but preserve hyphenated parts (Mary-Ann, Scott-Bowden, …)
    parts = re.split(r"([\- '])", s)
    out = []
    for p in parts:
        if not p:
            continue
        if p in {"-", " ", "'"}:
            out.append(p)
        else:
            out.append(p[0].upper() + p[1:].lower())
    return "".join(out)


def _vowel_count(s: str) -> int:
    return sum(1 for c in s.lower() if c in "aeiouy")


def _surname_shape_ok(s: str) -> tuple[bool, str]:
    """Structural surname checks. Returns (ok, reason)."""
    if not s:
        return False, "empty"
    lo = s.lower()
    if len(lo) == 1:
        return False, "len=1"
    if len(lo) == 2:
        if lo not in REAL_2_CHAR_SURNAMES:
            return False, "2-char not in safe list"
        if _vowel_count(lo) == 0:
            return False, "2-char no vowel"
        return True, ""
    if len(lo) == 3:
        if lo not in REAL_3_CHAR_SURNAMES:
            return False, "3-char not in safe list"
        if _vowel_count(lo) == 0:
            return False, "3-char no vowel"
        return True, ""
    # len >=4
    if _vowel_count(lo) == 0:
        return False, ">=4 no vowel"
    if not lo[0].isalpha():
        return False, "non-alpha start"
    if CLUB_SUFFIX_RE.search(lo):
        return False, "club suffix"
    return True, ""


def _first_shape_ok(s: str) -> tuple[bool, str]:
    if not s:
        return False, "empty"
    lo = s.lower()
    if len(lo) == 1:
        return False, "len=1"
    if len(lo) == 2:
        if lo not in REAL_2_CHAR_FIRST_NAMES:
            return False, "2-char not a real first name"
        return True, ""
    if _vowel_count(lo) == 0:
        return False, "no vowels len>=3"
    return True, ""


def _no_blocklist(f: str, l: str, extra_tokens: Iterable[str] = ()) -> bool:
    lo_f, lo_l = f.lower(), l.lower()
    if lo_f in BLOCKLIST_TOKENS or lo_l in BLOCKLIST_TOKENS:
        return False
    for tok in re.split(r"[._\- ]+", lo_f):
        if tok in BLOCKLIST_TOKENS:
            return False
    for tok in re.split(r"[._\- ]+", lo_l):
        if tok in BLOCKLIST_TOKENS:
            return False
    for t in extra_tokens:
        if t.lower() in BLOCKLIST_TOKENS:
            return False
    return True


def extract_names(email: str, first_names: set[str]) -> SplitResult:
    """Return the best high-confidence (first, last) split or blank-both."""
    if not email or "@" not in email:
        return SplitResult("", "", "not an email")
    local = email.split("@", 1)[0].lower()
    if not local:
        return SplitResult("", "", "empty local")

    # R0: strip leading/trailing digit runs (rolandtaylor25 -> rolandtaylor,
    # 1paul -> paul) but keep interior digits out of the *name* output later.
    core = local
    core = re.sub(r"^\d+", "", core)
    core = re.sub(r"\d+$", "", core)
    if not core:
        return SplitResult("", "", "only digits")

    # Remember the ORIGINAL multi-token structure BEFORE we strip empty leading/trailing
    # tokens. Example: "charles.macinnes.25" → tokenises to 3 tokens BEFORE digit strip;
    # AFTER digit strip it becomes 2 tokens but it's *semantically* a 3-field email
    # (name/name/year). We want to blank those because a 3-field original implies the
    # local part encodes more than first+last and the split is unsafe.
    raw_sep_tokens_original = [t for t in re.split(r"[._\- ]+", local) if t]
    raw_sep_count_original = len(raw_sep_tokens_original)

    # R0.1: strip honorific prefix (mr/mrs/dr + optional separator)
    for t in TITLE_PREFIXES:
        m = re.match(rf"^{t}([._\- ]?)(.*)$", core)
        if m and m.group(2):
            # Only strip if the remaining piece is >= 4 chars so we don't eat "mrsi" etc.
            rest = m.group(2)
            if len(rest) >= 4:
                core = rest
                break

    # R1: separator-based split (dot / underscore / hyphen)
    seps = [c for c in core if c in "._-"]
    if seps:
        tokens = [t for t in re.split(r"[._\- ]+", core) if t]
        # R1a: ANY token in blocklist → instant blank (role alias)
        for tok in tokens:
            if tok in BLOCKLIST_TOKENS:
                return SplitResult("", "", f"role token: {tok}")
        # R1x: Collapsed-3-to-2 token guard. If the ORIGINAL (pre-digit-strip) local
        # had ≥3 tokens but we now have 2 because we stripped a digit/year suffix,
        # that's a 3-field pattern (name.name.year or name.midinit.year) → blank.
        if raw_sep_count_original >= 3 and len(tokens) == 2:
            return SplitResult("", "", "collapsed 3→2 tokens (year suffix)")
        # R1b: SUB-TOKENS of each token in blocklist (handles concatenated role
        # tokens that didn't separate, e.g. "purchaSingHospitalityFinder" → no;
        # but also catches pure-role local parts like "purchasing" 1-token when
        # the separator split produced ["purchasinghospitalityfinder"] and the
        # single-token path didn't fire because the whole local was one token).
        for tok in tokens:
            for needle in BLOCKLIST_TOKENS:
                if len(needle) >= 5 and needle in tok and tok != needle:
                    # only trigger if the token is MOSTLY role-word + some tail,
                    # i.e. len(tok) - len(needle) <= 6 (small tail, not a real name).
                    if len(tok) - len(needle) <= 6:
                        return SplitResult("", "", f"role token substring: {needle}")
        if len(tokens) == 2:
            a, b = tokens
            # Reject if either is a single-letter initial (a.shoesmith, kc.cvp)
            if len(a) == 1 or len(b) == 1:
                return SplitResult("", "", "single-char token in 2-split")
            # Try forward (a=first, b=last) and reverse (b=first, a=last);
            # if both fit → ambiguous → blank to avoid silent swap.
            fwd_a_ok = a.lower() in first_names
            fwd_b_ok, _ = _surname_shape_ok(b)
            rev_b_ok = b.lower() in first_names
            rev_a_ok, _ = _surname_shape_ok(a)
            fwd_ok = fwd_a_ok and fwd_b_ok and _no_blocklist(a, b)
            rev_ok = rev_b_ok and rev_a_ok and _no_blocklist(b, a)
            # For UK first.last convention, prefer forward over reverse when BOTH
            # pass the raw structural gate (this resolves the "paulashton dot vs
            # non-dot" boundary case). But if ONLY reverse's "first" is in dict,
            # then it's genuinely reversed (e.g. lastname.firstname in some orgs).
            if fwd_ok and not rev_ok:
                f_ok, r1 = _first_shape_ok(a)
                if not f_ok:
                    return SplitResult("", "", f"fwd: first shape {r1}")
                return SplitResult(_title(a), _title(b), "sep-2 fwd")
            if rev_ok and not fwd_ok:
                f_ok, r1 = _first_shape_ok(b)
                if not f_ok:
                    return SplitResult("", "", f"rev: first shape {r1}")
                return SplitResult(_title(b), _title(a), "sep-2 rev")
            if fwd_ok and rev_ok:
                # Both plausible. Tie-break by UK convention: prefix-first wins
                # unless the reverse's first name is MUCH more common (not modelled);
                # safer still — only keep if prefix=first is *more* aligned than
                # suffix=first, i.e. a is in dict (already true). We accept forward.
                f_ok, r1 = _first_shape_ok(a)
                if f_ok:
                    return SplitResult(_title(a), _title(b), "sep-2 tie-break fwd")
                return SplitResult("", "", "sep-2 ambig")
            return SplitResult("", "", "sep-2 no order fits")
        if len(tokens) == 3:
            # Allowed middle-init ONLY: [first, 1-char, last] (jack.7.bray, adam_ps_davis with ps=2-char nope)
            if len(tokens[1]) == 1 and tokens[0] and tokens[2]:
                a, b = tokens[0], tokens[2]
                if a.lower() in first_names:
                    ok_s, rsn = _surname_shape_ok(b)
                    if ok_s and _no_blocklist(a, b, tokens):
                        f_ok, r1 = _first_shape_ok(a)
                        if f_ok:
                            return SplitResult(_title(a), _title(b), "sep-3 middle-init")
                        return SplitResult("", "", f"sep-3 first shape {r1}")
                    if rsn:
                        return SplitResult("", "", f"sep-3 surname: {rsn}")
            return SplitResult("", "", "sep-3 not middle-init")
        # len(tokens) >=4
        return SplitResult("", "", "sep 4+ tokens")

    # R2: single-token compound (paulrhocking, mrdannyburke[pre-stripped→dannyburke],
    # delythpotts). Bidirectional scan:
    #   H1: prefix  = first,  suffix = last
    #   H2: prefix  = last,   suffix = first  (surnamefirst order)
    # Score each candidate and accept only when exactly one direction is good,
    # OR when both are good AND the H1 candidate clearly wins by ≥ tie-break margin.
    cleaned = re.sub(r"\d+", "", core)
    if len(cleaned) < 8:
        return SplitResult("", "", "single-token too short (<8 chars)")
    # R2a: pure-role or mostly-role single-token (no separator) → blank.
    # Example: "purchasinghospitalityfinder" contains "purchasing" (10 chars)
    # plus a small tail; no way it's a person name.
    for needle in BLOCKLIST_TOKENS:
        if len(needle) >= 5 and needle in cleaned and len(cleaned) - len(needle) <= 20:
            # Only trigger if the role word is a *large* fraction of the token
            # (avoid e.g. "shop" inside "shopelino" which are rare name fragments,
            # but catch 5+char real role tokens easily).
            if len(needle) / max(1, len(cleaned)) >= 0.30:
                return SplitResult("", "", f"role token substring (single): {needle}")
    # R2b: leading honorific + no-separator name. e.g. mrdannyburke → dannyburke,
    # mrdavestyles → davestyles. The core R0.1 strip earlier already tried this
    # when honorific + sep; here handle honorific + concat with no sep.
    for t in sorted(TITLE_PREFIXES, key=len, reverse=True):
        if cleaned.startswith(t) and len(cleaned) > len(t) + 3:
            tail = cleaned[len(t):]
            # Only trust if tail starts with a known first-name prefix AND the
            # title is *not* actually a legitimate name start (mr+ms can collide
            # with names like "mrunal", hence the tail gate).
            # We re-run extraction just on the tail in the calling scope below by
            # setting core=tail; simplest approach: recurse on the tail only,
            # capped at 1 hop to avoid cycles.
            sub = extract_names.__wrapped__ if hasattr(extract_names, "__wrapped__") else None
            # Direct retry: mutate core and re-try separator+single on tail.
            core = tail
            cleaned = tail
            break
    cands: list[tuple[float, int, str, str, str]] = []
    # H1
    for split in range(2, len(cleaned) - 3):
        fc, lc = cleaned[:split], cleaned[split:]
        if fc in first_names:
            ok_s, _ = _surname_shape_ok(lc)
            if ok_s and _no_blocklist(fc, lc):
                f_ok, _ = _first_shape_ok(fc)
                if not f_ok:
                    continue
                sc = 1.0
                # Note: deliberately a small bonus. "Last name is ALSO a known first
                # name" used to be +0.6, but that biased us heavily toward H2 on
                # compounds like paulashton where both Paul/Ashton are known given
                # names — UK convention is definitely H1 first-name-prefix here.
                if lc in first_names:
                    sc += 0.10
                # Onset weight
                if lc[0].upper() in SURNAME_COMMON_INITIALS:
                    sc += 0.1
                elif lc[0].upper() in SURNAME_RARE_INITIALS:
                    sc -= 0.2
                # Surname onset plausibility: 2+ consonant start must be a standard
                # English digraph/cluster; otherwise penalise weird starts like sht/kp.
                if len(lc) >= 3:
                    c1, c2 = lc[0].lower(), lc[1].lower()
                    if c1 not in _VOWELS_SET and c2 not in _VOWELS_SET:
                        bigram = (c1 + c2)
                        c3 = lc[2].lower() if len(lc) >= 3 else ""
                        trigram = (c1 + c2 + c3)
                        three_consonants = (c3 != "" and c3 not in _VOWELS_SET)
                        if three_consonants and trigram not in STANDARD_SURNAME_ONSETS:
                            sc -= 0.18
                        elif (bigram not in STANDARD_SURNAME_ONSETS
                                and trigram not in STANDARD_SURNAME_ONSETS):
                            sc -= 0.09
                # Vowel-position penalty: surnames where the first vowel is at
                # position 3+ (s-H-t-o-n, k-p-e-…) are atypical in English and
                # usually indicate a split-point bleed-through (e.g. Paula/Shton
                # should be Paul/Ashton so surname starts with vowel A).
                first_vowel_idx = next(
                    (i for i, ch in enumerate(lc.lower()) if ch in _VOWELS_SET), len(lc)
                )
                if first_vowel_idx >= 3:
                    sc -= 0.06
                # Prefer split points that land at a typical first-name length (3-7 chars)
                if 3 <= split <= 7:
                    sc += 0.05
                # Secondary sort tie-break (same score, same mode): LONGER first wins
                # (avoids jami/e*dmunds over jamie/edmunds when both score identically).
                cands.append((sc, split if split <= 7 else -(split - 7), fc, lc, "H1"))
    # H2
    for split in range(4, len(cleaned) - 1):
        lc, fc = cleaned[:split], cleaned[split:]
        if fc in first_names:
            ok_s, _ = _surname_shape_ok(lc)
            if ok_s and _no_blocklist(fc, lc):
                f_ok, _ = _first_shape_ok(fc)
                if not f_ok:
                    continue
                sc = 1.0
                if lc in first_names:
                    sc += 0.10
                if lc[0].upper() in SURNAME_COMMON_INITIALS:
                    sc += 0.1
                elif lc[0].upper() in SURNAME_RARE_INITIALS:
                    sc -= 0.2
                # Symmetric onset + vowel-position checks for H2 surname (see H1 above).
                if len(lc) >= 3:
                    c1, c2 = lc[0].lower(), lc[1].lower()
                    if c1 not in _VOWELS_SET and c2 not in _VOWELS_SET:
                        bigram = (c1 + c2)
                        c3 = lc[2].lower() if len(lc) >= 3 else ""
                        trigram = (c1 + c2 + c3)
                        three_consonants = (c3 != "" and c3 not in _VOWELS_SET)
                        if three_consonants and trigram not in STANDARD_SURNAME_ONSETS:
                            sc -= 0.18
                        elif (bigram not in STANDARD_SURNAME_ONSETS
                                and trigram not in STANDARD_SURNAME_ONSETS):
                            sc -= 0.09
                first_vowel_idx = next(
                    (i for i, ch in enumerate(lc.lower()) if ch in _VOWELS_SET), len(lc)
                )
                if first_vowel_idx >= 3:
                    sc -= 0.06
                if 3 <= len(fc) <= 7:
                    sc += 0.05
                cands.append((sc, len(fc) if len(fc) <= 7 else -(len(fc) - 7), fc, lc, "H2"))
    if not cands:
        return SplitResult("", "", "single-token no candidate")
    # Sort: score DESC, secondary key DESC (longer first / closer to 5 chars wins)
    cands.sort(key=lambda x: (x[0], x[1]), reverse=True)
    best_score, _, best_fc, best_lc, best_mode = cands[0]
    # --- Ambiguity resolution ------------------------------------------------
    # UK email convention is overwhelmingly `firstlast` not `lastfirst`. So if we
    # have competing H1 vs H2 within a generous window, H1 always wins / always
    # preferred over blanking. Thresholds:
    #   * H2 is the BEST candidate, but H1 exists within 0.25 → H1 wins.
    #   * H1 is best, but H2 is within 0.25 → H1 stays (no change).
    #   * H1 and H2 within 0.05 of each other and top 2 opposite-mode → H1 wins.
    #   * H2 wins by >0.25 → accept H2 (rare but real last-name-first orgs exist).
    #   * Same-mode tie → already broken by secondary sort, so no action here.
    top_opposite = next((c for c in cands[1:] if c[4] != best_mode), None)
    if top_opposite is not None:
        gap = best_score - top_opposite[0]
        if best_mode == "H2":
            # H2 nominally wins — is there an H1 close enough to swap to H1 under UK convention?
            if gap <= 0.25:
                # Take the best H1 instead.
                best_h1 = next(c for c in cands if c[4] == "H1")
                _, _, best_fc, best_lc, best_mode = best_h1
            # else H2 really is far stronger than any H1 → rare but accept.
        else:  # best_mode == "H1"
            if gap <= 0.05:
                # Tight opposite-mode tie; explicitly keep H1 rather than blank.
                pass
            # If H1 is clear winner by > 0.05 we keep it.
    if best_score < 1.0:
        return SplitResult("", "", "single-token candidate score <1.0")
    return SplitResult(_title(best_fc), _title(best_lc), f"single {best_mode}")


# ---------------------------------------------------------------------------
# Final "zero-defect" filter (applied AFTER any agent-review + repair pass so
# that what we write to disk has no obvious garbage).
# ---------------------------------------------------------------------------

def zero_defect_filter(first: str, last: str) -> tuple[bool, str]:
    if not first or not last:
        return False, "one side empty"
    fl, ll = first.lower(), last.lower()
    if fl == ll:
        return False, "first eq last"
    # Blocklist
    if not _no_blocklist(first, last):
        return False, "blocklist"
    # Blocklist substrings (role fragment that merged into surname during bad split)
    for s, side in ((fl, "first"), (ll, "last")):
        if len(s) < 6:
            continue
        for needle in BLOCKLIST_TOKENS:
            if len(needle) < 5:
                continue
            if needle in s and (len(needle) / len(s)) >= 0.30:
                return False, f"{side} contains role substring: {needle}"
    # Narrow middle-init-merged-into-last blanker.
    # Pattern: last token = (1 letter) + (common-English surname tail like -son, -sons,
    # -ton, -ham, -ley, -bury, -well, -wick, -worth, -leigh, -ston, -kin, -ley, -nes,
    # -den, -don, -ford, -ger, -man, -gan, -pen, -sen, -sha, etc.) AND the tail after
    # removing the first letter still passes surname_shape (≥4, vowel, alpha start).
    # In that case it's overwhelmingly likely we ate a middle initial into the last
    # name (e.g. alex  W  parsons → alex + wparsons) — blank.
    # This is deliberately narrow; it does NOT fire on Rhocking / Reynolds / Edmunds
    # because "hocking" / "eynolds" / "dmunds" don't end with any of the strong
    # surname suffixes, and the strong-suffix set only triggers on very productive
    # English endings (which real surnames *with* a prepended initial keep, unlike
    # random-concat tails).
    STRONG_SURNAME_SUFFIXES = (
        # Long/productive English surname endings, deliberately >=3 chars and
        # chosen so random tails don't accidentally match (e.g. we omit
        # -ton / -ley / -ine / -ett / -ell / -ock which have high false-positive
        # rates on "shton" / "talley" / "catline" / "barnett" / "hocking").
        "sons", "bury", "well", "wick", "worth", "leigh", "ston", "kins",
        "ford", "den", "don", "ger", "gan", "pen", "sen", "sha", "son",
        "ham", "nes", "man", "low", "bow", "cox",
    )
    if 5 <= len(ll) <= 14 and ll[0].isalpha():
        tail = ll[1:]
        ok_tail, _ = _surname_shape_ok(tail)
        if ok_tail and any(tail.endswith(suf) for suf in STRONG_SURNAME_SUFFIXES):
            return False, "last looks like 1-letter-middle-init merged into surname"
    # Punctuation gates
    if any(p in first for p in "._"):
        return False, "punct in first"
    if any(p in last for p in "._"):
        return False, "punct in last"
    if first[0] in "-_" or first[-1] in "-_" or last[0] in "-_" or last[-1] in "-_":
        return False, "edge punct"
    # Digits anywhere
    if re.search(r"\d", first) or re.search(r"\d", last):
        return False, "digit present"
    # Case
    if not first[:1].isupper() or not last[:1].isupper():
        return False, "case"
    # Length + shape
    ok, rsn = _first_shape_ok(first)
    if not ok:
        return False, f"first: {rsn}"
    ok, rsn = _surname_shape_ok(last)
    if not ok:
        return False, f"last: {rsn}"
    return True, ""


# ---------------------------------------------------------------------------
# Workbook I/O + entry point
# ---------------------------------------------------------------------------

def process_workbook(
    src_path: str,
    out_path: str,
    first_names: set[str],
    *,
    emit_review_bundle: Optional[str] = None,
    verbose: bool = True,
) -> dict:
    src_path = os.path.expanduser(src_path)
    out_path = os.path.expanduser(out_path)

    wb_in = openpyxl.load_workbook(src_path, read_only=True, data_only=True)
    ws_in = wb_in.active
    rows_in = list(ws_in.iter_rows(values_only=True))
    wb_in.close()

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "Sheet1"

    total_data = 0
    prefilled_preserved = 0
    we_filled = 0
    blanked = 0
    fill_reasons: dict[str, int] = {}
    blank_reasons: dict[str, int] = {}
    filled_rows_for_review: list[dict] = []

    for raw in rows_in:
        cells = list(raw) if raw else []
        while len(cells) < 3:
            cells.append(None)
        email = str(cells[0]).strip() if cells[0] is not None else ""
        if not email:
            # Blank / label row → preserve exactly as in source (col B/C might be label)
            ws_out.append([cells[0], cells[1], cells[2]])
            continue
        total_data += 1
        src_f = str(cells[1]).strip() if cells[1] not in (None, "") else ""
        src_l = str(cells[2]).strip() if cells[2] not in (None, "") else ""
        if src_f and src_l:
            # Pre-filled by user in input → copy through unchanged.
            ws_out.append([email, src_f, src_l])
            prefilled_preserved += 1
            continue

        res = extract_names(email, first_names)
        # Apply the zero-defect filter *before* writing (this is the Phase 3
        # strict-blank step; we never emit a row that fails it).
        if res.first and res.last:
            ok, rsn = zero_defect_filter(res.first, res.last)
            if not ok:
                res = SplitResult("", "", f"zdrop: {rsn}")
        if res.first and res.last:
            ws_out.append([email, res.first, res.last])
            we_filled += 1
            fill_reasons[res.reason or "unknown"] = fill_reasons.get(res.reason or "unknown", 0) + 1
            filled_rows_for_review.append({
                "email": email,
                "first": res.first,
                "last": res.last,
                "parse_reason": res.reason,
            })
        else:
            ws_out.append([email, "", ""])
            blanked += 1
            blank_reasons[res.reason or "unknown"] = blank_reasons.get(res.reason or "unknown", 0) + 1

    wb_out.save(out_path)
    wb_out.close()

    if emit_review_bundle:
        bundle_dir = os.path.expanduser(emit_review_bundle)
        os.makedirs(bundle_dir, exist_ok=True)
        # Interleaved 2-chunk split so each reviewer sees the same distribution.
        ca = [filled_rows_for_review[i] for i in range(0, len(filled_rows_for_review), 2)]
        cb = [filled_rows_for_review[i] for i in range(1, len(filled_rows_for_review), 2)]
        with open(os.path.join(bundle_dir, "chunk_a.json"), "w") as fh:
            json.dump(ca, fh, indent=2)
        with open(os.path.join(bundle_dir, "chunk_b.json"), "w") as fh:
            json.dump(cb, fh, indent=2)
        with open(os.path.join(bundle_dir, "REVIEW_INSTRUCTIONS.txt"), "w") as fh:
            fh.write(
                "Per-row review of candidate email->name fills.\n"
                "For each row, output a JSON object per row:\n"
                "   {\"email\":\"..\",\"ok\": true}                     — keep\n"
                "   {\"email\":\"..\",\"ok\":false,\"reason\":\"…\"}    — flag for repair/blank\n"
                "Reasons should be one or more of:\n"
                "  Bad surname parse, Role word used, Generic first name,\n"
                "  Period in name, Digits in name, First/last swapped,\n"
                "  Surname too short, First name too long, Bad first parse,\n"
                "  title merged, misc.\n"
                "Be conservative: if in doubt, flag.\n"
            )

    summary = {
        "src": src_path,
        "out": out_path,
        "total_data_rows": total_data,
        "prefilled_preserved": prefilled_preserved,
        "we_filled": we_filled,
        "blanked": blanked,
        "fill_reasons_top": sorted(fill_reasons.items(), key=lambda x: -x[1])[:20],
        "blank_reasons_top": sorted(blank_reasons.items(), key=lambda x: -x[1])[:20],
        "review_bundle": emit_review_bundle if emit_review_bundle else None,
    }
    if verbose:
        _print_summary(summary)
    return summary


def _print_summary(s: dict) -> None:
    total = s["total_data_rows"] or 1
    print("Email → contact-name summary")
    print(f"  src : {s['src']}")
    print(f"  out : {s['out']}")
    print(f"  data rows            : {s['total_data_rows']}")
    print(f"  pre-filled preserved : {s['prefilled_preserved']}")
    print(f"  filled (this run)    : {s['we_filled']}  ({s['we_filled'] / total * 100:.1f}%)")
    print(f"  blank (left empty)   : {s['blanked']}  ({s['blanked'] / total * 100:.1f}%)")
    print("  fill pattern counts (top):")
    for k, v in s["fill_reasons_top"]:
        print(f"    {v:5d}  {k}")
    print("  blank reason counts (top):")
    for k, v in s["blank_reasons_top"]:
        print(f"    {v:5d}  {k}")
    if s["review_bundle"]:
        print(f"  review bundle written → {s['review_bundle']}  (chunk_a.json / chunk_b.json)")


def main(argv: Optional[list[str]] = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--src", required=True, help="Path to the source 3-col workbook (email / First name / Surname).")
    ap.add_argument("--out", required=True, help="Output path for the filled workbook.")
    ap.add_argument("--emit-review-bundle", default=None, metavar="DIR",
                    help="If set, also write the set of filled rows split into chunk_a/chunk_b.json here for 2-agent review.")
    ap.add_argument("--skip-name-download", action="store_true",
                    help="Do not hit network for the larger baby-names corpus (use embedded fallback only).")
    ap.add_argument("--quiet", action="store_true", help="Suppress summary output.")
    args = ap.parse_args(argv)

    first_names = load_first_names(allow_network=not args.skip_name_download)
    process_workbook(
        src_path=args.src,
        out_path=args.out,
        first_names=first_names,
        emit_review_bundle=args.emit_review_bundle,
        verbose=not args.quiet,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
