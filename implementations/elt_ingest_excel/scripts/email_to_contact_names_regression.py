"""Spot-check the canonical regression inputs for email_to_contact_names.py.

Runs against the ORIGINAL 2026-09 source file (4478 rows) OR against any supplied
PARSED workbook. Exits 0 if every spot case matches the legend in docs/email-contact-name-fill.md;
exits 1 with a diff otherwise. Used as a regression guard when the parse rules change.
"""

from __future__ import annotations

import os
import sys

import openpyxl

SRC_DEFAULT = os.path.expanduser("~/Downloads/Customer Contact Emails and Names - PARSED.xlsx")

LEGEND: list[tuple[str, str | None, str | None, str]] = [
    # (email_lowercase, expected_first_or_None, expected_last_or_None, reason_tag)
    # Strict expect-filled:
    ("lee.ashmore1@hotmail.co.uk",        "Lee",    "Ashmore",  "preserve prefilled"),
    ("paulrhocking@icloud.com",           "Paul",   "Rhocking", "single H1"),
    ("paulashton@a-aconcreterepairs.co.uk","Paul",  "Ashton",   "single tie-break H1"),
    ("adrianreynolds@pinnaclefp.co.uk",   "Adrian", "Reynolds", "single H1"),
    ("jamieedmunds@hotmail.com",          "Jamie",  "Edmunds",  "single H1"),
    ("mrdannyburke@hotmail.co.uk",        "Danny",  "Burke",    "honorific strip + 2-token"),
    ("mrdavestyles@yahoo.co.uk",          "Dave",   "Styles",   "honorific strip + 2-token"),
    ("chrisdharris@gmail.com",            "Chris",  "Dharris",  "middle-init H1"),
    ("jack.7.bray@gmail.com",             "Jack",   "Bray",     "sep-3 middle-init"),
    # Strict expect-BLANK:
    ("abi.invoicing@towergate.co.uk",     None, None, "blocklist R1"),
    ("accounts.payables@meccan.co.uk",    None, None, "blocklist R1"),
    ("purchase.ledger@mottmac.com",       None, None, "blocklist R1"),
    ("bpo-cbl.invoicereceipt@molsoncoors.com", None, None, "blocklist R1"),
    ("a.shoesmith@ntlworld.com",          None, None, "R5 1-char first"),
    ("alan.denis.smith1947@sky.com",      None, None, "R3 4-token + year bleed"),
    ("charles.macinnes.25@gmail.com",     None, None, "R3 4-token + digits"),
    ("rolandtaylor25@gmail.com",          None, None, "concat ambig"),
    ("wright.ian24@ntlworld.com",         None, None, "reversed-dot ambig"),
    ("docker.andrew@live.co.uk",          None, None, "reversed-dot ambig"),
    ("martin@1stbusstop.co.uk",           None, None, "single first-name only → no surname"),
    ("purchasing@hospitalityfinder.co.uk",None, None, "role mailbox, no local split"),
    # Borderline 2026-09 cases — expect BLANK after zero-defect gate:
    ("andrew7cartwright@yahoo.co.uk",     None, None, "digit-in-middle → blank"),
    ("alexwparsons67@gmail.com",          None, None, "concat + digit suffix → blank"),
]


def main(argv: list[str]) -> int:
    target = argv[0] if argv else SRC_DEFAULT
    target = os.path.expanduser(target)
    if not os.path.exists(target):
        print(f"REGRESSION FAIL: target file not found → {target}", file=sys.stderr)
        return 2
    wb = openpyxl.load_workbook(target, read_only=True, data_only=True)
    ws = wb.active
    by_email: dict[str, tuple[str, str]] = {}
    for r in ws.iter_rows(values_only=True):
        c = list(r) if r else []
        while len(c) < 3: c.append(None)
        a = str(c[0]).strip().lower() if c[0] is not None else ""
        if not a: continue
        f = str(c[1]).strip() if c[1] not in (None, "") else ""
        l = str(c[2]).strip() if c[2] not in (None, "") else ""
        by_email[a] = (f, l)
    wb.close()

    ok, fail = 0, 0
    failures: list[str] = []
    for em, ef, el, tag in LEGEND:
        k = em.lower()
        got_f, got_l = by_email.get(k, ("", ""))
        exp_f = "" if ef is None else ef
        exp_l = "" if el is None else el
        if got_f == exp_f and got_l == exp_l:
            ok += 1
        else:
            fail += 1
            failures.append(
                f"  [{tag:30s}] {em}\n"
                f"      expected : ({exp_f!r}, {exp_l!r})\n"
                f"      got      : ({got_f!r}, {got_l!r})"
            )

    print(f"Email→name regression check against: {target}")
    print(f"  cases : {len(LEGEND)}")
    print(f"  pass  : {ok}")
    print(f"  fail  : {fail}")
    if failures:
        print("\nFAILED cases:\n" + "\n".join(failures), file=sys.stderr)
        return 1
    print("  ✅ all spot cases match the legend.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
