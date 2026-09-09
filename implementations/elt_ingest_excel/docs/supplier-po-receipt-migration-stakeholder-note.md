# Supplier POs and receipts: Proactis → Workday

*A plain-English summary of what we built, the problems we hit with the Proactis data, and how we got the numbers to line up with finance.*

## What we set out to do

FA is moving off Proactis, and the open supplier commitments need to land in Workday cleanly. Two things have to come across at the switch:

- **Open purchase orders** — money FA has committed to suppliers but hasn't fully settled yet.
- **Receipts** — goods and services already received but not yet invoiced. In accounting terms that's a liability sitting on the books (it's usually called GRNI: *goods received, not invoiced*), so Workday has to know about it from day one.

Anything already invoiced and paid is closed history and doesn't move. We're only carrying what's still live.

We deliver this as two Workday-ready spreadsheets (the "EIBs" Workday loads from):

| Workbook | What it holds | Size today |
|---|---|---|
| **Purchase Order EIB** | Every open PO line with its full committed value | 10,833 lines across 6,971 POs |
| **Receipt EIB** | The received-not-invoiced portion, as receipts | 7,122 lines across 4,659 receipts |

## A quick word on how a PO is structured

A PO isn't one number. It's one agreement with a supplier, broken down:

- The **order** is the whole agreement (one PO number).
- Inside it are **lines** — the separate things being bought (a service on the 28th, an event on the 31st, and so on).
- Inside a line there can be **positions** — the same thing, but with its cost split across different budgets or cost centres for accounting.

Workday needs each of these at the lowest level, so every line-and-position becomes its own row. That's why one PO can produce several rows in the workbook.

Each line carries two kinds of value:

- **Soft** — ordered but not yet received. A promise; nothing has actually arrived.
- **Hard** — received but not yet invoiced. This is the real liability, and it's the only part that becomes a receipt.

The PO workbook carries the **full value (soft + hard)**. The receipt workbook carries the **hard part only**.

## The problem with the Proactis receipts data

The original plan was to build the receipts straight from Proactis's own receipts extract. That turned out not to be usable:

- **The "quantity received" column wasn't a value.** On more than half the rows it held a unit count, not an amount. On one PO alone, using it would have loaded a receipt of £280 against a real received value of £55,468. There was no reliable way to turn those quantities into money.
- **It contradicted the PO data.** What the receipts extract said had been invoiced didn't agree with what the commitment report said. The two couldn't be reconciled against each other.

So the receipts extract was a dead end. We couldn't get it to a state anyone could trust or reconcile — which led to the decision further down to build the receipts ourselves.

## The variance against the finance review data

When we reconciled our output against finance's review file, a small number of POs didn't match. We chased every one of them down, and they all traced back to the same thing: **the Proactis export itself was corrupted on those rows.**

The cause was stray line breaks buried inside the description text. When Proactis exported the file, those line breaks pushed the data out of alignment — a PO's values ended up sitting in the wrong columns, and in the worst cases an entire PO got swallowed into the description text of the PO above it. So the mismatch was never in our process; the source file had already scrambled those rows before we touched them.

A concrete example: PO `NFCPO300006489` came through carrying a value of £6,747 when finance's figure was £1,242 — because the descriptions of two *other* POs (`...6502` and `...6503`) had been absorbed into its description field, dragging their values with them. Those two POs then went missing entirely from the export.

In total this affected a handful of POs, worth low tens of thousands of pounds against a £97m book. Finance reviewed each one, agreed the differences were immaterial, and signed off on living with them rather than forcing Proactis to produce a clean re-export. We kept the text-cleaning in the build regardless, so the corruption can't break the Workday load even where finance has accepted the value difference.

## The decision: build the receipts ourselves

Working with finance, we agreed to stop using the Proactis receipts extract and instead **reconstruct the receipts from the one number we could trust** — the hard commitment value already sitting on each PO line in the commitment report.

In practice that means: for every PO line that has a received-not-invoiced value, we create a receipt line for exactly that amount, and bundle all of a PO's receipt lines into a single receipt document (numbered `REC-` plus the PO number).

The big advantage is that the receipts can't disagree with the POs, because they're drawn from the same source. A receipt can never be larger than the PO it came from, and every receipt points back to a real PO line. It's consistent by design, not by luck — and if the file were ever lost, we can rebuild it identically from the PO data.

The honest trade-off: these receipts are accurate on **value**, but they aren't a record of actual delivery events. There are no real receipt dates or quantities underneath them — we've collapsed all of that into one figure per line. That matters for how invoices match later (more on that at the end), but it gives finance a receipts position they can actually stand behind.

## Getting the numbers to reconcile with finance

We check our output in pounds against finance's own commitment report (the version dated 04/08 that they treat as the source of truth). Where things stand:

- **Purchase orders:** finance's total is **£97.87m**, ours is **£97.85m**. On a PO-by-PO basis, effectively **100% match** — the only non-matches are the corrupted-export POs described above, which finance has already signed off.
- **Receipts:** the receipt workbook totals **£26.49m**, and it ties back to the hard commitment on the POs to the penny.
- **The rest of the tiny residual** comes from the two Proactis exports being pulled on slightly different dates — a timing gap between source files, not an error in the build.

That last point is the important one for confidence: nothing in the way we transform the data introduces a discrepancy. Every difference is either a source-file corruption finance has accepted, or an explainable timing gap between two exports.

## The fixes we had to build in along the way

A number of practical issues came up during the build. Each was dealt with rather than worked around:

- **Filtering rules, agreed with finance.** We drop: POs with nothing left committed; old "ordered but never received" lines past an agreed cutoff; lines marked fully delivered and closed; and 7 specific POs finance flagged as adjustments. This is what takes the raw ~14,300 rows down to the ~10,800 that actually load.
- **Fully-paid POs drop out on their own.** Once a PO is fully invoiced, its outstanding value is already zero in the source, so our "nothing committed" filter removes it automatically — we don't need a separate rule for it.
- **Cleaning the description text.** Stray control characters and line breaks from the Proactis export were breaking the load. We strip those out while leaving legitimate text (quotes, dashes, symbols that appear in real descriptions) untouched.
- **Currency handling.** Japanese Yen has no decimal places, so those amounts load as whole numbers; everything else loads to two decimals. Both workbooks handle this the same way.
- **Supplier matching.** Some suppliers didn't map cleanly to a Workday supplier record. Rather than guess, we let those through flagged, so finance can confirm the right supplier before go-live. About 68 lines (~£42k) sit in this review bucket.
- **Dates: order date vs commitment date.** Two related issues surfaced here, because the source carries two dates — the PO's *order date* and a later *commitment date* — and on some rows the commitment date actually falls *before* the order date, which shouldn't happen.
  - **Due date column cleared on the PO load.** The due date field had been filled with the commitment date. Since some commitment dates sat before the order date (and a PO/receipt at this stage has no real payment due date anyway — that only arrives with the invoice), the consultant cleared this column out. Due dates will come in later with the actual invoices.
  - **Receipt date corrected.** The receipts were first dated on the commitment date, but Workday rejected 92 of them because that date fell before the PO's order date — you can't receipt goods before the PO exists. We changed the rule so each receipt is dated on the commitment date *or* the order date, whichever is later. That guarantees no receipt ever predates its PO, and all 92 now load.
  - **Worth noting for procurement:** about 1.7% of lines (185 of 10,833) carry a commitment date slightly earlier than the order date — small gaps, mostly within a month, and a chunk land on the 1st of the month. This most likely reflects retrospective POs (spend that was raised after the fact) or period-dated entries, rather than an error. It's handled safely by the receipt-date rule above, but it's a minor source data-quality point that procurement may want to be aware of.

## Where we are, and the one thing still to confirm

Both workbooks build cleanly, reconcile to finance, and load into Workday. What's left is a short list of review items (unmapped suppliers, a few status mismatches) that finance needs to eyeball — these are sign-off items, not blockers.

The one open question sits with the Workday consultant: **how invoice matching is configured.** Because our synthetic receipts are value-accurate but don't carry quantities, how cleanly real invoices match against them after go-live depends on whether Workday matches on amount or quantity, and what tolerance is set. It's worth pinning down before go-live, because it's the difference between receipts clearing automatically and needing a bit of manual matching.

---

*The full technical detail behind all of the above lives in the engineering documentation for this pipeline; this note is the summary for a non-technical audience.*
