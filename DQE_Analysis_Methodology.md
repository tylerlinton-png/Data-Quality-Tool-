# DQE — Duetto vs. PMS Data Quality Analysis Methodology

This document describes the data model, parsing rules, and analysis logic used by
Duetto's internal **Data Quality Engine (DQE)**. It is written so that the same
analysis can be reproduced manually (e.g. by an LLM) given the raw source files,
without running the original Flask application.

Attach this document along with the raw source files (DVA `.xlsx`, Duetto Bookings
`.tsv`/`.csv`, Duetto Blocks `.tsv`/`.csv`, Duetto Folio `.tsv`/`.csv`, and/or an
Opera Arrival Details Report in `.xml`, `.pdf`, or `.txt`) to run the same analysis.

---

## 1. Purpose

Duetto ingests reservation data from a hotel's Property Management System (PMS) via
an integration (e.g. Oracle OHIP/Opera). The **Daily Variance Analysis (DVA)** file
is a report that compares, for every stay date, the number of rooms and amount of
revenue Duetto has "committed" against what the PMS itself reports. When these
values diverge, the DQE identifies *why* using additional PMS/Duetto exports as
supporting evidence, and classifies the root cause using a fixed taxonomy of
diagnostic codes.

---

## 2. Source Files & Schemas

### 2.1 DVA File (required) — `.xlsx`

Produced by Duetto's PMS integration layer. Structure:

- Row 1: generation timestamp (ignore)
- Row 2: section super-headers (`TOTAL`, `NON_GROUP`, `GROUP` — merged cells)
- Row 3: column headers (repeated once per section; **only use the first / TOTAL
  occurrence of each column name**)
- Row 4+: one row per stay date

Required columns (first occurrence only):

| Column | Meaning |
|---|---|
| `Hotel` | Hotel name |
| `Stay Day` | Stay date. **Important**: this cell often contains an Excel formula string like `=DATE(2025,7,1)` rather than an evaluated date — you must parse the formula literally (regex `=DATE\((\d+),(\d+),(\d+)\)`) or read the raw cell value with a library that doesn't auto-evaluate it (e.g. `openpyxl`). Do not trust `pandas.read_excel` to give you real dates here. |
| `Duetto Commit` | Room count Duetto believes is booked for the date |
| `PMS Commit` | Room count the PMS itself reports for the date |
| `Duetto Revenue` | Duetto's committed room revenue |
| `PMS Revenue` | PMS-reported room revenue |
| `Commit Status` (optional) | Precomputed `PASS`/`FAIL` flag for rooms, already thresholded by Duetto |
| `Revenue Status` (optional) | Precomputed `PASS`/`FAIL` flag for revenue |

Derived fields:

- `RoomDiff = Duetto Commit − PMS Commit`
- `RevenueDiff = Duetto Revenue − PMS Revenue`
- `Period = "HISTORIC"` if `Stay Day < today` else `"FUTURE"`

Numeric cleaning rule (`clean_revenue`): strip currency symbols (`€$£`), commas,
whitespace; treat parenthesized numbers `(123.45)` as negative; non-numeric → `0.0`.

### 2.2 Duetto Bookings / Blocks Stay Date Report (optional) — `.tsv` / `.csv`

Flat, one-row-per-booking export. Tab or comma delimited (auto-detect by counting
tabs vs. commas in the header line). Column names are upper-cased on load.

Key columns used:

| Column | Meaning |
|---|---|
| `STAY_DATE` | Parsed to a date |
| `NUM_ROOMS` | Numeric, defaults to 0 if missing/invalid |
| `RATE` | Cleaned via the same currency-cleaning rule as above |
| `RESERVATION_STATUS` | e.g. `RESERVED`, `CHECKED_IN`, `CHECKED_OUT`, `CANCELLED`, `NO_SHOW` |
| `BOOKING_ID` | Used only for display in evidence lists |
| `ROOM_TYPE` / `ROOM_TYPE_CODE` | Used only for display |
| `SHARE_ID` | Non-empty indicates a shared reservation |
| `ALTERNATE_SOURCE_ID` | **This is the Opera/PMS confirmation number**, used to cross-reference against the Arrival Details Report (see §4). Do not assume a column literally named "confirmation number" exists — Duetto exports store it under `ALTERNATE_SOURCE_ID`. |

Blocks file uses the identical schema/parsing as Bookings.

### 2.3 Duetto Flat Folio Report (optional) — `.tsv` / `.csv`

Same delimiter auto-detection. Columns used:

| Column | Meaning |
|---|---|
| `STAY_DATE` | Parsed to a date |
| `REVENUE_USD` | Numeric |
| `RATE_AMOUNT` | Numeric |
| `CATEGORY_TYPE` | e.g. `ACTUAL_ROOM` (mapped room revenue) or `NONE` (unmapped transaction codes — the most useful for root-causing revenue diffs) |
| `REVENUE_TYPE` | The folio transaction code (e.g. resort fee, no-show fee) |

### 2.4 Opera Arrival Details Report (optional) — `.xml`, `.pdf`, or `.txt`

Dispatch by file extension. All three formats are normalized to the same schema
(one row **per stay-night**, i.e. a 3-night reservation produces 3 rows, one per
date between arrival (inclusive) and departure (exclusive)):

| Field | Meaning |
|---|---|
| `CONFIRMATION_NO` | Opera confirmation number (the join key back to Duetto's `ALTERNATE_SOURCE_ID`) |
| `STATUS` | Opera reservation status, e.g. `CKIN`, `CKOT`, `DUE_IN`, `RESV`, `NO_SHOW`, `CANCEL` |
| `STAY_DATE` | One row per night of the stay |
| `ARRIVAL` / `DEPARTURE` | Full stay range |
| `NO_OF_ROOMS` | Usually 1 |
| `RATE_AMOUNT` | Effective nightly rate |
| `IS_SHARED` | Y/N |
| `BLOCK_CODE` | Group block code, if any |
| `ROOM_CATEGORY` | Room type label |

**Format-specific parsing notes:**

- **XML**: iterate `G_RESERVATION` elements; dates in `DD/MM/YY` format.
- **TXT/TSV**: tab-delimited export with a repeating group-summary row acting as a
  record separator (regex: `\t{3}126\t71\t10\t Total\t71\t`). The address field
  contains embedded raw newlines that must be stripped before splitting on tabs.
  Dates are in `DD-MM-YY` format. Confirmation numbers are validated as 9–10 digit
  strings.
- **PDF**: requires layout-aware extraction (not flat-text regex, since some fields
  wrap across lines). Group text elements by Y-coordinate (±2pt tolerance) to
  reconstruct rows; the "primary" row (dates/status/rate) is identified by finding
  a known status token (`CKOT`, `CKIN`, `DUE_IN`, `RESV`, `NO_SHOW`, `CANCEL`) in
  the expected X-range; the confirmation number sits in a "secondary" row 8–25pt
  below the primary row. Dates are `DD-MM-YY`.

---

## 3. Scoping the Analysis

If any of Bookings / Blocks / Folio are supplied, restrict the DVA rows to the
**union of stay dates** present across those files. This lets an analyst upload a
narrow date-range export (e.g. one stay date) and only analyze that scope, rather
than the DVA's full date range.

---

## 4. Room & Revenue Accuracy

For each in-scope stay date:

- **Room failure**: prefer the DVA's own `Commit Status == 'FAIL'` column if
  present; otherwise fall back to `|RoomDiff| > 0` (zero-tolerance).
- **Revenue failure**: prefer `Revenue Status == 'FAIL'`; otherwise fall back to
  `|RevenueDiff| > 10` (a $10 tolerance band).

Headline metrics:

```
Room Accuracy    = (1 − failing_room_days    / total_days) × 100
Revenue Accuracy = (1 − failing_revenue_days / total_days) × 100
```

---

## 5. Root-Cause Classification

Every failing date gets a **diagnostic code** with a plain-English explanation and
a list of "contributing bookings" (evidence). The classifier is a deterministic
decision tree, evaluated in this priority order.

### 5.1 Room discrepancies (`classify_room`)

If `Duetto Commit > PMS Commit` (Duetto overstates):

1. **Historic** date + bookings still `RESERVED` (should be `CHECKED_OUT` by now)
   → `HO-O-10`
2. Cancelled bookings (status contains `CANCEL`) still carrying `NUM_ROOMS > 0`
   → `HO-O-11` (historic) / `FO-O-46` (future)
3. Non-empty `SHARE_ID` reservations found (possible orphaned share)
   → `HO-O-13` (historic) / `FO-O-47` (future)
4. Fallback: generic "possible cancelled/deleted booking not synced"
   → `HO-O-11` / `FO-O-46`

If `Duetto Commit < PMS Commit` (Duetto understates):

- Generic "bookings present in PMS not sent to Duetto — check integration
  publisher settings" → `HO-U-06` (historic) / `FO-U-40` (future)

### 5.2 Revenue discrepancies (`classify_revenue`)

If folio data is available for the date, it takes priority as direct evidence:

1. Sum folio rows where `CATEGORY_TYPE == 'ACTUAL_ROOM'` → compare to Duetto
   revenue (sanity check).
2. Group all `CATEGORY_TYPE == 'NONE'` (unmapped) transaction rows by
   `REVENUE_TYPE`, sum amounts.
3. Look for a **single unmapped code** whose absolute amount is within $1 of
   `|RevenueDiff|` → exact match, cite that code.
4. If none, look for a **pair of unmapped codes** whose sum is within $1 of
   `|RevenueDiff|` → combo match.
5. If still no match, list up to 8 unmapped codes for manual review.
   Codes used: `HR-U-21` / `HR-O-31` (historic) or `FR-U-51` / `FR-O-56` (future),
   selected by the sign of the diff.

If no folio data / no match, fall back to booking-level heuristics:

- **Understates** (`RevenueDiff < 0`):
  - `NO_SHOW` bookings with `RATE == 0` → `HR-U-14` (historic) / `FR-U-51` (future)
    — likely a folio-level no-show fee the PMS captured but Duetto's
    reservation-level integration didn't.
  - Any active-status booking (`RESERVED`, `CHECKED_IN`, `CHECKED_OUT`, `NO_SHOW`)
    with `RATE == 0` → `HR-U-21` / `FR-U-51`.
  - Fallback: generic rate-mismatch note.
- **Overstates** (`RevenueDiff > 0`):
  - `Duetto Commit rooms == 0` AND `PMS Revenue < 0` → PMS posted a negative folio
    adjustment (refund/cancellation fee) that a reservation-level integration
    can't see → `HR-O-24` / `FR-O-53`.
  - Cancelled bookings still carrying a non-zero rate → `HR-O-31` / `FR-O-56`.
  - Fallback: generic gross-vs-net rate note.

---

## 6. Arrival Details Cross-Reference & PMS Sync-Gap Detection

This is the most important recent addition — **do not skip it if an Arrival
Details Report is supplied.**

### 6.1 The join

The Opera Arrival Details Report's `CONFIRMATION_NO` is matched against Duetto's
`ALTERNATE_SOURCE_ID` column in the Bookings Stay Date Report. Only "active"
Duetto bookings are considered a match candidate — statuses in
`{CHECKED_IN, CHECKED_OUT, CKOT, CKIN, DUE_IN, RESV}` (i.e. exclude
`CANCELLED`/`DAY_CANCELLED`).

**Caveat (learned the hard way):** the Duetto Bookings export is scoped by
*stay date*, so it includes every reservation touching that date (e.g. a guest
staying June 15–19 shows up in the June 17 file). The Opera Arrival Details
Report is scoped by *arrival date* — a guest who arrived June 15 will not appear
in a June 17 arrivals export even though they are physically in-house. **Do not
treat "in Duetto but not in Opera Arrivals" as evidence of a data problem** — it
is usually just this scoping mismatch. For that reason, the DQE tracks
`missing_in_duetto` / `extra_in_duetto` counts but does **not** present a detailed
table of "extra in Duetto" reservations as an actionable discrepancy.

### 6.2 What IS meaningful: the three-way room-count comparison

For the stay date(s) covered by the arrivals report, compute three independent
room counts and compare them:

| Source | What it represents |
|---|---|
| **Opera Arrivals total** | Count of distinct `CONFIRMATION_NO` in the Opera export for the date — ground truth from the PMS itself (manual export). |
| **Duetto Bookings total** | Count of distinct active `ALTERNATE_SOURCE_ID` in the Duetto bookings export for the date. |
| **DVA PMS Commit total** | The `PMS Commit` figure from the DVA file for the same date — this number is fed by the automated OHIP sync integration, a *different pipeline* than the manual Opera export. |

**Key insight**: If Opera Arrivals and Duetto Bookings **agree** with each other
but **differ** from the DVA PMS Commit figure, this proves the discrepancy is
**not** a Duetto data-accuracy problem — Duetto's own book of business matches
the PMS's own arrivals list exactly. The gap instead means the **OHIP sync feed
that populates the DVA is incomplete or lagging**, i.e. an integration/sync
problem upstream of Duetto, not a Duetto integrity issue.

When this pattern is detected (`opera_total == duetto_total != dva_pms_total`),
re-classify that date's room discrepancy from a generic "Duetto overstates/
understates" verdict to a **`SYNC-GAP`** finding, with an explanation like:

> "PMS sync gap confirmed by Opera Arrivals: Opera=71, Duetto=71, DVA PMS=42.
> 29 room(s) missing from OHIP sync feed — Duetto and Opera agree."

### 6.3 Rate mismatches

For confirmation numbers that **do** match between Opera and Duetto, compare
`RATE_AMOUNT` (Opera) vs. the Duetto rate column. Flag a mismatch if the absolute
difference exceeds **$1.00**. Report confirmation number, room type, both rates,
and the signed difference.

---

## 7. Recommendations

Recommendations are generated by counting the frequency of each diagnostic code
across all discrepancies (most frequent first) and mapping each code to a fixed
remediation text (see table below). This is a static lookup, not something that
needs re-deriving analytically — reproduce it verbatim if the codes are cited.

| Code | Recommendation |
|---|---|
| `HO-O-10` | Request a historical reservation resync to clear stale RESERVED-status bookings on actualized dates. |
| `HO-O-11` | Request a historical resync including cancellations to remove phantom active bookings. |
| `HO-O-13` | Investigate orphaned share reservations — request a historical resync including profile data. |
| `HO-U-05` | Review leg-perm settings in the hotel back end; escalate to integration manager with booking XML logs. |
| `HO-U-06` | Investigate missing bookings — verify integration publisher settings and request a historical resync. |
| `FO-O-45` | Manually cancel phantom bookings in Duetto; verify integration receives delete/modify messages. |
| `FO-O-46` | Request a reservation resync including cancellations for affected future dates. |
| `FO-O-47` | Investigate orphaned share reservations on future dates; request a future reservation resync. |
| `FO-U-40` | Review integration setup and request a future reservation resync. |
| `HR-U-14` | No-show fee revenue is posted as a folio transaction in the PMS and not captured by Duetto's reservation-level integration. Inform the client of this limitation, or evaluate switching to a folio-level integration. |
| `HR-O-24` | PMS shows negative folio adjustments (refunds/cancellation charges) not reflected in Duetto. Inform the client that Duetto reads reservation-level revenue only — folio-level credits will not be subtracted. |
| `HR-U-21` | Review integration XML logs for zero-rate or incomplete rate messages; escalate to the Integration Partner Manager. |
| `HR-O-31` | Review integration XML logs for gross-rate messages; if amounts include tax or packages, escalate to the Integration Partner Manager. |
| `FR-U-51` | Check integration logs for zero-rate messages on future bookings; escalate to the Integration Partner Manager. |
| `FR-O-56` | Review integration logs for gross-rate messages on future bookings; escalate to the Integration Partner Manager. |
| `FR-O-53` | PMS folio adjustments are not forwarded to Duetto. Evaluate folio-level integration if full revenue fidelity is required. |
| `SYNC-GAP` | Opera Arrivals and Duetto agree with each other; escalate the gap to the OHIP integration/sync team rather than treating it as a Duetto data issue. |

---

## 8. Expected Output Structure

When asked to perform this analysis, structure the response as:

1. **Headline metrics** — Room Accuracy %, Revenue Accuracy %, date range, total
   days in scope.
2. **Discrepancy table** — one row per failing stay date, columns: Stay Date,
   Period (Historic/Future), Duetto Rooms, PMS Rooms, Diff, Room Code,
   Explanation, Contributing Bookings, and the equivalent revenue columns.
3. **Stay Date Reconciliation** (only if an Arrival Details Report was supplied)
   — the three-way comparison table (Opera / Duetto / DVA PMS) per date, called
   out explicitly as either "confirmed match" or "SYNC-GAP", plus any rate
   mismatches.
4. **Recommendations** — deduplicated, ordered by frequency of the underlying
   diagnostic code, using the fixed text above.

---

## 9. Known Limitations to Preserve

- Zero-tolerance on rooms (`|RoomDiff| > 0` is a fail) but a $10 tolerance on
  revenue — don't "fix" this asymmetry, it matches Duetto's actual DVA
  thresholds.
- The Opera Arrivals vs. Duetto Bookings join is confirmation-number-based and
  will show false "missing" bookings for multi-night stays if the two reports
  are scoped to different date windows — always check date-scope alignment before
  treating a mismatch as real.
- Rate mismatch threshold is a flat $1.00, not a percentage.
