# ARC — QuickBooks Setup Plan

**Prepared for:** Aldo Rojas & Somit Soni
**Date:** July 7, 2026 · **Updated:** July 15, 2026
**Purpose:** Shared plan for standing up QuickBooks Online so we can turn Xactimate estimates into customer invoices cleanly. Both of us can work off this.

> **What changed in this update (July 15):** We decided **how** to structure the catalog. Instead of mirroring every Xactimate line item into QuickBooks (which would force us to re-sync 383+ items every month when the price list rolls over), we're going with a **summary / category-level structure (Option A)**. Pricing authority stays in Xactimate; QuickBooks handles the invoice, A/R, tax, and books. See "The catalog decision" below.

---

## The big picture — why this matters

Two systems, two jobs:

- **Xactimate** builds the estimate/scope and **owns the pricing** (line items priced off TXHO8X_JUL26).
- **QuickBooks** produces the invoice the customer actually pays, and tracks A/R, payments, sales tax, and our books.

The handoff we want: build the scope in Xactimate → invoice the **category totals** in QuickBooks → attach the Xactimate PDF as the line-item backup. The Xactimate PDF is already the document adjusters read for line detail; QuickBooks is the payment document.

This connects to the master price list (`ARC_Price_List_Master.xlsx`), but note the change below: **we no longer load that file line-by-line into QuickBooks.** It stays a reference/source for Xactimate, not a feed into the QBO catalog.

---

## ⭐ The catalog decision — summary items, not line-item mirroring

**Question we resolved:** Should each Xactimate line item become its own priced Product/Service in QuickBooks?

**Answer: No.** Here's why.

- In QuickBooks, an item's price is only a **default** that pre-fills a line — you can always override it per invoice.
- The price list updates **every month**. If we mirror priced line items, then every month we either (a) re-import all items to refresh prices — a maintenance treadmill, or (b) leave stale prices and override every line by hand — which makes the stored catalog pointless.
- The price authority already lives in **Xactimate**. Mirroring it into QuickBooks creates a second source of truth that's wrong the moment the list rolls over. (Industry integration docs confirm Xactimate line items don't map cleanly to QuickBooks products — they don't sync without being summarized first.)

**What we're doing instead — Option A (summary / category-level):**

A small, stable set of items (~6–8) that map to income accounts by division. We invoice the **totals from the Xactimate estimate** and attach the Xactimate PDF for the line detail. **Prices never live in QuickBooks, so the monthly price-list update never touches our catalog.**

**Options we considered (for the record):**

| Option | What it is | Items to maintain | Monthly price-list impact |
|---|---|---|---|
| **A. Summary / category items** ✅ chosen | ~6–8 items by division; invoice Xactimate totals, attach PDF | ~8, names only (no prices) | **Zero** |
| B. Trade/phase items | ~20–50 items by trade, split labor vs. material; type phase subtotals | ~20–50 names | Zero |
| C. Full line-item mirror | Every Xactimate line as a priced item (the old plan) | 383 → up to 1,267, **with prices** | **Full re-sync every month** |
| D. Integration / middleware | Software pushes Xactimate → QBO (XactRemodel PRO, JobNimbus, DASH, etc.) | Varies | Most still summarize anyway |

If we ever want more internal reporting granularity, we can step up to Option B later — still with **no** monthly price maintenance. For per-job profitability, we use QuickBooks **Projects** (Plus/Advanced), independent of invoice granularity.

---

## Current state of our QuickBooks (audit, July 15)

Audited the live QBO account (RXS Group, LLC):

- **383 service items already loaded** — all "Service" type, in 4 groups: **Water Mitigation (100), Drying (100), Cleaning (100), Demo (82)**, under a parent "Services."
- Items are named by description (e.g., "Water Mitigation – Equipment setup, take down, and monitoring (hourly charge)"), **not** Xactimate SKU codes.
- **Prices are the stale JUN26 set.** The equipment-monitoring line reads **$66.75** — the exact June figure vs. **$78.34** on live JUL26. So the mirrored catalog is already out of date on day one — a concrete example of the treadmill problem.
- Payment terms exist: Due on receipt, Net 15, Net 30, Net 60.
- QuickBooks Payments enrollment status couldn't be read via the connector — needs a quick check in the QBO UI.

**Action:** Retire / deactivate the 383-item catalog (don't feed it monthly) and rebuild around the Option-A summary items below. We keep the old items inactive rather than deleting, so historical data isn't disturbed.

---

## The 5 setup pieces (revised for Option A)

### 1. Products/Services catalog — summary items  ← now a small task
Build ~6–8 summary items, each mapped to an income account and flagged for tax correctly:

| Item (Product/Service) | Type | Tax status | Income account |
|---|---|---|---|
| Water Mitigation | Service | Non-taxable (labor/equipment) | Water Mitigation Income |
| Structural Drying | Service | Non-taxable (labor/equipment) | Drying Income |
| Cleaning | Service | Non-taxable (labor) | Cleaning Income |
| Demolition | Service | Non-taxable (labor) | Demolition Income |
| Reconstruction / Repairs | Service | Non-taxable (labor) | Reconstruction Income |
| Contents | Service | Non-taxable (labor) | Contents Income |
| **Materials** | Service | **Taxable (8.25%)** | Materials Income |

The invoice for a job = one line per division (amount pulled from the Xactimate category total) **+ a taxable "Materials" line** for the material portion. Item names/accounts are stable; nothing here changes month to month. (Final account names/divisions to confirm with our CPA/bookkeeper.)

### 2. Tax setup  ← still the most important to get right
We run **separated contracts**, so:
- **Materials = taxable** (8.25% Houston-area rate)
- **Labor = non-taxable** (TX residential repair/remodel/restore labor is exempt)

**Option A makes this simpler:** we don't tax-flag hundreds of items — we put the **taxable base on one "Materials" line** and keep the division/labor lines non-taxable. What matters is that the **Materials subtotal in QuickBooks matches Xactimate's "Mat Only" 8.25% base.** Get it wrong and we either overcharge tax on labor or under-remit — both bad. **Worth a one-time CPA review.**

### 3. Invoice template
Branded layout, set up once, reused every job:
- ARC logo
- RXS Group LLC / DBA America's Restoration Contractor
- Tax ID: 41-3573136
- Payment terms + remit-to info
- Space/attachment for the Xactimate estimate PDF (our line-item backup)

### 4. QuickBooks Payments
Lets customers pay invoices directly (card / ACH) instead of chasing checks. Ties to our staged-payment model (deposit / progress / final). Confirm enrollment status in the QBO UI, then enable + test.

### 5. Customer records
Each job's customer (Gonzalez, Patel, Greer, etc.) as a QBO customer so invoices, payments, and history roll up per client. Use **Projects** under each customer for per-job profitability.

---

## Xactimate side — still worth standardizing

Even though we no longer load prices into QuickBooks, we still want Xactimate itself consistent so the totals we invoice are right:

- [ ] Confirm everyone estimates on **JUL26** (Aldo confirmed default is TXHO8X_JUL26). **Mike's estimates (Galindo, Guillote) showed JUN26** — standardize Mike on JUL26.
- [ ] Set the **Tax jurisdiction** default in Xactimate preferences (currently blank) to the 8.25% **Mat Only** jurisdiction so it applies automatically on every new estimate. This is what makes the Xactimate "Materials" subtotal we invoice against reliable.

Note: re-exporting the master price list is **no longer a blocker** for QuickBooks, since we don't load prices into QBO. It's only relevant to keeping Xactimate itself current.

---

## Task checklist (divide between us)

**Xactimate standardization**
- [ ] Confirm Mike is on JUL26 too
- [ ] Set Xactimate default tax jurisdiction to 8.25% Mat Only

**QuickBooks build (Option A)**
- [ ] Deactivate the 383-item legacy catalog (keep inactive, don't delete)
- [ ] Create ~6–8 summary items + income accounts (table above)
- [ ] Set tax status: Materials taxable / all labor & division lines non-taxable
- [ ] Build/finish invoice template (logo, DBA, Tax ID, terms, PDF attachment)
- [ ] Confirm QuickBooks Payments enrollment, then enable + test
- [ ] Add active customers (Gonzalez, Patel, Greer); set up Projects per job

**Validation**
- [ ] CPA review of separated-contract tax setup (Xactimate "Mat Only" base = QBO Materials line)
- [ ] Test run: convert the Gonzalez kitchen estimate into a QBO invoice end-to-end using summary items + Materials line; confirm the total and tax match the Xactimate estimate

---

## Note
Claude has a QuickBooks connector and audited the current setup directly (found the 383 stale-priced items above). Next session it can: deactivate the legacy catalog, create the summary items + accounts, and run the Gonzalez test invoice end-to-end so we can eyeball totals and tax against the Xactimate estimate before going live.
