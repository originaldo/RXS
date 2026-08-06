# RXS Group LLC / ARC — Project Context

## Who you are working with

**Aldo Rojas** — owner. Email `aldo@rxsgrp.com`, `(832) 650-6500`.

Sign customer-facing texts and emails as **Aldo**. Full name is **Aldo Rojas**.
**Never guess or infer a person's name from an email address or any other partial
source.** If a name is not stated, leave it blank and ask.

## Company facts

- **Legal / brand:** RXS Group LLC **dba America's Restoration Contractor**
  (SINGULAR — "Contractor", not "Contractors", even though the logo art renders
  the plural in the wordmark).
- **Location:** Houston, TX. Pricing basis **TXHO8X_JUL26** (Houston Xactimate).
- **Contact:** (832) 650-6500 · help@arcsvcs.com · arcsvcs.com · Licensed & Insured
- **Brand colors:** Navy `#16243F`, Red `#C41230`. Helvetica throughout.
- **Logo:** house-mark only in the letterhead. Do not recreate the
  "AMERICA'S / RESTORATION / CONTRACTOR" text wordmark as type; use the logo image.
  A copy lives at `estimates/3231_Appalachian_Trl/arc-logo.png`.

## Estimating

Full SOP is the `arc-estimating-sop` skill — invoke it for any estimate work.
Key standing rules that have bitten before:

- **Never invent Xactimate pricing.** Every rate comes from the master price list
  or Xactimate directly. Missing code means flag it `PULL`, not a plausible guess.
- **Set the category first.** Mold → HMR. Water only → WTR. Demo → DMO. Cleaning → CLN.
  HMR runs 10–40% above WTR on identical selectors.
- **Containment barrier (BARR) is priced on walls + ceiling SF**, never the doorway.
- **EQ monitoring** = 1.75 hr × techs × site days; site days = drying days + 1.
  This line carries drive time. Do not cut it to soften a total.
- **The rebuild is SUBBED.** Do not price it on bare per-SF production — that
  underquotes it and the sub eats the gap. Use the Xactimate minimum-charge lines
  (`DRY MNRP`, `PNT MNRP`, labor+material) as the realistic floor for small repairs.
  Target ~30–40% margin on the rebuild. Typical mitigation-to-rebuild ratio is
  roughly 1 : 1.5 to 1 : 2.
- **Tax:** Texas labor for residential repair/restoration is not taxable. Materials
  only, 8.25% Houston. Tax computed outside Xactimate is an estimate — say so.
- **No O&P markup** unless asked.
- **Never include mold testing** (TDLR conflict of interest).
- Always disclose a probable moisture source that falls outside scope, in writing.

## Customer communication voice (texts and emails)

- Casual, warm, human. Short sentences. Sign as Aldo.
- **No em-dashes and no " - " dashes.** They read as AI. Use periods, commas, parentheses.
- **No emojis.** "lmk", "plz" are fine.
- State facts and what needs to happen. No filler, no recapping things the customer
  was present for.
- On price: let the customer reveal the competitor's number first, then get close.
  Never pre-discount or negotiate against yourself.

## Deliverable formats

- **Internal workbook (.xlsx):** formula-driven off a yellow JOB ASSUMPTIONS block.
  No hardcoded totals. Columns: # / Cat / Sel / Description / Unit / Qty / Unit Price /
  Extension / Note.
- **Customer proposal (.pdf):** one page, ARC house style. Sell by scope SECTION,
  never itemize. No machine counts, crew sizes, or day counts.
  - Known bug in the skill's generator: checkmarks specified as ZapfDingbats glyph
    `3` fall back to a solid notdef box in viewers lacking that base-14 font.
    Draw them as vector paths instead (see
    `estimates/3231_Appalachian_Trl/build_proposal.py`).
  - Always verify: page count, bottom space, and **render and look at it**. Text
    extraction will not reveal a glyph rendering as the wrong character.
