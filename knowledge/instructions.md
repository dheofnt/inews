## ================================================================
## FILE: SPR_Station_Program_Rating_Skill.md
## ================================================================

# SPR (Station Program Rating & Share) Skill

## Overview
Generates the "SPR" (Station Program Rating & Share) Excel file — a minute-by-minute broadcast log grid for a single date, showing every program aired by INEWS, TVONE, KOMPASTV, and METRO side by side, each cell carrying that program's TVR/Share. This is a distinct deliverable from the Performance Dashboard skill (which reports aggregate share numbers) — SPR is the full per-program broadcast schedule for one day, styled as a printable station log.

**Synonyms (explicit user confirmation, 2026-08-26): "SPR", "Ratmap", and "SUSA" all refer to this exact same deliverable.** Treat a request using any of these three terms identically — same data sources, same grid algorithm, same layout, same output file. Don't ask which one they mean or treat them as different reports.

## Trigger
- **Single date**: "buatkan SPR" / "bikin SPR" / "Station Program Rating" / "buatkan Ratmap" / "bikin SUSA" (or any mix of these terms) for a specific date — e.g. "buatkan SPR untuk tanggal 16 Agustus 2026" → one workbook, one sheet, filename `SPR News Channel {D Mon YYYY}.xlsx`. The output filename always uses the "SPR News Channel" convention regardless of which of the three terms the user used to ask for it.
- **Full week**: "SPR Week {YYWW}" / "Ratmap Week {YYWW}" / "SUSA Week {YYWW}" (e.g. "SPR Week 2628") → one workbook covering all 7 days of that week, **one sheet per day** (see "Multi-Day / Weekly Workbooks" below). Don't ask which 7 dates — derive them from the data (see that section).

## Data Sources
- **`Program Detail REV2`** sheet — the ONLY source for the grid content. Columns used: `Date` (`DD/MM/YYYY`), `Channel`, `Program`, `Start time` (string `HH:MM`, extended past 24:00 up to `25:59` for next-day-early-morning slots), `End time` (same format), `TVR`, `Share`, `Dur` (minutes). Filter to the requested date and each of the 4 channels.
  - **Always filter out rows where `Dur == 0`** before building the grid — these are zero-duration placeholder/carryover rows in the source data (confirmed cases: "Kompas Sepekan L", "8Ersama 1Ndonesia L3" instantaneous duplicates) and must not appear in the log.
- **`PERFORMANCE DAILY`** sheet — the ONLY source for the header share numbers (the 4 numbers shown above each channel name), the plain `Week` number (used to build the YYWW label), and — for weekly requests — for **discovering which 7 calendar dates belong to a given `Week {YYWW}`** (filter `PERFORMANCE DAILY` where `Week == int(YYWW[2:])`, i.e. the last two digits; the resulting rows' `Date` values are the week's 7 dates, Sunday through Saturday). Per the project's `GLOBAL_National_Data_Source_Rule.md`, never derive the share numbers from Program Detail.

## Grid Construction Algorithm
The broadcast day runs from **02:00 to 25:59** (i.e. 02:00 today through 01:59 the next calendar day), in 5-minute rows — 288 rows total, mapped to spreadsheet rows 6 through 293 (row 6 = 02:00).

```python
def to_minutes(hhmm):
    h, m = hhmm.split(":")
    return int(h) * 60 + int(m)

def grid_row(total_min):
    offset = total_min - 120       # 02:00 baseline -> row 6
    r = 6 + round(offset / 5)      # ROUND to nearest 5-min row, not floor
    return max(6, min(293, r))
```

For each channel, sort programs by `Start time` (stable sort) after filtering `Dur > 0`, then compute `grid_row` for each program's start time and place it in a `{row: (name, tvr, share)}` dict — **iterate in start-time order and let later programs overwrite earlier ones at the same row**. This is intentional: when a very short program's rounded start collides with the next program's rounded start, the later program wins and the short one silently disappears from the visual grid — this exactly matches how the real SPR files behave (verified against two independent example files) and is not a bug to "fix".

Each channel's occupied rows are then sorted; a block's row-span runs from its own row to one row before the next occupied row (or to row 293 for the last block).

## Program Cell Text Formatting
- Program name: title-cased from the raw ALL-CAPS `Program` field (`str(name).title()`) — this reproduces the source display convention exactly, including quirks in the raw data (e.g. `"8ERSAMA 1NDONESIA L3"` is genuine source data, not an OCR error — do not "correct" it).
- Numbers (inside the program grid cells): format TVR and Share with `f"{x:.2f}".rstrip('0').rstrip('.')` — this drops trailing zeros exactly like the source (`0.10` → `"0.1"`, `0.00` → `"0"`, `2.71` stays `"2.71"`). **This trailing-zero-strip rule applies only inside the program grid cells, not to Row 4 — see Row 4 spec below, which is a deliberate exception.**
- **Span == 1 row (5 min)**: compact format, 2 spaces, small font — `f"{name}  {tvr}/{share}"`, font size 8.
- **Span >= 2 rows**: long format with a fixed 50-space gap between name and numbers (relies on `wrap_text=True` to visually separate them within the merged block) — `f"{name}" + " "*50 + f"{tvr}/{share}"`. Font size scales with span: `>=6 rows → 14pt`, `3-5 rows → 13pt`, `2 rows → 11pt`.
- Alignment for all program cells: `horizontal="center", vertical="center", wrap_text=True`.

## Layout (current template — columns, no spacers between channels)
This is the **current** revision as of 2026-08-25; an earlier revision had spacer columns between channels (see Revision History below) — always build to the layout described here unless the user explicitly asks to revert.

| Col | Width | Content |
|---|---|---|
| A | 7.71 | Time ruler (sparse, every 30 min) |
| B | 29.71 | INEWS |
| C | 29.71 | TVONE |
| D | 29.71 | KOMPASTV |
| E | 29.71 | METRO |
| F | 7.71 | Time ruler (sparse, every 30 min, mirrors column A) |
| G | 0.855 | spacer |
| H | 1.14 | spacer |
| I | 9.14 | reserved, **intentionally left blank** (the old dense per-5-minute ruler was removed by explicit user request — do not populate it) |
| J | 1.71 | spacer |

Row heights: row1=30.75, row2=27.0, row3=27.0, row4=39.0, row5=24.75, rows6-293=6.6 each.

Freeze panes at `B6`. Grid lines hidden (`ws.sheet_view.showGridLines = False`).

### Rows 1-5 (header block)
- Row1: `"Station Program Rating & Share - {WeekdayName}"` — Garamond 24 bold, `horizontal="centerContinuous"` across A:J (not a real merge — set the same alignment on the blank cells to its right).
- Row2: `"INEWS, TVONE, KOMPASTV, METRO"` — Arial Narrow 22, centerContinuous.
- Row3: `"Week {YY}{WW} - {Month D, YYYY} - National Urban,  20+ UM Terrestrial"` (note the double space before "20+") — Arial Narrow 22, centerContinuous. Week label is `f"26{week_num:02d}"` (year prefix `26` + the plain week number from `PERFORMANCE DAILY`'s `Week` column, zero-padded to 2 digits).
- **Row4**: the day's official share value per channel (from `PERFORMANCE DAILY`) in columns B-E — Arial 14 bold, centered. **Always formatted to exactly 2 decimal places** (`f"{value:.2f}"`, e.g. `"3.00"`, `"3.15"`) — explicit user request, 2026-08-26. This is the one place in the whole grid that does NOT use the trailing-zero-strip rule; every other TVR/Share number in the program cells still strips trailing zeros as described above. Don't let the two conventions bleed into each other.
- Row5 header cells:
  - Columns A and F: text `"Time"`, fill `#333F50` (dark navy — this is the Office theme's `dk2` color tinted -0.25, precomputed as a literal hex so no theme dependency is needed), font Arial 16 bold white, `border(l=thin, r=thin, t=thin, b=thin)` (full box).
  - Columns B-E: channel name, fill = brand color (`INEWS #C00000`, `TVONE #0070C0`, `KOMPASTV #00B0F0`, `METRO #FFC000` — METRO's is Office theme `accent4`, same literal hex), font Arial 20 bold, white text for INEWS/TVONE (dark fills), black text for KOMPASTV/METRO (light fills). Border: `top=thin, bottom=thin` only — no left/right border needed at the header since each cell's own fill color already delineates it.

### Time ruler columns (A and F), rows 6-293
48 half-hour blocks of 6 rows each. For block `i` (0-indexed), `row_start = 6 + i*6`, time label = `(120 + i*30)` minutes-from-midnight **wrapped modulo 24 hours** — i.e. past 24:00 the label rolls back to `00:00, 00:30, 01:00, 01:30...` (NOT an extended `24:00/24:30` notation). This was verified directly against the source example file.

Merge only the first 3 rows of each 6-row block (`row_start` to `row_start+2`) for the label text (`vertical="top"` so it sits near the top of the block). **Critical border rule** (this was the one revision the user explicitly corrected): the left/right border must run continuously through **all 6 rows** of the block, not just the 3 merged/labeled rows — only the very last of the 6 rows gets a `bottom=thin`, and only the first row gets `top=thin`. Getting this wrong makes the ruler look like disconnected floating boxes instead of one continuous column with dividers only at each half-hour mark. Both A and F get both `left=thin` and `right=thin` (full box sides) in the current no-spacer layout.

### Column I
Leave entirely blank (width reserved at 9.14 but no cell values, no borders) — do not populate a dense per-minute ruler here in the current template revision.

### Program columns (B-E), rows 6-293
For INEWS (column B, the leftmost channel column): no left border (relies on column A's right border as the divider). For TVONE/KOMPASTV/METRO (C, D, E): `left=thin` on every row of the block to create the divider between adjacent channels (no explicit right border needed — the next column's left border serves as the shared line). Every block additionally gets `top=thin` on its first row and `bottom=thin` on its last row.

### Footer / page setup
```python
ws.page_setup.orientation = "portrait"
ws.page_setup.fitToWidth = 2
ws.page_setup.scale = 40
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_options.horizontalCentered = True
ws.page_margins.left = ws.page_margins.right = 0.0
ws.page_margins.top = 0.25
ws.page_margins.bottom = 0.2
ws.page_margins.header = ws.page_margins.footer = 0.0
ws.oddFooter.left.text = "Programming Research & Development INEWSTV"
ws.oddFooter.center.text = "Confidential - Internal Use Only"
ws.oddFooter.right.text = "Source: Nielsen - Arianna"
ws.print_area = "A1:J296"
```

### Sheet naming (single-date request)
Workbook sheet tab = 3-letter day abbreviation (`Mon`, `Tue`, `Wed`, `Thu`, `Fri`, `Sat`, `Sun`) matching the target date's weekday. Output filename: `SPR News Channel {D Mon YYYY}.xlsx` (e.g. `SPR News Channel 16 Aug 2026.xlsx`) — always this naming convention, even if the user asked using "Ratmap" or "SUSA" instead of "SPR".

## Multi-Day / Weekly Workbooks
When the user asks for a whole week ("SPR Week 2628"), build **one workbook with 7 sheets, one per day of that week**, each sheet built exactly per the single-day spec above (own header block, own row4 share values, own program grid) — do not merge or stack days into one sheet.

- **Resolve the 7 dates**: filter `PERFORMANCE DAILY` for `Week == <last two digits of the YYWW label>` (e.g. Week 2628 → `Week == 28`) and take that filter's 7 `Date` rows — this gives the exact Sunday-through-Saturday date range without guessing offsets from a calendar. Verify you get exactly 7 rows before proceeding.
- **Sheet tabs**: name each sheet by its own weekday abbreviation (`Sun`, `Mon`, `Tue`, `Wed`, `Thu`, `Fri`, `Sat`) — since a week always has 7 distinct weekdays, no collision handling is needed. Build in date order so tabs read left-to-right Sun→Sat.
- **Output filename**: `SPR News Channel_Week {YYWW}.xlsx` (e.g. `SPR News Channel_Week 2628.xlsx`) — note the underscore before "Week" and no spaces removed elsewhere, matching the user's exact naming request.
- Implementation-wise, refactor the single-day build into a `build_sheet(wb, target_date)` function that creates and populates one sheet on a shared `openpyxl.Workbook()`, then loop it over the 7 resolved dates. Remove the default blank sheet openpyxl creates (`wb.remove(wb.active)`) before adding the 7 real ones.
- QA: after saving, reload and confirm `wb.sheetnames` has exactly the 7 expected day-abbreviations in order, and spot-check each sheet's row4 share values against `PERFORMANCE DAILY` for its date.

## Reference Implementation
Complete, verified-working Python scripts (openpyxl) implementing everything above exist in this session's history: a single-day version (used for 16 Aug 2026, 17 Aug 2026, 21 Mar 2026, and 10 Mei 2026) and a multi-sheet weekly version (`build_sheet(wb, target_date)` refactor, used for Week 2628 / 12-18 Jul 2026). Reconstruct directly from the algorithm and layout spec above — every design decision in this doc was reverse-engineered cell-by-cell from user-supplied real SPR files and corrected based on direct user feedback (see Revision History), so treat this doc as the source of truth over any older cached script.

## QA Before Delivery
1. Reload the saved file with `openpyxl.load_workbook()` to confirm it isn't corrupted; for weekly workbooks also check `wb.sheetnames` matches the expected 7 day-abbreviations.
2. Convert to PDF with `soffice --headless --convert-to pdf` and render a page (or one page per sheet, for weekly workbooks) with `pdftoppm` to visually inspect: header colors correct, title/date line correct, Time column borders continuous (no gaps), Row 4 share values show 2 decimals, no console/conversion errors.
3. Spot-check a handful of program cells' TVR/Share text against the source `Program Detail REV2` rows directly.

## Revision History (most recent first)
- **2026-08-26**: Confirmed "SPR", "Ratmap", and "SUSA" are user-interchangeable names for this exact same deliverable — no behavior difference, just vocabulary. Don't ask for clarification when any of the three terms comes up.
- **2026-08-26**: Row 4 (the share value shown above each channel name) now always formatted to exactly 2 decimal places (`"3.00"` not `"3"`), explicit user request. Every other number in the grid (the per-program TVR/Share pairs inside cells) still uses the trailing-zero-strip format — this is a deliberate two-convention split, not an inconsistency to "fix" later.
- **2026-08-25**: Added the multi-day/weekly workbook variant (`SPR Week {YYWW}` request → one workbook, 7 sheets Sun-Sat, filename `SPR News Channel_Week {YYWW}.xlsx`), including how to resolve a YYWW week label to its 7 calendar dates via `PERFORMANCE DAILY`.
- **2026-08-25**: Fixed the Time-ruler border bug — border must span the full 6-row half-hour block, not just the 3-row label merge (user caught this by screenshotting both Time columns and comparing).
- **2026-08-25**: Removed the spacer columns between channels (previously each channel column had a narrow blank column after it); channels B-E are now directly adjacent with thin left-border dividers instead. The dense per-5-minute time ruler column (previously populated) is now left blank. Time labels past 24:00 wrap to `00:00/00:30/...` instead of showing `24:00/24:30/...`.
- **2026-08-24**: Initial version built from a user-supplied example file for 17 Agustus 2026 (spacer-column layout, extended-hour time labels past 24:00, dense per-minute ruler populated). Superseded by the 2026-08-25 revision above — do not use this older layout unless the user explicitly asks to revert.


## ================================================================
## FILE: INEWS_Competitive_Daypart_Deepdive_Skill.md
## ================================================================

# INEWS Competitive Daypart Deep-Dive Skill

## Overview
Generate a detailed, graphic-forward analysis explaining WHY a competitor beat (or INEWS beat) a competitor on a specific day — broken down program-by-program, daypart-by-daypart. This is the skill to use for a "why did TVONE win on 17 Agustus, what programs were there" style investigation, not just a share-comparison number. Deliverable format is flexible (HTML infographic or PPT deck) depending on what the user asks for — the data-sourcing and structure rules below apply either way.

> **Implementation note (2026-08-26)**: the actual working, code-level implementation of this skill's PPT output path now lives as a proper Skill-tool skill, `inews-competitor-daypart-deck` — a `prepare_data.py` + `build_deck.js` pair, packaged as a `.skill` file and delivered to the user for account install. This project doc's own copy of that skill's `SKILL.md` and both scripts is mirrored at `claude/INEWS_Competitor_Daypart_Deck_Skill.md`, `claude/inews_competitor_daypart_deck_prepare_data.py`, and `claude/inews_competitor_daypart_deck_build_deck.js` — read those for the concrete, currently-correct slide structure (7 slides: Cover, Daypart chart, Head-to-Head, Top 25, optional Special Programs, Fresh vs Rerun, Executive Summary), which has diverged somewhat from the 12-section outline below as real revisions accumulated (INEWS-first Executive Summary framing, fixed INEWS/TVONE/KOMPASTV/METRO channel order everywhere per `GLOBAL_Competitor_Channel_Order_Rule.md`, rounded whole-hour Fresh/Rerun with no rounding caveat, a Rerun->Fresh reclassification fix for mistagged holiday-special content, and a hand-computed special-program Share comparison for dates where `Level 1=='Special'` tagging is absent). Treat that mirrored SKILL.md as the source of truth for how the PPT is actually built; treat this doc as the broader conceptual/methodology reference (data-sourcing tiers, HTML-infographic format, error-handling checklist) that still applies regardless of output format.

## Trigger
Use this skill when the user asks something like:
- "kenapa [kompetitor] menang/unggul tanggal [tanggal]? ada program apa?"
- "bandingkan program to program INEWS vs [kompetitor] di [tanggal]"
- "buatkan infografis/PPT kenapa [kompetitor] unggul di [tanggal]"
- "buat analisa performa INEWS [tanggal] vs kompetitor"
- Any follow-up asking to go deeper into a specific date's daily share result with program-level and/or daypart-level detail

This is a heavier, more investigative deliverable than the standard Performance Dashboard (`INEWS_Performance_Dashboard_Skill.md`) — only invoke it when the user is asking "why" for a specific day, not for a routine period share comparison.

## Input Requirements

### Data Sources (three-tier — do not mix them; see `GLOBAL_National_Data_Source_Rule.md`)
1. **`PERFORMANCE DAILY` sheet** — the ONLY source for the official headline daily share number (INEWS vs competitor, e.g. 3.79% vs 3.86%). Columns: `Date` (`DD/MM/YYYY`), `Week`, channel columns. This number is never recalculated or overridden by anything derived from Program Detail or the daypart sheets — Nielsen's official daily figure uses duration/minute-level weighting a simple average won't reproduce exactly.
2. **`BY DAYPART DAILY` sheet (single date) / `BY DAYPART WEEKLY` sheet (week or week range)** — the authoritative source for ALL per-daypart share/TVR/viewers numbers, added 2026-08-25. **Use this instead of averaging `Program Detail REV2` rows by `Daypart Custome`** — that approach is a real methodology hazard: on 17 Agustus 2026 it read the Sore daypart (16:00-17:59) as ~0% for 3 of 4 channels, when the official `BY DAYPART DAILY` figure was actually ~4.7-4.9%, because a long HUT RI special program that ran into that window had all its records tagged under the daypart it *started* in. Both sheets share the same 8 fixed time-range bins: `02:00 - 04:29`, `04:30 - 09:59`, `10:00 - 12:59`, `13:00 - 15:59`, `16:00 - 17:59`, `18:00 - 21:59`, `22:00 - 23:59`, `24:00 - 25:59`. Suggested display names (chronological, matching the SPR template's 02:00 day-start convention): Dini Hari 1, Pagi, Siang 1, Siang 2, Sore, Primetime, Malam, Dini Hari 2 — always show the underlying time range alongside the name (see `GLOBAL_Chart_Label_Style_Rule.md` two-line axis label convention), and always call the segment a **"Daypart"**, never "Blok"/"Blok Waktu" (`GLOBAL_Daypart_Terminology_Rule.md`).
3. **`Program Detail REV2` sheet** — the source for ALL program-level narrative/detail (this tier is unaffected by the daypart-sourcing fix above — it was already correct for program-level facts). Columns include `Date`, `Program`, `Channel`, `000s`, `TVR`, `Share`, `Dur`, `Daypart` (raw time range), `Daypart Custome` (named bucket, useful for a program's own label but NOT for computing a daypart aggregate), `Category` (Fresh/Rerun — see the mistagging gotcha in the mirrored SKILL.md for holiday-special content), `Level 1` (flag `Special` marks holiday/commemorative programs, but is not exhaustive for holiday content — see the same gotcha doc). Use this for: top programs (overall and per-channel), special/ceremonial program identification and comparison, and fresh/rerun composition.

### Parameters
- **Date** (or week/week-range): the period being investigated.
- **Competitor(s)**: the channel(s) to compare against — commonly TVONE, but the deliverable can compare against all three (TVONE, KOMPASTV, METRO) when the user just says "kompetitor" generically. Whenever multiple channels are shown together (chart, table, cards, prose list), always order them INEWS, TVONE, KOMPASTV, METRO — see `GLOBAL_Competitor_Channel_Order_Rule.md`.

## Special-Program Comparison (added 2026-08-25)
When the user asks to compare program-level performance "khususnya program spesial" (especially special programs — e.g. HUT RI ceremonial broadcasts), don't stop at listing them per channel. Build an aggregate comparison across channels: count of special programs aired, total viewers (000s) summed across them, average share per special program, and total duration invested. A channel can win on volume (most total special-program viewers) while another wins on efficiency (higher average share per special program aired) — surface both angles, since they tell different stories (e.g. one channel fielding 3 heavily-promoted specials vs another fielding 2 that each performed better individually). **If `Level 1=='Special'` returns nothing for the date (common for holiday content like Idul Fitri, where specials are just tagged under their normal News/Religious/Sport Level 1 with "SPESIAL"/"LEBARAN"/"IDUL"/"TAKBIR" only in the free-text Program name), fall back to a keyword search across `Program` for all channels rather than reporting "no special programs" — see the mirrored SKILL.md for the exact keyword list and worked example.**

## Output Format

Sections/slides, in order (adapt HTML-section vs PPT-slide framing to the requested format, but keep this order and content):

1. **Hero/cover** — gradient or dark background, a context badge if the date has a special occasion (e.g. "HUT RI KE-81"), one-line title ("INEWS vs Kompetitor — Analisa Performa [tanggal]").
2. **KPI row** — official INEWS share, competitor share(s), gap (all from `PERFORMANCE DAILY`).
3. **Ringkasan Eksekutif (Executive Summary)** — lead with INEWS specifically (which Daypart(s) it wins, its national share vs the nearest competitor), then the competitor's own Daypart wins, then the special-program win/lose verdict, then the single best program of the day. Keep it to a few short sentences, not a long numbered list — see the mirrored SKILL.md's Slide 7 for the current concise template (this was shortened from an earlier, longer draft per explicit user feedback 2026-08-26).
4. **Daypart share comparison chart** — grouped bar, one group per Daypart bin (all 8, in chronological order), INEWS vs competitor(s). X-axis tick labels two lines: Daypart name + time range. Data from `BY DAYPART DAILY`/`BY DAYPART WEEKLY` — never from Program Detail averages.
5. **Total Penonton per Daypart chart** — same X-axis structure, plotting total viewers (000s, the official per-block `000s` figure — not an average across programs) per channel. Label it "Total Penonton per Daypart", not "rata-rata", since the official sheet already gives the real aggregate.
6. **Daypart head-to-head table/detail** — INEWS vs main rival per Daypart bin: share, gap, winner (allow and label a tie/"Seri" when values match, e.g. Siang 1 on 17 Agustus 2026 was 4.38% vs 4.38%). Include a winner-tally summary.
7. **Program-level detail** (from `Program Detail REV2`): top programs overall and per channel (by `000s`), and — when the date has special/ceremonial programming — the special-program comparison table plus the aggregate volume-vs-efficiency comparison described above.
8. **Content composition** (optional, when relevant) — Fresh vs Rerun % by duration, per channel, plus total hours rounded to whole numbers (no "dibulatkan" caveat text — see `GLOBAL_Competitor_Channel_Order_Rule.md`'s sibling note in the mirrored SKILL.md). Double-check `Category` tagging on event/holiday dates before trusting it at face value — see the Rerun-mistagging gotcha.
9. **Verdict** — ONE condensed sentence stating the overall explanation, highlighted callout.
10. **Caveat line** — small text confirming daypart figures are sourced from `BY DAYPART DAILY`/`BY DAYPART WEEKLY` (official, Nielsen-weighted) and program-level detail from `Program Detail REV2`.
11. **Footer / Source Citation** — per `GLOBAL_Source_Footer_Rule.md`, labeled `Weekly Performance`: `Sumber: Nielsen Media Research; National Urban; Weekly Performance; [tanggal].`
12. **Lampiran (if PPT and an SPR exists for the date)** — embed the SPR grid image as a supporting appendix slide (see `SPR_Station_Program_Rating_Skill.md`), plus deliver the underlying SPR `.xlsx` as a companion file. (Note: the current `inews-competitor-daypart-deck` skill build does NOT include this appendix by default — it was explicitly removed from that deck's slides 8-10 on 2026-08-25 per user request. Build the SPR as a separate companion file instead, using the standalone `build_spr.py` pattern, unless the user specifically asks for it back inside the PPT.)

## Chart Data Labels
Follows `GLOBAL_Chart_Label_Style_Rule.md` — every bar shows its bold data value directly.

## Error Handling
- **Deriving daypart aggregates from Program Detail**: no longer acceptable now that `BY DAYPART DAILY`/`BY DAYPART WEEKLY` exist — always use the official sheet for daypart-level share/viewers/TVR.
- **Calling it "Blok"/"Blok Waktu"**: always "Daypart" per `GLOBAL_Daypart_Terminology_Rule.md`.
- **Overclaiming from Program Detail**: never state the headline "who won the day" number using a Program-Detail-derived average — only `PERFORMANCE DAILY` is authoritative for that.
- **Special-program comparison flattened to a list**: don't just list special programs — compute the cross-channel aggregate (count, total viewers, avg share, total duration) so volume vs efficiency is visible.
- **Trusting `Category` (Fresh/Rerun) blindly on event/holiday dates**: cross-check against the program-name repeat-marker convention (see mirrored SKILL.md) before reporting a channel's Fresh% — a first-airing holiday special can be mistagged Rerun.
- **Channel order drifting to alphabetical**: always INEWS, TVONE, KOMPASTV, METRO — see `GLOBAL_Competitor_Channel_Order_Rule.md`.
- **Too much text / not detailed enough**: this skill is graphic-forward — keep prose to the one Executive Summary block and the one Verdict sentence; don't summarize program-level detail down to only top-N when the user wants full program-by-program.

## Notes
- Iteration history: v1 was prose-heavy and rejected; v2 replaced prose with summary charts but was judged too summarized; v3 combined full program-by-program detail with chart-forward style, added Executive Summary + time-range axis labels (2026-08-24). v4 (2026-08-25): switched daypart-level sourcing from Program-Detail averages to the official `BY DAYPART DAILY`/`BY DAYPART WEEKLY` sheets after discovering the Program-Detail-average approach materially misrepresented the Sore daypart on 17 Agustus 2026; added the special-program volume-vs-efficiency comparison; renamed "Blok Waktu" to "Daypart" throughout; added PPT + SPR-appendix as a supported output format alongside the original HTML infographic. v5 (2026-08-26): the PPT output path was formalized into an actual reusable Skill-tool package (`inews-competitor-daypart-deck`, mirrored in this project — see the implementation note at the top); SPR appendix removed from the default PPT flow; Executive Summary shortened and re-ordered to lead with INEWS; fixed channel ordering rule established project-wide; Fresh/Rerun holiday-special mistagging and Level-1-tagging gaps documented and fixed.
- Created 2026-08-24 from the 17 Agustus 2026 (HUT RI ke-81) INEWS vs TVONE investigation. Revised 2026-08-25 after a raw-data update added the official daypart sheets. Revised 2026-08-26 after the Idul Fitri 2026 deck surfaced the Fresh/Rerun and special-program tagging gaps above.

## Related Skills
- `INEWS_Competitor_Daypart_Deck_Skill.md` — the concrete, currently-correct Skill-tool implementation (code + slide structure) of this deep-dive's PPT output path. Read this first for "how is it actually built today."
- `SPR_Station_Program_Rating_Skill.md` — the minute-by-minute program log, useful as supporting appendix material for this skill's PPT output.
- `GLOBAL_National_Data_Source_Rule.md` — full sheet-authority table across all granularities.
- `GLOBAL_Daypart_Terminology_Rule.md` — wording rule.
- `GLOBAL_Chart_Label_Style_Rule.md` / `GLOBAL_Source_Footer_Rule.md` — styling/citation rules.
- `GLOBAL_Competitor_Channel_Order_Rule.md` — channel ordering rule.


## ================================================================
## FILE: INEWS_Competitor_Daypart_Deck_Skill.md
## ================================================================

---
name: inews-competitor-daypart-deck
description: Build (or revise) a 7-slide Head-to-Head Daypart competitive analysis PowerPoint deck for INEWS vs one rival channel (TVONE/KOMPASTV/METRO), for one specific date, from RAW INEWS.xlsx. Covers national share, per-Daypart share with audience Kontribusi %, a Head-to-Head Daypart table, Top 25 programs, an optional per-Daypart special-event program comparison (e.g. HUT RI), Fresh vs Rerun composition, and an auto-composed Executive Summary. Use whenever the RDO/INEWS team asks for "analisa performa [channel] tanggal X vs kompetitor", "head to head daypart", "performa harian vs TVONE/KOMPASTV/METRO", or a refresh of that deck for a new date. Don't wait for the word "PowerPoint" or "deck" -- "performa 17 Agustus vs TVONE, buat pptnya" is enough.
---

# INEWS Competitor Head-to-Head Daypart Deck

## What this produces

A 7-slide, presentation-ready `.pptx`, always in this order:

1. **Cover** -- dark slide, optional event badge (e.g. "HUT RI KE-81"), national KPI cards for all 4 channels.
2. **Perbandingan Share per Daypart** -- grouped bar chart across the 8 official Daypart bins (English names), each category labelled with its Kontribusi % (bare number, no "Kontribusi" word in the chart itself -- see Design details), and a summary insight bar above the chart.
3. **Head To Head: {MAIN} vs {RIVAL} By Daypart** -- table (Daypart / Jam / Kontribusi / MAIN% / RIVAL% / Selisih / Unggul) + a "Skor Kemenangan Daypart" tally card, both grounded in an auto-composed insight bar.
4. **Top 25 Program Hari Ini** -- Filler excluded, sorted by 000s, always 000s+TVR+Share together, MAIN_CHANNEL's own rows highlighted.
5. **Program Spesial {EVENT_NAME} per Daypart** -- *only rendered when `CONFIG.EVENT_NAME` is set and Level 1=='Special' rows exist that date* -- per-Daypart mini-tables (same 000s/TVR/Share/highlight rule as slide 4), blue header (never gold).
6. **Komposisi Konten: Fresh vs Rerun** -- vertical stacked bar, channels in alphabetical order, blue=Fresh/orange=Rerun, **total hours for the whole day** (not a per-program average) alongside each channel's card.
7. **Executive Summary** -- dark slide, one auto-composed boxed paragraph.

This is a recurring, opinionated workflow tied to one project's data and conventions (RDO / INEWS), not a generic template -- it exists to skip re-deriving the Daypart data-sourcing rule, the Kontribusi normalisation, the special-program edge cases, and a couple of narrative-generation bugs every single time this request comes in. All of the above was hammered out interactively over several revision rounds on the 17 Agustus 2026 edition of this deck; the numbers and layout choices below are not guesses.

## Fixed project conventions -- do not re-derive these per run

- **Data file & sheets** (`RAW INEWS.xlsx`):
  - `PERFORMANCE DAILY` -- national daily share, the only authoritative source for the cover KPI cards and the Head-to-Head slide's headline gap.
  - `BY DAYPART DAILY` -- per-Daypart Share/TVR/000s (Date + Day Part key). **Never** derive a Daypart's share by averaging `Program Detail REV2` rows for that Daypart instead -- a single long special program can hide a whole Daypart's real audience behind its *start time's* bucket. (Concrete case: 16 Aug 2026's "Sore" Daypart showed 0% under a naive Program-Detail average, because a 4-hour flag-ceremony special starting at 14:00 was bucketed entirely under "Siang 2".) This sheet is the fix.
  - `Program Detail REV2` -- one row per aired program segment, the sole source for program-level detail: Top 25, special programs, Fresh/Rerun composition. Columns used: `Channel, Program, Level 1, Start time, End time, Daypart Custome, Category, TVR, Share, Dur, 000s`.
  - `Kontribusi By Daypart` -- **optional**, only present in RAW files from ~Aug 2026 onward. Channel == `"Total TV"` row gives each Daypart's share of that date's total TV audience (normalise to % of the 8-Daypart sum yourself -- the sheet stores raw 000s, not %). When this sheet is missing, `prepare_data.py` sets `kontribusi_daypart` to `null` and `build_deck.js` automatically falls back to the pre-Kontribusi wording ("bervolume besar" instead of "berkontribusi besar", no % in chart labels or tables) -- don't hand-patch this per run, the fallback is already wired.
- **"Daypart Custome" vocabulary is NOT stable across RAW file revisions.** An older export used Indonesian labels with a merged "Dini Hari" bucket covering both ends of the day; a newer export (confirmed Aug 2026) uses the same English labels as the 8 official Daypart bins, unmerged. `prepare_data.py`'s `canon_daypart_custome()` normalises either vocabulary to the canonical English 8-bin set and warns on stderr for anything it doesn't recognise -- read that warning if it fires, don't assume it's silently fine. This bit us once already: a mid-session RAW update silently changed this column's language, and a same-conversation spot-check of *other* sheets (which hadn't changed) was wrongly read as "nothing changed."
- **Daypart terminology**: always say **"Daypart"**, never "Blok" / "Blok Waktu". English display names (used everywhere -- chart, table, narrative) map from the 8 official bins:
  | Time range | English (display) |
  |---|---|
  | 02:00 - 04:29 | Early Morning |
  | 04:30 - 09:59 | Morning |
  | 10:00 - 12:59 | Noon 1 |
  | 13:00 - 15:59 | Noon 2 |
  | 16:00 - 17:59 | Evening |
  | 18:00 - 21:59 | Prime Time |
  | 22:00 - 23:59 | Night |
  | 24:00 - 25:59 | Midnight |
- **"Kontribusi" not "volume"**: when describing why a Daypart matters (a rival winning "the big" Dayparts), say "Daypart berkontribusi besar" with the real % from `Kontribusi By Daypart`, never "Daypart bervolume besar" -- that wording was explicitly replaced once the Kontribusi sheet became available. Keep "volume" only for literal counts (e.g. "TVONE paling agresif dari sisi volume (3 program)" -- that's about program *count*, a different concept, and stays as "volume").
- **Special-program "Breaking News" gap**: `Level 1=='Special'` is the primary tag for event-driven programs (e.g. HUT RI ceremonies) but is not exhaustive -- a channel's generic "Breaking News" bulletin covering the same event in the same Daypart is often tagged plain `Level 1=='News'`, not `Special`. Pass `--include-breaking-news` to `prepare_data.py` so every Daypart group that already has a Special-tagged program also pulls in each *other* channel's "Breaking News"-titled program in that same Daypart -- this is how the Primetime group went from a lopsided 2-channel comparison to a fair 4-channel one in this deck's actual history. Only applies within Dayparts that already have at least one genuine Special-tagged entry (an all-Breaking-News Daypart isn't auto-included, to avoid inventing an "event" comparison where there wasn't one).
- **Fresh vs Rerun, total hours, not average**: sum `Dur` (minutes) by `Category` (Fresh/Rerun) across the **whole day, all Level 1 types including Filler** (Filler segments are themselves tagged Fresh/Rerun and belong in the total -- don't exclude them, that was tried and it undercounts). Report both the % split and the absolute total hours (`fresh_hours`, `rerun_hours`) -- a request for "total dalam 1 hari, misal Fresh 20 jam, Rerun 4 jam" means literally that, not a per-program average.
- **Holiday-special content is sometimes mistagged `Category=="Rerun"` on its first-ever airing.** Found on 21/03/2026 (Idul Fitri): every genuine Rerun row that day carries an explicit repeat marker in the Program name (`(R)` for INEWS/KOMPASTV, `(F0#)` for TVONE, trailing `L#` loop numbers for METRO) -- except a handful of Lebaran-special blocks (INEWS's "INEWS PAGI/SIANG SPESIAL LEBARAN", TVONE's "APA KABAR INDONESIA PAGI/MALAM SPS LEBARAN", KOMPASTV's "SAPA INDONESIA MALAM SPS LEBARAN") which carry no marker and never air as "Fresh" on any other date either. `prepare_data.py`'s Fresh/Rerun block now reclassifies Rerun-tagged rows to Fresh when the Program name signals a holiday special (`LEBARAN`/`SPESIAL`/`SPS `) AND carries none of those repeat markers -- this only ever flips Rerun->Fresh and is a no-op (0 minutes) on a normal non-event day, so it's safe to leave on for every run. It logs what it reclassified to stderr (`NOTE: reclassified Rerun->Fresh...`) -- read that line whenever it fires and sanity-check it before trusting the Fresh/Rerun slide, the way the "Kontribusi" numbers above were checked. This is the reason a first pass at 21/03/2026 showed INEWS with the *lowest* Fresh% of the 4 channels (62.77%) when it was actually roughly tied for highest (85.18%) once corrected.
- **Holiday-special content is also not reliably tagged `Level 1=='Special'`.** On 21/03/2026, `special_agg`/`special_by_daypart` came back completely empty even though real Idul Fitri special programming existed for all 4 channels -- it's just tagged under the program's normal Level 1 (News/Religious/Sport) with "SPESIAL"/"LEBARAN"/"IDUL"/"TAKBIR" only in the free-text Program name, not the Level 1 field. `--include-breaking-news` doesn't help here either since it only *adds to* an existing Special-tagged group, and there wasn't one. When this happens for a genuinely special-event date, don't just leave the Executive Summary's special-program sentence out (the auto-fallback used to do this) -- do a manual keyword search across `Program` (`LEBARAN|SPESIAL|IDUL|TAKBIR|EID|SPS `) for all 4 channels on that date, compute each channel's duration-weighted average Share across its matches, and hand-write the comparison into the Slide 7 `summaryText` (with a comment, like the current file has) rather than silently omitting it -- this was an explicit user request ("kamu tonjolin program spesial dihari itu, unggul/kalah gak inews vs kompetitor").
- **Channel order is always INEWS -> TVONE -> KOMPASTV -> METRO, never alphabetical or share-sorted.** This is a standing GLOBAL project rule (`GLOBAL_Competitor_Channel_Order_Rule.md`), not specific to this skill -- it applies to every competitor comparison in the whole project. `CHANNELS` at the top of `build_deck.js` already holds this order and every slide should read from it directly. The Fresh vs Rerun slide (Slide 6) used to build its own `ORDER_ALPHA = [...CHANNELS].sort()` and use that instead -- fixed 2026-08-26. If you ever see a `.sort()` applied to a channel list anywhere in this script, that's a bug, not a style choice.
- **Colors** (hex, no leading `#`, pptxgenjs convention): INEWS `c41e3a` -> TVONE `1976d2` -> KOMPASTV `00bcd4` -> METRO `ff9800`. Dark navy `1e2761` for headers/cover/exec-summary background. Special-program table headers are **blue** (`1976d2`/`1e2761`), never gold -- gold was the first draft and was explicitly rejected.
- **Fresh/Rerun hours are shown rounded to whole numbers, with no "rounded"/"pembulatan" caveat text anywhere on the slide** -- an explicit user request (2026-08-26), the opposite of the Duration Break skill's convention (which does show a rounding footnote). Don't copy that footnote pattern into this deck.
- **Table row highlight**: MAIN_CHANNEL's own rows get a light fill (`FDECEA`) plus bold program-title text, in every "top program" style table (Top 25, special-program groups) -- this was a specific, repeated request ("Higlight Program iNewsnya") and applies to *every* such table, not just one slide.
- **Footer** (every slide, exact wording, only the date changes): `Sumber: Nielsen Media Research; National Urban; Daily Performance; [tanggal]` -- never the word "Resmi" or "Official" anywhere in the deck (footer, chart titles, or prose). If you extend this skill to a weekly/monthly cut, swap "Daily Performance" for "Weekly Performance" / "Monthly Performance" and the date for the period range -- see the sibling project rule doc `GLOBAL_Source_Footer_Rule.md` if present.
- **`LAYOUT_WIDE`** (13.33x7.5in) is the base layout, matching every other INEWS deck skill in this project.

## Step 1 -- Run `scripts/prepare_data.py`

```
python3 scripts/prepare_data.py "RAW INEWS.xlsx" --date 17/08/2026 \
  --main INEWS --rival TVONE \
  --event-name "HUT RI ke-81" --include-breaking-news \
  --out data.json
```

- `--date` is `DD/MM/YYYY` (matches `PERFORMANCE DAILY`'s own date format).
- `--main`/`--rival` default to `INEWS`/`TVONE` -- change `--rival` to compare against KOMPASTV or METRO instead.
- `--event-name` and `--include-breaking-news` are both optional. Omit `--event-name` for a routine (non-event) date -- slide 5 is skipped entirely in that case, and the cover's event badge doesn't render.
- Read the script's stdout summary every run (`official_share`, `daypart tally`, `top25` INEWS count, `special_by_daypart` groups) as a first sanity check before opening PowerPoint. It also prints a `WARNING` on stderr for any unrecognised `Daypart Custome` value -- don't ignore that.
- If `Kontribusi By Daypart` isn't in the workbook, the script prints a note and proceeds with `kontribusi_daypart: null` -- this is expected for older RAW files, not a bug.

This script only covers **daily** granularity (one specific date) -- Top 25 / special programs / Fresh-Rerun are inherently per-date data (`Program Detail REV2` is a per-episode log, not a weekly aggregate). If a request is explicitly *weekly*, slides 2-3 could in principle be re-pointed at `BY DAYPART WEEKLY` (Week + Day Part key, same per-Daypart shape as the daily sheet) and a weekly `Kontribusi By Daypart` sum, but slides 4-6 don't have a natural weekly equivalent -- treat a weekly ask as out of scope for this skill's current script and say so, rather than force-averaging program-level data across a week.

## Step 2 -- `scripts/build_deck.js`

Edit the `CONFIG` block at the top (`DATA_PATH`, `OUT`, `EVENT_NAME` -- keep `EVENT_NAME` in sync with what you passed `prepare_data.py`, it drives the cover badge and slide 5's title/visibility), then:

```
node build_deck.js
```

Everything else -- every number in every insight bar, the tally, the Kontribusi phrasing, the Executive Summary paragraph -- is computed live from `data.json` inside the script (see the "Derive the small set of narrative facts" block near the top). You should not normally need to hand-edit sentence templates between runs; if a run's story genuinely needs an extra nuance the auto-text can't express, edit the `summaryText` template in the Slide 7 section directly, but check it still reads correctly for the *next* run too (don't leave one run's hardcoded specifics behind).

**Executive Summary ordering rule**: point 1 of the summary paragraph always leads with `MAIN_CHANNEL` (INEWS) -- which Daypart(s) it wins plus its national share vs `RIVAL_CHANNEL` -- then follows with the rival's own Daypart wins and the top overall program. This is a deliberate, explicit convention (not derived from who actually won that day): the insight bars on slides 2-3 and 7's tally sentence still correctly credit whichever channel nationally won via `dailyWinner`/`otherChannel`/`winnerWinListStr`, but the Slide 7 headline paragraph itself always uses `MAIN_CHANNEL`/`mainWinDayparts`/`mainWinListStr` first regardless of `dailyWinner`, since this deck is an INEWS-perspective tool. Don't revert this to `dailyWinner`-first phrasing on a future edit without checking with the user first.

## Step 3 -- QA pipeline, every time, before delivering

1. Run the pptx skill's `validate.py` against the built file.
2. Convert to PDF via the pptx skill's `soffice.py --convert-to pdf`, then `pdftoppm -jpeg` each slide.
3. Actually open (Read) each rendered slide image and look at it -- text overflow, clipped headers, an empty slide 5 when an event was expected, a highlight color that didn't apply. `validate.py` does not catch layout problems.
4. Run `markitdown` on the pptx and grep for `resmi`, `official`, `\bblok\b`, `bervolume` (should only ever appear as "kontribusi" once the Kontribusi sheet is present), `lorem`, `ipsum`, `placeholder`.
5. Only after all four pass, deliver the file.

## Design details worth reusing

- **Chart category labels carry the Kontribusi %, but without the word "Kontribusi" printed under the bars** -- that word already appears once in the insight bar above the chart; repeating it under every single bar reads as clutter and was explicitly cut. Just the bare `NN%` on its own line under the time range.
- **Winner-first phrasing**: any sentence naming a "%, vs %" pair should put the actual winner's number first (`dailyWinner`'s share, then the other channel's) -- looks obviously wrong swapped, and was a real bug caught mid-build (`daily WINNER's own winning Dayparts` comment in `build_deck.js` marks where this matters).
- **`Daypart Custome` normalisation must run before grouping**, not after -- group first by the raw string and you'll silently split what should be one "Prime Time" group into two differently-spelled groups if the RAW file mixes vocabularies within itself (unlikely but not impossible mid-migration).
- **Special-program slide layout adapts to the group count** (1 row of up to 3 cards, or a 2x2 grid for exactly 4) rather than assuming a fixed 4 groups -- the very first version of this deck hardcoded a 2x2 grid and broke when a "Dini Hari"/Early Morning group with only one entry was later dropped, down to 3 groups.
- **Top 25's two side-by-side tables split evenly** (`Math.ceil(length/2)`) rather than a hardcoded 13/12 -- keeps working if a future run's Top N differs from 25.

## Files in this skill

- `scripts/prepare_data.py` -- reads `RAW INEWS.xlsx` for one target date and writes `data.json`: national share, per-Daypart summary (from `BY DAYPART DAILY`), Kontribusi % (from `Kontribusi By Daypart`, gracefully absent), Top 25, special programs (grouped per Daypart, with the Breaking News cross-check), and Fresh/Rerun totals.
- `scripts/build_deck.js` -- pptxgenjs script for the full 7-slide deck. `CONFIG` at the top is the only thing that normally changes per run; every sentence in the deck is derived from `data.json` at build time.


## ================================================================
## FILE: inews_competitor_daypart_deck_prepare_data.py
## ================================================================

```py
#!/usr/bin/env python3
"""
prepare_data.py -- extract every JSON input the Head-to-Head Daypart deck
(build_deck.js in this skill) needs, for ONE target date, from RAW INEWS.xlsx.

Usage:
    python3 prepare_data.py "RAW INEWS.xlsx" --date 17/08/2026 \
        --main INEWS --rival TVONE \
        --event-name "HUT RI ke-81" --include-breaking-news \
        --out data.json

Notes on sourcing (see SKILL.md "Fixed project conventions" for the full
rationale -- this is the short version):

  * Daily national share            -> PERFORMANCE DAILY (Date row)
  * Daily per-Daypart share/TVR     -> BY DAYPART DAILY (Date + Day Part key)
                                        NEVER derive this by averaging
                                        Program Detail REV2 rows -- a single
                                        long special program can hide a
                                        whole Daypart's real audience behind
                                        its start-time bucket.
  * Program-level detail (Top 25,
    special programs, Fresh/Rerun)  -> Program Detail REV2 (one row per
                                        aired program segment, that date)
  * Daypart audience contribution   -> "Kontribusi By Daypart" sheet,
    ("how much of the day's total    Channel == "Total TV" row, normalised
    TV audience sits in this          to % of that date's 8-Daypart total.
    Daypart")                        This sheet may not exist in older RAW
                                      files -- script degrades gracefully
                                      (kontribusi_daypart becomes null) if so.
"""
import argparse
import json
import re
import sys

import pandas as pd

CHANNELS = ["INEWS", "TVONE", "KOMPASTV", "METRO"]

DAYPART_ORDER = [
    "02:00 - 04:29", "04:30 - 09:59", "10:00 - 12:59", "13:00 - 15:59",
    "16:00 - 17:59", "18:00 - 21:59", "22:00 - 23:59", "24:00 - 25:59",
]
DAYPART_NAME_ID = {
    "02:00 - 04:29": "Dini Hari 1", "04:30 - 09:59": "Pagi",
    "10:00 - 12:59": "Siang 1", "13:00 - 15:59": "Siang 2",
    "16:00 - 17:59": "Sore", "18:00 - 21:59": "Primetime",
    "22:00 - 23:59": "Malam", "24:00 - 25:59": "Dini Hari 2",
}
DAYPART_NAME_EN = {
    "02:00 - 04:29": "Early Morning", "04:30 - 09:59": "Morning",
    "10:00 - 12:59": "Noon 1", "13:00 - 15:59": "Noon 2",
    "16:00 - 17:59": "Evening", "18:00 - 21:59": "Prime Time",
    "22:00 - 23:59": "Night", "24:00 - 25:59": "Midnight",
}
# "Daypart Custome" (Program Detail REV2) is a program-level bucketing that,
# in practice, has changed vocabulary between RAW file revisions -- an older
# export used Indonesian labels with "Dini Hari" merging both ends of the day
# into one bucket, a newer export (Aug 2026 onwards) uses the same English
# labels as the 8 official Daypart bins, unmerged. Don't assume either one;
# normalise whatever the current file uses through this alias table so the
# rest of the script (and this skill's build_deck.js) always sees the
# canonical English 8-bin vocabulary. If a RAW file ever surfaces a value not
# listed here, this script warns on stderr and falls back to Early Morning's
# sort position rather than crashing -- check that warning by hand.
DAYPART_CUSTOME_CANON = [
    "Early Morning", "Morning", "Noon 1", "Noon 2",
    "Evening", "Prime Time", "Night", "Midnight",
]
DAYPART_CUSTOME_ALIASES = {
    "dini hari 1": "Early Morning", "early morning": "Early Morning",
    "pagi": "Morning", "morning": "Morning",
    "siang 1": "Noon 1", "noon 1": "Noon 1",
    "siang 2": "Noon 2", "noon 2": "Noon 2",
    "sore": "Evening", "evening": "Evening",
    "primetime": "Prime Time", "prime time": "Prime Time",
    "malam": "Night", "night": "Night",
    "dini hari 2": "Midnight", "midnight": "Midnight",
    "dini hari": "Early Morning",  # legacy merged bucket -- best-effort fallback, see note above
}


def canon_daypart_custome(raw_value):
    key = str(raw_value).strip().lower()
    canon = DAYPART_CUSTOME_ALIASES.get(key)
    if canon is None:
        print(f"WARNING: unrecognised Daypart Custome value {raw_value!r} -- "
              f"treating as 'Early Morning' for sort purposes. Add it to "
              f"DAYPART_CUSTOME_ALIASES in prepare_data.py.", file=sys.stderr)
        return "Early Morning"
    return canon


def parse_date_col(df, col="Date"):
    return pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce")


def fmt_id_date(ts):
    bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
             "Agustus", "September", "Oktober", "November", "Desember"]
    return f"{ts.day} {bulan[ts.month - 1]} {ts.year}"


def title_program(name):
    return str(name).title()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("raw", help="Path to RAW INEWS.xlsx")
    ap.add_argument("--date", required=True, help="Target date, DD/MM/YYYY")
    ap.add_argument("--main", default="INEWS")
    ap.add_argument("--rival", default="TVONE")
    ap.add_argument("--event-name", default=None,
                     help='e.g. "HUT RI ke-81" -- used only for your own slide titles, not required by this script')
    ap.add_argument("--include-breaking-news", action="store_true",
                     help="For every Daypart that already has a Level 1=='Special' program, also pull in "
                          "each channel's program whose name contains 'breaking news' in that same Daypart "
                          "Custome bucket, so the special-program comparison isn't missing an unflagged but "
                          "clearly event-relevant bulletin. This was a real, user-requested fix in this deck's "
                          "history -- see SKILL.md.")
    ap.add_argument("--out", default="data.json")
    args = ap.parse_args()

    target = pd.to_datetime(args.date, format="%d/%m/%Y", errors="coerce")
    if pd.isna(target):
        target = pd.to_datetime(args.date, errors="coerce")
    if pd.isna(target):
        sys.exit(f"Could not parse --date {args.date!r}; use DD/MM/YYYY")

    MAIN, RIVAL = args.main.upper(), args.rival.upper()

    xl = pd.ExcelFile(args.raw)

    # -- PERFORMANCE DAILY: national share -----------------------------------
    daily = pd.read_excel(args.raw, sheet_name="PERFORMANCE DAILY")
    daily["DateParsed"] = parse_date_col(daily)
    drow = daily[daily["DateParsed"] == target]
    if drow.empty:
        sys.exit(f"{args.date} not found in PERFORMANCE DAILY -- check the date or that RAW is up to date.")
    drow = drow.iloc[0]
    week_num = int(drow["Week"])
    week_label = f"{str(target.year)[2:]}{week_num:02d}"
    official_share = {ch: round(float(drow[ch]), 2) for ch in CHANNELS}

    # -- BY DAYPART DAILY: per-Daypart share/TVR/viewers ---------------------
    dpd = pd.read_excel(args.raw, sheet_name="BY DAYPART DAILY")
    dpd["DateParsed"] = parse_date_col(dpd)
    dsub = dpd[dpd["DateParsed"] == target]
    if dsub.empty:
        sys.exit(f"{args.date} not found in BY DAYPART DAILY.")
    summary = {}
    for dp in DAYPART_ORDER:
        row = {}
        for ch in CHANNELS:
            r = dsub[(dsub["Day Part"] == dp) & (dsub["Channel"] == ch)]
            if r.empty:
                row[ch] = {"share": None, "viewers": None, "tvr": None}
            else:
                r = r.iloc[0]
                row[ch] = {"share": round(float(r["Share"]), 2),
                            "viewers": float(r["000s"]) if "000s" in r else None,
                            "tvr": round(float(r["TVR"]), 2)}
        summary[dp] = row

    winners, tally = {}, {}
    for dp in DAYPART_ORDER:
        i_share = summary[dp][MAIN]["share"]
        r_share = summary[dp][RIVAL]["share"]
        winner = MAIN if i_share >= r_share else RIVAL
        winners[dp] = winner
        tally[winner] = tally.get(winner, 0) + 1

    # -- Kontribusi By Daypart: contribution of each Daypart to Total TV -----
    kontribusi_daypart = None
    if "Kontribusi By Daypart" in xl.sheet_names:
        ktb = pd.read_excel(args.raw, sheet_name="Kontribusi By Daypart")
        ktb["DateParsed"] = parse_date_col(ktb)
        ksub = ktb[(ktb["DateParsed"] == target) & (ktb["Channel"] == "Total TV")]
        if not ksub.empty:
            row = ksub.iloc[0]
            total = sum(float(row[dp]) for dp in DAYPART_ORDER)
            kontribusi_daypart = {}
            for dp in DAYPART_ORDER:
                val = float(row[dp])
                pct = val / total * 100
                kontribusi_daypart[dp] = {
                    "total_tv": val, "pct": round(pct, 2), "pct_round": round(pct),
                }

    # -- Program Detail REV2: program-level detail for this date -------------
    rev2 = pd.read_excel(args.raw, sheet_name="Program Detail REV2")
    rev2["DateParsed"] = parse_date_col(rev2)
    psub = rev2[(rev2["DateParsed"] == target) & (rev2["Dur"] > 0)].copy()
    if psub.empty:
        sys.exit(f"{args.date} not found in Program Detail REV2 (or all Dur==0).")

    def row_dict(r, extra=None):
        d = {
            "Channel": r["Channel"],
            "ProgramTitle": title_program(r["Program"]),
            "Daypart Custome": canon_daypart_custome(r["Daypart Custome"]),
            "000s": int(round(float(r["000s"]))),
            "TVR": round(float(r["TVR"]), 2),
            "Share": round(float(r["Share"]), 2),
        }
        if extra:
            d.update(extra)
        return d

    # Top 25 -- exclude Filler, sort by 000s desc
    top_pool = psub[psub["Level 1"] != "Filler"].sort_values("000s", ascending=False)
    top25 = [row_dict(r) for _, r in top_pool.head(25).iterrows()]
    inews_count_top25 = sum(1 for p in top25 if p["Channel"] == MAIN)

    top_per_channel = {}
    for ch in CHANNELS:
        chrows = psub[psub["Channel"] == ch].sort_values("000s", ascending=False)
        if not chrows.empty:
            r = chrows.iloc[0]
            top_per_channel[ch] = {
                "program": title_program(r["Program"]), "viewers": int(round(float(r["000s"]))),
                "tvr": round(float(r["TVR"]), 2), "share": round(float(r["Share"]), 2),
                "daypart": canon_daypart_custome(r["Daypart Custome"]),
            }

    # Special programs (Level 1 == 'Special')
    spec_rows = psub[psub["Level 1"] == "Special"].sort_values(["Channel", "Start time"])
    special = [row_dict(r, {"Start time": r["Start time"], "End time": r["End time"], "Dur": int(r["Dur"])})
               for _, r in spec_rows.iterrows()]

    special_agg = {}
    for ch in CHANNELS:
        chrows = [s for s in special if s["Channel"] == ch]
        if chrows:
            special_agg[ch] = {
                "n_programs": len(chrows),
                "total_viewers": sum(s["000s"] for s in chrows),
                "avg_share": round(sum(s["Share"] for s in chrows) / len(chrows), 2),
                "total_dur": sum(s["Dur"] for s in chrows),
            }

    # canon_daypart_custome() has already normalised every row's "Daypart Custome"
    # to the English 8-bin vocabulary, so group by that directly -- no need to
    # re-normalise the source DataFrame column here for the exact-match grouping.
    psub_canon_dp = psub["Daypart Custome"].apply(canon_daypart_custome)

    special_by_daypart = {}
    for dp in DAYPART_CUSTOME_CANON:
        grp = [s for s in special if s["Daypart Custome"] == dp]
        if args.include_breaking_news and grp:
            have = {s["Channel"] for s in grp}
            bn_rows = psub[(psub_canon_dp == dp) &
                           (psub["Program"].str.contains("breaking news", case=False, na=False))]
            for _, r in bn_rows.iterrows():
                if r["Channel"] not in have:
                    grp.append(row_dict(r, {"Start time": r["Start time"], "End time": r["End time"], "Dur": int(r["Dur"])}))
                    have.add(r["Channel"])
        if grp:
            grp.sort(key=lambda s: s["000s"], reverse=True)
            special_by_daypart[dp] = [
                {k: v for k, v in s.items() if k != "Daypart Custome" and k != "Dur"} for s in grp
            ]

    # Fresh vs Rerun -- total minutes for the whole day, ALL Level 1 types
    # (Filler segments are themselves tagged Fresh/Rerun and belong in the total).
    #
    # KNOWN RAW-DATA GAP (found on 21/03/2026, Idul Fitri): holiday/event
    # special-programming blocks are sometimes left tagged Category=="Rerun"
    # even on their first-ever airing. Evidence: every OTHER Rerun row that
    # day carries an explicit repeat marker in the Program name -- "(R)"
    # (INEWS, KOMPASTV), "(F0#)" (TVONE), or a trailing "L#" loop number
    # (METRO) -- but a handful of Lebaran-special blocks (e.g. INEWS's
    # "INEWS PAGI SPESIAL LEBARAN"/"INEWS SIANG SPESIAL LEBARAN", TVONE's
    # "APA KABAR INDONESIA PAGI/MALAM SPS LEBARAN", KOMPASTV's "SAPA
    # INDONESIA MALAM SPS LEBARAN") carry no such marker and never appear as
    # "Fresh" on any other date either -- there's no earlier airing for them
    # to legitimately be a repeat OF. That combination (Rerun-tagged, no
    # repeat marker, name signals a one-off holiday special) is a strong
    # signal of a raw Category tagging gap, not genuinely low fresh output.
    # Reclassify those specific rows to Fresh so the Fresh/Rerun comparison
    # isn't skewed against whichever channel's holiday special happens to be
    # mistagged. This only ever flips Rerun -> Fresh, never the reverse, and
    # only for rows matching BOTH conditions below -- it will simply find
    # nothing to reclassify (0-minute effect) on a normal non-event day.
    import re as _re

    def _has_repeat_marker(name):
        return bool(_re.search(r"\(R\)|\(F\d|L\d\b", str(name).upper()))

    def _is_holiday_special(name):
        u = str(name).upper()
        return "LEBARAN" in u or "SPESIAL" in u or " SPS " in (" " + u + " ")

    fresh_rerun = {}
    fresh_rerun_reclass_notes = {}
    for ch in CHANNELS:
        s = psub[psub["Channel"] == ch]
        total = float(s["Dur"].sum())
        fresh = float(s[s["Category"] == "Fresh"]["Dur"].sum())
        rerun_mask = s["Category"] == "Rerun"
        reclass_mask = rerun_mask & s["Program"].apply(_is_holiday_special) & (~s["Program"].apply(_has_repeat_marker))
        reclass_min = float(s[reclass_mask]["Dur"].sum())
        fresh += reclass_min
        rerun = float(s[rerun_mask]["Dur"].sum()) - reclass_min
        fresh_rerun[ch] = {
            "fresh_pct": round(fresh / total * 100, 2) if total else None,
            "rerun_pct": round(rerun / total * 100, 2) if total else None,
            "fresh_hours": round(fresh / 60, 2),
            "rerun_hours": round(rerun / 60, 2),
        }
        if reclass_min:
            fresh_rerun_reclass_notes[ch] = {
                "reclassified_minutes": int(reclass_min),
                "programs": sorted(set(s[reclass_mask]["Program"].astype(str))),
            }
    if fresh_rerun_reclass_notes:
        import sys as _sys
        print(
            "NOTE: reclassified Rerun->Fresh (unmarked holiday-special blocks): "
            + json.dumps(fresh_rerun_reclass_notes),
            file=_sys.stderr,
        )

    data = {
        "target_date": fmt_id_date(target),
        "week_num": week_num,
        "week_label": week_label,
        "main_channel": MAIN,
        "rival_channel": RIVAL,
        "official_share": official_share,
        "daypart_order": DAYPART_ORDER,
        "daypart_name": DAYPART_NAME_ID,
        "summary": summary,
        "winners": winners,
        "tally": tally,
        "kontribusi_daypart": kontribusi_daypart,
        "top25": top25,
        "inews_count_top25": inews_count_top25,
        "top_per_channel": top_per_channel,
        "special": special,
        "special_agg": special_agg,
        "special_by_daypart": special_by_daypart,
        "fresh_rerun": fresh_rerun,
    }

    with open(args.out, "w") as f:
        json.dump(data, f, indent=2)

    print(f"Wrote {args.out}")
    print(f"  target_date={data['target_date']}  week={week_label}")
    print(f"  official_share={official_share}")
    print(f"  daypart tally={tally}")
    print(f"  top25: {MAIN} in {inews_count_top25}/25")
    print(f"  special_by_daypart groups: {list(special_by_daypart.keys())}")
    if kontribusi_daypart is None:
        print("  NOTE: 'Kontribusi By Daypart' sheet not found in this RAW file -- "
              "kontribusi_daypart is null. The deck script falls back to the old "
              "'bervolume besar' wording when this happens (see SKILL.md).")


if __name__ == "__main__":
    main()
```


## ================================================================
## FILE: inews_competitor_daypart_deck_build_deck.js
## ================================================================

```js
// Head-to-Head Daypart competitive analysis deck (7 slides), pptxgenjs.
// Reads data.json produced by prepare_data.py. See SKILL.md for the full
// per-run checklist -- this file's CONFIG block is the only thing you
// normally need to touch between runs.

const pptxgen = require("pptxgenjs");
const fs = require("fs");

// ============================================================================
// CONFIG -- edit these per run
// ============================================================================
const CONFIG = {
  DATA_PATH: "./data.json",
  OUT: "./Analisa Performa.pptx",
  // EVENT_NAME drives the cover badge and the special-program slide title.
  // Set to null to skip both (and the special-program slide entirely) for a
  // routine, non-event comparison.
  EVENT_NAME: null, // e.g. "HUT RI ke-81"
};

const DATA = JSON.parse(fs.readFileSync(CONFIG.DATA_PATH, "utf8"));
const MAIN_CHANNEL = DATA.main_channel;
const RIVAL_CHANNEL = DATA.rival_channel;
const OUT = CONFIG.OUT;

// ---------------------------------------------------------------------------
// Fixed project conventions (see SKILL.md) -- colors, footer wording, fonts
// ---------------------------------------------------------------------------
const COLOR = {
  INEWS: "C41E3A",
  TVONE: "1976D2",
  KOMPASTV: "00BCD4",
  METRO: "FF9800",
};
const CHANNELS = ["INEWS", "TVONE", "KOMPASTV", "METRO"];
const DARK = "1E2761";
const BLUE = "1976D2";
const ICE = "CADCFC";
const WHITE = "FFFFFF";
const TEXT_DARK = "222222";
const MUTED = "6B7280";
const HILITE = "FDECEA"; // MAIN_CHANNEL row highlight, all tables

const FOOTER = `Sumber: Nielsen Media Research; National Urban; Daily Performance; ${DATA.target_date}`;

const DAYPART_EN = {
  "02:00 - 04:29": "Early Morning",
  "04:30 - 09:59": "Morning",
  "10:00 - 12:59": "Noon 1",
  "13:00 - 15:59": "Noon 2",
  "16:00 - 17:59": "Evening",
  "18:00 - 21:59": "Prime Time",
  "22:00 - 23:59": "Night",
  "24:00 - 25:59": "Midnight",
};

const KTB = DATA.kontribusi_daypart; // may be null -- see prepare_data.py note
function kpct(dp) {
  return KTB && KTB[dp] ? KTB[dp].pct_round : null;
}
function fmtNum(n) {
  return Math.round(n).toLocaleString("en-US");
}

// ---------------------------------------------------------------------------
// Derive the small set of narrative facts every insight bar / the exec
// summary needs, straight from DATA -- so this script runs end-to-end without
// hand-editing sentences every time (only the CONFIG block above changes).
// ---------------------------------------------------------------------------
const nationalGap = +(DATA.official_share[MAIN_CHANNEL] - DATA.official_share[RIVAL_CHANNEL]).toFixed(2);
const dailyWinner = nationalGap >= 0 ? MAIN_CHANNEL : RIVAL_CHANNEL;
const otherChannel = dailyWinner === MAIN_CHANNEL ? RIVAL_CHANNEL : MAIN_CHANNEL;
const rivalWinDayparts = DATA.daypart_order.filter((dp) => DATA.winners[dp] === RIVAL_CHANNEL);
const mainWinDayparts = DATA.daypart_order.filter((dp) => DATA.winners[dp] === MAIN_CHANNEL);
// The daily WINNER's own winning Dayparts -- this is what causally explains
// the national daily result, and must be recomputed from `dailyWinner`
// rather than hardcoded to either channel: MAIN doesn't always win, RIVAL
// doesn't always win either.
const winnerWinDayparts = dailyWinner === MAIN_CHANNEL ? mainWinDayparts : rivalWinDayparts;

// "Berkontribusi besar" dayparts = the daily winner's winning dayparts,
// described with their real Kontribusi % when the sheet is available; falls
// back to a plain daypart-name list (no percentages, no "kontribusi" word)
// when it isn't -- see prepare_data.py's note on "Kontribusi By Daypart"
// sometimes being absent from older RAW files.
function daypartListSentence(dayparts, withPct) {
  return dayparts
    .map((dp) => (withPct && kpct(dp) != null ? `${DAYPART_EN[dp]} ${kpct(dp)}%` : DAYPART_EN[dp]))
    .join(", ");
}
const HAS_KONTRIBUSI = !!KTB;
const rivalWinListStr = daypartListSentence(rivalWinDayparts, HAS_KONTRIBUSI);
const winnerWinListStr = daypartListSentence(winnerWinDayparts, HAS_KONTRIBUSI);
const mainWinListStr = daypartListSentence(mainWinDayparts, HAS_KONTRIBUSI);
const winReasonWord = HAS_KONTRIBUSI ? "berkontribusi besar" : "bervolume besar";

let bestEfficiencyChannel = null, bestEfficiencyShare = -1;
let mostAggressiveChannel = null, mostAggressiveN = -1;
Object.entries(DATA.special_agg || {}).forEach(([ch, agg]) => {
  if (agg.avg_share > bestEfficiencyShare) { bestEfficiencyShare = agg.avg_share; bestEfficiencyChannel = ch; }
  if (agg.n_programs > mostAggressiveN) { mostAggressiveN = agg.n_programs; mostAggressiveChannel = ch; }
});

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "RDO";
pres.title = `Analisa Performa ${MAIN_CHANNEL} vs Kompetitor - ${DATA.target_date}`;

const PW = 13.33, PH = 7.5;
let pageNum = 0;

function addFooter(slide, label) {
  slide.addText(FOOTER, {
    x: 0.5, y: PH - 0.38, w: PW - 1.0, h: 0.3,
    fontFace: "Calibri", fontSize: 9, color: MUTED, align: "left",
  });
  if (label) {
    slide.addText(label, {
      x: PW - 3.2, y: PH - 0.38, w: 2.7, h: 0.3,
      fontFace: "Calibri", fontSize: 9, color: MUTED, align: "right",
    });
  }
}

function addSlideTitle(slide, title, subtitle) {
  slide.addText(title, {
    x: 0.5, y: 0.3, w: PW - 1.0, h: 0.6,
    fontFace: "Cambria", fontSize: 26, bold: true, color: DARK,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.5, y: 0.88, w: PW - 1.0, h: 0.4,
      fontFace: "Calibri", fontSize: 13, bold: true, color: TEXT_DARK,
    });
  }
}

function addInsightBar(slide, x, y, w, h, richText) {
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.05, fill: { color: "F4F5F9" }, line: { color: "DDDDE3", width: 0.75 } });
  slide.addShape("rect", { x, y, w: 0.07, h, fill: { color: DARK }, line: { type: "none" } });
  slide.addText(richText, { x: x + 0.28, y, w: w - 0.5, h, valign: "middle", fontFace: "Calibri", fontSize: 12, color: TEXT_DARK });
}

// ---------------------------------------------------------------------------
// SLIDE 1: COVER
// ---------------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: DARK };
  s.addShape("rect", { x: 0, y: 0, w: PW, h: PH, fill: { color: DARK } });
  if (CONFIG.EVENT_NAME) {
    s.addText(CONFIG.EVENT_NAME.toUpperCase(), {
      x: 0.9, y: 1.5, w: 8, h: 0.5,
      fontFace: "Calibri", fontSize: 14, bold: true, color: "FFC107", charSpacing: 2,
    });
  }
  s.addText(`Analisa Performa ${MAIN_CHANNEL}\nvs Kompetitor`, {
    x: 0.85, y: 2.0, w: 10.5, h: 2.0,
    fontFace: "Cambria", fontSize: 44, bold: true, color: WHITE, lineSpacingMultiple: 1.05,
  });
  s.addText(`${DATA.target_date}  •  National Urban, 20+ UM Terrestrial  •  Sumber Daypart: BY DAYPART DAILY`, {
    x: 0.9, y: 4.05, w: 11, h: 0.5, fontFace: "Calibri", fontSize: 15, color: ICE,
  });

  const cardW = 2.6, gap = 0.25, startX = 0.9;
  CHANNELS.forEach((ch, i) => {
    const x = startX + i * (cardW + gap);
    s.addShape("roundRect", { x, y: 5.05, w: cardW, h: 1.55, rectRadius: 0.08, fill: { color: "2A3470" }, line: { type: "none" } });
    s.addShape("rect", { x, y: 5.05, w: 0.08, h: 1.55, fill: { color: COLOR[ch] }, line: { type: "none" } });
    s.addText(ch, { x: x + 0.2, y: 5.15, w: cardW - 0.3, h: 0.35, fontFace: "Calibri", fontSize: 13, bold: true, color: ICE });
    s.addText(DATA.official_share[ch].toFixed(2) + "%", { x: x + 0.2, y: 5.45, w: cardW - 0.3, h: 0.9, fontFace: "Calibri", fontSize: 32, bold: true, color: WHITE });
  });

  s.addText(FOOTER, { x: 0.9, y: 7.05, w: 11, h: 0.3, fontFace: "Calibri", fontSize: 10, color: "8892C8" });
}

// ---------------------------------------------------------------------------
// SLIDE 2: DAYPART SHARE COMPARISON
// Category labels: English Daypart name + time range + (if available) the
// Kontribusi % as a bare number under it -- no "Kontribusi" word in the
// chart itself, per explicit feedback: it reads as clutter once the concept
// is already introduced in the insight bar above.
// ---------------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "Perbandingan Share per Daypart", `Share (%) per Daypart (BY DAYPART DAILY), ${MAIN_CHANNEL} vs Kompetitor — ${DATA.target_date}`);

  addInsightBar(s, 0.5, 1.32, 12.35, 0.55, [
    { text: `${RIVAL_CHANNEL} `, options: { bold: true, color: COLOR[RIVAL_CHANNEL] || TEXT_DARK } },
    { text: `unggul di ${rivalWinDayparts.length} Daypart ${winReasonWord} (${rivalWinListStr}); `, options: {} },
    { text: `${MAIN_CHANNEL} `, options: { bold: true, color: COLOR[MAIN_CHANNEL] || TEXT_DARK } },
    { text: `unggul di Daypart lain (${daypartListSentence(mainWinDayparts, HAS_KONTRIBUSI)}).`, options: {} },
  ]);

  const cats = DATA.daypart_order.map((dp) => {
    const line3 = kpct(dp) != null ? `\n${kpct(dp)}%` : "";
    return `${DAYPART_EN[dp]}\n(${dp.replace(/ /g, "")})${line3}`;
  });
  const chartData = CHANNELS.map((ch) => ({
    name: ch,
    labels: cats,
    values: DATA.daypart_order.map((dp) => DATA.summary[dp][ch].share),
  }));

  s.addChart("bar", chartData, {
    x: 0.4, y: 2.05, w: 12.55, h: 4.9,
    barDir: "col",
    barGapWidthPct: 25,
    chartColors: CHANNELS.map((ch) => COLOR[ch]),
    showTitle: false,
    showLegend: true,
    legendPos: "b",
    legendFontSize: 12,
    legendFontFace: "Calibri",
    legendFontBold: true,
    catAxisLabelFontSize: 12,
    catAxisLabelFontFace: "Calibri",
    catAxisLabelFontBold: true,
    catAxisLabelColor: TEXT_DARK,
    valAxisTitle: "Share (%)",
    showValAxisTitle: true,
    valAxisTitleFontSize: 11,
    valAxisLabelFontSize: 10,
    valAxisLabelFormatCode: "0.0",
    valGridLine: { color: "E5E5E5", size: 0.75 },
    catGridLine: { style: "none" },
    showValue: true,
    dataLabelPosition: "outEnd",
    dataLabelFontSize: 8.5,
    dataLabelFontBold: true,
    dataLabelFormatCode: "0.00",
    dataLabelColor: TEXT_DARK,
  });

  pageNum++;
  addFooter(s, "Halaman " + pageNum);
}

// ---------------------------------------------------------------------------
// SLIDE 3: HEAD-TO-HEAD TABLE + tally + Kontribusi column
// ---------------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, `Head To Head: ${MAIN_CHANNEL} vs ${RIVAL_CHANNEL} By Daypart`, `Rival utama nasional — sumber BY DAYPART DAILY`);

  addInsightBar(s, 0.5, 1.32, 12.35, 0.55, [
    { text: `${MAIN_CHANNEL} dan ${RIVAL_CHANNEL} ${mainWinDayparts.length}-${rivalWinDayparts.length} dari ${DATA.daypart_order.length} Daypart — `, options: {} },
    { text: `${dailyWinner} unggul tipis secara harian (${DATA.official_share[dailyWinner].toFixed(2)}% vs ${DATA.official_share[otherChannel].toFixed(2)}%) `, options: {} },
    { text: `karena menang di Daypart ${winReasonWord} (${winnerWinListStr}).`, options: { bold: true } },
  ]);

  const showKontribusi = HAS_KONTRIBUSI;
  const headerLabels = ["Daypart", "Jam"].concat(showKontribusi ? ["Kontribusi"] : [], [`${MAIN_CHANNEL} (%)`, `${RIVAL_CHANNEL} (%)`, "Selisih (pp)", "Unggul"]);
  const headerRow = headerLabels.map((t) => ({
    text: t, options: { bold: true, color: WHITE, fill: { color: DARK }, fontSize: 11, align: "center", valign: "middle" },
  }));
  const rows = [headerRow];
  DATA.daypart_order.forEach((dp) => {
    const i = DATA.summary[dp][MAIN_CHANNEL].share;
    const t = DATA.summary[dp][RIVAL_CHANNEL].share;
    const gap = +(i - t).toFixed(2);
    const winner = DATA.winners[dp];
    const isTie = Math.abs(gap) < 0.005;
    const winnerLabel = isTie ? "Seri" : winner;
    const winnerColor = isTie ? "9E9E9E" : winner === MAIN_CHANNEL ? COLOR[MAIN_CHANNEL] : winner === RIVAL_CHANNEL ? COLOR[RIVAL_CHANNEL] : "757575";
    const row = [
      { text: DAYPART_EN[dp], options: { bold: true, color: TEXT_DARK, fontSize: 11, align: "left" } },
      { text: dp, options: { color: MUTED, fontSize: 9.5, align: "center" } },
    ];
    if (showKontribusi) row.push({ text: kpct(dp) + "%", options: { color: MUTED, bold: true, fontSize: 10.5, align: "center" } });
    row.push(
      { text: i.toFixed(2), options: { color: COLOR[MAIN_CHANNEL], bold: true, fontSize: 11, align: "center" } },
      { text: t.toFixed(2), options: { color: COLOR[RIVAL_CHANNEL], bold: true, fontSize: 11, align: "center" } },
      { text: (gap >= 0 ? "+" : "") + gap.toFixed(2), options: { color: gap > 0 ? "2E7D32" : gap < 0 ? "C41E3A" : "757575", bold: true, fontSize: 11, align: "center" } },
      { text: winnerLabel, options: { color: WHITE, fill: { color: winnerColor }, bold: true, fontSize: 10, align: "center", valign: "middle" } },
    );
    rows.push(row);
  });

  const tableW = showKontribusi ? 9.13 : 8.6;
  const colW = showKontribusi ? [1.55, 1.3, 1.05, 1.2, 1.2, 1.63, 1.2] : [1.85, 1.55, 1.4, 1.4, 1.2, 1.2];
  s.addTable(rows, {
    x: 0.5, y: 2.05, w: tableW, h: 4.6,
    colW,
    fontFace: "Calibri",
    border: { type: "solid", color: "E0E0E0", pt: 0.75 },
    autoPage: false,
    rowH: 0.55,
    valign: "middle",
  });

  const tx = 0.5 + tableW + 0.2, tw = PW - 0.5 - tx;
  s.addShape("roundRect", { x: tx, y: 2.05, w: tw, h: 4.6, rectRadius: 0.08, fill: { color: "F4F5F9" }, line: { color: "DDDDE3", width: 0.75 } });
  s.addText("Skor Kemenangan Daypart", { x: tx + 0.2, y: 2.25, w: tw - 0.4, h: 0.4, fontFace: "Calibri", fontSize: 13, bold: true, color: DARK });
  const tallyList = CHANNELS.map((ch) => ({ ch, n: DATA.tally[ch] || 0 }));
  let ty = 2.8;
  tallyList.forEach((t) => {
    s.addShape("rect", { x: tx + 0.2, y: ty, w: 0.12, h: 0.5, fill: { color: COLOR[t.ch] }, line: { type: "none" } });
    s.addText(t.ch, { x: tx + 0.42, y: ty, w: 1.3, h: 0.5, fontFace: "Calibri", fontSize: 12, bold: true, color: TEXT_DARK, valign: "middle" });
    s.addText(String(t.n) + " daypart", { x: tx + 1.55, y: ty, w: tw - 1.75, h: 0.5, fontFace: "Calibri", fontSize: 12.5, bold: true, color: COLOR[t.ch], align: "right", valign: "middle" });
    ty += 0.66;
  });
  s.addShape("roundRect", { x: tx + 0.2, y: ty + 0.15, w: tw - 0.4, h: 1.55, rectRadius: 0.05, fill: { color: "FFF3CD" }, line: { type: "none" } });
  s.addText(
    `${MAIN_CHANNEL} & ${RIVAL_CHANNEL} ${mainWinDayparts.length}-${rivalWinDayparts.length} dari ${DATA.daypart_order.length} Daypart. Kemenangan harian ${dailyWinner} ditentukan oleh Daypart ${winReasonWord} (${winnerWinListStr}), bukan oleh jumlah Daypart yang dimenangkan.`,
    { x: tx + 0.35, y: ty + 0.27, w: tw - 0.7, h: 1.3, fontFace: "Calibri", fontSize: 9.5, color: "6B5900", valign: "top" }
  );

  pageNum++;
  addFooter(s, "Halaman " + pageNum);
}

// ---------------------------------------------------------------------------
// SLIDE 4: TOP 25 PROGRAM (excl. Filler) — 000s / TVR / Share, sorted by
// 000s, MAIN_CHANNEL rows highlighted.
// ---------------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "Top 25 Program Hari Ini", `Berdasarkan Penonton (000s), Filler dikecualikan — ${DATA.target_date}`);

  s.addShape("roundRect", { x: 0.5, y: 1.35, w: 12.35, h: 0.55, rectRadius: 0.05, fill: { color: "FDECEA" }, line: { color: COLOR[MAIN_CHANNEL] || DARK, width: 1 } });
  s.addText([
    { text: `${MAIN_CHANNEL} `, options: { bold: true, color: COLOR[MAIN_CHANNEL] || DARK, fontSize: 14 } },
    { text: `masuk ${DATA.inews_count_top25} dari ${DATA.top25.length} program teratas hari ini.`, options: { color: TEXT_DARK, fontSize: 13 } },
  ], { x: 0.7, y: 1.35, w: 12, h: 0.55, fontFace: "Calibri", valign: "middle" });

  function buildTable(rowsData, startRank) {
    const headerRow = ["#", "Program", "Ch", "000s", "TVR", "Sh%"].map((t) => ({
      text: t, options: { bold: true, color: WHITE, fill: { color: DARK }, fontSize: 10, align: "center", valign: "middle" },
    }));
    const rows = [headerRow];
    rowsData.forEach((r, idx) => {
      const isMain = r.Channel === MAIN_CHANNEL;
      const fill = isMain ? { color: HILITE } : undefined;
      rows.push([
        { text: String(startRank + idx), options: { bold: true, align: "center", fontSize: 9.5, fill } },
        { text: r.ProgramTitle, options: { align: "left", fontSize: 9, bold: isMain, fill } },
        { text: r.Channel, options: { bold: true, color: COLOR[r.Channel] || TEXT_DARK, align: "center", fontSize: 8.5, fill } },
        { text: fmtNum(r["000s"]), options: { align: "center", fontSize: 9, fill } },
        { text: r.TVR.toFixed(2), options: { align: "center", fontSize: 9, fill } },
        { text: r.Share.toFixed(2), options: { align: "center", fontSize: 9, bold: true, fill } },
      ]);
    });
    return rows;
  }

  const half = Math.ceil(DATA.top25.length / 2);
  const left = DATA.top25.slice(0, half);
  const right = DATA.top25.slice(half);

  s.addTable(buildTable(left, 1), {
    x: 0.5, y: 2.05, w: 6.15, h: 4.9,
    colW: [0.35, 2.85, 0.55, 0.85, 0.65, 0.7],
    fontFace: "Calibri",
    border: { type: "solid", color: "E0E0E0", pt: 0.75 },
    autoPage: false,
    rowH: 4.9 / (half + 1),
    margin: [1, 3, 1, 3],
    valign: "middle",
  });

  s.addTable(buildTable(right, half + 1), {
    x: 6.85, y: 2.05, w: 5.95, h: 4.9,
    colW: [0.35, 3.0, 0.6, 0.9, 0.65, 0.65],
    fontFace: "Calibri",
    border: { type: "solid", color: "E0E0E0", pt: 0.75 },
    autoPage: false,
    rowH: 4.9 / (right.length + 1),
    margin: [1, 3, 1, 3],
    valign: "middle",
  });

  pageNum++;
  addFooter(s, "Halaman " + pageNum);
}

// ---------------------------------------------------------------------------
// SLIDE 5: SPECIAL PROGRAMS, grouped per Daypart -- only rendered when
// CONFIG.EVENT_NAME is set AND prepare_data.py found Level 1=='Special' rows.
// Layout adapts to how many Daypart groups exist (1-4): a single row of up
// to 3 cards, or a 2x2 grid for 4.
// ---------------------------------------------------------------------------
const specialGroups = Object.keys(DATA.special_by_daypart || {});
if (CONFIG.EVENT_NAME && specialGroups.length > 0) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, `Program Spesial ${CONFIG.EVENT_NAME} per Daypart`, `Perbandingan slot & performa siaran spesial antar channel, per Daypart — ${DATA.target_date}`);

  if (mostAggressiveChannel && bestEfficiencyChannel) {
    s.addText([
      { text: "Volume: ", options: { bold: true, color: TEXT_DARK } },
      { text: `${mostAggressiveChannel} paling agresif (${mostAggressiveN} program, ${fmtNum(DATA.special_agg[mostAggressiveChannel].total_viewers)} ribu total penonton).  `, options: { color: TEXT_DARK } },
      { text: "Efisiensi: ", options: { bold: true, color: TEXT_DARK } },
      { text: `${bestEfficiencyChannel} paling tinggi rata-rata share per program spesial (${bestEfficiencyShare.toFixed(2)}%).`, options: { color: TEXT_DARK } },
    ], { x: 0.5, y: 1.35, w: 12.35, h: 0.35, fontFace: "Calibri", fontSize: 11.5, italic: true });
  }

  const nGroups = specialGroups.length;
  const cols = nGroups >= 4 ? 2 : nGroups;
  const rowsN = Math.ceil(nGroups / cols);
  const gx0 = 0.5, gy0 = 1.85, ggap = 0.25;
  const availW = PW - 1.0, availH = 6.35 - gy0;
  const gw = (availW - (cols - 1) * ggap) / cols;
  const gh = rowsN === 1 ? availH : (availH - (rowsN - 1) * 0.2) / rowsN;

  specialGroups.forEach((dp, gi) => {
    const col = gi % cols, row = Math.floor(gi / cols);
    const pos = { x: gx0 + col * (gw + ggap), y: gy0 + row * (gh + 0.2) };
    const rowsData = DATA.special_by_daypart[dp];

    s.addShape("roundRect", { x: pos.x, y: pos.y, w: gw, h: gh, rectRadius: 0.05, fill: { color: "FFFFFF" }, line: { color: "DDDDE3", width: 1 } });
    s.addShape("rect", { x: pos.x, y: pos.y, w: gw, h: 0.5, fill: { color: DARK }, line: { type: "none" } });
    s.addText(DAYPART_EN[Object.keys(DAYPART_EN).find((k) => DAYPART_EN[k] === dp)] || dp, { x: pos.x + 0.15, y: pos.y, w: gw - 0.3, h: 0.5, fontFace: "Calibri", fontSize: 11, bold: true, color: WHITE, valign: "middle" });

    const headerRow = ["Program", "Ch", "Jam", "000s", "TVR", "Sh%"].map((t) => ({
      text: t, options: { bold: true, color: WHITE, fill: { color: BLUE }, fontSize: 8, align: "center", valign: "middle" },
    }));
    const rows = [headerRow];
    rowsData.forEach((r) => {
      const isMain = r.Channel === MAIN_CHANNEL;
      const fill = isMain ? { color: HILITE } : undefined;
      rows.push([
        { text: r.ProgramTitle, options: { align: "left", fontSize: 8, bold: isMain, fill } },
        { text: r.Channel, options: { bold: true, color: COLOR[r.Channel] || TEXT_DARK, align: "center", fontSize: 7.5, fill } },
        { text: r["Start time"] + "-" + r["End time"], options: { align: "center", fontSize: 7, color: MUTED, fill } },
        { text: fmtNum(r["000s"]), options: { align: "center", fontSize: 7.5, fill } },
        { text: r.TVR.toFixed(2), options: { align: "center", fontSize: 7.5, fill } },
        { text: r.Share.toFixed(2), options: { align: "center", fontSize: 7.5, bold: true, fill } },
      ]);
    });
    s.addTable(rows, {
      x: pos.x + 0.05, y: pos.y + 0.58, w: gw - 0.1, h: gh - 0.68,
      colW: [(gw - 0.1) * 0.37, (gw - 0.1) * 0.10, (gw - 0.1) * 0.19, (gw - 0.1) * 0.13, (gw - 0.1) * 0.10, (gw - 0.1) * 0.11],
      fontFace: "Calibri",
      border: { type: "solid", color: "E8E8E8", pt: 0.5 },
      autoPage: false,
      rowH: (gh - 0.68) / (rowsData.length + 1),
      margin: [1, 2, 1, 2],
      valign: "middle",
    });
  });

  pageNum++;
  addFooter(s, "Halaman " + pageNum);
}

// ---------------------------------------------------------------------------
// SLIDE 6: FRESH VS RERUN -- vertical stacked bar, FIXED channel order
// (INEWS, TVONE, KOMPASTV, METRO -- same as every other competitor
// comparison in this deck; never alphabetical -- see GLOBAL_Competitor_
// Channel_Order_Rule.md, this slide used to violate it via `.sort()`),
// blue = Fresh / orange = Rerun, total daily hours (not per-program
// average -- an explicit correction requested in this deck's history),
// rounded to whole hours with NO "rounded" caveat text anywhere on the
// slide -- an explicit user request (2026-08-26).
// ---------------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "Komposisi Konten: Fresh vs Rerun", `Persentase & total durasi siaran per channel, sepanjang hari — ${DATA.target_date}`);

  const fresh = CHANNELS.map((ch) => DATA.fresh_rerun[ch].fresh_pct);
  const rerun = CHANNELS.map((ch) => DATA.fresh_rerun[ch].rerun_pct);

  const chartData = [
    { name: "Fresh", labels: CHANNELS, values: fresh },
    { name: "Rerun", labels: CHANNELS, values: rerun },
  ];

  s.addChart("bar", chartData, {
    x: 0.6, y: 1.55, w: 8.1, h: 4.9,
    barDir: "col",
    barGrouping: "stacked",
    barGapWidthPct: 45,
    chartColors: ["1976D2", "FB8C00"],
    showTitle: false,
    showLegend: true,
    legendPos: "t",
    legendFontSize: 12,
    legendFontFace: "Calibri",
    legendFontBold: true,
    catAxisLabelFontSize: 13,
    catAxisLabelFontFace: "Calibri",
    catAxisLabelColor: TEXT_DARK,
    catAxisLabelFontBold: true,
    valAxisLabelFontSize: 10,
    valAxisLabelFormatCode: "0",
    valGridLine: { color: "E5E5E5", size: 0.75 },
    catGridLine: { style: "none" },
    showValue: true,
    dataLabelPosition: "ctr",
    dataLabelFontSize: 12,
    dataLabelFontBold: true,
    dataLabelFormatCode: '0.00"%"',
    dataLabelColor: WHITE,
  });

  const cx = 9.05, cw = 3.8, cy0 = 1.75, ch = 1.15, cg = 0.13;
  CHANNELS.forEach((c, i) => {
    const y = cy0 + i * (ch + cg);
    const fr = DATA.fresh_rerun[c];
    s.addShape("roundRect", { x: cx, y, w: cw, h: ch, rectRadius: 0.05, fill: { color: "F4F5F9" }, line: { color: "DDDDE3", width: 0.75 } });
    s.addShape("rect", { x: cx, y, w: 0.08, h: ch, fill: { color: COLOR[c] || DARK }, line: { type: "none" } });
    s.addText(c, { x: cx + 0.2, y: y + 0.08, w: cw - 0.4, h: 0.3, fontFace: "Calibri", fontSize: 12, bold: true, color: TEXT_DARK });
    s.addText([
      { text: "Total Fresh: ", options: { fontSize: 10.5, color: MUTED } },
      { text: Math.round(fr.fresh_hours) + " jam", options: { fontSize: 12, bold: true, color: "1976D2" } },
      { text: "   Total Rerun: ", options: { fontSize: 10.5, color: MUTED } },
      { text: Math.round(fr.rerun_hours) + " jam", options: { fontSize: 12, bold: true, color: "FB8C00" } },
    ], { x: cx + 0.2, y: y + 0.42, w: cw - 0.4, h: 0.3, fontFace: "Calibri" });
    s.addText(`Fresh ${fr.fresh_pct.toFixed(2)}%  •  Rerun ${fr.rerun_pct.toFixed(2)}%`, {
      x: cx + 0.2, y: y + 0.74, w: cw - 0.4, h: 0.3, fontFace: "Calibri", fontSize: 9, color: MUTED,
    });
  });

  s.addText(`Total durasi dihitung dari seluruh durasi siaran sepanjang hari, ${DATA.target_date}.`, {
    x: 0.6, y: 6.65, w: 11.5, h: 0.3, fontFace: "Calibri", fontSize: 9.5, italic: true, color: MUTED,
  });

  pageNum++;
  addFooter(s, "Halaman " + pageNum);
}

// ---------------------------------------------------------------------------
// SLIDE 7: EXECUTIVE SUMMARY -- one auto-composed paragraph. Edit the
// template string below by hand if a run needs a specific extra nuance;
// everything it references is computed above from DATA, so the default is
// safe to ship as-is.
// ---------------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: DARK };
  s.addText("Executive Summary", { x: 0.7, y: 0.5, w: 12, h: 0.6, fontFace: "Cambria", fontSize: 27, bold: true, color: WHITE });

  const topOverall = DATA.top25[0];
  // Exec-summary point 1 is always MAIN_CHANNEL-first (this deck is an
  // INEWS-perspective tool): which Dayparts MAIN_CHANNEL itself wins + its
  // national share vs RIVAL_CHANNEL, THEN the nearest competitor's own
  // Daypart wins. Deliberately does not reuse dailyWinner/otherChannel here
  // (those exist for the insight-bar sentences on earlier slides, which
  // correctly credit whichever channel actually won nationally) -- this
  // slide's ordering was an explicit user request to always lead with
  // MAIN_CHANNEL regardless of who won that day.
  const mainVsRivalWord = DATA.official_share[MAIN_CHANNEL] >= DATA.official_share[RIVAL_CHANNEL] ? "unggul atas" : "tertinggal dari";
  let summaryText =
    `${MAIN_CHANNEL} ${mainWinDayparts.length ? `unggul di ${mainWinDayparts.length} Daypart ${winReasonWord} (${mainWinListStr})` : "tidak unggul di Daypart manapun"}, ` +
    `dengan share nasional ${DATA.official_share[MAIN_CHANNEL].toFixed(2)}% — ${mainVsRivalWord} kompetitor terdekatnya, ${RIVAL_CHANNEL}, di ${DATA.official_share[RIVAL_CHANNEL].toFixed(2)}%. ` +
    `${RIVAL_CHANNEL} sendiri ${rivalWinDayparts.length ? "unggul di " + rivalWinListStr : "tidak unggul di Daypart manapun"} ` +
    `dan program terbaik hari ini adalah ${topOverall.ProgramTitle} (${topOverall.Channel}, ${fmtNum(topOverall["000s"])} ribu penonton).`;
  if (mostAggressiveChannel && bestEfficiencyChannel) {
    summaryText += ` Dari sisi program spesial, ${mostAggressiveChannel} paling agresif dari sisi volume (${mostAggressiveN} program) namun ${bestEfficiencyChannel} paling efisien dari sisi share rata-rata per program (${bestEfficiencyShare.toFixed(2)}%).`;
  }

  s.addShape("roundRect", { x: 0.9, y: 2.1, w: 11.5, h: 3.3, rectRadius: 0.08, fill: { color: "2A3470" }, line: { color: "FFC107", width: 1.5 } });
  s.addText(summaryText, { x: 1.3, y: 2.35, w: 10.5, h: 2.8, fontFace: "Calibri", fontSize: 15, bold: true, italic: true, color: WHITE, valign: "middle" });

  s.addText(FOOTER, { x: 0.9, y: PH - 0.4, w: 10, h: 0.3, fontFace: "Calibri", fontSize: 9, color: "8892C8" });
}

pres.writeFile({ fileName: OUT }).then(() => {
  console.log("Saved:", OUT);
});
```


## ================================================================
## FILE: INEWS_Performance_Dashboard_Skill.md
## ================================================================

# INEWS Performance Dashboard Skill

## Overview
Generate a comprehensive INEWS Performance Dashboard that visualizes TV news channel performance across 4 major Indonesian TV news channels (INEWS, TVONE, KOMPASTV, METRO) for any specified week, month, or day/date range.

## Purpose
This skill creates an interactive, multi-section dashboard showing:
- Key performance metrics (share, top program performance, content composition)
- Share comparison across channels
- Content composition (Fresh vs Rerun) breakdown
- Commercial break distribution by daypart (NPT2, NPT1, PT)
- Top 25 programs ranked by viewership
- Top 15 brand advertisers with channel breakdown
- Grand total advertising expenditure summary

For a quick share-comparison question that doesn't need the full multi-section dashboard (e.g. "performa INEWS vs 4 tv berita tanggal 20-22 Agustus"), it's fine to answer directly with the relevant numbers (and optionally a lightweight chart) rather than generating every section below — see "National Share Data Source by Period" for which sheet to pull from at each granularity.

## Input Requirements

### Data Source
- **File**: RAW INEWS.xlsx (Excel workbook with multiple sheets)
- **Required Sheets** (use whichever match the request's granularity — see "National Share Data Source by Period" below):
  - `PERFORMANCE DAILY`: daily national share. Columns: `Date` (format `DD/MM/YYYY`), `Week`, `INEWS`, `TVONE`, `KOMPASTV`, `METRO`, plus 10 other monitored channels (RTV, MDTV, RCTI, SCTV, MNCTV, TRANS, TRANS7, IVM, GTV, ANTV). Added to RAW INEWS.xlsx on/around 2026-08-24 — use this whenever the request is for a specific date or a short date range ("performa harian", "tanggal X-Y").
  - `PERFORMANCE WEEKLY`: weekly national share, `Week` as a YYWW code (e.g. 2631 = week 31/2026). Use for week-based requests.
  - `PERFORMANCE MONTHLY`: monthly national share, `MONTH` as a month name. Use for month-based requests.
  - `Program Detail`: Contains program information with columns: Week, Program, Channel, 000s (viewers), TVR, Share, Dur (duration in minutes), Category (Fresh/Rerun)
  - `Advertiser`: Contains advertiser data with columns: Week, Brand/Advertiser, Channel, Adex (advertising expenditure)
  - `CB COMMERCIAL DAILY`: Contains commercial break data with columns: Week, Channel, Daypart (NPT2/NPT1/PT), Count

### National Share Data Source by Period
Follows this project's standing rule — see the `GLOBAL_National_Data_Source_Rule.md` project doc. In short: **daily request → `PERFORMANCE DAILY`; weekly request → `PERFORMANCE WEEKLY`; monthly request → `PERFORMANCE MONTHLY`.** Never derive a daily/weekly/monthly national figure from a different-granularity or city-level/program-level sheet when the matching period-native sheet already has the row.

### Parameters
- **Week Number**: Format YYXX where:
  - YY = Last 2 digits of year (e.g., 26 for 2026)
  - XX = Week number (01-52, e.g., 31 for week 31)
  - Example: 2631 = Week 31 of 2026
- **Month**: month name + year, for monthly requests (e.g. "Maret 2026") — pulls from `PERFORMANCE MONTHLY`.
- **Date / date range**: a single date or a short range (e.g. "20-22 Agustus 2026") — pulls from `PERFORMANCE DAILY`, filtering the `Date` column (format `DD/MM/YYYY`).

## Output
An interactive HTML dashboard file (e.g., `inews_performance_dashboard_week2631.html`) containing:

### Dashboard Sections

1. **Header**
   - INEWS branding with performance icon
   - Week badge showing week number in YYXX format

2. **Metric Cards** (4 KPIs)
   - Share Terbaik (Best Share): Average share percentage for INEWS
   - Program Teratas (Top Program): Highest share among INEWS programs
   - Dominasi Program (Program Dominance): Count of INEWS programs in top 25
   - Konten Terfresh (Fresh Content): Percentage of fresh content for INEWS (calculated by DURATION — see Content Composition rule below)

3. **Share Comparison Chart**
   - Vertical bar chart showing share % for all 4 channels
   - Color-coded: INEWS (red #c41e3a), TVONE (blue #4A7DB8), KOMPASTV (cyan #00bcd4), METRO (orange #ff9800)
   - Data labels showing exact percentages — **bold, comfortably large font on every bar** (see "Chart Data Labels" under Technical Specifications below), not just a tooltip-on-hover

4. **Content Composition (Fresh vs Rerun)**
   - Stacked horizontal bars for each channel
   - Shows Fresh vs Rerun percentage distribution
   - **IMPORTANT RULE — always calculate by DURATION, not episode count**: Fresh % = (sum of `Dur` where Category=='Fresh') / (sum of `Dur` for all programs) × 100, and likewise for Rerun %. Do NOT compute this as a simple count of episodes/rows — a program's duration matters, not how many times it aired. This was explicitly confirmed by the user and must be remembered for every future Fresh/Rerun composition request (weekly, monthly, or any period).
   - Interactive visual representation

5. **CB Commercial by Daypart**
   - Grid display for each channel showing 3 dayparts:
     - NPT2: Late Night period
     - NPT1: Daytime/Evening period
     - PT: Prime Time period
   - Color-coded boxes matching channel colors
   - Percentage values for each daypart

6. **Key Insights**
   - Bullet-point summary of major findings
   - Customizable based on data trends

7. **Top 25 Programs Table**
   - Ranked by 000s (viewers in thousands)
   - Columns: Rank, Program, Channel, 000s, TVR, Share %
   - INEWS programs highlighted in red background
   - Interactive hover effects

8. **Top 15 Brand Advertisers Table**
   - Ranked by total advertising expenditure
   - Columns: Rank, Brand, INEWS, TVONE, KOMPASTV, METRO, Grand Total
   - Grand Total row highlighted in yellow (#ffeb3b)
   - Shows total Adex across all channels

## Data Processing Steps

### 1. Period Filtering
Extract the requested period and filter the matching sheet (see "National Share Data Source by Period"):
```
# Daily:   PERFORMANCE_DAILY[(Date >= start) & (Date <= end)]
# Weekly:  PERFORMANCE_WEEKLY[Week == week_input]        # week_input in YYXX form
# Monthly: PERFORMANCE_MONTHLY[MONTH == month_input]
```

### 2. Program Data Extraction (Program Detail sheet)
- Filter by week (or by Month, for monthly requests) and each channel
- Sort by 000s (descending)
- Calculate:
  - Average share for each channel
  - Top program by 000s
  - Count of programs per channel in top 25
  - **Fresh/Rerun percentage for each channel — MUST be computed using total duration (`Dur` column sum), NOT count of programs/episodes.** Example: `fresh_pct = df[df.Category=='Fresh']['Dur'].sum() / df['Dur'].sum() * 100`

### 3. Advertiser Data Extraction (Advertiser sheet)
- Filter by week
- Group by brand and channel
- Calculate:
  - Top 15 brands by total Adex
  - Channel breakdown for each brand
  - Grand totals

### 4. CB Commercial Data Extraction (CB COMMERCIAL DAILY sheet)
- Filter by week
- Group by channel and daypart
- Calculate percentage of total commercials per daypart

### 5. Dashboard Generation
- Create HTML with embedded CSS and JavaScript
- Use Chart.js for share comparison visualization
- Apply consistent color scheme across all sections
- Maintain responsive design for different screen sizes

## Technical Specifications

### Color Scheme
- **INEWS**: #c41e3a (Red)
- **TVONE**: #4A7DB8 (Blue)
- **KOMPASTV**: #00bcd4 (Cyan)
- **METRO**: #ff9800 (Orange)
- **Highlight**: #ffeb3b (Yellow for totals)
- **Text**: #333 (Dark gray/black)

### Chart Library
- Chart.js 3.9.1 (CDN hosted) — or bundle Chart.js inline (no `<script src>` to a CDN) if the dashboard must be guaranteed to render offline/without external network access; either is acceptable, but verify with a headless render before delivering (see the by-cities/audience-profile skills' QA pipelines for the same principle).
- ChartJS DataLabels plugin for value display
- Responsive layout with proper scaling

### Chart Data Labels
Follows this project's standing rule (see the `GLOBAL_Chart_Label_Style_Rule.md` project doc) — every chart in this dashboard (Share Comparison bar chart, Content Composition bars, any others) must show the actual data value directly on the bar/point, in a **bold** font that is **comfortably legible, never shrunk down to save space**. This applies to every channel shown, not just INEWS. As a concrete baseline: a simple 4-channel bar chart's value labels should be sized around 16-18px bold; multi-series charts can use a slightly smaller size for secondary/context series but must stay bold and legible. This supersedes any older, unstated assumption that data labels are optional or can default to a small/thin font.

### Data Format Standards
- Numbers: Use thousands separator (comma) for display
- Currency: Rupiah values displayed as-is (no currency symbol in dashboard)
- Percentages: Display to 2 decimal places (e.g., 3.27%)
- Empty values: Display as "-" (dash)

### Footer / Source Citation
The dashboard's data-source attribution line follows this project's standing rule (see the `GLOBAL_Source_Footer_Rule.md` project doc) — `Sumber: Nielsen Media Research; National Urban; [Jenis Laporan]; [Date/Week yang dipilih]`. This skill's `[Jenis Laporan]` label is **`Weekly Performance`** for week-based or daily-based requests (unless the user asks for a distinct "Daily Performance" label), or **`Monthly Performance`** for month-based requests. Example: `Sumber: Nielsen Media Research; National Urban; Weekly Performance; Week 2631` or `Sumber: Nielsen Media Research; National Urban; Weekly Performance; 20-22 Agustus 2026` (daily range). This supersedes any older/generic "Data source attribution" wording used before 2026-08-24.

## Usage Example

**Input**: Week 2631 (August 2-8, 2026)

**Process**:
1. Load RAW INEWS.xlsx
2. Filter all sheets where Week == "2631"
3. Extract and calculate metrics (Fresh/Rerun by duration, not count)
4. Generate dashboard HTML with data
5. Save as `inews_performance_dashboard_week2631.html`

**Output**: Interactive HTML dashboard showing complete week analysis

**Daily example**: "performa INEWS vs 4 tv berita tanggal 20-22 Agustus 2026" → filter `PERFORMANCE DAILY` where `Date` is between 20/08/2026 and 22/08/2026, report/chart INEWS/TVONE/KOMPASTV/METRO's daily share (and the period average) — a direct answer or a lightweight chart is fine; the full multi-section dashboard isn't required unless asked for.

## Customization Options

### Dynamic Elements
- Week number, month, or date/date-range (input parameter, depending on requested granularity)
- Date range (calculated from week number, or given directly for daily requests)
- Data source attribution — see "Footer / Source Citation" above for the exact required wording
- Generation timestamp

### Optional Enhancements
- Export to PDF functionality
- Additional metrics/KPIs
- Comparison with previous weeks
- Trend analysis charts
- Custom theme colors

## Error Handling

### Common Issues
- **Missing sheets**: Verify all required sheets (Program Detail, Advertiser, CB COMMERCIAL DAILY, and the relevant PERFORMANCE DAILY/WEEKLY/MONTHLY sheet for the requested granularity) exist
- **Wrong week format**: Ensure week parameter follows YYXX format
- **Wrong date format**: `PERFORMANCE DAILY`'s `Date` column is `DD/MM/YYYY` — parse with that format explicitly (`pd.to_datetime(..., format='%d/%m/%Y')`), don't rely on pandas' default date inference.
- **Missing data**: Handle empty sections gracefully with "No data available" messages
- **Channel order**: Always maintain fixed order: INEWS → TVONE → KOMPASTV → METRO
- **Fresh/Rerun miscalculation**: Never use `.count()` or row count for Fresh/Rerun percentages — always use `.sum()` on the `Dur` column. This is a recurring mistake to avoid.
- **Wrong-granularity national figures**: Never derive a daily/weekly/monthly national share by aggregating a different sheet (e.g. don't average `PERFORMANCE DAILY` rows to answer a weekly question) — use the period-native sheet per `GLOBAL_National_Data_Source_Rule.md`.

## Notes
- Channel order is fixed across all reports
- INEWS programs are always highlighted for easy identification
- Grand Total row uses yellow background for emphasis
- All data is self-contained in HTML (no external data dependencies)
- Dashboard is optimized for 1600+ pixel width displays
- Mobile responsive design included
- **Fresh vs Rerun composition is ALWAYS duration-weighted (sum of `Dur`), never episode-count-based. This applies to weekly, monthly, or any other period requested.**
- Footer source citation follows the project-wide `GLOBAL_Source_Footer_Rule.md` standing rule (added 2026-08-24), labeled `Weekly Performance` or `Monthly Performance` as appropriate.
- Chart data-label styling (bold, legible, every series shown) follows the project-wide `GLOBAL_Chart_Label_Style_Rule.md` standing rule (added 2026-08-24).
- **Daily requests use the `PERFORMANCE DAILY` sheet** (added to RAW INEWS.xlsx 2026-08-24) — this is a project-wide standing rule, see `GLOBAL_National_Data_Source_Rule.md`. Don't derive daily figures from `Program Detail` or any other sheet.

## File Outputs
- Main dashboard: `ineis_performance_dashboard_week[YYXX].html`
- Individual reports (optional):
  - `top_program_week[YYXX].html` - Top 25 programs detailed view
  - `advertiser_comparison_week[YYXX].html` - Advertiser comparison
  - `fresh_rerun_week[YYXX].html` / `fresh_rerun_[month][year].html` - Content composition analysis (by duration)


## ================================================================
## FILE: INEWS_By_Cities_Performance_Skill.md
## ================================================================

# INEWS By-Cities Performance Skill

## Overview
Generate (or refresh) the standard 4-slide "Performance by Cities" PowerPoint deck comparing INEWS vs TVONE/KOMPASTV/METRO — national weekly trend, by-cities trend split into high/low contribution groups, and an executive summary infographic — for any requested week range.

## Trigger Phrases
Use this skill whenever the user asks things like:
- "performa by cities [week X-Y]"
- "performance by cities"
- "trend rating per kota/market"
- "raw data sudah update, refresh performance by cities"

## Purpose
Produces a `.pptx`, always 4 slides in this order:
1. **Weekly Performance News Channel** — national weekly trend line chart (INEWS vs 3 competitors) across the requested week range, Share & Δ% table (latest vs prior week), gap-to-nearest-competitor callout.
2. **Performance By Cities - News Channel** (high-contribution cities, Kontribusi > threshold%) — small-multiples trend charts + Kontribusi & Δ% table.
3. Same layout, **low-contribution cities** (Kontribusi ≤ threshold%).
4. **Siapa Unggul?** — dark executive-summary infographic: national scoreboard, gap callout, 4 insight cards.

## Input Requirements

### Data Source
- **File**: RAW INEWS.xlsx
- **Required Sheets**:
  - `PERFORMANCE WEEKLY`: columns Week (YYWW code, e.g. 2633 = week 33/2026), INEWS, TVONE, KOMPASTV, METRO — the ONLY authoritative source for national "average weekly" figures (slides 1 & 4). Never derive/contribution-weight a national average from city-level data when this sheet has the row you need; if the requested week isn't there yet, use the latest available week instead and say so on the slide.
  - `PERFORMANCE MONTHLY`: same idea, for monthly-basis requests (not yet automated by `prepare_data.py` — pull by hand if asked).
  - `Performance By Cities`: columns Week (plain week number, NOT the YYWW code), Market, INEWS, TVONE, KOMPASTV, METRO, Rank INEWS, Kontri. % — drives the by-city trend charts (slides 2-3) and the slide-4 insight cards.
  - `Kontribusi By Cities`: columns Market, Kontri. % — the canonical, de-duplicated contribution table (static per market; doesn't vary by week).

### Parameters
- **Week range**: start week + end week (e.g. Week 27-33), plus year prefix (e.g. 26 for 2026).

## Data Processing Steps

### 1. Segment cities by contribution
Split by Kontribusi % into a high group (slide 2) and low group (slide 3), sorted descending within each. The threshold floats around ~4% — find the *natural gap* in that run's numbers near 4% (one run's real gap was between 3.96% and 3.00%; a naive ">4.0" cutoff would wrongly exclude a ~4.0% city). Sanity-check the resulting counts look reasonable (roughly 8-10 cities in the high group).

### 2. National weekly figures (slides 1 & 4)
Pull directly from `PERFORMANCE WEEKLY` at the latest available week in/near the requested range. Compute:
- Scoreboard: each channel's share at that week.
- Δ% table: latest week vs the week immediately before it.
- Gap callout: INEWS's point-gap to its nearest competitor, at the requested range's first week vs the latest week (to show whether the gap is widening or narrowing).

### 3. By-city figures (slides 2-3)
From `Performance By Cities`, per city: weekly series per channel across the range, contribution % (from `Kontribusi By Cities`), and Δ% (latest week vs prior week) per channel.

### 4. Slide 4 insight cards
From the latest available week's by-city summary, compute per city: which channel leads, and INEWS's gap to the nearest competitor. Build 4 cards: (1) where INEWS leads — biggest leads by gap descending; (2) where competitors dominate — grouped by which competitor leads; (3) closest races — INEWS leads but narrowly (at-risk); (4) who's the real rival — tally of "nearest competitor" across all cities.

**Note**: `PERFORMANCE WEEKLY`'s latest available week can differ from `Performance By Cities`' latest available week (the two sheets aren't always updated in lockstep). When they differ, keep slide 1/4's *national* figures pinned to `PERFORMANCE WEEKLY`'s week and the by-city insight cards on `Performance By Cities`' week — state this explicitly on the slide (e.g. different "Sumber" period text) rather than silently mixing periods.

## Output Format

### Fixed Conventions (always, every run)
- **Channel order & colors** (no leading '#'): INEWS `c41e3a` (red, primary) → TVONE `1976d2` (blue) → KOMPASTV `00bcd4` (cyan) → METRO `ff9800` (orange).
- **Titles**: Slide 1 = "Weekly Performance News Channel", subtitle "National Average Weekly Share (ALL Markets) — Week [start]-[end] [year]". Slides 2 & 3 = "PERFORMANCE BY CITIES - NEWS CHANNEL" (identical), subtitles "Trend Share by Cities — Kota Kontribusi > X%" / "≤ X%". All subtitles **bold, black** (not muted grey).
- **Footer**: follows this project's standing rule (see the `GLOBAL_Source_Footer_Rule.md` project doc) — `Sumber: Nielsen Media Research; National Urban; [Jenis Laporan]; [period]`. In this deck: slides 1 and 4 use `Weekly Performance`; slides 2-3 use `Performance by Cities`. Example: `Sumber: Nielsen Media Research; National Urban; Weekly Performance; Week 33 2026` (slide 4) or `Sumber: Nielsen Media Research; National Urban; Performance by Cities; Week 27-33 2026` (slides 2-3). This supersedes the older `20+ UM Terestrial; National urban` wording used before 2026-08-24.
- **Chart data labels (all charts, every slide)**: follows this project's standing rule (see the `GLOBAL_Chart_Label_Style_Rule.md` project doc) — every bar/point's actual value must be shown directly on the chart, in a bold font, sized comfortably legible (never shrunk down just because a series is secondary). On the slide 1 national chart specifically: bigger font AND bold for all 4 channels, each in its own line color (not muted/small for competitors) — explicit user feedback, don't revert to a smaller/thin label style here. Apply the same "bold, legible, every series" spirit to the slide 2-3 by-cities small-multiples charts too, scaled down proportionally for the smaller chart size.
- **Δ% basis, everywhere**: latest available week vs the week immediately before it, labelled explicitly on the table.
- **Rounding**: keep 1-2 decimals on share values (this deck is NOT whole-number-rounded, unlike the Duration Break skill).
- **4 channels (1 primary + 3 competitors)**: the delta table needs to be wider than a 2-competitor default layout, and the top-right legend chip row needs tight spacing or the 4th chip clips off the slide edge.

### Technical Build Notes
- **Never use pptxgenjs native charts (`addChart()`)** — it can pass every validator and LibreOffice render, yet still show "needs repair" in real Microsoft PowerPoint. Render every chart (national trend, by-cities grids, scoreboard) as a matplotlib PNG and embed via `addImage()`. Tables/text/shapes stay native.
- Match matplotlib `figsize` ratio exactly to the target pptx image box's w/h (no `bbox_inches='tight'`) or the embedded image stretches/squashes.
- Sort same-x data points by y-value and enforce a minimum vertical gap before annotating, to avoid overlapping value labels.
- `axes.flat` is a one-shot iterator — `list()` it before zipping against data, or "hide unused axes" logic silently breaks.
- Each slide's executive-summary bar must cite only that slide's own table/chart — never a different slide's data source.

## QA Pipeline (run every time before delivering)
1. `validate.py` (schema/chart-part/relationship checks).
2. Convert to PDF (`soffice.py --headless --convert-to pdf`), then `pdftoppm -jpeg` per slide.
3. Actually view each rendered slide image — check text overflow, clipped titles/legend chips, blank grid cells. This catches what validators miss (every real bug in this deck's history was caught here, not by validate.py).
4. `markitdown` + grep for placeholder text (lorem/ipsum/TODO/xxx/[insert/edit me).
5. Only then deliver.

## File Outputs
- Deck filename pattern: `By_Cities_Trend_Week[start]-[end].pptx` — e.g. `By_Cities_Trend_Week27-33.pptx`.

## Notes
- This is a project-specific, opinionated workflow (fixed to INEWS's own 4-channel comparison and data conventions) — not a generic template requiring channel/color setup each time.
- An installable Claude Skill package (`inews-by-cities-performance-deck.skill`, containing SKILL.md + `scripts/prepare_data.py` + `scripts/generate_charts.py` + `scripts/build_ppt_template.js` + `scripts/leadership_analysis.py`) was delivered to the user on 2026-08-24 for them to optionally install account-wide — `prepare_data.py` automates Steps 1-3 above directly from RAW INEWS.xlsx and prints the exact numbers to paste into the chart/deck scripts.
- The footer wording was updated on 2026-08-24 to follow the new project-wide `GLOBAL_Source_Footer_Rule.md` standing rule; the skill package's `build_ppt_template.js` and SKILL.md were updated and re-delivered the same day.
- The chart data-label styling (bold, legible font, every series) is also a project-wide standing rule as of 2026-08-24 — see `GLOBAL_Chart_Label_Style_Rule.md`.

## Related Skills
- INEWS Performance Dashboard Skill: weekly/monthly share performance HTML dashboard (single-week snapshot, not city-level).
- CB Duration Break Skill / CB Commercial Breakdown Skill: Commercial Break composition tables (unrelated data domain, but same project/channel conventions).


## ================================================================
## FILE: INEWS_Audience_Profile_Skill.md
## ================================================================

# INEWS Audience Profile Skill

## Overview
Generate the "Profil Penonton TV Berita (Index)" audience-demographic comparison table — SEX/SEC/AGE index values for two weeks side by side, across a set of TV news channels — from the "Weekly profile" sheet in RAW INEWS.xlsx.

## Trigger Phrases
Use this skill whenever the user asks things like:
- "profil penonton week X vs Y"
- "audience profile [minggu]"
- "profile penonton TV berita"
- comparison of viewer demographics (gender/SEC/age) between two weeks

## Purpose
Produces a single self-contained HTML report:
- A table comparing two weeks side by side (colored header per week), rows grouped into SEX (Male/Female), SEC (Upper 1/Upper 2/Middle 1/Middle 2/Lower), AGE (5-9 through 60+), one column per requested channel within each week.
- Cells with index value **> 99** highlighted bold red on pink background (exact threshold, not approximate).
- A "Kesimpulan Profil Penonton" section: one paragraph per channel describing which segments over-index, **without quoting index numbers** in the prose.

This is a fixed, user-approved visual template (colors, borders, highlight rule) — not to be restyled per run.

## Input Requirements

### Data Source
- **File**: RAW INEWS.xlsx
- **Required Sheet**: `Weekly profile` — columns Week (plain week number, e.g. 33 — NOT the YYWW code used by PERFORMANCE WEEKLY), Channel, SEX(MALE), SEX(FEMALE), SEC(UPPER 1), SEC(UPPER 2), SEC(MIDDLE 1), SEC(MIDDLE 2), SEC(LOWER), AGE(5-9 YEARS), AGE(10-19), AGE(20-29), AGE(30-39 YEARS), AGE(40-49 YEARS), AGE(50-59 YEARS), AGE(60 + YEARS).
- Values are Nielsen-style audience indices against a Total TV population baseline: 100 = same as average population, >100 = over-represented ("over-index").
- 10 possible Channel values: INEWS, TVONE, KOMPASTV, METRO (4 fixed news channels), GARUDA TV, NUSANTARATV, CNN INDONESIA, SINPO (4 more news channels), NEWS CHANNEL (pre-computed average across all 8 news channels), Total TV (average across every monitored channel).

### Parameters
- **Two week numbers** to compare (plain week numbers, e.g. 32 and 33).
- **Channel list** (optional): defaults to the 4 fixed news channels (INEWS, TVONE, KOMPASTV, METRO) + NEWS CHANNEL average. Confirm with the user if they mean a different subset (e.g. "4 tv berita" = the 4 fixed ones; "semua tv berita" = all 8 + average; can also include Total TV).

## Data Processing Steps

### 1. Build the table
For each of the two weeks, for each requested channel, pull all 14 index columns from the `Weekly profile` sheet. Highlight any value > 99.

### 2. Auto-write the summary (no index numbers in prose)
For each channel, using the LATER of the two weeks (more current snapshot), determine which categories are over-index (>99) per group and render as compact Indonesian phrases:
- **SEX**: "Male", "Female", "kedua gender, Male maupun Female" (both over-index), or "tidak ada gender yang menonjol (merata)".
- **SEC** (ordered Upper 1 → Upper 2 → Middle 1 → Middle 2 → Lower): contiguous over-index runs as "Upper 1 hingga Middle 1"; disjoint runs joined with ", ditambah "; all 5 as "seluruh kelas SEC, dari Upper 1 sampai Lower".
- **AGE** (ordered 5-9 → ... → 60+): same contiguous-run logic, but a run reaching the open-ended top category (60+) renders as "[lower bound] tahun ke atas" (e.g. "50 tahun ke atas") — NOT the full range spelled out (avoids a grammar bug like "50-59 tahun tahun ke atas").
- NEWS CHANNEL / Total TV aggregate rows get a natural subject phrase ("TV Berita (rata-rata)" / "seluruh TV (rata-rata)") instead of literally "Penonton NEWS CHANNEL...".

## Output Format

### Visual Style (fixed template, match exactly)
- Table border: 3px solid dark navy (#2c3347), rounded corners.
- "Profile" header cell: dark navy background, white bold text, left-aligned, rowspan 2.
- Week header row: one color per week block (e.g. blue #1f6fb2 for the earlier week, green #4c7a2e for the later week), spanning all channel columns for that week.
- Channel sub-header row: lighter tint of the week's color per named channel; the NEWS CHANNEL / average column gets a distinct yellow (#ffe28a) background regardless of which week block it's in.
- Row labels: left-aligned, bold, light-grey background (#f7f7f7) — same style for EVERY row including the first row of each group (a duplicate-`class`-attribute bug once made "SEC (Upper 1)" and "AGE (5-9 Years)" render centered/unstyled — watch for two `class="..."` attributes landing on one `<td>`).
- Row-group separators: thicker top border (2.5px, #2c3347) between the SEX/SEC/AGE groups.
- Data cells: value > 99 → bold red text (#b0231f) on pink background (#f9c6cb); else black text on white.
- Legend + footer note below the table (see "Footer" below for the exact citation text) plus the index-100 baseline definition.
- "Kesimpulan Profil Penonton" box below: white background, olive border (#8a8a3c), one heading + paragraph per channel.

### Footer (source citation)
Follow `GLOBAL_Source_Footer_Rule.md` — this skill's [Jenis Laporan] label is **`Profile`**. Resolved example: `Sumber: Nielsen Media Research; National Urban; Profile; Week 32-33 2026`. (Supersedes the older `Sumber: Nielsen Media Research (sheet "Weekly profile"), 20+ UM Terestrial, National urban; Week 32-33.` wording used before 2026-08-24 — drop "20+ UM Terestrial" and use this exact 4-segment format instead.)

## QA Checklist
1. Every cell >99 is pink/red, every other cell white/black — spot-check a few against the raw sheet.
2. Row-group borders present; all row labels have identical styling (no duplicate-class bug regression).
3. Kesimpulan prose has zero index numbers and no grammar bugs (e.g. doubled "tahun tahun").
4. Wide channel lists (8-10 columns per week) don't wrap awkwardly or overflow.
5. Footer matches the exact `GLOBAL_Source_Footer_Rule.md` format (see above) — not the pre-2026-08-24 wording.

## File Outputs
- Filename pattern: `profil_penonton_tv_berita_week[N]_vs_[M].html`.

## Notes
- Different data domain from the By-Cities Performance skill (Share/Kontribusi %, not audience index) — don't conflate the two, even though both read from RAW INEWS.xlsx.
- An installable Claude Skill package (`inews-audience-profile-report.skill`, containing SKILL.md + `scripts/generate_profile_report.py`) was delivered to the user on 2026-08-24 for them to optionally install account-wide — the script does data extraction, table HTML, and the auto-written summary all in one step; only `--week1`/`--week2`/`--channels`/`--out` vary per run. The script's footer text was updated the same day to match `GLOBAL_Source_Footer_Rule.md` — a corresponding updated `.skill` package was redelivered.

## Related Skills
- INEWS By-Cities Performance Skill: Share/Kontribusi %-based deck comparing INEWS vs competitors by city (separate data domain, same RAW INEWS.xlsx file).
- INEWS Performance Dashboard Skill / CB Duration Break Skill / CB Commercial Breakdown Skill: other RAW INEWS.xlsx-based reports, unrelated data domain.
- GLOBAL Source Footer Rule: the standing source-citation format every skill in this project follows.


## ================================================================
## FILE: CB_Commercial_Breakdown_Skill.md
## ================================================================

# CB Commercial Breakdown Skill

## Overview
Generate a comprehensive CB Commercial Breakdown dashboard that visualizes commercial break distribution across dayparts (NPT1, PT, NPT2) for all 4 TV news channels for any specified month.

## Purpose
This skill creates an interactive dashboard showing:
- CB Commercial composition by daypart (NPT1, PT, NPT2)
- Total spots and duration per daypart
- Channel breakdown within each daypart
- Commercial percentage distribution
- Comparative analysis across channels
- Visual representation with pie charts and progress bars

## Input Requirements

### Data Source
- **File**: RAW INEWS.xlsx (Excel workbook with multiple sheets)
- **Required Sheet**:
  - `CB COMMERCIAL DAILY`: Contains columns: Year, Month, Channel, Day part, Custome (Daypart category), No. Of Spots, Duration, Commercial (%)

### Parameters
- **Month**: Month name (e.g., "August", "July", "June")
- **Year**: 4-digit year (e.g., 2026)

## Output
An interactive HTML dashboard file (e.g., `cb_commercial_breakdown_august2026.html`) containing:

### Dashboard Sections

1. **Header**
   - CB Commercial title with broadcast icon
   - Month and year display

2. **Summary Cards** (3 KPIs by Daypart)
   - NPT 1 (06:00-17:59, 22:00-23:59): Total spots and duration in minutes
   - PT (18:00-21:59): Total spots and duration in minutes
   - NPT 2 (24:00-25:59, 02:00-05:59): Total spots and duration in minutes

3. **Composition Charts**
   - Doughnut chart 1: Spots distribution by daypart
   - Doughnut chart 2: Duration distribution by daypart
   - Interactive legend with percentages

4. **Detailed Breakdown Table**
   - Organized by daypart sections with color-coded headers
   - Columns: Daypart, Channel, No. Of Spots, Duration (seconds), Commercial %
   - Subsection totals for each daypart
   - Grand total row at bottom

5. **Channel Comparison Section**
   - Separate subsections for each daypart
   - Progress bars showing relative spot distribution
   - Percentage and absolute values displayed
   - Color-coded by daypart (NPT1: Blue, PT: Orange, NPT2: Green)

6. **Key Insights**
   - Summary of major findings
   - Daypart dominance analysis
   - Channel ranking consistency
   - Commercial strategy observations

## Data Processing Steps

### 1. Month Filtering
Filter CB COMMERCIAL DAILY sheet by specified Year and Month:
```
if sheet['Year'] == year_input AND sheet['Month'] == month_input
```

### 2. Daypart Grouping
Group data by Custome (daypart category):
- NPT 1: 06:00-17:59, 22:00-23:59
- PT: 18:00-21:59
- NPT 2: 24:00-25:59, 02:00-05:59

### 3. Aggregation
For each daypart and channel combination, calculate:
- Total No. Of Spots (sum)
- Total Duration (sum in seconds)
- Average Commercial %
- Percentage of total spots

### 4. Channel Ranking
Maintain fixed channel order: INEWS → TVONE → KOMPASTV → METRO

### 5. Dashboard Generation
- Create HTML with embedded CSS and JavaScript
- Use Chart.js for doughnut charts
- Apply consistent color scheme
- Responsive design for multiple screen sizes

## Technical Specifications

### Color Scheme
- **NPT 1**: #1976d2 (Blue) | Gradient: #42a5f5
- **PT**: #ff9800 (Orange) | Gradient: #ffb74d
- **NPT 2**: #4caf50 (Green) | Gradient: #66bb6a
- **Channel Colors** (for reference):
  - INEWS: #c41e3a (Red)
  - TVONE: #1976d2 (Blue)
  - KOMPASTV: #00bcd4 (Cyan)
  - METRO: #ff9800 (Orange)
- **Text**: #333 (Dark gray/black)

### Chart Library
- Chart.js 3.9.1 (CDN hosted) — or bundle Chart.js inline (no `<script src>` to a CDN) if the dashboard must be guaranteed to render offline/without external network access; either is acceptable, but verify with a headless render before delivering.
- Doughnut charts for daypart composition
- Responsive layout

### Chart Data Labels
Follows this project's standing rule (see the `GLOBAL_Chart_Label_Style_Rule.md` project doc) — every chart (doughnut charts, progress bars) must show the actual data value/percentage directly on the chart, in a **bold**, comfortably legible font — never a small/thin label, and never hidden behind hover-only tooltips. Applies to every daypart/channel slice shown, not just one.

### Data Format Standards
- Numbers: Use thousands separator (comma) for display
- Duration: Display in seconds in table, minutes in stat cards
- Percentages: Display to 4 decimal places in table, 1 decimal in progress bars
- Empty values: Display as "-" (dash)

### Footer / Source Citation
The dashboard's data-source attribution line follows this project's standing rule (see the `GLOBAL_Source_Footer_Rule.md` project doc) — `Sumber: Nielsen Media Research; National Urban; [Jenis Laporan]; [Date/Week yang dipilih]`. This skill's `[Jenis Laporan]` label is **`Commercial Breakdown`**. Example: `Sumber: Nielsen Media Research; National Urban; Commercial Breakdown; Agustus 2026`. This supersedes any older/generic "Data source attribution" wording used before 2026-08-24.

## Usage Example

**Input**: August 2026

**Process**:
1. Load RAW INEWS.xlsx
2. Filter CB COMMERCIAL DAILY sheet where Year == 2026 AND Month == "August"
3. Group data by Custome (daypart) and Channel
4. Calculate totals and percentages
5. Generate dashboard HTML with visualizations
6. Save as `cb_commercial_breakdown_august2026.html`

**Output**: Interactive HTML dashboard showing CB Commercial composition by daypart

## Data Structure

### Expected Data Format (CB COMMERCIAL DAILY)
| Year | Month | Day part | Channel | Custome | No. Of Spots | Duration | Commercial (%) |
|------|-------|----------|---------|---------|--------------|----------|----------------|
| 2026 | August | 06:00-17:59 | INEWS | NPT 1 | 4955 | 89175 | 0.158866 |
| 2026 | August | 18:00-21:59 | INEWS | PT | 1795 | 30054 | 0.208708 |
| 2026 | August | 24:00-25:59 | INEWS | NPT 2 | 82 | 2071 | 0.010960 |

## Customization Options

### Dynamic Elements
- Month and year (input parameters)
- Data source attribution — see "Footer / Source Citation" above for the exact required wording
- Generation timestamp
- Daypart definitions (adjustable if needed)

### Optional Enhancements
- Weekly breakdown instead of monthly
- Trend comparison with previous months
- Year-to-date aggregation
- Custom daypart definitions
- Export to PDF functionality

## Error Handling

### Common Issues
- **Missing data**: Handle missing months gracefully with "No data available" messages
- **Empty dayparts**: Show 0 values or "-" for dayparts without data
- **Channel missing**: Skip channels with no data for that period
- **Zero values**: Display correctly without division errors

## Notes
- Channel order is fixed: INEWS → TVONE → KOMPASTV → METRO
- Daypart order: NPT 1 → PT → NPT 2 (consistent display)
- All data is self-contained in HTML (no external dependencies)
- Dashboard optimized for 1400+ pixel width displays
- Mobile responsive design included
- Color scheme clearly distinguishes between dayparts
- Footer source citation follows the project-wide `GLOBAL_Source_Footer_Rule.md` standing rule (added 2026-08-24) with label `Commercial Breakdown`.
- Chart data-label styling (bold, legible, every series shown) follows the project-wide `GLOBAL_Chart_Label_Style_Rule.md` standing rule (added 2026-08-24).

## File Outputs
- Main dashboard: `cb_commercial_breakdown_[MONTH][YEAR].html`
- Example: `cb_commercial_breakdown_august2026.html`

## Related Skills
- INEWS Performance Dashboard Skill: For weekly/monthly performance metrics
- BY DAYPART WEEKLY Skill: For detailed weekly daypart analysis


## ================================================================
## FILE: CB_Duration_Break_Skill.md
## ================================================================

# CB Duration Break Skill (Promo & Commercial)

## Overview
Generate a "Duration Break (Promo & Commercial)" table showing the percentage of broadcast duration occupied by Commercial Breaks and Promo spots, broken down by raw Day Part (actual time ranges) and by Channel, for any specified period (month OR week). This is the standard template used whenever CB duration composition is requested.

## Trigger Phrases
Use this skill whenever the user asks things like:
- "perlihatkan durasi CB [bulan/minggu]"
- "komposisi cb commercial dan promo [bulan/minggu]"
- "duration break promo dan commercial"
- "cb commercial vs promo by daypart"

## Purpose
Unlike the CB Commercial Breakdown Skill (which groups by Custome/NPT category: NPT1, PT, NPT2), this skill groups by the RAW "Day part" column exactly as recorded in the data:
- 02:00 - 05:59
- 06:00 - 17:59
- 18:00 - 21:59
- 22:00 - 23:59
- 24:00 - 25:59

And shows BOTH Promo (%) and Commercial (%) side-by-side for each channel, in one combined table (not two separate dashboards).

## Input Requirements

### Data Source
- **File**: RAW INEWS.xlsx
- **Required Sheets**:
  - `CB COMMERCIAL DAILY`: columns Year, Month, Week, Day part, Channel, Commercial (%), Duration Minutes/Day, Duration Daypart
  - `CB PROMO DAILY`: columns Year, Month, Week, Day part, Channel, Promo (%), Duration Minutes/Day, Duration Daypart

### Parameters
- **Period**: either a Month name + Year (e.g., "August 2026") OR a specific Week number (e.g., Week 31 / Week 2631 format) — filter using whichever the user specifies. When filtering by week, use the `Week` column directly (do not attempt to parse/compare `Date` values, since some rows may have mixed date formats).

## Data Processing Steps

### 1. Filter by Period
```
comm_df = CB_COMMERCIAL_DAILY[(Year==year) & (Month==month)]   # for month requests
comm_df = CB_COMMERCIAL_DAILY[Week==week_number]                 # for week requests
promo_df = CB_PROMO_DAILY[... same filter ...]
```

### 2. Per Day Part & Channel (row-level values)
Group by raw `Day part` (NOT Custome) and `Channel`, taking the MEAN of the `Commercial (%)` / `Promo (%)` columns across all days in that period:
```python
comm_grp = comm_df.groupby(['Day part','Channel'])['Commercial (%)'].mean().unstack()
promo_grp = promo_df.groupby(['Day part','Channel'])['Promo (%)'].mean().unstack()
```
Row order (fixed, always in this sequence):
1. 02:00 - 05:59
2. 06:00 - 17:59
3. 18:00 - 21:59
4. 22:00 - 23:59
5. 24:00 - 25:59

### 3. Grand Total Row
Grand Total per channel = pooled mean of ALL raw daily rows for that channel across the whole period (NOT an average-of-the-5-daypart-averages — it is a straight mean over every individual day+daypart observation, which naturally gives more weight to dayparts with more complete daily records):
```python
comm_grand = comm_df.groupby('Channel')['Commercial (%)'].mean()
promo_grand = promo_df.groupby('Channel')['Promo (%)'].mean()
```

### 4. Channel Order (FIXED — always this sequence)
INEWS → TVONE → KOMPASTV → METRO

### 5. Rounding — ALWAYS WHOLE NUMBERS, NO DECIMALS
Round every percentage value (all 5 daypart rows AND the Grand Total row) to the nearest whole number for display — e.g. show "18%", never "18.42%". Use a simple helper like `round(v)` and format as `f"{r(v)}%"`. This applies by default to every Duration Break table generated going forward, for both month and week versions, unless the user explicitly asks for decimals back. Add a footer note stating the numbers are rounded, e.g. "Angka dibulatkan ke bilangan bulat terdekat".

## Output Format

### Table Structure
A single bordered HTML table titled **"Duration Break (Promo & Commercial)"**:
- Top-right grey label row: "Channel / Promo (%) / Commercial (%)"
- Header row 1: "Day part" (rowspan 2) + one merged column per channel (colspan 2) in fixed order INEWS, TVONE, KOMPASTV, METRO
- Header row 2: "Promo (%)" / "Commercial (%)" sub-columns under each channel
- Body rows: one per Day part (5 rows), alternating light-yellow (#fdf3d0) / white row background
- Final row: "Grand total" — light blue background (#c9dff2), bold, with a thicker top border

### Visual Style (match exactly)
- Outer container: 3px solid olive/tan border (#8a8a3c), 6px border radius, 4px padding
- Header cells: light blue background (#c9dff2)
- Odd data rows: light yellow (#fdf3d0)
- Even data rows: white (#ffffff)
- Grand total row: light blue (#c9dff2), bold text, 2px top border matching outer border color
- Font: Calibri / Segoe UI, 13-14px
- Title centered above the table
- Footer note below table citing data source, sheet names, period (month/year or week number), the rounding note, and a data coverage caveat if the period is incomplete (e.g., "data s.d. 10 Agustus" if only partial month data exists)

### Footer (source citation)
The footer's source citation line follows this project's standing rule (see the `GLOBAL_Source_Footer_Rule.md` project doc) — `Sumber: Nielsen Media Research; National Urban; [Jenis Laporan]; [Date/Week yang dipilih]`. This skill's `[Jenis Laporan]` label is **`Duration Break`**. Example: `Sumber: Nielsen Media Research; National Urban; Duration Break; Week 2631` or `Sumber: Nielsen Media Research; National Urban; Duration Break; Agustus 2026`. This citation line sits alongside (does not replace) the rounding note and any data-coverage caveat described above. This supersedes any older/looser footer wording used before 2026-08-24.

## Data Completeness Caveat
Always check `Date.nunique()` (or day count) per Day part + Channel before generating. If the days recorded are uneven across dayparts/channels (common for late-night slots like 24:00-25:59 which may have sparse recording), state this explicitly in the footer or as a note beneath the table — do NOT silently present partial-period data as if it were complete.

## File Output
- Filename pattern (month): `duration_break_promo_commercial_[month][year].html` — e.g. `duration_break_promo_commercial_august2026.html`
- Filename pattern (week): `duration_break_promo_commercial_week[weekcode].html` — e.g. `duration_break_promo_commercial_week2631.html`

## Notes
- This is a SINGLE combined table (not separate Commercial and Promo dashboards) — the whole point of this skill is the side-by-side Promo/Commercial comparison per daypart per channel.
- Do not substitute Custome (NPT1/PT/NPT2) grouping here — this skill specifically uses the raw Day part time ranges. If the user wants NPT1/PT/NPT2 grouping instead, use the CB Commercial Breakdown Skill.
- Channel order is always INEWS → TVONE → KOMPASTV → METRO regardless of magnitude, even if a reference image shows a different order.
- Numbers are always rounded to whole numbers (no decimals) by default — this was an explicit, permanent user instruction (confirmed for Week 2631 request), not a one-off.
- Footer source citation follows the project-wide `GLOBAL_Source_Footer_Rule.md` standing rule (added 2026-08-24) with label `Duration Break`.

## Related Skills
- CB Commercial Breakdown Skill: NPT1/PT/NPT2 grouped dashboard (charts + detail table), Commercial only
- INEWS Performance Dashboard Skill: weekly/monthly share performance


## ================================================================
## FILE: GLOBAL_Competitor_Channel_Order_Rule.md
## ================================================================

# GLOBAL Rule — Competitor Channel Order

## Status
Standing, project-wide rule confirmed by the user on 2026-08-26. Applies to every deliverable in this project ("CHAT BOT AI" / RDO-INEWS reporting) that compares INEWS against competitor channels — PPT decks, HTML infographics/dashboards, tables, charts, and prose — regardless of skill or report type.

## The rule
Whenever channels are listed side by side in a competitor comparison, always use this fixed order:

**INEWS → TVONE → KOMPASTV → METRO**

Never alphabetical, never sorted by share/rank, never any other ordering — even when a channel other than INEWS wins that particular metric. This applies to:
- Chart category/series order and x-axis ordering (bar charts, stacked charts, small multiples)
- Table column order and table row order (when rows are per-channel)
- Card/panel layouts (e.g. one summary card per channel)
- Legend order
- Narrative lists in prose ("INEWS, TVONE, KOMPASTV, dan METRO ...")

## Known instance this bit us
The `inews-competitor-daypart-deck` skill's Slide 6 (Fresh vs Rerun) used `[...CHANNELS].sort()` (alphabetical: INEWS, KOMPASTV, METRO, TVONE) for its bar chart and per-channel cards, while every other slide in the same deck already used the fixed `CHANNELS` array order. Fixed 2026-08-26 — `build_deck.js` now uses `CHANNELS` directly everywhere, no alphabetical sort anywhere in the script. When building or reviewing any new skill/script in this project, grep for `.sort()` applied to a channel list and remove it.

## Where this applies
- `inews-competitor-daypart-deck` skill (Head-to-Head table, Daypart chart, Fresh vs Rerun, Top 25 "Ch" column context)
- `INEWS_By_Cities_Performance_Skill` and any other multi-channel comparison skill/deck in this project
- CB Commercial/Duration Break tables (already used this fixed order correctly — no change needed there, kept as the reference example)

## Related
- `GLOBAL_Daypart_Terminology_Rule.md` — wording rule, independent of this ordering rule; apply both together.
- `GLOBAL_Chart_Label_Style_Rule.md` — chart label styling, independent of this ordering rule.


## ================================================================
## FILE: GLOBAL_Source_Footer_Rule.md
## ================================================================

# Global Rule: Source Footer (applies to every report/skill in this project)

## The rule
Every output produced in this project (HTML dashboard, HTML table, PowerPoint deck — any deliverable built from RAW INEWS.xlsx) must carry a footer citing the source, in exactly this format:

```
Sumber: Nielsen Media Research; National Urban; [Jenis Laporan]; [Date/Week yang dipilih]
```

Four semicolon-separated segments, always in this order:
1. `Nielsen Media Research` — literal, never changes.
2. `National Urban` — literal, never changes.
3. **[Jenis Laporan]** — a short label naming what the specific report/section covers, e.g. `Profile`, `Daily Performance`, `Weekly Performance`, `Performance by Cities`, `Duration Break`, `Commercial Breakdown`. Pick the label that matches that report's own name/subject **and granularity** — see the per-skill mapping below. If one deck has multiple sections covering different report types (e.g. a national performance slide plus a by-city slide in the same deck), each section's footer uses ITS OWN matching label, not one label for the whole deck.
4. **[Date/Week yang dipilih]** — the actual period requested/shown, written as a specific date or date range, e.g. `17 Agustus 2026`, `Week 27-33 2026`, `Week 33 2026`, `August 2026`. Match the period actually covered by that section/table, not a different section's period.

Never use the word "Resmi" (or "Official") anywhere in the footer or in chart/slide titles when referencing these sheets — confirmed by the user on 2026-08-25. Simply name the source sheet/label and the period; don't editorialize that it's the official number. (Applies project-wide, not just the footer — e.g. a chart titled "Share (%) Resmi (BY DAYPART DAILY)" should read "Share (%) per Daypart (BY DAYPART DAILY)" or similar, dropping "Resmi".)

This **replaces** any earlier footer wording used before 2026-08-24 (e.g. the by-cities deck's older `Sumber: Nielsen Media Research; 20+ UM Terestrial; National urban; ...` — drop the `20+ UM Terestrial` segment and the different capitalization of "National urban"; the new standard is `National Urban`, capitalized, and no UM Terestrial segment).

## Daily vs Weekly Performance — resolved 2026-08-25
Previously "Daily Performance" as a distinct label was unconfirmed and every daily/single-date deliverable defaulted to `Weekly Performance`. This is now resolved: **use whichever label matches the actual granularity of the period shown**, the same way `GLOBAL_National_Data_Source_Rule.md` picks the sheet by granularity:
- A single specific date (e.g. "17 Agustus 2026") → **`Daily Performance`**.
- A week or range of weeks (e.g. "Week 27-33") → **`Weekly Performance`**.
- A month → **`Monthly Performance`**.

Don't default to `Weekly Performance` for a daily request anymore — match the case.

## Per-skill [Jenis Laporan] mapping (current projects)

- **INEWS Audience Profile Skill** → `Profile`
- **INEWS By-Cities Performance Skill**:
  - Slide 1 "Weekly Performance News Channel" (national trend) → `Weekly Performance`
  - Slides 2-3 "Performance By Cities" (by-city trend) → `Performance by Cities`
  - Slide 4 "Siapa Unggul?" (national scoreboard/gap, sourced from PERFORMANCE WEEKLY) → `Weekly Performance`
- **INEWS Performance Dashboard Skill** → `Daily Performance` / `Weekly Performance` / `Monthly Performance` matching the request's granularity (see rule above)
- **INEWS Competitive Daypart Deep-Dive Skill** → `Daily Performance` for a single-date deep-dive (e.g. "17 Agustus 2026"), `Weekly Performance` if the request spans a week/range of weeks
- **CB Duration Break Skill** → `Duration Break`
- **CB Commercial Breakdown Skill** → `Commercial Breakdown`

When a new report type is created later, add its label to this list rather than inventing ad-hoc wording each time — keep this doc as the single source of truth for footer text across every skill in this project.

## Why this is a separate doc
Every skill doc in this project (`INEWS_*_Skill.md`, `CB_*_Skill.md`) references this rule instead of repeating the full footer text — that way, if the exact wording needs to change again, it's a one-place edit rather than five. Each individual skill doc's Output Format / Footer section states its own resolved [Jenis Laporan] label and gives one concrete example string, but should say "see GLOBAL_Source_Footer_Rule.md" for the general rule itself.


## ================================================================
## FILE: GLOBAL_National_Data_Source_Rule.md
## ================================================================

# GLOBAL Rule — Which Sheet Is Authoritative for National Performance, by Period

## Status
Standing, project-wide rule. Originally established for weekly/monthly national share (INEWS vs TVONE/KOMPASTV/METRO), extended on 2026-08-24 to cover daily requests once the `PERFORMANCE DAILY` sheet appeared in RAW INEWS.xlsx, and extended again on 2026-08-25 to cover daypart-level requests once `BY DAYPART DAILY` and `BY DAYPART WEEKLY` appeared.

## The rule
When asked for INEWS's national performance/share vs competitors (TVONE, KOMPASTV, METRO — or a subset), **always pull from the period-native sheet that matches the request's granularity** — never derive or aggregate a national figure from a different granularity or from city-level/program-level data when the matching sheet already has the row you need.

| Request granularity | Authoritative sheet | Key/index column | Notes |
|---|---|---|---|
| **Daily** (a specific date or a short date range, e.g. "20-22 Agustus 2026") | `PERFORMANCE DAILY` | `Date` (format `DD/MM/YYYY`), plus a `Week` column for cross-reference | Added to RAW INEWS.xlsx on/around 2026-08-24. Columns: `Date`, `Week`, `INEWS`, `TVONE`, `KOMPASTV`, `METRO`, plus 10 other monitored channels (RTV, MDTV, RCTI, SCTV, MNCTV, TRANS, TRANS7, IVM, GTV, ANTV) — filter/select only the 4 news channels unless the user asks for the wider roster. |
| **Weekly** (a week number or week range, e.g. "Week 27-33", "week kemarin") | `PERFORMANCE WEEKLY` | `Week` (YYWW code, e.g. 2633 = week 33/2026) | Established earlier for the by-cities-performance-deck skill (slides 1 & 4) and the INEWS Performance Dashboard skill. |
| **Monthly** (a month name, e.g. "Maret 2026") | `PERFORMANCE MONTHLY` | `MONTH` (month name, e.g. `MARCH`) | Also has a yearly rollup row (`MONTH == 2026`) — exclude that row from any month-by-month trend/list unless a full-year figure is explicitly wanted. |
| **Daypart, daily** (per-daypart share/TVR/viewers for one specific date) | `BY DAYPART DAILY` | `Date` (`DD/MM/YYYY`) + `Day Part` (8 fixed time-range bins: `02:00 - 04:29`, `04:30 - 09:59`, `10:00 - 12:59`, `13:00 - 15:59`, `16:00 - 17:59`, `18:00 - 21:59`, `22:00 - 23:59`, `24:00 - 25:59`) | Added 2026-08-25. Columns: `Year`, `Month`, `Week`, `Date`, `Day Part`, `Channel`, `TVR`, `Share`, `000s`, `Reach 000s`. This is Nielsen-weighted per calendar time-block — **use this, not an average of `Program Detail REV2` rows grouped by `Daypart Custome`**, whenever a request is about a single date's daypart/Daypart breakdown. Confirmed materially more accurate on 17 Agustus 2026: the Sore daypart (16:00-17:59) reads as ~4.7-4.9% share here, vs ~0% when derived from a simple Program-Detail average (a long-running HUT RI special program starting in the prior daypart absorbed all its records under one earlier `Daypart Custome` tag, hiding real audience inside the 16:00-17:59 window). |
| **Daypart, weekly** (per-daypart share/TVR/viewers rolled up for a week or range of weeks) | `BY DAYPART WEEKLY` | `Week` (plain week number, e.g. 33) + `Day Part` (same 8 bins as above) | Added 2026-08-25 (briefly named `BY DAYPART WEEKLYREV` mid-revision, renamed back to `BY DAYPART WEEKLY`). Columns: `Year`, `Month`, `Week`, `Day Part`, `Channel`, `TVR`, `Share`, `000s`, `Reach 000s`. Use this instead of averaging daily daypart rows or Program Detail when the request spans a week or several weeks. |

**Never** compute a "daily" or "weekly" national OR daypart-level average by filtering/aggregating `Performance By Cities` (city-level) or `Program Detail` / `Program Detail REV2` (program-level) data — those sheets exist for different purposes (by-city breakdowns, program-level narrative/detail) and are not guaranteed to reconcile exactly with the period-native national or daypart sheets. `Program Detail REV2` remains the correct and only source for program-level detail (top programs, special/ceremonial program identification via `Level 1 == 'Special'`, fresh/rerun composition) — use it alongside `BY DAYPART DAILY`/`BY DAYPART WEEKLY`, not instead of them, when a deliverable needs both the official daypart numbers AND the program-by-program narrative. If the period-native sheet doesn't yet have the requested date/week/month, say so explicitly and offer the latest available period instead of silently substituting a derived number.

## Why this matters
Each of these sheets is Nielsen's own pre-aggregated "official" figure for that period/segment. Deriving a same-period number a different way (e.g. averaging daily rows to get a "weekly" figure, or averaging program-level rows to get a "daypart" figure, instead of reading the matching official sheet directly) risks small-to-large discrepancies from Nielsen's own per-minute weighting/methodology and breaks the "one source of truth per period" principle this project has followed since the by-cities-performance-deck work. The 17 Agustus 2026 Sore-daypart case above is a concrete example of how large that discrepancy can get when a derived figure is used instead of the official one.

## Related
- Complements `GLOBAL_Source_Footer_Rule.md` (footer wording — the `[Jenis Laporan]` label for a daily request should typically read `Weekly Performance` unless a more specific label is agreed, since "Daily Performance" as a distinct label hasn't been requested yet — ask if unsure), `GLOBAL_Chart_Label_Style_Rule.md` (chart label styling), and `GLOBAL_Daypart_Terminology_Rule.md` (always say "Daypart", never "Blok"/"Blok Waktu").
- Referenced by: INEWS Performance Dashboard Skill, INEWS By-Cities Performance Skill, INEWS Competitive Daypart Deep-Dive Skill.


## ================================================================
## FILE: GLOBAL_Daypart_Terminology_Rule.md
## ================================================================

# GLOBAL Rule — Daypart Terminology

## Status
Standing, project-wide rule confirmed by the user on 2026-08-25. Applies to every deliverable in this project ("CHAT BOT AI" / RDO-INEWS reporting) that references TV time-block segments — PPT decks, HTML infographics/dashboards, tables, chart labels, and prose — regardless of skill or report type.

## The rule
Always call it **"Daypart"** (or "Daypart [Nama]", e.g. "Daypart Primetime", "Daypart Pagi") — never "Blok Waktu" or "Blok" as the generic term for a TV schedule time segment.

- Correct: "Perbandingan Share per Daypart", "Daypart Primetime (18:00-21:59)", "Rata-Rata Penonton per Daypart", table column header "Daypart".
- Incorrect: "Perbandingan Share per Blok Waktu", "Blok Primetime", column header "Blok Waktu" / "Blok".

This applies to slide titles, subtitles, chart axis labels/titles, table headers, card labels, and body prose — anywhere the concept of a time-segment is named. It does not require translating the daypart NAMES themselves (Dini Hari, Pagi, Siang 1, Siang 2, Sore, Primetime, Malam stay as-is, or the official `Day Part` time-range labels from the `BY DAYPART DAILY` / `BY DAYPART WEEKLYREV` sheets) — only the generic category word changes from "Blok (Waktu)" to "Daypart".

## Where this applies
- INEWS Competitive Daypart Deep-Dive Skill (already uses "Daypart" in its own name — make sure generated output text matches, not just the skill's title)
- Any PPT/HTML deck with a per-time-segment breakdown (e.g. the 17 Agustus 2026 INEWS vs Kompetitor analysis deck — "Perbandingan Share per Blok Waktu" / "Rata-Rata Penonton per Blok Waktu" slide titles, the "Blok Waktu" / "Unggul" table column, and the "Skor Kemenangan Blok" card on the head-to-head slide all need correcting to "Daypart" on next rebuild)
- INEWS Performance Dashboard Skill, CB Duration Break Skill, or any other skill/report that segments the day into time buckets

## Related
- `GLOBAL_National_Data_Source_Rule.md` — governs which sheet is authoritative for daily/daypart figures (see 2026-08-25 update: `BY DAYPART DAILY` / `BY DAYPART WEEKLYREV` are now the authoritative per-daypart sources, superseding daypart averages derived from Program Detail).
- `GLOBAL_Chart_Label_Style_Rule.md` — chart label styling (bold, sized) is independent of this wording rule; apply both together.


## ================================================================
## FILE: GLOBAL_Chart_Label_Style_Rule.md
## ================================================================

# GLOBAL Rule — Chart Data Labels & Font Size

## Status
Standing, project-wide rule confirmed by the user on 2026-08-24. Applies to every chart (bar, line, or any other chart type) built for this project ("CHAT BOT AI" / RDO-INEWS reporting), in any output format (HTML/Chart.js dashboards, matplotlib PNGs embedded in PowerPoint decks, etc.) — not specific to one skill or one report.

## The rule
1. **Always show the actual data value directly on the chart** — as a label on every bar / every point on a line — not hidden behind a hover tooltip only. The person reading a static screenshot or a printed/exported version must be able to read the exact number without interacting with the chart.
2. **Data label font must be bold.** Never render value labels in a thin/regular weight.
3. **Data label font must be reasonably large — never "too small" to read comfortably.** As a concrete baseline (validated in practice, adjust proportionally to the chart's own size):
   - Single-series or few-category bar chart (e.g. a 4-channel comparison bar chart): value label font size ≈ 16-18px bold.
   - Multi-series line/trend chart, primary series being compared (e.g. INEWS vs TVONE): ≈ 13px bold.
   - Multi-series line/trend chart, secondary/context series (e.g. KOMPASTV, METRO shown for context, not the main comparison): may be slightly smaller than the primary series (≈ 11px bold) but still bold and still legible — never shrunk down to a barely-readable size just because the series is secondary.
   - Axis tick labels and legend text should also be reasonably sized and bold, not just the value labels — the whole chart should read comfortably at a glance, not just the datapoints.
4. **This applies to every channel/series shown, not just the primary channel.** (This generalizes the earlier per-report instruction — originally given for the by-cities deck's Slide 1 national chart, where the user asked for all 4 channels' labels to be enlarged and bolded, not just INEWS's.)
5. When a chart has many close/overlapping data points (e.g. two competitor lines running close together), stagger label positions (alternate above/below the line, add small offsets) to avoid label collisions — but do not solve overlap by shrinking the font below the size floors above. Reducing clutter should come from label placement/offset logic, not from making text smaller.

## Where this has been applied so far
- The standalone "Share INEWS vs TVONE" HTML chart/dashboard (Maret 2026 request, built 2026-08-24) — bar chart value labels bumped to 18px bold; line chart given per-point value labels for all 4 channels (13px bold for INEWS/TVONE, 11px bold for KOMPASTV/METRO), with alternating top/bottom label alignment per series to avoid collisions.
- The by-cities-performance-deck's Slide 1 national weekly trend chart already had this convention (established earlier) — all 4 channels' data labels enlarged and bolded, each in its own line color.

## Related
- This is a visual-styling rule, independent of (and complementary to) `GLOBAL_Source_Footer_Rule.md`, which governs the footer/source-citation text rather than chart labels.
- Applies to any current or future skill/report in this project that renders a chart: INEWS Performance Dashboard Skill (Share Comparison chart, Content Composition bars), CB Commercial Breakdown Skill (doughnut charts / progress bars), INEWS By-Cities Performance Skill (national + by-city trend charts), and any one-off chart/dashboard built on request (e.g. ad-hoc "share X vs Y" comparisons).
