# RCV Election Simulation — Research Notes

Working notes on building this tool with Claude as a pair. Captures the goal, the architecture that emerged, every data source we evaluated, the data-quality issues that surfaced and how we addressed them, and the calls we deferred or ruled out.

The intent is that anyone (including future-us) reading this can understand *why* the code looks the way it does without re-doing the investigation.

## Goal

Find Pennsylvania races where no candidate cleared a majority — the universe where ranked-choice voting (instant-runoff) would have changed the outcome. The output is a set of Excel workbooks, each grouping non-majority races by jurisdiction or category, so the user can scan for the most RCV-relevant historical contests.

Initial scope was Philadelphia primaries; expanded over the session to PA statewide (federal + state offices via OpenElections) and PA local (Allegheny via WPRDC, then all 67 counties for 2025 via OpenElections).

## High-level architecture

One file: `primary_scraper.py`. Three layers:

1. **`ElectionSource` subclasses** — each one takes a URL (or set of URLs) and produces a tidy DataFrame `[Candidate, Race_Name, Votes]`. Subclassing per file *shape*, not per jurisdiction, so the same parser is reused across cities/states with the same upstream format.
2. **Pipeline filters** — pure functions that take a tidy DataFrame and return a tidy DataFrame. Source-agnostic. Composable.
3. **Writers** — two output modes:
   - `write_workbook(sources, out_path)` — one Excel sheet per source (used for Philly: one sheet per election year).
   - `write_workbook_pooled_by_category(sources, out_path)` — sources grouped by their `category` attribute, all years pooled into one sheet per category (used for the OE/Allegheny workbooks).

Adding a new election source is one new class + one entry in a hardcoded source list at the bottom of the file. No CLI; the file's bottom block configures everything.

### Why this shape

- Single-file is enough at current scale (~1k lines); split is deferred until something concrete demands it.
- Class-per-shape (vs. function-per-source) makes shared knobs (sheet names, party prefix handling, etc.) declarative — a new Philly-style file is `WideCandidateRaceCsvSource(name="...", url="...")`, not a new function.
- The pipeline-of-pure-functions style means every transformation is independently testable without network mocking. 74 unit tests, all on pure functions, sub-second.

## Data sources evaluated

| Source | Coverage | What it gives us | Status |
|---|---|---|---|
| `vote.phila.gov` | Philly primaries 2007, 2011, 2015, 2019, 2023 | Local-only races (mayor, council, district council) | ✅ Integrated. 5 file shapes (3 source classes). |
| OpenElections PA (`<year>/*.csv` at root) | 2018 / 2020 / 2024 PA general | Federal + state offices statewide | ✅ Integrated as `OpenElectionsCsvSource`. |
| OpenElections PA (`<year>/counties/*.csv`) | 2020 primary (48/67), 2022 general (24/67), 2024 primary (10/67), **2025 general (67/67)** | Statewide if you stitch; partial in many years | ✅ Integrated as `StitchedOpenElectionsCountiesSource`. |
| WPRDC (`data.wprdc.org`) | Allegheny County 2012–2025, every primary + general | Pre-aggregated to county totals, includes **local races** | ✅ Integrated as `WprdcCsvSource` for 2017+. |
| **Clarity Elections** (`results.enr.clarityelections.com`) | **York 2023–2025, Delaware 2024–2025** | JSON-API for mid-size county results, summary endpoint pre-aggregated | ✅ Integrated as `ClaritySummaryJsonSource`. CloudFront requires browser User-Agent. |
| **Electionware PDFs** (multiple county sites) | **Berks 2023/2025; Chester, Northampton, Mercer, Northumberland 2021/2023/2025; Centre 2025; Lebanon 2021/2025; Indiana 2025; Schuylkill 2023/2025; Lehigh 2021/2025** | Reading, West Chester, Bethlehem/Easton, State College, City of Lebanon, Mercer/Sharon/Hermitage, Sunbury/Shamokin, Indiana borough, Pottsville, Allentown + surrounding municipalities. Snyder is confirmed Electionware too (per OE parser docstrings), but its older primary PDFs have been rotated off the public site since the last election cycle. | ✅ Integrated as `ElectionwarePdfSource` (pdfplumber + regex, no API key needed). `LlmPdfSource` retained as an alternative for novel PDF layouts. |
| **Lycoming PDFs** (`lycomingcountypa.gov`) | **Lycoming 2021/2023/2025 primaries** | Williamsport + boroughs + townships. Different vendor — inline contest headers with party in parens (e.g., `"District Attorney (Dem) (Vote for 1), ..."`) and a VOTE % candidate column. | ✅ Integrated as `LycomingPdfSource`. Same `_pdf_url_to_lines` download helper as the Electionware path. |
| Lehigh / Bucks / Chester / Montgomery county portals | Variable | Local races for each county | ❌ Deferred — PDF-only, would need per-county work. Could be added via `LlmPdfSource` later. |
| PA Department of State (`electionreturns.pa.gov`) | All 67 counties statewide | Federal/state races only, not local | ❌ Not pursued — duplicates OE statewide coverage; doesn't add local. |
| OpenElections PA pre-2018 | 2000–2016 fixed-width files | Older races | ❌ Deferred — different schema; would need new parser. |

### Why these and not others

We made these choices in roughly the order shown above. The progression was:

1. **Philly first** because the user started there. Five different file shapes across the years revealed the need for the class-per-shape abstraction.
2. **OpenElections second** because it's the obvious "give me PA data" source. The schema is the same across years, so one parser handles 2018+. Coverage is uneven (some years have state rollups, others need stitching from per-county files), so we added the stitching source.
3. **WPRDC third** for Allegheny because the user wanted local races and Allegheny is the second-largest PA jurisdiction with the cleanest data hub outside Philly.
4. **OpenElections 2025 county files fourth** — discovered late in the session that OE has now done the per-county PDF/Excel parsing work for the entire state for 2025, producing tidy CSVs we could ingest with our existing parser. This was the highest-leverage addition: one new source unlocks all 67 PA counties for 2025 local races.
5. **Clarity Elections fifth (York + Delaware)** — when the user asked to broaden coverage across PA counties, we revisited Clarity. The user provided URLs for several mid-size counties, which lets us pick known-Clarity instances directly without needing a discovery API. York and Delaware together give two different mid-size county profiles with strong competitive races in both major-party primaries. CloudFront blocks plain `curl` so we send a Safari User-Agent on every request.
6. **Berks PDFs via Electionware parser sixth** — Berks publishes PDF-only results back to 2003. We first built a generic `LlmPdfSource` (Claude tool_use + disk cache) on the assumption PDFs were too varied to parse with regex; then learned the user didn't have an Anthropic API key, so we wrote a non-LLM alternative. The Berks PDFs are generated by the Electionware vendor (same as Bedford, Blair, Bradford, Carbon, Centre, Chester, Crawford, Elk, Erie, Franklin, etc. per OpenElections' parser scripts) with a consistent format: party-prefixed contest header, "Vote For N" line, then candidate rows with TOTAL + per-method columns. A 40-line state machine over `pdfplumber`'s layout-preserving text extraction handles it cleanly. The same `ElectionwarePdfSource` class will work for many other PA counties as one-line registry additions. `LlmPdfSource` is kept in the codebase as a fallback for novel PDF layouts that don't match any vendor's standard form.

### Why we ruled out the alternatives

- **PDF parsing for Bucks/Lehigh/Chester/Montgomery**: Each county's PDF layout is different. OpenElections itself maintains 49 county-specific parser scripts to handle PA. That work has been done — we get it for free for 2025 via their `counties/` directory. Once `LlmPdfSource` is proven on Berks, the same class can be pointed at other counties' PDFs.
- **State portal (`electionreturns.pa.gov`)**: Doesn't have local races; only adds noise relative to OE which we already use.

## Output workbooks

| Workbook | Source | Layout | Scope |
|---|---|---|---|
| `Philadelphia_Primary_Mayor_DistrictCouncil.xlsx` | vote.phila.gov | One sheet per year (2007–2023) | Philly primaries, mayor/council only |
| `Pennsylvania_NonMajority_2018plus.xlsx` | OpenElections | Two pooled sheets (Primaries, Generals) | PA federal + state, 2018–2024 |
| `Allegheny_NonMajority_Local.xlsx` | WPRDC | Two pooled sheets (Primaries, Generals) | Allegheny County local races, 2017–2025 |
| `Pennsylvania_NonMajority_Local2025.xlsx` | OpenElections `2025/counties/` | One pooled sheet (Generals) | All 67 PA counties local races, 2025 only |
| `Pennsylvania_NonMajority_MidCounties.xlsx` | Clarity (York/Delaware) + Electionware PDF (Berks + 9 more) + Lycoming PDF | Two pooled sheets (Primaries, Generals) | York/Delaware 2023+; Electionware-format PDF for 10 PA counties 2021–2025; Lycoming 2021–2025 |
| `Top_RCV_Races.xlsx` | derived from the five workbooks above | Four sheets: closest two-way, crowded fields, notable local offices, all non-majority | One row per race; sorted to put the most RCV-relevant cases first |

The workbooks are checked into the repo so anyone can read the findings without re-running the pipeline. Regeneration (`python primary_scraper.py`) overwrites them in place.

## Data quality issues encountered and how we addressed them

This was the bulk of the work. Election data published by counties is messy in surprisingly consistent ways. Each issue below was a discovered surprise that produced false positives in the output before we fixed it.

### 1. Candidate name canonicalization

The biggest source of false non-majority hits. Different counties spell the same candidate's name differently, splitting their vote total across multiple "candidates" and falsely flagging races as non-majority.

Cleanups applied (in order, inside `_normalize_candidate`):

| Variant | Example | Fix |
|---|---|---|
| Case | `"Josh Shapiro"` vs `"JOSH SHAPIRO"` | Uppercase |
| Periods | `"Richard L. Weiss"` vs `"Richard L Weiss"` | Strip periods |
| Commas in suffixes | `"Langerholc, Jr"` vs `"Langerholc Jr"` | Strip commas |
| Apostrophes | `"O'Brien"` vs `"OBrien"` | Strip apostrophes |
| Trailing party tags | `"Dave Sunday REP"` vs `"Dave Sunday"` | Regex-strip a known party-code suffix |
| Middle initials | `"Ryan E Mackenzie"` vs `"Ryan Mackenzie"` | Drop single-char middle tokens (preserve first/last) |
| Typos | `"Shaun Doughherty"` vs `"Shaun Dougherty"` | `difflib.SequenceMatcher.ratio() >= 0.92`, min 6 chars; merge lower-vote spelling into higher-vote |

The thresholds (0.92 ratio, min 6 chars) are deliberately conservative. They catch the `DOUGHHERTY/DOUGHERTY` case (ratio 0.97) but leave `JOHN/JOAN` (0.75) and `TRUMP / TRUMP / VANCE` (0.55) alone.

### 2. Tickets — the issue we couldn't auto-fix

Gubernatorial and presidential races report each ticket multiple ways across counties: `TRUMP`, `TRUMP / VANCE`, `DONALD J TRUMP`, `DONALD J TRUMP, PRESIDENT / JD VANCE, VICE-PRESIDENT`, `DONALD JOHN TRUMP AND JD VANCE`, etc. Same with `JOSH SHAPIRO` vs `SHAPIRO / DAVIS` vs `JOSH SHAPIRO AUSTIN DAVIS`.

These can't be merged with edit-distance heuristics (the strings are too different) and the first-word/last-word approach is ambiguous (which surname is the head of the ticket?). The honest fix would require an external candidate roster.

**Our choice**: exclude the President race entirely from the PA statewide workbook via `race_exclude_pattern=r"president"`. 2022 Governor still appears with ticket fragmentation but is flagged via partial coverage (24/67 counties). Better to surface a noisy entry with a coverage caveat than silently filter a real race.

### 3. Pseudo-candidates inflating denominators

Some sources include `OVER VOTES`, `UNDER VOTES`, `NOT ASSIGNED`, `WRITE-IN TOTALS` as if they were candidates. Counting them in the race total falsely lowers every real candidate's percentage and flags races as non-majority that actually had a majority winner.

**Fix**: hardcoded `_NON_CANDIDATE_NAMES` set, dropped early in `_parse_openelections_df`.

### 4. Aggregate write-in rows double-counting

Discovered when running the 2025 PA-wide workbook: one borough mayoral race appeared to have 38 candidates with the "leader" at 6.1%. Investigation: the file contained one aggregated `WRITE-INS` row PLUS 37 individual named write-in rows. The aggregate is the *sum* of the individual ones, so including both double-counted votes and inflated the field.

**Fix**: added `WRITE-INS` / `WRITE-IN` / `WRITE INS` / `WRITE IN` variants to `_NON_CANDIDATE_NAMES`. Lose the data point for races that have only an aggregate (no individual breakdown), but those are sub-1% anyway and don't affect majority-winner detection.

### 5. Missing-district silent merge

The 2024 OpenElections PA file lists all Philadelphia state-house candidates with `district=""`. Without intervention, our parser builds `Race_Name = "STATE HOUSE"` for all of them, merging 27 unrelated reps from different districts into one bogus race.

**Fix**: `_NEEDS_DISTRICT_OFFICES = {"STATE HOUSE", "STATE SENATE", "U.S. HOUSE"}` — drop rows where the office is in this set AND district is blank.

### 6. Multi-seat (Vote For N>1) races

IRV's single-winner-majority concept doesn't apply to multi-winner contests (those would use STV, a different system). Our analysis only makes sense for single-seat races.

**Fix**: parse the `(Vote For N)` tail from WPRDC contest names; drop rows where N > 1.

### 7. Write-in-only chaos races

After all the canonicalization above, the 2025 PA workbook still surfaced ~100 races where the "leader" got under 10% and the "field" was 20+ named write-ins. These are uncontested races on the ballot — no listed candidate, just write-in chaos. Not RCV-relevant.

**Fix**: `filter_min_leader_percent` (default 10%). Drops races where no candidate cleared 10%. The threshold leaves legitimate competitive multi-way races intact (most have leaders well above 10% even in 6+-way fields).

### 8. WPRDC primary-naming convention change

WPRDC's 2017/2019 primary files don't include the party prefix in `contest_name` — party is only in `party_name`. Their 2021+ files include it (`DEM Mayor Pittsburgh`). Without intervention, the 2017/2019 parser would merge DEM and REP primaries for the same office into one fake race.

**Fix**: in `_parse_wprdc_summary_df`, if `is_primary=True` and the contest name's first word isn't a known party code, prepend `party_name`. Dropped 2017 false positives from 35 to 14 and 2019 from 28 to 11.

### 9. Coverage tracking

Stitched sources can have partial county coverage (e.g., OE 2022 PA general has only 24 of 67 counties' files). A race that looks non-majority across those 24 counties may have a majority winner statewide.

**Fix**: every output row carries a `Coverage` column showing `"24 of 67 counties (stitched)"`. The number is derived from the data (number of distinct counties in the fetched frame), not hardcoded — so misrepresentations like "67/67 (state rollup)" can't slip in when the rollup file is actually partial. (Discovered: the 2020 PA general "county rollup" file is only 13/67. Switched to the precinct file which is 67/67.)

### 10. CloudFront User-Agent blocking (Clarity)

Surfaced when probing Clarity Elections endpoints: `curl` and the bare `urllib.request` UA get a `403 Forbidden` from CloudFront. The data itself is public, but the front-end blocks anything that looks like a bot.

**Fix**: `ClaritySummaryJsonSource` sets `User-Agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15` on every request. The header is stored as a class attribute so it can be overridden per-instance if a Clarity URL ever needs something different.

### 11. Write-in variant proliferation

After integrating Clarity, the workbook revealed many more write-in spellings the existing literal set didn't catch: `WRITE-IN (TOTAL)` (Clarity's preferred form, with the parenthetical), `UNASSIGNED WRITE-INS` (Delaware), `UNRESOLVED WRITE-IN` and `WRITE-IN: SCATTERED` (Allegheny WPRDC). The literal set was getting unwieldy and missing variants.

**Fix**: replaced the literal `~candidate.isin(_NON_CANDIDATE_NAMES)` check with a helper `_is_non_candidate()` that combines the small literal set (`OVER VOTES`, `UNDER VOTES`, `NOT ASSIGNED`, `SCATTERED`) with a regex `WRITE.?IN` that catches every write-in variant in one rule. Real candidate names should never contain the substring `WRITE-IN` (or `WRITE IN`/`WRITEIN`), so the regex is safe.

### 12. Clarity row-ID candidate prefixes (Delaware)

Delaware's Clarity export prefixes candidate names with a numeric row ID: `"(112) MIKE HIGGINS"`, `"(82) MIKE JOHNSON"`. Cross-filed candidates (e.g., a magistrate running on both DEM and REP lines) get *different* row IDs and would otherwise appear as two separate "candidates".

**Fix**: added `_CANDIDATE_ID_PREFIX_RE = r'^\(\d+\)\s*'` to `_normalize_candidate`. Strips the prefix early so the same person collapses to one row regardless of which party line they appeared on.

### 13. Per-county OE conflation bugs (Northumberland Mayor Marion Heights)

Surfaced when a spot-check noticed `MAYOR MARION HEIGHTS BOROUGH` (Northumberland County, 2025 General) in the curated top-races output. The actual November 2025 PDF lists exactly one candidate (DEM John Wargo, 105 votes) — the race was unopposed. But OpenElections' 2025 Northumberland CSV groups three candidates under that office name (Wargo + Joseph Petrovich + John O Lear, all DEM, ~110 votes each). Petrovich and OLear are actually Borough *Council* winners from the next race on the page; OE's parser apparently failed to detect the section boundary between Mayor and Borough Council.

**Fix**: route around the OE bug by using the source PDF directly. Added `filename_exclude_substrs` to `StitchedOpenElectionsCountiesSource` so we can skip specific counties from the OE 2025 stitch; Northumberland 2025 General is then ingested via the existing `ElectionwarePdfSource` pointed at Northumberland's official `overall.pdf`. The direct parse correctly identifies Wargo as the sole listed candidate (filtered as a majority winner) and also surfaces a real finding the OE bug had been masking: the 2025 Sunbury City Council 50/50 tie (Rosancrans D / Ramos R, 752 votes each).

This isn't a fully general fix — other PA counties may have similar OE-parsing bugs. A broader heuristic (drop OE general-election Vote-For-1 races with 3+ same-party candidates, since PA primary law makes that essentially impossible) is deferred until we encounter more cases; it would require preserving party info through the pipeline, a moderate refactor.

### 14. Multi-seat OE races masquerading as Vote-For-1

Surfaced when a spot-check flagged `SCHOOL DIRECTOR DEER LAKES` (Allegheny, 2025) in the curated top-races output. WPRDC has the same race as `"School Director Deer Lakes (Vote For 4)"` — voters pick 4 of 5 candidates and the top 4 win. OpenElections strips the `(Vote For N)` metadata when converting county PDFs to their tidy CSVs, so `_parse_openelections_df` had no way to know it was multi-seat; the five candidates at ~20% each looked like a wide-open Vote For 1 race.

This was inflating the 2025 all-PA workbook significantly: 389 SCHOOL DIRECTOR races and 455 COUNCIL races, most of them multi-seat.

**Fix**: heuristic filter `_is_likely_multiseat_oe_race` that flags office names which are almost always multi-seat in PA — School Director / School Board, County Commissioner (Vote For 2 by law), at-large Borough/City/Township Council. District-numbered races (`"Council District 9"`) are explicitly excluded from the filter since those are Vote For 1. Applied only inside `_parse_openelections_df`; the WPRDC and Clarity sources preserve `(Vote For N)` correctly and don't need this. Dropped the 2025 all-PA workbook from 1,141 to 551 races; the curated `Top_RCV_Races.xlsx` from ~1,350 to ~760.

### 15. Electionware percent-column column-eating bug

When extending the parser from Berks to Chester, we discovered Chester 2021 PDFs include a `VOTE %` column between TOTAL and Election Day (e.g., `"MARIA MCLAUGHLIN  42,302  99.33%  23,107  19,055   140"`). The original regex `(.+?)\s{2,}([\d,]+)…` greedily backtracked: it expanded the non-greedy name capture to include the actual total and the percent, then captured the Election Day count as "TOTAL". The result was ~5,800 corrupt candidate rows from one PDF.

**Fix**: tightened the name capture to `[A-Za-z][^\d%]*?` (must start with a letter, may not contain digits or `%`). The first numeric column after the name is then unambiguously TOTAL. The trailing column class was widened to `[\d,%.]+` to accept the percent column harmlessly. Added a `test_electionware_handles_layout_with_percent_column` regression test so we don't backslide. Also added `TOTAL VOTES CAST` to `_NON_CANDIDATE_NAMES` to drop the per-contest totals row Chester emits.

### 16. System Python SSL cert store sometimes lacks intermediate certs

Surfaced when fetching PDFs directly from county sites (berkspa.gov): plain `urllib.request.urlopen` failed with `SSLCertVerificationError: unable to get local issuer certificate`, while the same `curl` and `pandas.read_csv` calls worked because they use bundled CA stores. The system Python's default trust store didn't accept the county's cert chain.

**Fix**: `_extract_electionware_pdf` builds an `ssl.create_default_context(cafile=certifi.where())` and passes it to `urlopen`. `certifi` is already a transitive dep via `pandas`. Only applies to the PDF download path; other `urllib` calls (GitHub API) happen to work without it.

### 17. Removing write-ins changes majority math

After tightening the write-in filter, some races that previously looked non-majority (e.g., 49.8% vs 49.5% vs 0.7% write-ins) flipped to majority once write-ins were dropped (49.8/99.3 ≈ 50.15%). This is actually the correct IRV interpretation — under IRV, eliminated write-ins don't transfer further, so the remaining candidates compete for the reduced pool. The previously-surfaced 50/50 Delaware races correctly disappeared.

**Note**: this changed a few Allegheny/2025-PA counts as a side effect. Comparing before/after, no real RCV-relevant race was lost — only races where the "non-majority" classification was an artifact of including write-in inflation in the denominator.

## Filtering thresholds and rationale

| Filter | Default | Why |
|---|---|---|
| `filter_non_majority(threshold=50.0)` | leader percent ≤ 50% | Keeps 50/50 ties (canonical RCV case). Original was strict `<` which dropped ties; flipped to `<=`. |
| `filter_min_winner_votes(min_votes=100)` | leader ≥ 100 votes | Drop sub-100-vote precinct judges and other micro-races. |
| `filter_min_leader_percent(min_percent=10.0)` | leader ≥ 10% | Drop write-in-only chaos races. Threshold conservative enough to keep legitimate 6+ way competitive races. |
| `filter_min_candidate_percent(min_percent=1.0)` | candidate row ≥ 1% | Display filter — strips long-tail noise. Doesn't recompute percentages. |
| fuzzy merge `ratio_threshold=0.92`, `min_length=6` | — | Merges `DOUGHHERTY`/`DOUGHERTY` (0.97). Leaves `JOHN`/`JOAN` (0.75), `TRUMP`/`TRUMP/VANCE` (0.55), all short strings alone. |

All are tunable per call. The defaults reflect what we've found works for PA data.

## Notable findings the tool surfaces

Selection of races that came out of the cleanup as genuine RCV-relevant cases:

**Statewide (`Pennsylvania_NonMajority_2018plus.xlsx`)**
- 2020 PA State Senate 45: Jim Brewster 49.97% vs Nicole Ziccarelli 49.91% — actual razor-thin race
- 2024 PA U.S. Senate: McCormick 48.4% / Casey 47.9% / Thomas 1.3% — Libertarian splitter on the canonical decisive race

**Allegheny County (`Allegheny_NonMajority_Local.xlsx`)**
- 2021 Pittsburgh DEM mayoral primary: Gainey 46.4% in a 4-way against incumbent Peduto (the textbook PA local RCV case)
- 2023 Allegheny County Chief Executive DEM: Sara Innamorato 37.6% in a 6-way primary
- 2017 Mt. Lebanon School Director DEM: 7-way, leader at 27.2%

**All-PA 2025 local (`Pennsylvania_NonMajority_Local2025.xlsx`)**
- ~1,140 non-majority races across 67 counties
- Multiple exact 50/50 ties (Auditor Springfield, McConnellsburg Council, Commissioner East Deer Ward 2)
- Many 50.0/49.x school director and council races

**Mid-counties (`Pennsylvania_NonMajority_MidCounties.xlsx`)**
- 2023 York REP Magisterial District Judge 19-3-09: 4-way (Spadaccino 44.5%, Farren 28.3%, Dehart 22.2%, Ruth 5.0%)
- 2023 York REP Township Supervisor Newberry Township: 41.1 / 40.3 / 18.6 — razor-thin race with a third-candidate splitter
- 2023 York REP Township Supervisor West Manheim: 4-way (Hoffman 49.2%, Franks 23.8%, Staaf 14.6%, OConnor 12.5%)
- 2023 York REP Magisterial District Judge 19-2-03: 48.9% / 38.6% / 12.6%
- 2025 York REP Codorus Township Supervisor: Gross 49.7% / Maxwell 39.8% / Bupp 10.1%
- 2025 Delaware DEM Chester Twp Auditor: 34.0% / 33.3% / 32.7% — essentially a 3-way tie
- 2025 Berks DEM Council President City of Reading: Reed 40.7% / Baez Jr 40.4% / Campos 19.0% — extremely close mayoral-equivalent
- 2025 Berks REP Township Supervisor Maxatawny Twp: 4-way (Wilson 28.2% / Weil 26.8% / Reynolds 24.3% / Turner 20.6%) — every candidate within ~8 pts
- 2025 Berks REP Judge of the Court of Common Pleas: Lehman 41.9% / Marks 39.1% / Taylor 18.9%
- 2023 Berks REP Magisterial District Judge District 23-3-07: Book 42.6% / Dye 40.5% / Zimmerman 9.1% / Raup-Konsavage 7.8%
- 2023 Berks DEM School Director Boyertown Area Region 1: 4-way (Arndt 31.2% / Sweisfort 25.8% / Neiman 19.2% / Scott 14.7%)
- 2023 Chester DEM Magisterial District Judge District 15-3-06: 4-way (Hutton 36.7% / Colon 23.2% / McDermott 21.5% / Hashem 11.6%)
- 2023 Chester REP Magisterial District Judge District 15-3-06: 4-way with the same Hashem/McDermott/Colon names (cross-filed) plus Tim Arndt 43.1%
- 2025 Chester REP Township Supervisor West Caln: Mininger 43.1% / Martin 33.8% / Hutton 23.1%
- 2021 Northampton DEM Magisterial District Judge 03-2-09: 4-way at 39.4% / 39.2% / 12.4% / 9.0% — top two within 0.2 points
- 2021 Northampton REP same district: 36.3% / 28.0% / 22.9% / 12.8% — cross-filed counterpart
- 2023 Northampton DEM Hellertown Borough Council 2yr: 4-way at 31.5% / 30.9% / 21.0% / 16.6%
- 2023 Northampton REP Commissioner Bethlehem Twsp 1st Ward: 46.2% / 45.9% / 7.9% — extremely close 2-way with a spoiler
- 2025 Northampton REP Judge of the Court of Common Pleas: Fuller 47.2% / Clark 34.0% / Eyer 18.8%
- 2025 Northampton DEM School Director Northampton Area Region III: Flamisch 47.1% / Gogel 32.9% / Marchiano 20.0%
- 2021 Lebanon DEM Magisterial District Judge 52-2-01: 4-way at 40.0 / 31.8 / 18.3 / 9.9 — and a cross-filed REP counterpart (Capello 37.1 / Itzen 32.8 / Maguire 17.9 / Magaro 8.8) for the same district.
- 2021 Mercer DEM Judge of the Court of Common Pleas: 42.8 / 29.9 / 27.3 — and a cross-filed REP counterpart (McEwen 41.7 / McConnell Jr 40.8 / Joanow 17.5) with the same three candidates in different rank order.
- 2021 Mercer DEM/REP Magisterial District Judge 35-03-02: same three candidates (Straub / Osborne / Gerwick) on both party lines, both non-majority.
- 2023 Northumberland REP Council Region 2: Hetherington 41.4 / Walker 29.7 / Pfeil 28.9.
- 2023 Northumberland REP Supervisor Delaware Twp: Smith 45.4 / Hertzler 38.9 / Heater 15.6.
- 2025 Indiana REP Tax Collector Center Township: 4-way at 38.7 / 38.0 / 17.8 / 5.5 — top two within 0.7 pts.
- 2023 Schuylkill: same five candidates contest both DEM and REP North Schuylkill School District 2-yr primaries (Woodward, Kiehl, Green, Reichwein appear in both), and both come out non-majority — another cross-filed school-board echo of the magistrate pattern.
- 2025 Schuylkill REP Magisterial District Judge 21-3-04: 5-way at 39.3 / 34.6 / 13.8 / 8.1 / 4.2.
- **2021 Lehigh DEM Mayor Allentown** — the canonical PA RCV case: Tuerk 26.6 / O'Connell 25.1 (incumbent) / Guridy 24.9 / Gerlach 23.4. Top four all within 3.2 percentage points; the eventual winner unseated an incumbent by 1.5 pts with just 26.6% support.
- 2025 Lehigh DEM and REP Salisbury School Director 2-yr: another cross-filed pair, with the same three candidates (Glenister, McKelvey, Gnall) competing on both party lines and both primaries coming out non-majority.
- 2023 Lycoming REP Magisterial District Judge 29-3-03: 3-way at Gardner 46.4 / Reitz 44.4 / Sees 9.2 — top two within 2 points.

The cross-filed Magistrate/Judge of the Court of Common Pleas pattern keeps recurring across counties: same candidates compete in both party primaries because PA judicial candidates can file in both. It's now visible in Chester (2023 District 15-3-06), Northampton (2021 District 03-2-09), Lebanon (2021 District 52-2-01), and Mercer (2021 District 35-03-02 + Court of Common Pleas). When neither party clears 50% with these crowded fields, RCV would be especially decisive.

## Out of scope (deferred)

Items we considered and did not pursue, with the reason:

- **2018 PA primary in OpenElections**: only the legacy fixed-width file exists. Would need a separate parser. Low payoff for one year.
- **Pre-2025 odd-year PA data outside Allegheny + York/Delaware**: OpenElections doesn't have 2017/2019/2021/2023 PA municipal cycle data. Each remaining county would be a custom integration. With `LlmPdfSource` proven on Berks, additional counties become tractable: one new registry entry per county PDF.
- **Pre-2023 Berks via PDF**: 2021 Berks publishes only a 5-page statewide-summary PDF for top-level offices; the corresponding local-race data is in a precinct-only PDF that would need a precinct-aware variant of the Electionware parser. Earlier years (2003–2019) likely similar.
- **`LlmPdfSource`**: kept in the codebase as a fallback for PDF layouts that don't match a standard vendor format (e.g., scanned reports, ad-hoc spreadsheets-converted-to-PDF). Not currently invoked from the main script. Requires `ANTHROPIC_API_KEY` if used.
- **Lancaster 2017–2023**: data is in an interactive HTML portal (no file download), would need scraping.
- **Pre-2023 Delaware**: hosted on `election.co.delaware.pa.us` HTML pages, separate scraper.
- **Pre-2023 York**: PDFs on the county's DocumentCenter — drop-in target for `LlmPdfSource` once enabled.
- **Bucks, Lehigh, Montgomery, Dauphin, Cumberland**: all PDF-based; same pattern as Berks. Add via `LlmPdfSource` registry entries.
- **Renaming `primary_scraper.py`** to something more accurate like `rcv_finder.py`: cosmetic; defer.
- **Caching downloaded files between runs**: only LLM responses are cached (`.cache/llm/`); CSV/JSON downloads are cheap enough to refetch.
- **CLI / argparse**: would be nice; not needed — the bottom of the file is editable in seconds.
- **Splitting into a package**: not needed at current scale.
- **Cross-jurisdiction summary sheet**: a single "top-100 PA non-majority races" view across all workbooks would be useful but the user hasn't asked for it.
- **Merging Allegheny + PA 2025 + MidCounties workbooks**: would create some overlap. Defer until asked.

## Curated top-races view

`Top_RCV_Races.xlsx` is generated by `make_top_races_workbook()` after the five per-source workbooks. It reads them all back in, collapses to one row per race (with the leader, runner-up, third candidate, top-two gap, and a full inline candidate list), and writes four ranked views:

- **Closest two-way races** (~1,000 rows): top-two gap ≤ 5 points, sorted by gap ascending. The closest examples are exact ties (e.g., a 4-way borough council where the top three are within fractions of a percent).
- **Crowded fields** (~960 rows): 4+ candidates with the leader under 40%, sorted by leader % ascending. Brings up the cases where the winner has the weakest mandate — a 19-candidate Judge of Election ward race where the leader gets 10.24% sits at the top.
- **Notable local offices** (~656 rows): filtered to mayor / council / DA / sheriff / controller / county exec / commissioner / Court of Common Pleas / magisterial district judge / tax collector / township supervisor. Useful for narrowing to races with broader public relevance.
- **All non-majority races** (~1,350 rows): the full union for completeness.

The 2021 Allentown DEM mayoral primary (Tuerk 26.6% / O'Connell 25.1% / Guridy 24.9% / Gerlach 23.4%) is the headline example in both Closest two-way and Crowded fields. Browse those sheets for many more in the same shape.

## Repository state at time of writing

- `primary_scraper.py` — single-file pipeline (~1700 lines).
- `test_primary_scraper.py` — 98 unit tests, all pure functions, runs in <1s.
- `.gitignore` — excludes the LLM response cache, pytest cache, IDE configs. The output workbooks are now tracked.
- 6 output workbooks committed (Philly, PA statewide, Allegheny local, 2025 PA all-counties, mid-counties, and the curated `Top_RCV_Races.xlsx`).
- Dependencies: `pandas`, `xlsxwriter`, `openpyxl` (core); `pytest` (testing); `pdfplumber` + `certifi` (PDF parsing); `anthropic` (only for `LlmPdfSource`, not currently invoked).

## If we were starting over

A few things we'd do differently with hindsight:

- **Validate the coverage assumption earlier**. We trusted the OE filename `__county.csv` as "complete state rollup" and only caught the 13/67 problem when totals looked wrong. Should have counted distinct counties in every source on first ingest.
- **Treat candidate-name canonicalization as a core concern, not an add-on**. We added it incrementally as data-quality issues surfaced. The right framing from the start would be: every source is going to have inconsistent candidate names; build the canonicalization layer once and apply it everywhere.
- **Output schema first**. We grew the output shape (Candidate/Race_Name/Votes/Percent/Coverage/Year) iteratively. Starting from the desired sheet layout would have made the source classes' contract clearer.

But these are 20/20 hindsight points; the iterative approach worked and the code is in good shape.
