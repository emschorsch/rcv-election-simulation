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
| Lehigh / Bucks / Chester / Delaware / Montgomery county portals | Variable | Local races for each county | ❌ Deferred — PDF-only. Each county is a custom parser. |
| Clarity Elections platform | Allegheny + handful of others | Mid-sized counties' results in CSV | ❌ Ruled out — too few PA counties use it; no public registry. |
| PA Department of State (`electionreturns.pa.gov`) | All 67 counties statewide | Federal/state races only, not local | ❌ Not pursued — duplicates OE statewide coverage; doesn't add local. |
| OpenElections PA pre-2018 | 2000–2016 fixed-width files | Older races | ❌ Deferred — different schema; would need new parser. |

### Why these and not others

We made these choices in roughly the order shown above. The progression was:

1. **Philly first** because the user started there. Five different file shapes across the years revealed the need for the class-per-shape abstraction.
2. **OpenElections second** because it's the obvious "give me PA data" source. The schema is the same across years, so one parser handles 2018+. Coverage is uneven (some years have state rollups, others need stitching from per-county files), so we added the stitching source.
3. **WPRDC third** for Allegheny because the user wanted local races and Allegheny is the second-largest PA jurisdiction with the cleanest data hub outside Philly.
4. **OpenElections 2025 county files fourth** — discovered late in the session that OE has now done the per-county PDF/Excel parsing work for the entire state for 2025, producing tidy CSVs we could ingest with our existing parser. This was the highest-leverage addition: one new source unlocks all 67 PA counties for 2025 local races.

### Why we ruled out the alternatives

- **Clarity Elections**: We initially thought this would be a broad parser (one parser → many PA counties). Investigation showed only Allegheny + a small handful of mid-sized counties use it, and Clarity URLs are per-election numeric IDs with no discovery endpoint. Plus, we already cover Allegheny via WPRDC. Net: low payoff.
- **PDF parsing for Bucks/Lehigh/Delaware/Montgomery**: Each county's PDF layout is different. OpenElections itself maintains 49 county-specific parser scripts to handle PA. That work has been done — we get it for free for 2025 via their `counties/` directory. Doing it ourselves for older years would be substantial effort with low expected yield.
- **State portal (`electionreturns.pa.gov`)**: Doesn't have local races; only adds noise relative to OE which we already use.

## Output workbooks

| Workbook | Source | Layout | Scope |
|---|---|---|---|
| `Philadelphia_Primary_Mayor_DistrictCouncil.xlsx` | vote.phila.gov | One sheet per year (2007–2023) | Philly primaries, mayor/council only |
| `Pennsylvania_NonMajority_2018plus.xlsx` | OpenElections | Two pooled sheets (Primaries, Generals) | PA federal + state, 2018–2024 |
| `Allegheny_NonMajority_Local.xlsx` | WPRDC | Two pooled sheets (Primaries, Generals) | Allegheny County local races, 2017–2025 |
| `Pennsylvania_NonMajority_Local2025.xlsx` | OpenElections `2025/counties/` | One pooled sheet (Generals) | All 67 PA counties local races, 2025 only |

The output files are gitignored. They're regenerated by running `python primary_scraper.py`.

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
- 1,154 non-majority races across 67 counties
- Multiple exact 50/50 ties (Auditor Springfield, McConnellsburg Council, Commissioner East Deer Ward 2)
- Many 50.0/49.x school director and council races

## Out of scope (deferred)

Items we considered and did not pursue, with the reason:

- **2018 PA primary in OpenElections**: only the legacy fixed-width file exists. Would need a separate parser. Low payoff for one year.
- **Pre-2025 odd-year PA data outside Allegheny**: OpenElections doesn't have 2017/2019/2021/2023 PA municipal cycle data. Each county would be a custom integration. Substantial work; current scope (Allegheny + 2025 statewide) covers most of the interesting cases.
- **PDF parsing for Bucks/Lehigh/Delaware/Montgomery**: PDFs vary per county. OpenElections' 2025 county CSVs already cover all of them for 2025. For earlier years, the cost-per-county is high.
- **Renaming `primary_scraper.py`** to something more accurate like `rcv_finder.py`: cosmetic; defer.
- **Caching downloaded files between runs**: not needed at current scale.
- **CLI / argparse**: would be nice; not needed — the bottom of the file is editable in seconds.
- **Splitting into a package**: not needed at current scale.
- **Cross-jurisdiction summary sheet**: a single "top-100 PA non-majority races" view across all workbooks would be useful but the user hasn't asked for it.
- **Merging Allegheny + PA 2025 workbooks**: would create some overlap (Allegheny appears in both) but a consolidated view might be more useful for browsing. Defer until asked.

## Repository state at time of writing

- `primary_scraper.py` — single-file pipeline (~750 lines).
- `test_primary_scraper.py` — 74 unit tests, all pure functions, runs in <1s.
- `.gitignore` — excludes generated workbooks, pytest cache, IDE configs.
- 4 output workbooks generated by running the script.

## If we were starting over

A few things we'd do differently with hindsight:

- **Validate the coverage assumption earlier**. We trusted the OE filename `__county.csv` as "complete state rollup" and only caught the 13/67 problem when totals looked wrong. Should have counted distinct counties in every source on first ingest.
- **Treat candidate-name canonicalization as a core concern, not an add-on**. We added it incrementally as data-quality issues surfaced. The right framing from the start would be: every source is going to have inconsistent candidate names; build the canonicalization layer once and apply it everywhere.
- **Output schema first**. We grew the output shape (Candidate/Race_Name/Votes/Percent/Coverage/Year) iteratively. Starting from the desired sheet layout would have made the source classes' contract clearer.

But these are 20/20 hindsight points; the iterative approach worked and the code is in good shape.
