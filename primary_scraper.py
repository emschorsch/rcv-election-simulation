# pip install requests pandas xlsxwriter openpyxl
"""Find races where no candidate cleared a majority — i.e., races where RCV would
have changed the outcome.

Each input file shape is modeled as an `ElectionSource` subclass that knows how to
turn its file into a tidy `[Candidate, Race_Name, Votes]` DataFrame. The pipeline
below is source-agnostic: add percentages, drop races with a majority winner,
optionally filter by race name, format, and write to a sheet.

Adding a new election: pick the subclass whose file shape matches and append an
instance to the source list at the bottom. If no shape matches, write a new
`ElectionSource` subclass.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import urllib.error
import urllib.request
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Optional

import pandas as pd


# --- Source base class and subclasses ---------------------------------------

@dataclass(kw_only=True)
class ElectionSource(ABC):
    name: str  # used as the Excel sheet name in per-source layout
    url: str = ""  # primary URL; stitched sources leave blank and override fetch_tidy
    write_in_label: str = "write in"
    # Optional metadata, only consulted by the pooled-by-category writer:
    year: Optional[int] = None
    category: Optional[str] = None  # e.g. "primary" or "general" → pooled sheet name
    coverage_note: Optional[str] = None  # e.g. "48 of 67 counties"

    @abstractmethod
    def fetch_tidy(self) -> pd.DataFrame:
        """Return a DataFrame with columns: Candidate, Race_Name, Votes."""


def _wide_columns_to_tidy(
    df: pd.DataFrame,
    vote_cols: list[str],
    *,
    candidate_first: bool,
) -> pd.DataFrame:
    """Sum each vote column and split its name into (Candidate, Race_Name).

    `candidate_first=True`  → column names look like `"CANDIDATE - RACE"`.
    `candidate_first=False` → column names look like `"RACE - CANDIDATE"`; when
    there are >2 dash-separated parts, the first two are joined as the race name
    and the rest as the candidate (matches the 2023 Philly format).
    """
    df = df.copy()
    df[vote_cols] = df[vote_cols].replace({',': ''}, regex=True)
    df[vote_cols] = df[vote_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
    totals = df[vote_cols].sum()

    rows = []
    for col, votes in totals.items():
        if '-' in col:
            parts = [p.strip() for p in col.split('-')]
            if candidate_first:
                candidate = parts[0]
                race_name = ' - '.join(parts[1:])
            elif len(parts) == 2:
                race_name, candidate = parts[0], parts[1]
            else:
                race_name = ' - '.join(parts[:2])
                candidate = ' - '.join(parts[2:])
        else:
            candidate = 'Unknown'
            race_name = col
        rows.append({
            'Candidate': candidate.strip(),
            'Race_Name': race_name.strip(),
            'Votes': votes,
        })
    return pd.DataFrame(rows)


@dataclass(kw_only=True)
class WideCandidateRaceCsvSource(ElectionSource):
    """Wide CSV: one row per division, one column per `"CANDIDATE - RACE"`."""
    id_cols: list[str] = field(default_factory=lambda: ['WARD', 'DIVISION'])

    def fetch_tidy(self) -> pd.DataFrame:
        df = pd.read_csv(self.url)
        vote_cols = [c for c in df.columns if c not in self.id_cols]
        return _wide_columns_to_tidy(df, vote_cols, candidate_first=True)


@dataclass(kw_only=True)
class WideRaceCandidateExcelSource(ElectionSource):
    """Wide Excel: columns are `"RACE - CANDIDATE"`. Drops everything before
    `start_column` (whose own column holds candidate names, not votes). Applies
    `header_replacements` in order before parsing, so jurisdiction-specific
    quirks (newlines, hyphenation variants) are declarative knobs, not code.
    """
    sheet_name: str = 'Totals'
    start_column: str = 'Council'
    header_replacements: list[tuple[str, str]] = field(default_factory=list)

    def fetch_tidy(self) -> pd.DataFrame:
        df = pd.read_excel(self.url, sheet_name=self.sheet_name)
        df.columns = df.columns.str.strip()
        for needle, replacement in self.header_replacements:
            df.columns = df.columns.str.replace(needle, replacement, regex=False)
        start_idx = df.columns.get_loc(self.start_column)
        df_votes = df.iloc[:, start_idx:]
        vote_cols = list(df_votes.columns[1:])
        return _wide_columns_to_tidy(df_votes, vote_cols, candidate_first=False)


@dataclass(kw_only=True)
class LongCategorySelectionExcelSource(ElectionSource):
    """Long Excel: one row per (race, candidate), with a numeric votes column.
    Race name comes from `category_col`; if `party_col` is set the race name is
    `"<CATEGORY> - <PARTY>"` (the 2015 Philly shape).
    """
    votes_col: str
    category_col: str = 'CATEGORY'
    selection_col: str = 'SELECTION'
    party_col: Optional[str] = None
    sheet_name: Optional[str] = None

    def fetch_tidy(self) -> pd.DataFrame:
        kwargs = {'sheet_name': self.sheet_name} if self.sheet_name else {}
        df = pd.read_excel(self.url, **kwargs)
        df.columns = df.columns.str.strip()

        race = df[self.category_col].astype(str).str.strip()
        if self.party_col:
            race = race + " - " + df[self.party_col].astype(str).str.strip()
        candidate = df[self.selection_col].astype(str).str.strip()
        votes = pd.to_numeric(df[self.votes_col], errors='coerce').fillna(0)

        tidy = pd.DataFrame({
            'Candidate': candidate,
            'Race_Name': race,
            'Votes': votes,
        })
        return tidy.groupby(['Race_Name', 'Candidate'], as_index=False)['Votes'].sum()


def _normalize_text(s: pd.Series) -> pd.Series:
    """Upper-case, strip, and collapse internal whitespace. Used for office and
    party fields so case/whitespace differences across counties don't fragment
    a race or party label."""
    return (
        s.fillna('').astype(str)
        .str.upper()
        .str.replace(r'\s+', ' ', regex=True)
        .str.strip()
    )


# Common PA party codes seen as trailing tokens on a candidate name (e.g.
# "DAVE SUNDAY REP" — the same candidate cross-filed on a different ballot line).
_PARTY_CODES = ('DEM', 'REP', 'LIB', 'GRN', 'IND', 'CON', 'WFP', 'SOC', 'NTL',
                'FOR', 'NEN', 'NPP')
_PARTY_SUFFIX_RE = re.compile(r'\s+(?:' + '|'.join(_PARTY_CODES) + r')\s*$')

# Pseudo-"candidates" that are really election-administration metadata. Counting
# them as candidates inflates the race's denominator and falsely lowers every
# real candidate's share, producing false non-majority hits.
_NON_CANDIDATE_NAMES = frozenset({
    'OVER VOTES', 'UNDER VOTES', 'OVERVOTES', 'UNDERVOTES',
    'NOT ASSIGNED', 'SCATTERED',
    'TOTAL VOTES CAST',  # Chester Electionware footer per contest
    'TOTAL VOTES',  # Montgomery per-contest total row (different label, same role)
    'CONTEST TOTALS',    # Mercer / Northumberland: total ballots cast in the contest
    'TIMES BLANK VOTED', 'TIMES OVER VOTED', 'TIMES UNDER VOTED',  # Electionware stat rows
    'TOTAL',  # Lycoming per-contest sum row
    # Bare-token meta rows seen in OE 2025 county.csv rollups (Blair has
    # "Not" with small vote counts in every race — likely "Not Voted" /
    # "Not Cast" mis-extracted from the source PDF; Cumberland's precinct
    # CSV has "OVER" alone in Inspector-of-Elections rows). Adding the
    # singular tokens here is safe because real candidate names are
    # multi-token.
    'NOT', 'OVER', 'UNDER',
})

# Anything matching this is treated as a write-in aggregate/variant and
# dropped — covers `WRITE-IN`, `WRITE-INS`, `WRITE-IN (TOTAL)`,
# `UNASSIGNED WRITE-INS`, `UNRESOLVED WRITE-IN`, `WRITE-IN: SCATTERED`,
# `WRITEIN`, etc. Real candidate names should never contain this substring.
_WRITE_IN_RE = re.compile(r'WRITE.?IN', re.IGNORECASE)


def _is_non_candidate(s: pd.Series) -> pd.Series:
    """Boolean mask: True where the candidate name should be dropped because
    it's election-admin metadata (over/under votes, write-in aggregate, etc.)
    rather than a real listed candidate."""
    return s.isin(_NON_CANDIDATE_NAMES) | s.str.contains(_WRITE_IN_RE, regex=True, na=False)


def _strip_middle_initials(name: str) -> str:
    """Drop single-letter middle tokens. `RYAN E MACKENZIE` → `RYAN MACKENZIE`;
    `CHRISTINA M HARTMAN` → `CHRISTINA HARTMAN`. Leaves the first and last
    tokens alone (so `R KELLY` stays — first token is preserved)."""
    parts = name.split()
    if len(parts) <= 2:
        return name
    middle = [p for p in parts[1:-1] if len(p) > 1]
    return ' '.join([parts[0]] + middle + [parts[-1]])


_CANDIDATE_ID_PREFIX_RE = re.compile(r'^\(\d+\)\s*')


def _normalize_candidate(s: pd.Series) -> pd.Series:
    """Canonicalize candidate names so cross-county variants collapse to one row.

    Steps, in order: upper-case, strip periods/commas/apostrophes, collapse
    whitespace, drop a leading `(NNN)` row-ID prefix (Clarity Delaware lists
    candidates as `"(82) MIKE JOHNSON"` — cross-filed entries get different
    IDs but should be merged), drop a trailing party code (`DAVE SUNDAY REP`
    → `DAVE SUNDAY`), and drop single-letter middle tokens (`RYAN E MACKENZIE`
    → `RYAN MACKENZIE`).

    Does NOT attempt to merge gov/president ticket variants like
    `SHAPIRO / DAVIS` vs `JOSH SHAPIRO` vs `JOSH SHAPIRO AUSTIN DAVIS` — those
    would require an external roster and are flagged as a known limitation.
    """
    out = (
        s.fillna('').astype(str)
        .str.upper()
        .str.replace(r"[.,']", '', regex=True)
        .str.replace(r'\s+', ' ', regex=True)
        .str.strip()
    )
    out = out.str.replace(_CANDIDATE_ID_PREFIX_RE, '', regex=True).str.strip()
    out = out.str.replace(_PARTY_SUFFIX_RE, '', regex=True).str.strip()
    out = out.map(_strip_middle_initials)
    return out


def _is_nickname_variant(name: str, canonical: str) -> bool:
    """True if `name` looks like a nickname / short-form variant of
    `canonical` (or vice versa). The pairs we want to merge:

      * "GREG BOWER"      ↔ "GREGORY BOWER"    (first-name prefix)
      * "CLIFF ACKMAN"    ↔ "CLIFFORD ACKMAN"  (first-name prefix)
      * "KEUERLEBER"      ↔ "RICHARD KEUERLEBER"  (single-token last name)

    Requires the last token (last name) to match exactly; first-name match
    is either prefix (≥3 chars) or implied by the shorter name being a
    single token equal to the longer name's last token. SequenceMatcher
    misses all three of these (ratios 0.71–0.89), so they slip through the
    generic fuzzy threshold.
    """
    n = name.split()
    c = canonical.split()
    if not n or not c or n[-1] != c[-1]:
        return False
    # Single-token (last-name only) vs multi-token: definitely the same person.
    if len(n) == 1 or len(c) == 1:
        return True
    # Both multi-token: require first-name prefix relationship (≥3 chars to
    # avoid collapsing JO/JOHN/JOANNE-type unrelated short names).
    shorter, longer = sorted([n[0], c[0]], key=len)
    return len(shorter) >= 3 and longer.startswith(shorter) and shorter != longer


def fuzzy_canonicalize_candidates(
    df: pd.DataFrame,
    *,
    ratio_threshold: float = 0.92,
    min_length: int = 6,
) -> pd.DataFrame:
    """Merge near-duplicate candidate names *within each race* (likely typos
    like `DOUGHHERTY` vs `DOUGHERTY`, or nickname variants like `GREG BOWER`
    vs `GREGORY BOWER`). Within a race, names are processed in vote-count
    order; the highest-vote spelling becomes canonical, and any later
    spelling that either (a) is within `ratio_threshold` by SequenceMatcher
    or (b) is a nickname/short-form variant per `_is_nickname_variant`
    gets remapped to it.

    `min_length` (default 6) prevents short distinct names like `JOHN`/`JOAN`
    from being collapsed via the SequenceMatcher path. The 0.92 threshold
    leaves gov/president ticket variants like `TRUMP` vs `TRUMP / VANCE`
    (~0.55) alone — fuzzy matching is the wrong tool for that problem.
    """
    if df.empty:
        return df

    parts: list[pd.DataFrame] = []
    for _, group in df.groupby('Race_Name', sort=False):
        group = group.sort_values('Votes', ascending=False)
        mapping: dict[str, str] = {}
        canonicals: list[str] = []
        for name in group['Candidate']:
            if name in mapping:
                continue
            matched = next(
                (c for c in canonicals if _is_nickname_variant(name, c)),
                None,
            )
            if matched is None and len(name) >= min_length:
                matched = next(
                    (c for c in canonicals
                     if SequenceMatcher(None, name, c).ratio() >= ratio_threshold),
                    None,
                )
            if matched is None:
                canonicals.append(name)
                mapping[name] = name
            else:
                mapping[name] = matched
        group = group.copy()
        group['Candidate'] = group['Candidate'].map(mapping)
        parts.append(group)

    combined = pd.concat(parts, ignore_index=True)
    return combined.groupby(['Race_Name', 'Candidate'], as_index=False)['Votes'].sum()


def _parse_openelections_df(
    df: pd.DataFrame,
    *,
    is_primary: bool,
    scope_by_county: bool = False,
) -> pd.DataFrame:
    """Aggregate an OpenElections-format (2018+) tidy CSV into [Candidate, Race_Name, Votes].

    Expected columns: county, office, district, party, candidate, votes (extras
    like precinct/election_day/absentee/mail are ignored). Rows with no candidate
    are dropped — those are meta rows (Registered Voters, Ballots Cast, etc.).
    Race_Name is `<office> <district>` for generals; primaries prepend `<party>`
    so each party's contest becomes its own race. Candidate / office / party
    strings are upper-cased before grouping so case differences across counties
    don't split a single candidate's votes.

    `scope_by_county=True` prepends the county name to every Race_Name. Use
    this for purely-local source data (township / borough / city races)
    where PA has many duplicate municipality names — e.g. Worth Township
    exists in Butler, Centre, Lawrence, and Mercer counties, and without
    scoping all four tax-collector races would be merged into one. Don't
    enable this for state/federal races, which span multiple counties and
    rely on the cross-county merge.

    `vote_for` column (present in OE precinct CSVs from 2025+) is consulted
    authoritatively when available: rows with vote_for > 1 are dropped as
    multi-seat. This replaces our heuristic regex for sources that publish
    the metadata.
    """
    df = df.copy()
    df['_candidate'] = _normalize_candidate(df['candidate'])
    df = df[df['_candidate'] != '']
    df = df[~_is_non_candidate(df['_candidate'])].reset_index(drop=True)

    # Drop multi-seat rows authoritatively when the source publishes vote_for
    # (OE precinct CSVs from 2025 onward). The heuristic regex below catches
    # the rest for older / county-rollup sources that lack this column.
    if 'vote_for' in df.columns:
        vf = pd.to_numeric(df['vote_for'], errors='coerce')
        df = df[(vf.isna()) | (vf <= 1)].reset_index(drop=True)

    office = _normalize_text(df['office'])
    district = df['district'].fillna('').astype(str).str.strip()
    party_col = df['party'] if 'party' in df.columns else pd.Series('', index=df.index)
    party = _normalize_text(party_col)
    county_col = (
        df['county'] if 'county' in df.columns else pd.Series('', index=df.index)
    )
    county = _normalize_text(county_col)

    # Drop rows for office types that always require a district but where the
    # district column is blank — leaving them in would silently merge candidates
    # from many different districts into one bogus race.
    bad = office.isin(_NEEDS_DISTRICT_OFFICES) & (district == '')
    df = df[~bad].reset_index(drop=True)
    office = office[~bad].reset_index(drop=True)
    district = district[~bad].reset_index(drop=True)
    party = party[~bad].reset_index(drop=True)
    county = county[~bad].reset_index(drop=True)

    if is_primary:
        race_name = party + ' ' + office + ' ' + district
    else:
        race_name = office + ' ' + district
    if scope_by_county:
        race_name = county + ' ' + race_name
    race_name = race_name.str.replace(r'\s+', ' ', regex=True).str.strip()

    # Drop races that look like multi-seat — OE source data doesn't preserve
    # Vote For N, so without this filter School Director / County Commissioner
    # / at-large borough council races get treated as Vote For 1 and surface
    # as bogus non-majority RCV findings. The same filter is applied
    # post-pipeline to every source (see filter_likely_multiseat_races);
    # doing it here too saves a downstream pass over the OE rows.
    multiseat = race_name.map(_is_likely_multiseat_race)
    df = df[~multiseat].reset_index(drop=True)
    race_name = race_name[~multiseat].reset_index(drop=True)
    party = party[~multiseat].reset_index(drop=True)

    tidy = pd.DataFrame({
        'Candidate': df['_candidate'],
        'Race_Name': race_name,
        'Party': party,
        'Votes': pd.to_numeric(df['votes'], errors='coerce').fillna(0),
    })
    # Aggregate by (race, candidate). Party is collapsed via mode — when the
    # same candidate appears across multiple counties, take the most common
    # party label they were reported under.
    grouped = (
        tidy.groupby(['Race_Name', 'Candidate'], as_index=False)
        .agg({'Votes': 'sum', 'Party': lambda s: _most_common(s)})
    )

    # General-election sanity checks. Two heuristics, applied only for
    # non-primary OE data:
    #
    # 1) 3+ same-party candidates: under PA primary law each party gets at
    #    most one nominee per single-seat office, so this is essentially
    #    impossible and signals an OE conflation (e.g., Mayor + Borough
    #    Council fused under one office name, as in 2025 Northumberland).
    #
    # 2) Extreme vote-spread (max/min > 20x): real single-seat races have
    #    candidates whose totals are within roughly an order of magnitude
    #    of each other; a 75x spread (as in 2025 Centre's "Mayor UNIONVILLE
    #    BOROUGH" where 20 candidates from clearly-different actual races
    #    got merged) means rows from multiple distinct races were lumped
    #    together. Catches cases where party info is empty so heuristic
    #    (1) can't fire.
    if not is_primary:
        bad_races = (
            _races_with_too_many_same_party(grouped, threshold=3)
            | _races_with_extreme_vote_spread(grouped, max_ratio=20.0)
        )
        if bad_races:
            grouped = grouped[~grouped['Race_Name'].isin(bad_races)]

    return grouped[['Race_Name', 'Candidate', 'Votes']]


def _most_common(s: pd.Series) -> str:
    """Return the most common value in a Series, or '' if empty."""
    m = s.mode()
    return m.iloc[0] if not m.empty else ''


def _races_with_too_many_same_party(
    grouped: pd.DataFrame, *, threshold: int = 3,
) -> set[str]:
    """Set of race names where some non-empty party has `threshold`+ distinct
    candidates. PA general elections legally cap each party at one nominee
    per single-seat office, so this signals upstream conflation.
    """
    bad: set[str] = set()
    for race, g in grouped.groupby('Race_Name'):
        non_empty = g[g['Party'].astype(str) != '']
        if non_empty.empty:
            continue
        counts = non_empty['Party'].value_counts()
        if (counts >= threshold).any():
            bad.add(race)
    return bad


def _races_with_extreme_vote_spread(
    grouped: pd.DataFrame, *,
    max_ratio: float = 20.0,
    min_candidates: int = 5,
) -> set[str]:
    """Set of race names where max candidate votes / min candidate votes
    exceeds `max_ratio` AND the race has at least `min_candidates`. Real
    single-seat races with many candidates have a graceful decay; a 20x+
    spread across 5+ candidates signals multiple distinct races got merged
    into one office name upstream (as in 2025 Centre's Unionville Borough
    case, where 20+ candidates spanned 61–4588 votes).

    The 5-candidate minimum prevents flagging legitimate 2-or-3-candidate
    races with a dominant winner and one minor-party challenger; those can
    easily span 20x without being conflated. Considers only candidates
    with strictly positive vote counts.
    """
    bad: set[str] = set()
    for race, g in grouped.groupby('Race_Name'):
        positive = g[g['Votes'].astype(float) > 0]
        if len(positive) < min_candidates:
            continue
        v_min = float(positive['Votes'].min())
        v_max = float(positive['Votes'].max())
        if v_min > 0 and (v_max / v_min) > max_ratio:
            bad.add(race)
    return bad


def _coverage_from_counties(df: pd.DataFrame, expected: int = 67) -> str:
    """Build a coverage note like `"63 of 67 counties"` from the data itself."""
    n = df['county'].nunique() if 'county' in df.columns else 0
    return f"{n} of {expected} counties"


# Offices that are always reported by district in PA. If a row has one of these
# office names but no district, the row is unusable — leaving it in would
# collapse candidates from many different districts into one bogus race (e.g.,
# Philadelphia's 2024 State House rows have blank district, which would merge
# 27+ unrelated reps into one entry).
_NEEDS_DISTRICT_OFFICES = frozenset({'STATE HOUSE', 'STATE SENATE', 'U.S. HOUSE'})


# OpenElections strips the `(Vote For N)` metadata when converting source PDFs
# to their tidy county CSVs. That means we can't tell multi-seat races from
# single-seat ones in OE data. These office-name patterns are *almost always*
# multi-seat in PA — School Director seats are by region (3-9 candidates
# typically competing for 4 seats); County Commissioner is Vote For 2 by law
# (3 commissioners, vote for at most 2 to ensure minority representation);
# at-large borough/city/township councils are usually Vote For 3+.
#
# A district-numbered race (e.g. "Member of Council District 9") is the
# exception — those are Vote For 1 and we keep them.
# Multi-seat office-name patterns in PA. School Director and School Board
# elect 4+ directors per district (also written "SCH DIR", "SCH DIRS",
# "SCHOOL DIRECTORS"). County Commissioner is Vote For 2 by law.
# "Delegate" (national/state convention) is always 3-5 per district.
# "Study Commission" is a 7-9 member elected body. "SD" is the shorthand
# Lycoming uses for School District in "<district> SD".
_LIKELY_MULTISEAT_KEYWORDS_RE = re.compile(
    r"\b(?:"
    r"(?:school|sch\.?)\s+(?:directors?|dirs?|boards?)"
    r"|county commissioners?"
    r"|delegates?"
    r"|study commissions?"
    r")\b",
    re.IGNORECASE,
)
_COUNCIL_RE = re.compile(r"\bcouncil\b", re.IGNORECASE)
_SD_ABBREV_RE = re.compile(r"\bsd\b", re.IGNORECASE)
# Race name that ends in "<Place> School District" with no further qualifier
# is an at-large school-board election (e.g. Dauphin's 2025 "Halifax Area
# School District" with 5 candidates competing for Vote For 4 seats). A
# race that has a region/seat designator after — "<District> School
# District III", "<District> School District Region 2" — is Vote For 1
# by region and excluded by the _DISTRICT_NUMBERED_RE check above.
_OFFICE_IS_BARE_SCHOOL_DISTRICT_RE = re.compile(
    r"school\s+district\s*$", re.IGNORECASE,
)
# District-numbered race (e.g. "Council District 9") — Vote For 1.
_DISTRICT_NUMBERED_RE = re.compile(r"\bdistrict\s+\d", re.IGNORECASE)


def _is_likely_multiseat_race(race_name: str) -> bool:
    """True if the race name looks like a multi-seat (Vote For N>1) contest.

    Source-agnostic — we apply this universally because OE, WPRDC, and
    Clarity data don't preserve Vote For N, so without this filter
    multi-seat races get treated as Vote For 1 and produce bogus
    non-majority RCV findings.

    Heuristic: any "Council" race in PA without a district number is
    at-large multi-seat (3-5 council seats per borough/township/city,
    voters pick multiple). Plus the specific multi-seat keywords
    (school director(s), county commissioner, delegate, study commission)
    and the "SD" abbreviation Lycoming uses. District-numbered races
    (Vote For 1 by region) are explicitly excluded.
    """
    if _DISTRICT_NUMBERED_RE.search(race_name):
        return False
    if _LIKELY_MULTISEAT_KEYWORDS_RE.search(race_name):
        return True
    if _COUNCIL_RE.search(race_name):
        return True
    if _SD_ABBREV_RE.search(race_name):
        return True
    if _OFFICE_IS_BARE_SCHOOL_DISTRICT_RE.search(race_name):
        return True
    return False


# Backwards-compatible alias for the (previously OE-only) name. Tests and
# OE-internal code still import this; the function itself is source-agnostic.
_is_likely_multiseat_oe_race = _is_likely_multiseat_race


@dataclass(kw_only=True)
class OpenElectionsCsvSource(ElectionSource):
    """Single OpenElections PA tidy CSV (2018+ format).

    Works for state-level rollups (county or precinct) and single-county precinct
    files alike — all share the columns we read by name.
    """
    is_primary: bool = False

    def fetch_tidy(self) -> pd.DataFrame:
        df = pd.read_csv(self.url, dtype=str)
        self.coverage_note = _coverage_from_counties(df)
        return _parse_openelections_df(df, is_primary=self.is_primary)


@dataclass(kw_only=True)
class StitchedOpenElectionsCountiesSource(ElectionSource):
    """Stitch per-county OpenElections CSVs into a single statewide frame.

    Lists files at `listing_api_url` (a GitHub Contents API endpoint), keeps
    files whose name contains `filename_substr` AND ends with `filename_suffix`
    (if set), then fetches each from `raw_base_url + filename` and concatenates.
    Populates `coverage_note` with the actual file count so partial-coverage
    years are visibly flagged.

    `filename_suffix` is needed when a year's `counties/` directory contains
    both `__county.csv` summary rollups and `__precinct.csv` files for the
    same county — without it, both get stitched and votes are double-counted.
    """
    listing_api_url: str = ""
    raw_base_url: str = ""
    filename_substr: str = ""
    filename_suffix: str = ""
    filename_exclude_substrs: tuple[str, ...] = ()
    is_primary: bool = False
    expected_counties: int = 67
    # Prepend the county name to every Race_Name (see _parse_openelections_df).
    # Enable for purely-local sources where PA's many duplicate township /
    # borough names would otherwise be merged across counties.
    scope_by_county: bool = False

    def fetch_tidy(self) -> pd.DataFrame:
        filenames = self._list_matching_files()
        if not filenames:
            self.coverage_note = f"0 of {self.expected_counties} counties (stitched)"
            return pd.DataFrame(columns=['Candidate', 'Race_Name', 'Votes'])

        frames = []
        for f in filenames:
            print(f"  fetching {f}")
            frames.append(pd.read_csv(self.raw_base_url + f, dtype=str))
        combined = pd.concat(frames, ignore_index=True)
        self.coverage_note = (
            f"{_coverage_from_counties(combined, self.expected_counties)} (stitched)"
        )
        return _parse_openelections_df(
            combined,
            is_primary=self.is_primary,
            scope_by_county=self.scope_by_county,
        )

    def _list_matching_files(self) -> list[str]:
        req = urllib.request.Request(
            self.listing_api_url,
            headers={
                'Accept': 'application/vnd.github+json',
                'User-Agent': 'rcv-finder',
            },
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = json.load(resp)
        except urllib.error.URLError as e:
            raise RuntimeError(f"Failed to list {self.listing_api_url}: {e}") from e
        if not isinstance(data, list):
            raise RuntimeError(
                f"Unexpected listing response from {self.listing_api_url}: {data!r}"
            )
        return _filter_oe_listing(
            data, self.filename_substr, self.filename_suffix,
            self.filename_exclude_substrs,
        )


def _filter_oe_listing(
    listing: list[dict],
    substr: str,
    suffix: str,
    exclude_substrs: tuple[str, ...] = (),
) -> list[str]:
    """Filter a GitHub Contents API listing to file entries whose `name`
    contains `substr` AND ends with `suffix` AND doesn't contain any
    `exclude_substrs`. Returns sorted file names.

    `exclude_substrs` lets us skip specific counties whose OE-parsed data
    has known conflation bugs (e.g., the 2025 Northumberland file groups
    Mayor + Borough Council candidates under one office name); we fetch
    those counties from the source PDF directly instead.
    """
    return sorted(
        d['name'] for d in listing
        if d.get('type') == 'file'
           and substr in d.get('name', '')
           and d.get('name', '').endswith(suffix)
           and not any(ex in d.get('name', '') for ex in exclude_substrs)
    )


# Trailing "(Vote For N)" metadata on WPRDC contest names. Multi-seat (N>1)
# races have no single-winner majority concept, so we skip them entirely for
# IRV-style analysis.
_VOTE_FOR_RE = re.compile(r'\s*\(Vote For\s+(\d+)\)\s*$', re.IGNORECASE)
_WPRDC_META_CONTESTS = frozenset({'BALLOTS CAST', 'REGISTERED VOTERS', 'TURNOUT'})


def _parse_wprdc_summary_df(df: pd.DataFrame, *, is_primary: bool) -> pd.DataFrame:
    """Aggregate a WPRDC (Allegheny County) summary CSV into [Candidate, Race_Name, Votes].

    Expected columns: contest_name, choice_name, party_name, total_votes
    (extras ignored). The 2017+ files are pre-aggregated to county totals.

    WPRDC changed its primary-contest naming convention around 2021: newer
    files include the party in the contest name (`DEM Mayor Pittsburgh`),
    older ones don't (`Mayor Pittsburgh`, with party only in `party_name`).
    For primaries we conditionally prepend `party_name` when the contest
    doesn't already start with a known party code — otherwise DEM and REP
    primaries for the same office would merge into a single fake race.

    `(Vote For N)` is stripped from contest names; N>1 (multi-seat) rows are
    dropped because IRV's majority concept doesn't apply. `BALLOTS CAST` /
    `REGISTERED VOTERS` meta-contest rows are dropped. Candidate strings
    pass through the same `_normalize_candidate` used for OpenElections.
    """
    df = df.copy()
    contest = df['contest_name'].fillna('').astype(str)

    # Parse and strip the (Vote For N) tail.
    vf = contest.str.extract(_VOTE_FOR_RE, expand=False)
    vote_for_n = pd.to_numeric(vf, errors='coerce')
    contest_base = contest.str.replace(_VOTE_FOR_RE, '', regex=True).str.strip()

    # Build mask of rows to keep.
    keep = pd.Series(True, index=df.index)
    keep &= (vote_for_n.isna() | (vote_for_n == 1))  # single-seat only
    keep &= ~contest_base.str.upper().isin(_WPRDC_META_CONTESTS)

    df = df[keep].reset_index(drop=True)
    contest_base = contest_base[keep].reset_index(drop=True)

    # For primaries: prepend party_name if not already prefixed (older WPRDC files).
    if is_primary and 'party_name' in df.columns:
        party = df['party_name'].fillna('').astype(str).str.strip().str.upper()
        first_word = contest_base.str.split(' ', n=1).str[0].str.upper()
        already_prefixed = first_word.isin(_PARTY_CODES)
        contest_base = contest_base.where(
            already_prefixed | (party == ''),
            party + ' ' + contest_base,
        )

    candidate = _normalize_candidate(df['choice_name'])
    valid = (candidate != '') & ~_is_non_candidate(candidate)

    tidy = pd.DataFrame({
        'Candidate': candidate[valid].reset_index(drop=True),
        'Race_Name': contest_base[valid].reset_index(drop=True),
        'Votes': pd.to_numeric(df.loc[valid, 'total_votes'], errors='coerce')
                    .fillna(0).reset_index(drop=True),
    })
    return tidy.groupby(['Race_Name', 'Candidate'], as_index=False)['Votes'].sum()


@dataclass(kw_only=True)
class WprdcCsvSource(ElectionSource):
    """Single Allegheny County (WPRDC) summary CSV. The 2017+ files are
    pre-aggregated to county totals — no precinct stitching needed.

    `coverage_note` should be set explicitly at the source instance level
    since WPRDC files don't include a `county` column to derive it from.
    """
    is_primary: bool = False

    def fetch_tidy(self) -> pd.DataFrame:
        df = pd.read_csv(self.url, dtype=str)
        return _parse_wprdc_summary_df(df, is_primary=self.is_primary)


# Map from Clarity's verbose "CAT" category label to the short party code we
# prepend to primary contest names so behavior matches the WPRDC parser.
_CLARITY_CAT_TO_PARTY = {
    'DEMOCRATIC': 'DEM',
    'REPUBLICAN': 'REP',
    'LIBERTARIAN': 'LIB',
    'GREEN': 'GRN',
    'INDEPENDENT': 'IND',
    'CONSTITUTION': 'CON',
}

# Trailing "(DEM)" / "(REP)" / etc. parenthetical on Clarity contest names.
# In York's data this duplicates the CAT field (`CAT="Democratic"`); in
# Luzerne's data it's the ONLY signal of party — Luzerne uses CAT="Results"
# uniformly so we have to pull the party from the contest name. Erie uses
# the full party word ("DEMOCRATIC") instead of the short code.
_CLARITY_PARTY_PAREN_RE = re.compile(
    r'\s*\((DEMOCRATIC|REPUBLICAN|LIBERTARIAN|GREEN|INDEPENDENT|CONSTITUTION'
    r'|' + '|'.join(_PARTY_CODES) + r')\)\s*$',
    re.IGNORECASE,
)
# Long-form party word -> short code, for when a contest name uses the
# verbose party-paren ("(DEMOCRATIC)" -> "DEM ").
_CLARITY_PARTY_WORD_TO_CODE = {
    'DEMOCRATIC': 'DEM', 'REPUBLICAN': 'REP', 'LIBERTARIAN': 'LIB',
    'GREEN': 'GRN', 'INDEPENDENT': 'IND', 'CONSTITUTION': 'CON',
}


def _parse_clarity_summary_json(data: list[dict], *, is_primary: bool) -> pd.DataFrame:
    """Aggregate Clarity Elections summary.json into [Candidate, Race_Name, Votes].

    Each entry in `data` is a contest with `CAT` (category, often a party
    label), `C` (contest name), `VF` (vote-for-N), `CH` (parallel list of
    candidate names), `V` (parallel list of vote counts), and `P` (parallel
    list of per-candidate party tags). Multi-seat contests (VF>1) are dropped
    because IRV's majority concept doesn't apply.

    For primaries we prepend the short party code derived from CAT to the
    contest name when not already prefixed — same conditional-prepend logic
    as the WPRDC parser. The trailing "(DEM)" duplicate is also stripped.

    For primary contests where no candidate has a `P` value matching the
    contest's `CAT` party (i.e., nobody filed for that party's nomination),
    the entire contest is dropped: any "candidates" present are write-ins
    breaking down individual write-in spellings, often the *other* party's
    unopposed incumbent. Auditing the York County 2025 DEM primary
    surfaced this in District Attorney, County Controller, County Coroner,
    Recorder of Deeds, Clerk of Courts, and Sheriff races.
    """
    rows = []
    for contest in data:
        if int(contest.get('VF') or 1) > 1:
            continue  # multi-seat — skip
        contest_name = str(contest.get('C') or '').strip()
        # Pull the party from a trailing "(DEM)"/"(REP)" paren BEFORE stripping
        # — this is the only party signal for counties like Luzerne that use
        # CAT="Results" uniformly. York's CAT="Democratic" still wins below
        # if both are present.
        paren_match = _CLARITY_PARTY_PAREN_RE.search(contest_name)
        if paren_match:
            raw = paren_match.group(1).upper()
            name_party = _CLARITY_PARTY_WORD_TO_CODE.get(raw, raw)
        else:
            name_party = ''
        contest_name = _CLARITY_PARTY_PAREN_RE.sub('', contest_name).strip()
        if not contest_name:
            continue
        cat = str(contest.get('CAT') or '').strip().upper()
        party_prefix = _CLARITY_CAT_TO_PARTY.get(cat, '') or name_party
        parties = contest.get('P') or []
        # Drop primary contests with no filed-for-this-party candidates.
        # All votes are then write-ins (often for the opposite-party
        # incumbent), and the "race" doesn't exist in any real sense.
        if is_primary and party_prefix:
            if not any(str(p).strip().upper() == party_prefix for p in parties):
                continue
        if is_primary and party_prefix:
            first_word = contest_name.split(' ', 1)[0].upper()
            if first_word not in _PARTY_CODES:
                contest_name = f"{party_prefix} {contest_name}"
        choices = contest.get('CH') or []
        votes = contest.get('V') or []
        for choice, v in zip(choices, votes):
            rows.append({
                'choice_name': choice,
                'contest_name': contest_name,
                'votes': v,
            })

    if not rows:
        return pd.DataFrame(columns=['Candidate', 'Race_Name', 'Votes'])

    df = pd.DataFrame(rows)
    candidate = _normalize_candidate(df['choice_name'])
    valid = (candidate != '') & ~_is_non_candidate(candidate)

    tidy = pd.DataFrame({
        'Candidate': candidate[valid].reset_index(drop=True),
        'Race_Name': df.loc[valid, 'contest_name'].reset_index(drop=True),
        'Votes': pd.to_numeric(df.loc[valid, 'votes'], errors='coerce')
                    .fillna(0).reset_index(drop=True),
    })
    return tidy.groupby(['Race_Name', 'Candidate'], as_index=False)['Votes'].sum()


# Browser User-Agent for CloudFront-fronted services (Clarity Elections).
# CloudFront 403s requests with non-browser UAs.
_BROWSER_UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15"
)


@dataclass(kw_only=True)
class ClaritySummaryJsonSource(ElectionSource):
    """Pull election results from a Clarity Elections summary.json endpoint.

    Base URL is e.g. https://results.enr.clarityelections.com/PA/York/117795/
    — we fetch <base>/current_ver.txt to discover the live version, then
    <base>/<ver>/json/en/summary.json. CloudFront blocks non-browser UAs so
    we send a Safari UA on every request.
    """
    is_primary: bool = False
    user_agent: str = _BROWSER_UA

    def fetch_tidy(self) -> pd.DataFrame:
        base = self.url.rstrip('/')
        version = self._get(f"{base}/current_ver.txt").decode().strip()
        data = json.loads(self._get(f"{base}/{version}/json/en/summary.json"))
        return _parse_clarity_summary_json(data, is_primary=self.is_primary)

    def _get(self, url: str) -> bytes:
        req = urllib.request.Request(url, headers={'User-Agent': self.user_agent})
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                return resp.read()
        except urllib.error.URLError as e:
            raise RuntimeError(f"Failed to fetch Clarity URL {url}: {e}") from e


# Cache directory for LLM PDF extractions. First-extraction calls Claude;
# subsequent runs read from disk so reruns are free and offline-capable.
_LLM_CACHE_DIR = Path(__file__).parent / '.cache' / 'llm'


def _llm_cache_path(url: str) -> Path:
    """Stable per-URL cache filename."""
    h = hashlib.sha256(url.encode()).hexdigest()[:16]
    return _LLM_CACHE_DIR / f"{h}.json"


_LLM_EXTRACTION_PROMPT = (
    "Extract every contest and candidate from this election results PDF. "
    "For each (contest, candidate) pair output one row. Skip non-candidate "
    "rows like 'Registered Voters', 'Ballots Cast', 'Over Votes', 'Under "
    "Votes', and aggregated 'Write-In Totals'. For primary elections, the "
    "contest name typically starts with the party (DEM/REP/etc.) — preserve "
    "that prefix. Use the candidate's name exactly as printed. Use the "
    "total vote count summed across all voting methods (in-person + "
    "absentee + mail-in + provisional). If a contest has 'Vote for N' "
    "where N>1, still include all its candidates."
)


def _extract_pdf_via_claude(pdf_url: str, *, model: str = "claude-sonnet-4-6") -> list[dict]:
    """Extract election results from a PDF via Claude with tool_use for
    structured output. Results are cached to .cache/llm/<hash>.json keyed by
    URL — first call hits the API, subsequent calls are free.

    Returns a list of dicts: [{contest_name, candidate, party, votes}, ...].
    """
    cache_path = _llm_cache_path(pdf_url)
    if cache_path.exists():
        return json.loads(cache_path.read_text())

    from anthropic import Anthropic  # lazy import — Phase A doesn't need it
    client = Anthropic()

    tool_schema = {
        "name": "submit_election_results",
        "description": "Submit the extracted election results.",
        "input_schema": {
            "type": "object",
            "properties": {
                "rows": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "contest_name": {"type": "string"},
                            "candidate": {"type": "string"},
                            "party": {"type": "string"},
                            "votes": {"type": "integer"},
                        },
                        "required": ["contest_name", "candidate", "votes"],
                    },
                }
            },
            "required": ["rows"],
        },
    }

    response = client.messages.create(
        model=model,
        max_tokens=16000,
        tools=[tool_schema],
        tool_choice={"type": "tool", "name": "submit_election_results"},
        messages=[{
            "role": "user",
            "content": [
                {"type": "document", "source": {"type": "url", "url": pdf_url}},
                {"type": "text", "text": _LLM_EXTRACTION_PROMPT},
            ],
        }],
    )
    rows = next(
        block.input["rows"]
        for block in response.content
        if block.type == "tool_use"
    )
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    cache_path.write_text(json.dumps(rows, indent=2))
    return rows


def _rows_to_tidy(rows: list[dict], *, is_primary: bool) -> pd.DataFrame:
    """Convert a list of raw {contest_name, candidate, party, votes} dicts
    (from any extractor — LLM, PDF, etc.) into the standard tidy frame.

    Applies the same candidate canonicalization and meta-row filtering used
    elsewhere. If the source forgot to include a party prefix on a primary
    contest, prepend it from the row's `party` field — same conditional
    safety net as the WPRDC parser.
    """
    if not rows:
        return pd.DataFrame(columns=['Candidate', 'Race_Name', 'Votes'])

    df = pd.DataFrame(rows)
    contest_base = df['contest_name'].fillna('').astype(str).str.strip()

    if is_primary and 'party' in df.columns:
        party = df['party'].fillna('').astype(str).str.strip().str.upper()
        first_word = contest_base.str.split(' ', n=1).str[0].str.upper()
        already_prefixed = first_word.isin(_PARTY_CODES)
        contest_base = contest_base.where(
            already_prefixed | (party == ''),
            party + ' ' + contest_base,
        )

    candidate = _normalize_candidate(df['candidate'])
    valid = (candidate != '') & ~_is_non_candidate(candidate)

    tidy = pd.DataFrame({
        'Candidate': candidate[valid].reset_index(drop=True),
        'Race_Name': contest_base[valid].reset_index(drop=True),
        'Votes': pd.to_numeric(df.loc[valid, 'votes'], errors='coerce')
                    .fillna(0).reset_index(drop=True),
    })
    return tidy.groupby(['Race_Name', 'Candidate'], as_index=False)['Votes'].sum()


@dataclass(kw_only=True)
class LlmPdfSource(ElectionSource):
    """Extract election results from a PDF using Claude (tool_use) for
    structured output. Results are cached to .cache/llm/ so reruns don't
    incur API cost. Generic — works for any election-results PDF, though
    tuned for the official county summary PDFs (Berks, etc.)."""
    is_primary: bool = False
    model: str = "claude-sonnet-4-6"

    def fetch_tidy(self) -> pd.DataFrame:
        # Skip gracefully without an API key — the extraction is the only step
        # that hits Anthropic. Cached results in .cache/llm/ still work fine
        # offline, so we only skip when both the cache miss AND the key absent.
        cache_path = _llm_cache_path(self.url)
        if not cache_path.exists() and not os.environ.get('ANTHROPIC_API_KEY'):
            print(f"  skipping {self.name}: ANTHROPIC_API_KEY not set and no cache hit")
            return pd.DataFrame(columns=['Candidate', 'Race_Name', 'Votes'])
        print(f"  extracting via Claude ({self.model})...")
        rows = _extract_pdf_via_claude(self.url, model=self.model)
        return _rows_to_tidy(rows, is_primary=self.is_primary)


# Electionware-PDF parsing (used by Berks, and applicable to many other PA
# counties that publish results via Electionware). The Summary Results
# Report has a highly consistent format:
#
#     <Contest header — possibly party-prefixed>
#     Vote For N
#                              Election
#                     TOTAL    Day     Mail Provisional
#     Candidate Name              <total>  <ed>  <mail>  <prov>
#     ...
#     Write-In Totals             <n>     <n>   <n>     <n>
#      Not Assigned               <n>     <n>   <n>     <n>
#
# We use pdfplumber's layout-preserving text extraction and a small state
# machine to walk the page text, anchoring on each "Vote For N" line.

_ELECTIONWARE_VOTE_FOR_RE = re.compile(r'^\s*Vote For\s+(\d+)\s*$')
# Some Electionware reports (Montgomery County) inline the vote-for-N marker
# inside the contest header line as "(Vote for N)". Capture both the N and
# the contest name itself.
_ELECTIONWARE_INLINE_VF_RE = re.compile(
    r'^(.*?)\s*\(Vote for\s+(\d+)\)\s*$', re.IGNORECASE,
)
# Header tokens (case-insensitive). Covers every variant we've seen across PA
# Electionware-published PDFs: Berks uses {Election, TOTAL, Day, Mail,
# Provisional}; Chester uses Election/TOTAL/Day/Absentee/Mail-In/Provisional;
# some counties have Military columns. We compare lowercased tokens against
# this set.
_ELECTIONWARE_HEADER_TOKENS = frozenset({
    'election', 'total', 'day', 'mail', 'provisional',
    'mail-in', 'absentee', 'absentee/', 'absentee/mail-in',
    'military', 'extra',
    # Northampton wraps the Provisional/Mail headers as
    # "Mail Votes" and "Provisional Votes". Including 'votes' is safe
    # because a real contest title that's only header-set words is
    # extremely unlikely.
    'votes',
})
# Candidate row: indented name (no digits, no % signs — that constraint is
# important; without it, `(.+?)` would happily swallow the percent column
# from Chester-style PDFs and we'd capture the wrong number as TOTAL),
# two-or-more spaces, then 1-6 columns of numbers (digits + commas + optional
# % and . for the VOTE % column some counties include). The first captured
# number is always TOTAL.
#
# Layouts encountered:
#   Berks 2021: TOTAL only
#   Berks 2023+: TOTAL, Election Day, Mail, Provisional
#   Chester 2021: TOTAL, VOTE %, Election Day, Absentee/Mail-In, Provisional
#   Chester 2023+: TOTAL, Election Day, Absentee/Mail-In, Provisional
_ELECTIONWARE_CAND_RE = re.compile(
    r'^\s+([A-Za-z][^\d%]*?)\s{2,}([\d,]+)(?:\s+[\d,%.]+){0,5}\s*$'
)
# Montgomery County variant: candidate rows have the party code (DEM/REP/etc.)
# inline between name and numbers, AND the Total column is LAST instead of
# FIRST. Example: "Daniel McCaffery DEM 37,520 34,002 205 71,727" — 4 numbers,
# last is total. We require at least 2 numbers so we don't false-match a
# header line like "Candidate Party Election Day...".
_PARTY_CODE_GROUP = '|'.join(_PARTY_CODES)
_ELECTIONWARE_CAND_PARTY_RE = re.compile(
    r'^\s*([A-Za-z][^\d%]*?)\s+(?:' + _PARTY_CODE_GROUP + r')'
    r'(?:\s+[\d,]+){1,}\s+([\d,]+)\s*$'
)


def _parse_electionware_lines(lines: list[str]) -> list[dict]:
    """Parse Electionware-formatted PDF text lines into raw result rows.

    Pure function — no I/O. Takes the layout-preserved line list and emits
    a list of `{contest_name, candidate, votes}` dicts. Skips multi-seat
    races (Vote For N>1) entirely. The downstream `_rows_to_tidy` handles
    candidate canonicalization and meta-row filtering.
    """
    rows: list[dict] = []
    i = 0
    while i < len(lines):
        # Two header layouts we've seen across PA Electionware reports:
        # (a) standalone "Vote For N" line below a contest name line (Berks,
        #     Chester, Lehigh, etc.)
        # (b) inline "(Vote for N)" on the contest name line itself
        #     (Montgomery)
        m_standalone = _ELECTIONWARE_VOTE_FOR_RE.match(lines[i])
        m_inline = _ELECTIONWARE_INLINE_VF_RE.match(lines[i])
        if m_standalone:
            vote_for_n = int(m_standalone.group(1))
            # Walk backward (over blank lines and column-header lines) to
            # find the contest header — the most recent non-blank,
            # non-header line. "Column header" = a line whose tokens are
            # all in the known set (Election / TOTAL / Day / Mail /
            # Provisional, in any whitespace arrangement).
            contest_name = None
            for j in range(i - 1, -1, -1):
                stripped = lines[j].strip()
                if not stripped:
                    continue
                tokens = stripped.lower().split()
                if tokens and all(t in _ELECTIONWARE_HEADER_TOKENS for t in tokens):
                    continue
                contest_name = stripped
                break
        elif m_inline:
            contest_name = m_inline.group(1).strip()
            vote_for_n = int(m_inline.group(2))
        else:
            i += 1
            continue
        if not contest_name:
            i += 1
            continue
        # PDF text extraction sometimes preserves column-aligned spacing
        # inside the contest header line ("Lower  Merion Township"). Collapse
        # runs of internal whitespace to a single space so race names group
        # consistently downstream.
        contest_name = re.sub(r'\s+', ' ', contest_name).strip()

        # Multi-seat: skip the whole contest.
        if vote_for_n > 1:
            i += 1
            continue

        # Scan forward for candidate rows; stop at the next contest header
        # (either layout).
        i += 1
        while i < len(lines):
            line = lines[i]
            if (_ELECTIONWARE_VOTE_FOR_RE.match(line)
                or _ELECTIONWARE_INLINE_VF_RE.match(line)):
                break
            # Try Montgomery's "Name PARTY ... Total" layout first — if
            # there's a party code between name and numbers, the row is
            # Total-column-last. Otherwise fall back to the Berks-style
            # "Name [PARTY] Total ED Mail Prov" layout (Total first).
            cand_match = _ELECTIONWARE_CAND_PARTY_RE.match(line)
            if cand_match:
                name = cand_match.group(1).strip()
                total = int(cand_match.group(2).replace(',', ''))
                rows.append({
                    'contest_name': contest_name,
                    'candidate': name,
                    'votes': total,
                })
            else:
                cand_match = _ELECTIONWARE_CAND_RE.match(line)
                if cand_match:
                    name = cand_match.group(1).strip()
                    total = int(cand_match.group(2).replace(',', ''))
                    rows.append({
                        'contest_name': contest_name,
                        'candidate': name,
                        'votes': total,
                    })
            i += 1
    return rows


def _pdf_url_to_lines(pdf_url: str) -> list[str]:
    """Download a PDF and return its text as a list of layout-preserved lines.

    Shared helper used by every PDF source. `certifi` provides the CA
    bundle explicitly because the system Python's default trust store
    doesn't accept some county sites' cert chains. We send a friendly
    `rcv-finder` User-Agent — CloudFront-fronted sites are handled
    separately by `ClaritySummaryJsonSource` with a browser UA.
    """
    import ssl
    import tempfile
    import certifi  # lazy imports — only needed when a PDF source is used
    import pdfplumber

    ctx = ssl.create_default_context(cafile=certifi.where())
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
        req = urllib.request.Request(
            pdf_url, headers={'User-Agent': 'rcv-finder'}
        )
        try:
            with urllib.request.urlopen(req, timeout=60, context=ctx) as resp:
                tmp.write(resp.read())
        except urllib.error.URLError as e:
            # Some county servers (e.g. lehighcounty.org as of mid-2026)
            # omit the intermediate cert from their TLS chain, so neither
            # certifi nor the macOS system store can verify them — `curl`
            # works because of its caching of intermediates from prior
            # connections. Public election PDFs aren't credentialed, so
            # fall back to an unverified SSL context with a warning.
            if not (isinstance(e.reason, ssl.SSLCertVerificationError)):
                raise
            print(f"  WARNING: cert verification failed for {pdf_url}; "
                  f"retrying with unverified SSL context")
            unverified = ssl._create_unverified_context()
            with urllib.request.urlopen(req, timeout=60, context=unverified) as resp:
                tmp.write(resp.read())
        tmp_path = tmp.name

    all_lines: list[str] = []
    with pdfplumber.open(tmp_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text(layout=True)
            if text:
                all_lines.extend(text.split('\n'))
    return all_lines


def _extract_electionware_pdf(pdf_url: str) -> list[dict]:
    """Download an Electionware-format PDF and extract raw result rows."""
    return _parse_electionware_lines(_pdf_url_to_lines(pdf_url))


@dataclass(kw_only=True)
class ElectionwarePdfSource(ElectionSource):
    """Extract election results from an Electionware-format Summary Results
    Report PDF using pdfplumber + regex. No LLM, no API key. Works for
    Berks and many other PA counties that publish via Electionware
    (Bedford, Blair, Bradford, Carbon, Centre, Chester, Crawford, Elk,
    Erie, Franklin, etc. all use the same vendor)."""
    is_primary: bool = False

    def fetch_tidy(self) -> pd.DataFrame:
        print(f"  extracting Electionware PDF via pdfplumber...")
        rows = _extract_electionware_pdf(self.url)
        return _rows_to_tidy(rows, is_primary=self.is_primary)


# Lycoming's PDFs use a different vendor format. Each contest is announced
# on ONE line with its name, party-in-parens, vote-for-N, and turnout stats:
#     "District Attorney (Dem) (Vote for 1), 18804 registered voters, ..."
# Candidate rows follow, indented, with columns:
#     <name>  <total votes>  <vote%>%  <election day>  <mail in>  <provisional>
# A "Total" row closes each contest (sum of candidates — dropped via
# `_NON_CANDIDATE_NAMES` since 'TOTAL' is now in the literal set).

# Partisan contest header: name + party in parens + "Vote for N", followed
# by a comma and turnout stats.
_LYCOMING_CONTEST_RE = re.compile(
    r'^\s*(.+?)\s+\(([A-Za-z]+)\)\s+\(Vote for (\d+)\),'
)
# Any contest header (partisan or non-partisan ballot question). Used to
# detect when a *new* race starts so we don't keep bundling YES/NO votes
# from a ballot question into the previously-seen partisan contest.
_LYCOMING_ANY_CONTEST_RE = re.compile(
    r'^\s*(.+?)\s+\(Vote for (\d+)\),'
)
_LYCOMING_CAND_RE = re.compile(
    r'^\s+([A-Za-z][^\d%]*?)\s+([\d,]+)\s+[\d.]+%\s+[\d,]+\s+[\d,]+\s+[\d,]+\s*$'
)


def _parse_lycoming_pdf_lines(lines: list[str]) -> list[dict]:
    """Parse Lycoming-format PDF text lines into raw result rows.

    Pure function — no I/O. Skips multi-seat races (Vote for N>1) and
    non-partisan contests like ballot questions (the YES/NO rows would
    otherwise be miscredited to whatever partisan contest preceded them).
    Contest names are emitted with the party prefix from the parens,
    matching the convention used by `_rows_to_tidy` for primaries.
    """
    rows: list[dict] = []
    current_contest: Optional[str] = None
    for line in lines:
        any_match = _LYCOMING_ANY_CONTEST_RE.match(line)
        if any_match:
            partisan = _LYCOMING_CONTEST_RE.match(line)
            if partisan:
                base = partisan.group(1).strip()
                party = partisan.group(2).strip().upper()
                vote_for = int(partisan.group(3))
                current_contest = None if vote_for > 1 else f"{party} {base}"
            else:
                # Non-partisan contest (ballot question, etc.) — drop entirely
                current_contest = None
            continue
        if current_contest is None:
            continue
        cm = _LYCOMING_CAND_RE.match(line)
        if cm:
            name = cm.group(1).strip()
            total = int(cm.group(2).replace(',', ''))
            rows.append({
                'contest_name': current_contest,
                'candidate': name,
                'votes': total,
            })
    return rows


@dataclass(kw_only=True)
class LycomingPdfSource(ElectionSource):
    """Extract election results from Lycoming County's PDF format.

    Different vendor than Electionware — uses inline contest headers
    (`<contest> (Dem) (Vote for 1), <stats>`) and candidate rows with
    a percent column. Multi-seat races are skipped.
    """
    is_primary: bool = False

    def fetch_tidy(self) -> pd.DataFrame:
        print("  extracting Lycoming PDF via pdfplumber...")
        rows = _parse_lycoming_pdf_lines(_pdf_url_to_lines(self.url))
        return _rows_to_tidy(rows, is_primary=self.is_primary)


# Bucks County uses an "ElectionSource"-vendor certified results PDF whose
# layout is distinct from both Electionware and Lycoming:
#
#   Contest header:  "<Office Name> - D (Dem) (Vote for N)"   (party letter + paren)
#   Candidate row:   "<NAME>  <Total>  <ED>  <MI>  <PR>"      (Total FIRST, single-
#                                                              space separator)
#   Closing row:     "Total  <T>  <ED>  <MI>  <PR>"           (dropped by NCN filter)
#
# Multi-word names like "BRANDON NEUMAN" or "KAREN M.S. KRIEGER" mean the
# name-and-numbers boundary needs the trailing four numbers as the anchor.
# Non-greedy `.+?` then captures whatever's before them as the name.

_BUCKS_CONTEST_RE = re.compile(
    r'^(.+?)\s+-\s+([DR])\s+\((?:Dem|Rep)\)\s+\(Vote for\s+(\d+)\)\s*$',
    re.IGNORECASE,
)
# Any contest header (partisan or ballot-question / non-partisan), used to
# reset current_contest so YES/NO votes from a referendum don't leak into
# the preceding partisan race.
_BUCKS_ANY_CONTEST_RE = re.compile(r'\(Vote for\s+\d+\)\s*$', re.IGNORECASE)
_BUCKS_CAND_RE = re.compile(
    r'^(.+?)\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s*$'
)


def _parse_bucks_pdf_lines(lines: list[str]) -> list[dict]:
    """Parse Bucks PDF text lines into raw result rows. Pure function.

    Skips multi-seat (Vote for N>1) contests entirely. Contest names are
    emitted with a "DEM "/"REP " prefix so they match the convention used
    by `_rows_to_tidy` and the multi-seat-race filter downstream. Ballot
    questions / non-partisan contests are recognized by their "(Vote for
    N)" header (without the partisan " - D (Dem) " suffix) and used only
    to reset current_contest — their YES/NO rows are discarded.
    """
    rows: list[dict] = []
    current_contest: Optional[str] = None
    for line in lines:
        stripped = line.strip()
        m = _BUCKS_CONTEST_RE.match(stripped)
        if m:
            office, party_letter, vf_str = m.groups()
            vote_for_n = int(vf_str)
            if vote_for_n > 1:
                current_contest = None
                continue
            prefix = 'DEM' if party_letter.upper() == 'D' else 'REP'
            current_contest = f"{prefix} {office.strip()}"
            continue
        # Non-partisan / ballot-question contest header — reset so the
        # following YES/NO rows aren't credited to the prior partisan race.
        if _BUCKS_ANY_CONTEST_RE.search(stripped):
            current_contest = None
            continue
        if current_contest is None:
            continue
        cm = _BUCKS_CAND_RE.match(stripped)
        if cm:
            name = cm.group(1).strip()
            total = int(cm.group(2).replace(',', ''))
            rows.append({
                'contest_name': current_contest,
                'candidate': name,
                'votes': total,
            })
    return rows


# Lehigh County's pre-2021 primary PDFs use a custom Election-Source-like
# layout that doesn't match Electionware or the modern Bucks parser:
#
#   <Jurisdiction>             e.g. "CITY OF ALLENTOWN"
#   <Office>                   e.g. "MAYOR"
#   (Vote for One - 4 Year Term)
#   <PARTY> <Name> <ED> Ballot <Mail> <Prov> <Other> <Total>
#   <PARTY> <Name> <WriteIn-tot> Write-In <Total>
#
# The contest header spans 1-2 lines preceding the "(Vote for X)" line.
# "Vote for One/Two/Three/..." uses a word in place of a digit. Multi-
# seat contests are skipped at the caller.
_LEHIGH_OLD_VOTE_FOR_RE = re.compile(
    r"\(Vote for\s+(?:not more than\s+)?(\w+)", re.IGNORECASE,
)
_LEHIGH_OLD_CAND_RE = re.compile(
    r"^\s*(" + _PARTY_CODE_GROUP + r")\s+([A-Z][^0-9]*?)\s+"
    r"(\d[\d,]*)\s+Ballot\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s*$"
)
_LEHIGH_OLD_WRITEIN_RE = re.compile(
    r"^\s*(" + _PARTY_CODE_GROUP + r")\s+([A-Z][^0-9]*?)\s+"
    r"(\d[\d,]*)\s+Write-In(?:\s+\d[\d,]*){0,3}\s+(\d[\d,]*)\s*$"
)
# "Empty" candidate rows — placeholder lines like "DEM     0 Ballot 0 0 0 0"
# that appear when no candidate filed for that party. They aren't real
# candidates and don't match _LEHIGH_OLD_CAND_RE (which requires a name);
# we recognize them here so the state machine doesn't mistake them for
# jurisdiction headers.
_LEHIGH_OLD_EMPTY_CAND_RE = re.compile(
    r"^\s*(" + _PARTY_CODE_GROUP + r")\s+\d", re.IGNORECASE,
)
# Section-divider lines that group contests by category (e.g. "BOROUGHS"
# divides borough contests from township ones). These are NOT jurisdictions
# — the actual borough/township name appears on the line right after, and
# subsequent contests inherit that jurisdiction until the next divider.
_LEHIGH_OLD_SECTION_HEADERS = frozenset({
    'BOROUGHS', 'TOWNSHIPS', 'WARDS', 'MAGISTERIAL DISTRICTS',
    'COUNTY WIDE', 'COUNTYWIDE', 'DISTRICT WIDE', 'JUDICIAL OFFICES',
})


def _parse_lehigh_old_pdf_lines(lines: list[str]) -> list[dict]:
    """Parse Lehigh's pre-2021 primary PDF format into raw result rows.

    The file groups contests by municipality:

        CITY OF ALLENTOWN          (jurisdiction)
        MAYOR                      (office)
        (Vote for One - 4 Year Term)
        ... candidates ...
        CITY COUNCIL               (office; same jurisdiction inherited)
        (Vote for ...)
        ...
        BOROUGHS                   (section divider)
        ALBURTIS                   (new jurisdiction)
        MAYOR                      (office)
        (Vote for ...)
        ...

    Anchors on each "(Vote for X)" line; takes the immediately-preceding
    non-blank line as the office. The line BEFORE that is either:
      (a) a candidate row from a previous contest in the same jurisdiction
          — keep the tracked jurisdiction, OR
      (b) a new jurisdiction header — update the tracked jurisdiction.
    Section dividers like "BOROUGHS" are skipped explicitly.

    Multi-seat contests (N>1) are dropped. Both "PARTY Name X Ballot ED
    Mail Prov Total" and the shorter "PARTY Name X Write-In Total" rows
    are recognized.
    """
    rows: list[dict] = []
    current_jurisdiction = ""
    i = 0
    while i < len(lines):
        m = _LEHIGH_OLD_VOTE_FOR_RE.search(lines[i])
        if not m:
            i += 1
            continue
        raw = m.group(1).lower()
        try:
            vote_for_n = int(raw)
        except ValueError:
            vote_for_n = _NUMBER_WORDS.get(raw, 1)

        # Walk back: most recent non-blank line is the office.
        j = i - 1
        while j >= 0 and not lines[j].strip():
            j -= 1
        if j < 0:
            i += 1
            continue
        office = lines[j].strip()
        # Sanity: a candidate row would be misidentified as office; bail.
        if (_LEHIGH_OLD_CAND_RE.match(lines[j])
            or _LEHIGH_OLD_WRITEIN_RE.match(lines[j])):
            i += 1
            continue

        # Walk back further: the line above the office is either a new
        # jurisdiction header (if this is the first contest in a new
        # municipality) or a candidate row from the previous contest
        # (in which case we keep the inherited jurisdiction).
        k = j - 1
        while k >= 0 and not lines[k].strip():
            k -= 1
        if k >= 0:
            above = lines[k].strip()
            is_candidate_row = (
                _LEHIGH_OLD_CAND_RE.match(lines[k])
                or _LEHIGH_OLD_WRITEIN_RE.match(lines[k])
                or _LEHIGH_OLD_EMPTY_CAND_RE.match(lines[k])
            )
            if (not is_candidate_row
                and above.upper() not in _LEHIGH_OLD_SECTION_HEADERS):
                current_jurisdiction = above

        if vote_for_n > 1 or not office:
            i += 1
            continue

        contest_office = (f"{current_jurisdiction} {office}"
                          if current_jurisdiction else office)

        # Walk forward, collect candidate rows.
        i += 1
        while i < len(lines):
            line = lines[i]
            if _LEHIGH_OLD_VOTE_FOR_RE.search(line):
                break
            cm = _LEHIGH_OLD_CAND_RE.match(line)
            if cm:
                rows.append({
                    'contest_name': f"{cm.group(1)} {contest_office}",
                    'candidate': cm.group(2).strip(),
                    'votes': int(cm.group(7).replace(',', '')),
                })
            else:
                wm = _LEHIGH_OLD_WRITEIN_RE.match(line)
                if wm:
                    rows.append({
                        'contest_name': f"{wm.group(1)} {contest_office}",
                        'candidate': wm.group(2).strip(),
                        'votes': int(wm.group(4).replace(',', '')),
                    })
            i += 1
    return rows


@dataclass(kw_only=True)
class LehighOldPdfSource(ElectionSource):
    """Extract pre-2021 Lehigh County primary results from their custom PDF
    layout ("CITY OF ALLENTOWN / MAYOR / (Vote for One) / PARTY Name X
    Ballot ED Mail Prov Total" rows). Used specifically to capture the
    canonical 2017 Allentown DEM mayoral primary (Pawlowski 28.4% in a
    7-way contest that motivated Allentown's later RCV interest)."""
    is_primary: bool = True

    def fetch_tidy(self) -> pd.DataFrame:
        print(f"  extracting Lehigh-old PDF via pdfplumber...")
        rows = _parse_lehigh_old_pdf_lines(_pdf_url_to_lines(self.url))
        return _rows_to_tidy(rows, is_primary=self.is_primary)


@dataclass(kw_only=True)
class BucksPdfSource(ElectionSource):
    """Extract election results from Bucks County's "ElectionSource"-vendor
    PDF. Different layout from Electionware / Lycoming: Total column FIRST,
    single-space separator between name and numbers, party encoded in the
    contest header rather than per-candidate."""
    is_primary: bool = False

    def fetch_tidy(self) -> pd.DataFrame:
        print("  extracting Bucks PDF via pdfplumber...")
        rows = _parse_bucks_pdf_lines(_pdf_url_to_lines(self.url))
        return _rows_to_tidy(rows, is_primary=self.is_primary)


# Lancaster County publishes 2021/2023 (and earlier) election results as a
# hierarchical HTML site: a SelectParty page → per-party CatX index pages
# (CatC = countywide, CatJ = judiciary, CatS = schools, CatL = Lancaster
# City, CatB = boroughs, CatT = townships, CatP = statewide) → individual
# race-result pages with vote-count tables. The leaf pages have a stable
# layout — bold contest name, "Vote for not more than N" indicator, and
# a 5-column table (Candidate / Election Day / Mail In / Provisional /
# Total). The TOTAL column is the LAST per row.
#
# 2025+ results moved to a JavaScript SPA whose CDN-hosted PDFs aren't
# downloadable without an in-browser session; that year is deferred.

_LANC_RACE_LINK_RE = re.compile(r"href='([^']*/(\d+)\.html)'", re.IGNORECASE)
_LANC_CONTEST_RE = re.compile(
    r"<td[^>]*font-weight:\s*bold[^>]*>([^<]+)</td>", re.IGNORECASE,
)
_LANC_VOTE_FOR_RE = re.compile(
    r"Vote\s+for\s+not\s+more\s+than\s+(\w+)", re.IGNORECASE,
)
# Candidate row in the bordered table: 5 cells (name + 4 vote totals).
# The name <td> may contain an <a> tag for write-in / by-precinct links.
_LANC_ROW_RE = re.compile(
    r"<tr>\s*"
    r"<td[^>]*>(?:<[^>]+>)?\s*([^<]+?)\s*(?:</[^>]+>)?\s*</td>\s*"
    r"<td[^>]*align='right'[^>]*>([\d,]+)</td>\s*"
    r"<td[^>]*align='right'[^>]*>([\d,]+)</td>\s*"
    r"<td[^>]*align='right'[^>]*>([\d,]+)</td>\s*"
    r"<td[^>]*align='right'[^>]*>([\d,]+)</td>\s*"
    r"</tr>",
    re.IGNORECASE | re.DOTALL,
)
_NUMBER_WORDS = {
    'one': 1, 'two': 2, 'three': 3, 'four': 4, 'five': 5, 'six': 6,
    'seven': 7, 'eight': 8, 'nine': 9, 'ten': 10,
}


def _parse_lancaster_race_page(html: str) -> tuple[Optional[str], int, list[tuple[str, int]]]:
    """Parse one Lancaster race-result HTML page into (contest_name, vote_for_n, [(candidate, votes), ...]).

    Returns (None, 0, []) if the page isn't a race-result page (e.g. category
    index pages have no bold contest header). Multi-seat races still return
    their data; the caller is responsible for skipping them.
    """
    cm = _LANC_CONTEST_RE.search(html)
    if not cm:
        return None, 0, []
    contest = cm.group(1).strip()
    vote_for_n = 1
    vm = _LANC_VOTE_FOR_RE.search(html)
    if vm:
        raw = vm.group(1).lower()
        try:
            vote_for_n = int(raw)
        except ValueError:
            vote_for_n = _NUMBER_WORDS.get(raw, 1)
    rows: list[tuple[str, int]] = []
    for rm in _LANC_ROW_RE.finditer(html):
        name = rm.group(1).strip()
        total = int(rm.group(5).replace(',', ''))
        rows.append((name, total))
    return contest, vote_for_n, rows


def _walk_lancaster_year(year_path: str, *, max_workers: int = 8) -> list[dict]:
    """Walk Lancaster's per-party CatX index pages and return raw result rows.

    `year_path` is the URL segment that identifies the election, e.g.
    "May_16,_2023_-_Municipal_Primary". We visit each party's index pages
    (DEMCategories.html, etc.), follow the links to per-category index
    pages, then to individual race pages, and parse each. Multi-seat
    contests are skipped. Write-in detail pages and by-precinct pages
    have URLs with letters after the number (e.g. "11WriteIn.html") and
    are filtered out by the integer-only race-id regex.
    """
    from concurrent.futures import ThreadPoolExecutor

    base = f"https://vr.co.lancaster.pa.us/ElectionReturns/{year_path}/"
    # Walk index pages serially (small fan-out) then race pages in parallel.

    def fetch(url: str) -> str:
        req = urllib.request.Request(url, headers={'User-Agent': 'rcv-finder'})
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.read().decode('utf-8', errors='ignore')

    parties = [('DEM', 'DEM'), ('REP', 'REP'), ('NON', '')]
    cats = ['C', 'J', 'S', 'L', 'B', 'T']  # skip 'P' = statewide (already covered)
    # Boroughs / Townships / Schools have a second level: each
    # municipality has its own sub-index page (e.g. "DEMCatBADAMSTOWN.html")
    # before the race-result pages. Detect those by a "<letter>Cat<c>{NAME}.html"
    # link in the category index, fetch each, then collect race-page links.
    race_urls: list[tuple[str, str]] = []  # (party_prefix, race_url)
    seen: set[str] = set()
    for letter, prefix in parties:
        try:
            cat_index_html = fetch(f"{base}{letter}Categories.html")
        except urllib.error.HTTPError:
            continue
        for c in cats:
            top_cat_re = re.compile(
                r"href='[^']*/" + letter + "Cat" + c + r"\.html'", re.IGNORECASE,
            )
            if not top_cat_re.search(cat_index_html):
                continue
            queue = [f"{base}{letter}Cat{c}.html"]
            visited_idx: set[str] = set()
            while queue:
                idx_url = queue.pop()
                if idx_url in visited_idx:
                    continue
                visited_idx.add(idx_url)
                try:
                    idx_html = fetch(idx_url)
                except urllib.error.HTTPError:
                    continue
                # Sub-category pages (e.g. DEMCatBADAMSTOWN.html) link to
                # additional sub-indices — queue them for traversal.
                sub_cat_re = re.compile(
                    r"href='([^']*/" + letter + "Cat" + c + r"[A-Z0-9_]+\.html)'",
                    re.IGNORECASE,
                )
                for sm in sub_cat_re.finditer(idx_html):
                    queue.append(sm.group(1))
                # Numeric race-page links.
                for m in _LANC_RACE_LINK_RE.finditer(idx_html):
                    race_url = m.group(1)
                    if race_url in seen:
                        continue
                    seen.add(race_url)
                    race_urls.append((prefix, race_url))

    print(f"  walking {len(race_urls)} Lancaster race pages...")
    rows: list[dict] = []

    def parse_one(item):
        prefix, url = item
        try:
            html = fetch(url)
        except urllib.error.HTTPError:
            return []
        contest, vote_for_n, cand_rows = _parse_lancaster_race_page(html)
        if not contest or vote_for_n > 1:
            return []
        contest_name = f"{prefix} {contest}".strip()
        return [{'contest_name': contest_name, 'candidate': n, 'votes': v}
                for n, v in cand_rows]

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        for batch in ex.map(parse_one, race_urls):
            rows.extend(batch)
    return rows


@dataclass(kw_only=True)
class LancasterHtmlSource(ElectionSource):
    """Scrape Lancaster County's hierarchical HTML election-returns site.

    Walks the per-party CatX index pages (Countywide, Judiciary, Schools,
    City, Boroughs, Townships) and parses each leaf race page. 2025+
    results moved to a JS-rendered SPA and aren't accessible this way —
    only 2019/2021/2023 (and earlier) elections work.
    """
    year_path: str = ""  # e.g. "May_16,_2023_-_Municipal_Primary"
    is_primary: bool = False

    def fetch_tidy(self) -> pd.DataFrame:
        print(f"  scraping Lancaster HTML site ({self.year_path})...")
        rows = _walk_lancaster_year(self.year_path)
        return _rows_to_tidy(rows, is_primary=self.is_primary)


# --- Pipeline ---------------------------------------------------------------

def add_percentages(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df['Votes'] = pd.to_numeric(df['Votes'], errors='coerce').fillna(0)
    df['Total_Votes_Race'] = df.groupby('Race_Name')['Votes'].transform('sum')
    df['Percent'] = df['Votes'] / df['Total_Votes_Race'] * 100
    return df


def filter_non_majority(df: pd.DataFrame, threshold: float = 50.0) -> pd.DataFrame:
    # Keep races where the leader did not *exceed* the threshold — a 50/50 tie
    # is a non-majority outcome RCV would have to resolve, so it stays in.
    max_percent = df.groupby('Race_Name')['Percent'].transform('max')
    return df[max_percent <= threshold].copy()


def filter_min_winner_votes(df: pd.DataFrame, min_votes: int = 100) -> pd.DataFrame:
    """Drop races where the top vote-getter received fewer than `min_votes`,
    so tiny contests (precinct judges, sparsely-reported races, etc.) don't
    clutter the output."""
    max_votes = df.groupby('Race_Name')['Votes'].transform('max')
    return df[max_votes >= min_votes].copy()


def filter_min_leader_percent(df: pd.DataFrame, min_percent: float = 10.0) -> pd.DataFrame:
    """Drop races whose top finisher got less than `min_percent` of the vote.

    Such races are almost always write-in-only contests where no candidate was
    listed on the ballot — the "field" is just a sprawl of write-in names, none
    of which has a real shot. They aren't meaningfully RCV-relevant (there's no
    candidate field for IRV to operate on) and they crowd the output."""
    max_percent = df.groupby('Race_Name')['Percent'].transform('max')
    return df[max_percent >= min_percent].copy()


def filter_min_candidate_percent(df: pd.DataFrame, min_percent: float = 1.0) -> pd.DataFrame:
    """Drop individual candidate rows below `min_percent`. Display filter only —
    percentages are not recomputed, so the surviving rows still reflect each
    candidate's share of the *full* race total."""
    return df[df['Percent'] >= min_percent].copy()


def filter_race_names(df: pd.DataFrame, pattern: Optional[str]) -> pd.DataFrame:
    if pattern is None:
        return df
    mask = df['Race_Name'].str.lower().str.contains(pattern, regex=True, na=False)
    return df[mask].copy()


def filter_exclude_race_names(df: pd.DataFrame, pattern: Optional[str]) -> pd.DataFrame:
    """Symmetric to filter_race_names: drop races whose name matches `pattern`
    (case-insensitive regex). Used e.g. to drop President from the PA workbook
    because cross-county ticket fragmentation makes its totals unreliable."""
    if pattern is None:
        return df
    mask = df['Race_Name'].str.lower().str.contains(pattern, regex=True, na=False)
    return df[~mask].copy()


def filter_rcv_useful_races(
    df: pd.DataFrame, *, min_candidates: int = 3,
) -> pd.DataFrame:
    """Narrow to races where RCV would have been genuinely useful.

    Two conditions, both must hold:

    1. At least `min_candidates` candidates survive the upstream
       per-candidate threshold (`filter_min_candidate_percent`, default
       1%). Two-candidate races — including 50/50 ties — get dropped
       because RCV can't redistribute anything with only two choices.

    2. The top two candidates aren't exactly tied. A 50/50 outcome is
       a different problem (resolved by coin flip / marble draw under
       PA election code); RCV doesn't help unless there's a third
       candidate whose voters' second choices could break the tie,
       which our data doesn't currently surface.
    """
    if df.empty or 'Race_Name' not in df.columns:
        return df
    bad: set[str] = set()
    for race, g in df.groupby('Race_Name', sort=False):
        if g['Candidate'].nunique() < min_candidates:
            bad.add(race)
            continue
        if 'Percent' in g.columns:
            top = sorted(g['Percent'].dropna(), reverse=True)
            if len(top) >= 2 and top[0] == top[1]:
                bad.add(race)
    return df[~df['Race_Name'].isin(bad)].copy()


def filter_write_in_fragmentation(
    df: pd.DataFrame, *, max_ratio: float = 3.0,
) -> pd.DataFrame:
    """Drop races where the leader has more than `max_ratio` times the
    runner-up's vote share. Real competitive non-majority primaries have
    leader/runner-up < 2 (Tuerk 26.6 / O'Connell 25.1 = 1.06; Parker
    32.7 / Rhynhart 22.8 = 1.43); ratios above 3 reliably signal an
    uncontested party primary where voters wrote in various spellings
    of the unopposed other-party incumbent's name.

    Audited examples that this filter drops (all York County 2023/2025
    DEM primaries for offices where the Republican incumbent ran
    unopposed in the general):

      * DEM Recorder of Deeds 2025 — Shue 39.96 / Riston 1.45 (ratio 27.6)
      * DEM Clerk of Courts 2023 — Byrnes 32.75 / Supler 3.77 (ratio 8.7)
      * DEM Sheriff 2023 — Keuerleber 29.84 / Becker 4.03 (ratio 7.4)
      * DEM County Coroner 2025 — Zech 43.37 / Gay 7.40 (ratio 5.9)
      * DEM District Attorney 2025 — Barker 35.92 / Graybill 11.29 (ratio 3.2)
      * DEM County Controller 2025 — Bower 21.56 / Ackman 9.32 (ratio 3.2,
        post-canonicalization of GREG/GREGORY BOWER + CLIFF/CLIFFORD ACKMAN)
    """
    if df.empty or 'Race_Name' not in df.columns or 'Percent' not in df.columns:
        return df
    bad: set[str] = set()
    for race, g in df.groupby('Race_Name', sort=False):
        sorted_pct = sorted(g['Percent'].dropna(), reverse=True)
        if len(sorted_pct) >= 2 and sorted_pct[1] > 0:
            if sorted_pct[0] / sorted_pct[1] > max_ratio:
                bad.add(race)
    return df[~df['Race_Name'].isin(bad)].copy()


def filter_likely_multiseat_races(df: pd.DataFrame) -> pd.DataFrame:
    """Drop races whose names look like Vote For N>1 contests.

    Source-agnostic — applied to every parsed source because none of our
    data sources (OE, WPRDC, Clarity, Electionware/Boyer PDFs, Berks LLM
    extracts) preserve Vote For N. Without this filter, school director,
    county commissioner, at-large council, delegates, and study-commission
    races get treated as Vote For 1 and surface as bogus non-majority
    findings (e.g. "leader has 23.78%" for a 5-candidate Vote For 4 race).
    """
    if df.empty or 'Race_Name' not in df.columns:
        return df
    mask = df['Race_Name'].astype(str).map(_is_likely_multiseat_race)
    return df[~mask].copy()


def format_for_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Return Candidate/Race_Name/Votes/Percent sorted with a blank row between races."""
    df = df[['Candidate', 'Race_Name', 'Votes', 'Percent']].copy()
    df = df.sort_values(['Race_Name', 'Votes'], ascending=[True, False])

    blank = pd.DataFrame([{'Candidate': '', 'Race_Name': '', 'Votes': '', 'Percent': ''}])
    chunks: list[pd.DataFrame] = []
    for _, group in df.groupby('Race_Name', sort=True):
        chunks.append(group)
        chunks.append(blank)

    if not chunks:
        return pd.DataFrame(columns=['Candidate', 'Race_Name', 'Votes', 'Percent'])
    return pd.concat(chunks, ignore_index=True)


def write_workbook(
    sources: list[ElectionSource],
    out_path: str,
    *,
    race_pattern: Optional[str] = None,
    race_exclude_pattern: Optional[str] = None,
    threshold: float = 50.0,
    min_winner_votes: int = 100,
    min_leader_percent: float = 10.0,
    min_candidate_percent: float = 1.0,
    fuzzy_merge: bool = True,
) -> None:
    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        for source in sources:
            print(f"Processing {source.name}...")
            tidy = source.fetch_tidy()
            if fuzzy_merge:
                tidy = fuzzy_canonicalize_candidates(tidy)
            tidy = add_percentages(tidy)
            tidy = filter_non_majority(tidy, threshold=threshold)
            tidy = filter_min_winner_votes(tidy, min_votes=min_winner_votes)
            tidy = filter_min_leader_percent(tidy, min_percent=min_leader_percent)
            tidy = filter_min_candidate_percent(tidy, min_percent=min_candidate_percent)
            tidy = filter_race_names(tidy, race_pattern)
            tidy = filter_exclude_race_names(tidy, race_exclude_pattern)
            tidy = filter_likely_multiseat_races(tidy)
            tidy = filter_write_in_fragmentation(tidy)
            tidy = filter_rcv_useful_races(tidy)
            sheet = format_for_sheet(tidy)
            sheet.to_excel(writer, sheet_name=source.name, index=False)


_POOLED_COLS = ['Year', 'Race_Name', 'Candidate', 'Votes', 'Percent', 'Coverage']


def format_pooled_for_sheet(
    items: list[tuple[ElectionSource, pd.DataFrame]],
) -> pd.DataFrame:
    """Pool already-filtered races from multiple sources into one sheet.

    Adds `Year` and `Coverage` columns (sourced from each `ElectionSource`),
    sorts by Year asc, Race_Name asc, Votes desc, and inserts a blank separator
    row between each (Year, Race_Name) group.
    """
    blocks = []
    for source, df in items:
        if df.empty:
            continue
        df = df[['Race_Name', 'Candidate', 'Votes', 'Percent']].copy()
        df.insert(0, 'Year', source.year)
        df['Coverage'] = source.coverage_note or ''
        blocks.append(df[_POOLED_COLS])

    if not blocks:
        return pd.DataFrame(columns=_POOLED_COLS)

    combined = pd.concat(blocks, ignore_index=True).sort_values(
        ['Year', 'Race_Name', 'Votes'],
        ascending=[True, True, False],
        kind='mergesort',
    )

    blank = pd.DataFrame([{c: '' for c in _POOLED_COLS}])
    chunks: list[pd.DataFrame] = []
    for _, group in combined.groupby(['Year', 'Race_Name'], sort=False):
        chunks.append(group)
        chunks.append(blank)
    return pd.concat(chunks, ignore_index=True)


def write_workbook_pooled_by_category(
    sources: list[ElectionSource],
    out_path: str,
    *,
    race_pattern: Optional[str] = None,
    race_exclude_pattern: Optional[str] = None,
    threshold: float = 50.0,
    min_winner_votes: int = 100,
    min_leader_percent: float = 10.0,
    min_candidate_percent: float = 1.0,
    fuzzy_merge: bool = True,
) -> None:
    """Group sources by `category`; emit one sheet per category, all years pooled.

    Sources must set both `category` and `year`. `coverage_note` is propagated
    onto every row as a `Coverage` column so partial-coverage sources are
    visibly flagged in the output.
    """
    by_cat: dict[str, list[tuple[ElectionSource, pd.DataFrame]]] = {}
    for source in sources:
        if not source.category or source.year is None:
            raise ValueError(
                f"Source {source.name!r} needs both `category` and `year` for pooled output"
            )
        print(f"Processing {source.name}...")
        tidy = source.fetch_tidy()
        if fuzzy_merge:
            tidy = fuzzy_canonicalize_candidates(tidy)
        tidy = add_percentages(tidy)
        tidy = filter_non_majority(tidy, threshold=threshold)
        tidy = filter_min_winner_votes(tidy, min_votes=min_winner_votes)
        tidy = filter_min_leader_percent(tidy, min_percent=min_leader_percent)
        tidy = filter_min_candidate_percent(tidy, min_percent=min_candidate_percent)
        tidy = filter_race_names(tidy, race_pattern)
        tidy = filter_exclude_race_names(tidy, race_exclude_pattern)
        tidy = filter_likely_multiseat_races(tidy)
        tidy = filter_write_in_fragmentation(tidy)
        tidy = filter_rcv_useful_races(tidy)
        race_count = tidy['Race_Name'].nunique() if not tidy.empty else 0
        print(f"  -> {race_count} non-majority races")
        by_cat.setdefault(source.category, []).append((source, tidy))

    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        for category in sorted(by_cat):
            sheet_df = format_pooled_for_sheet(by_cat[category])
            sheet_df.to_excel(writer, sheet_name=category, index=False)


# --- Source registry --------------------------------------------------------
# URLs from https://vote.phila.gov/resources-data/past-election-results/archived-data-sets/

PHILLY_PRIMARY_SOURCES: list[ElectionSource] = [
    WideCandidateRaceCsvSource(
        name="2007",
        url="https://vote.phila.gov/files/election-results/2007_Primary_WD.csv",
    ),
    WideCandidateRaceCsvSource(
        name="2011",
        url="https://vote.phila.gov/files/election-results/2011_Primary_WD.csv",
    ),
    LongCategorySelectionExcelSource(
        name="2015",
        url="https://vote.phila.gov/files/election-results/2015_PRIMARY/2015_PRIMARY_-_WARD_RESULTS_-_PUBLIC.xlsx",
        votes_col="TOTAL",
        party_col="PARTY",
    ),
    # 2017 — Larry Krasner's 7-way DEM District Attorney primary, also
    # Rebecca Rhynhart's challenge to incumbent City Controller Butkovitz.
    # The XLSX has three sheets; the "Precinct-level Data" sheet is the
    # long-format one. OFFICE column already includes the party suffix
    # ("DISTRICT ATTORNEY-DEM"), so party_col stays unset.
    LongCategorySelectionExcelSource(
        name="2017",
        url="https://files7.philadelphiavotes.com/election-results/"
            "2017_PRIMARY/2017_PRIMARY_-_RESULTS_.xlsx",
        sheet_name="Precinct-level Data",
        category_col="OFFICE",
        selection_col="CANDIDATE",
        votes_col="VOTES",
    ),
    LongCategorySelectionExcelSource(
        name="2019",
        url="https://vote.phila.gov/files/election-results/2019_PRIMARY/2019_PRIMARY_RESULTS_BY_TYPE_BY_PRECINCT.xlsx",
        sheet_name="2019 PRIMARY",
        votes_col="VOTE COUNT",
    ),
    WideRaceCandidateExcelSource(
        name="2023",
        url="https://vote.phila.gov/media/2023_Primary_Results.xlsx",
        header_replacements=[
            ("\n", " - "),
            ("AT-LARGE", "AT LARGE"),
            ("COUNCIL - ", "COUNCIL"),
            ("Write-in", "write in"),
        ],
    ),
]


# OpenElections PA (2018+). Files at https://github.com/openelections/openelections-data-pa
# Skipped: 2018 primary (only legacy fixed-width format), 2022 primary (no files).

_OE_RAW = "https://raw.githubusercontent.com/openelections/openelections-data-pa/master/"
_OE_API = "https://api.github.com/repos/openelections/openelections-data-pa/contents/"

OPENELECTIONS_PA_SOURCES: list[ElectionSource] = [
    # 2018: state-level county rollup (67/67 verified).
    OpenElectionsCsvSource(
        name="2018 PA General", year=2018, category="Generals",
        url=_OE_RAW + "2018/20181106__pa__general__county.csv",
    ),
    # 2020: county-rollup file is only 13/67; use the precinct file (67/67).
    OpenElectionsCsvSource(
        name="2020 PA General", year=2020, category="Generals",
        url=_OE_RAW + "2020/20201103__pa__general__precinct.csv",
    ),
    # 2024: state precinct rollup (63/67 at time of writing).
    OpenElectionsCsvSource(
        name="2024 PA General", year=2024, category="Generals",
        url=_OE_RAW + "2024/20241105__pa__general__precinct.csv",
    ),
    StitchedOpenElectionsCountiesSource(
        name="2022 PA General", year=2022, category="Generals", is_primary=False,
        listing_api_url=_OE_API + "2022/counties",
        raw_base_url=_OE_RAW + "2022/counties/",
        filename_substr="__general__",
    ),
    StitchedOpenElectionsCountiesSource(
        name="2020 PA Primary", year=2020, category="Primaries", is_primary=True,
        listing_api_url=_OE_API + "2020/counties",
        raw_base_url=_OE_RAW + "2020/counties/",
        filename_substr="__primary__",
    ),
    StitchedOpenElectionsCountiesSource(
        name="2024 PA Primary", year=2024, category="Primaries", is_primary=True,
        listing_api_url=_OE_API + "2024/counties",
        raw_base_url=_OE_RAW + "2024/counties/",
        filename_substr="__primary__",
    ),
]


# Allegheny County (Pittsburgh) local races, from the Western PA Regional Data
# Center: https://data.wprdc.org/dataset/election-results. Off-cycle odd-year
# primaries+generals are where PA local races (DA, mayor, council, etc.) live.

_WPRDC_DUMP = "https://data.wprdc.org/datastore/dump/"
_AC_COVERAGE = "Allegheny County (Pittsburgh + boroughs)"

ALLEGHENY_LOCAL_SOURCES: list[ElectionSource] = [
    WprdcCsvSource(name="2017 Allegheny Primary", year=2017, category="Primaries",
                   is_primary=True, coverage_note=_AC_COVERAGE,
                   url=_WPRDC_DUMP + "abfa5c73-4152-4d37-b8d4-c7df517c3b49"),
    WprdcCsvSource(name="2017 Allegheny General", year=2017, category="Generals",
                   coverage_note=_AC_COVERAGE,
                   url=_WPRDC_DUMP + "542c655c-b3f5-43ec-8626-e2cd9253af9e"),
    WprdcCsvSource(name="2019 Allegheny Primary", year=2019, category="Primaries",
                   is_primary=True, coverage_note=_AC_COVERAGE,
                   url=_WPRDC_DUMP + "153b6de5-06f5-4061-a417-384331b842c3"),
    WprdcCsvSource(name="2019 Allegheny General", year=2019, category="Generals",
                   coverage_note=_AC_COVERAGE,
                   url=_WPRDC_DUMP + "91c50f7e-8b49-4835-9fef-b7b3f4f0ce82"),
    WprdcCsvSource(name="2021 Allegheny Primary", year=2021, category="Primaries",
                   is_primary=True, coverage_note=_AC_COVERAGE,
                   url=_WPRDC_DUMP + "1d37bc6b-b5ab-41ca-a006-da7de348b883"),
    WprdcCsvSource(name="2021 Allegheny General", year=2021, category="Generals",
                   coverage_note=_AC_COVERAGE,
                   url=_WPRDC_DUMP + "73b61240-e3c5-444c-bd1b-68d19d5eb4d5"),
    WprdcCsvSource(name="2023 Allegheny Primary", year=2023, category="Primaries",
                   is_primary=True, coverage_note=_AC_COVERAGE,
                   url=_WPRDC_DUMP + "4b38a2aa-09e8-4cbe-9fdd-40fd84012178"),
    WprdcCsvSource(name="2023 Allegheny General", year=2023, category="Generals",
                   coverage_note=_AC_COVERAGE,
                   url=_WPRDC_DUMP + "6010ffdd-6d2d-4962-94eb-24d38b0a245b"),
    WprdcCsvSource(name="2025 Allegheny Primary", year=2025, category="Primaries",
                   is_primary=True, coverage_note=_AC_COVERAGE,
                   url=_WPRDC_DUMP + "6051ca15-fb5d-425b-890c-a25c8c9b0c73"),
    WprdcCsvSource(name="2025 Allegheny General", year=2025, category="Generals",
                   coverage_note=_AC_COVERAGE,
                   url=_WPRDC_DUMP + "5b094e63-1659-45d9-b249-34cf3a06c50f"),
]


# 2025 PA general statewide local-races source: stitched from OpenElections's
# per-county tidy CSVs. They published one `__county.csv` summary per PA county
# for the 2025 municipal cycle, with all 67 counties present. The 2017/2019/
# 2021/2023 odd-year cycles aren't in OpenElections (verified — those year
# directories don't exist), so this is single-year coverage for the rest of PA.
# Allegheny's 2025 data also appears here (alongside its WPRDC copy in the
# Allegheny workbook); cross-validation is useful, not a bug.

PA_LOCAL_2025_SOURCES: list[ElectionSource] = [
    StitchedOpenElectionsCountiesSource(
        name="2025 PA Local (All Counties)",
        year=2025, category="Generals", is_primary=False,
        listing_api_url=_OE_API + "2025/counties",
        raw_base_url=_OE_RAW + "2025/counties/",
        filename_substr="__general__",
        # Use precinct files (not county.csv rollups): OE's 2025 county
        # rollups concatenate office+district into one buggy `office`
        # string ("Supervisor JUNIATA TWP"), which merges adjacent races
        # whose section boundaries the PDF→CSV converter mis-detected
        # (e.g., Township Supervisor + Township Auditor in Blair County).
        # The precinct files keep office and district properly separated
        # and additionally include a `vote_for` column we use to drop
        # multi-seat races authoritatively.
        filename_suffix="__precinct.csv",
        # Prepend county name to every Race_Name. PA has many duplicate
        # local-government names (Worth Twp in 4 counties, Springfield in
        # 7+, Juniata Twp in 4+) and without scoping all of those races
        # collapse into one bogus pseudo-race.
        scope_by_county=True,
        # Skip counties whose OE-parsed data has known conflation bugs;
        # those counties are integrated separately via their source PDFs
        # (see ELECTIONWARE_PDF_SOURCES). Northumberland: OE's 2025 file
        # groups Mayor + Borough Council candidates under one office name.
        filename_exclude_substrs=("__northumberland__",),
    ),
]


# Clarity Elections sources for mid-sized PA counties. Election IDs are
# per-county per-election numeric identifiers Clarity assigns; discovered
# from each county's official results page. The base URL pattern is
# https://results.enr.clarityelections.com/PA/<County>/<id>/.

_CLARITY_BASE = "https://results.enr.clarityelections.com/PA/"

CLARITY_PA_SOURCES: list[ElectionSource] = [
    # York County (R-leaning) — 3 primaries on Clarity. Strong source of
    # competitive REP primary races for bipartisan motivating examples.
    ClaritySummaryJsonSource(
        name="2023 York Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="York County (city + boroughs + townships)",
        url=_CLARITY_BASE + "York/117795/",
    ),
    ClaritySummaryJsonSource(
        name="2024 York Primary", year=2024, category="Primaries", is_primary=True,
        coverage_note="York County (city + boroughs + townships)",
        url=_CLARITY_BASE + "York/120831/",
    ),
    ClaritySummaryJsonSource(
        name="2025 York Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="York County (city + boroughs + townships)",
        url=_CLARITY_BASE + "York/123816/",
    ),
    # Delaware County (D-leaning Philly suburb) — 2024-2025 on Clarity. Pre-2024
    # is HTML-based on a different subdomain; deferred.
    ClaritySummaryJsonSource(
        name="2024 Delaware Primary", year=2024, category="Primaries", is_primary=True,
        coverage_note="Delaware County (Media + boroughs + townships)",
        url=_CLARITY_BASE + "Delaware/120839/",
    ),
    ClaritySummaryJsonSource(
        name="2024 Delaware General", year=2024, category="Generals",
        coverage_note="Delaware County (Media + boroughs + townships)",
        url=_CLARITY_BASE + "Delaware/122488/",
    ),
    ClaritySummaryJsonSource(
        name="2025 Delaware Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Delaware County (Media + boroughs + townships)",
        url=_CLARITY_BASE + "Delaware/123824/",
    ),
    ClaritySummaryJsonSource(
        name="2025 Delaware General", year=2025, category="Generals",
        coverage_note="Delaware County (Media + boroughs + townships)",
        url=_CLARITY_BASE + "Delaware/125145/",
    ),
    # Westmoreland County (suburban Pittsburgh) — full Clarity archive back
    # to 2019. Contest names already include the "DEM "/"REP " prefix in the
    # `C` field; the parser leaves them alone.
    ClaritySummaryJsonSource(
        name="2019 Westmoreland Primary", year=2019, category="Primaries", is_primary=True,
        coverage_note="Westmoreland County (city + boroughs + townships)",
        url=_CLARITY_BASE + "Westmoreland/95683/",
    ),
    ClaritySummaryJsonSource(
        name="2021 Westmoreland Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Westmoreland County (city + boroughs + townships)",
        url=_CLARITY_BASE + "Westmoreland/109366/",
    ),
    ClaritySummaryJsonSource(
        name="2023 Westmoreland Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Westmoreland County (city + boroughs + townships)",
        url=_CLARITY_BASE + "Westmoreland/117764/",
    ),
    ClaritySummaryJsonSource(
        name="2025 Westmoreland Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Westmoreland County (city + boroughs + townships)",
        url=_CLARITY_BASE + "Westmoreland/123823/",
    ),
    # Luzerne County (Wilkes-Barre / Hazleton). Contest names use a trailing
    # "(DEM)" / "(REP)" parenthetical and CAT="Results" uniformly, so the
    # parser pulls party from the paren before stripping it.
    ClaritySummaryJsonSource(
        name="2019 Luzerne Primary", year=2019, category="Primaries", is_primary=True,
        coverage_note="Luzerne County (Wilkes-Barre + boroughs + townships)",
        url=_CLARITY_BASE + "Luzerne/95689/",
    ),
    ClaritySummaryJsonSource(
        name="2021 Luzerne Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Luzerne County (Wilkes-Barre + boroughs + townships)",
        url=_CLARITY_BASE + "Luzerne/109364/",
    ),
    ClaritySummaryJsonSource(
        name="2025 Luzerne Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Luzerne County (Wilkes-Barre + boroughs + townships)",
        url=_CLARITY_BASE + "Luzerne/123839/",
    ),
    # Erie County — Clarity covers 2024+. Pre-2024 primaries are on the
    # county website as Electionware-format PDFs (parsed by the same
    # ElectionwarePdfSource — see below).
    ClaritySummaryJsonSource(
        name="2025 Erie Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Erie County (city + boroughs + townships)",
        url=_CLARITY_BASE + "Erie/123825/",
    ),
    # Cambria County (Johnstown + suburbs). Contest names already include the
    # "DEM "/"REP " prefix (Westmoreland-style), so the parser handles them
    # without further changes. Pre-2022 results aren't on Clarity.
    ClaritySummaryJsonSource(
        name="2023 Cambria Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Cambria County (Johnstown + boroughs + townships)",
        url=_CLARITY_BASE + "Cambria/117757/",
    ),
    ClaritySummaryJsonSource(
        name="2025 Cambria Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Cambria County (Johnstown + boroughs + townships)",
        url=_CLARITY_BASE + "Cambria/123841/",
    ),
    # --- General elections for the Clarity counties. Odd-year generals
    # (2021, 2023) are valuable because OpenElections doesn't cover those
    # years for PA — these are the only place local mayor/council/sup-
    # ervisor head-to-head races surface.
    ClaritySummaryJsonSource(
        name="2021 Westmoreland General", year=2021, category="Generals",
        coverage_note="Westmoreland County (city + boroughs + townships)",
        url=_CLARITY_BASE + "Westmoreland/111510/",
    ),
    ClaritySummaryJsonSource(
        name="2023 Westmoreland General", year=2023, category="Generals",
        coverage_note="Westmoreland County (city + boroughs + townships)",
        url=_CLARITY_BASE + "Westmoreland/119030/",
    ),
    ClaritySummaryJsonSource(
        name="2023 Cambria General", year=2023, category="Generals",
        coverage_note="Cambria County (Johnstown + boroughs + townships)",
        url=_CLARITY_BASE + "Cambria/119020/",
    ),
    # Fayette County (Uniontown + southwestern PA). 2025 primary is on
    # Clarity (Luzerne-style trailing party paren); pre-2025 primaries
    # are Electionware PDFs on the county DocumentCenter.
    ClaritySummaryJsonSource(
        name="2025 Fayette Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Fayette County (Uniontown + boroughs + townships)",
        url=_CLARITY_BASE + "Fayette/123836/",
    ),
]


# PA counties that publish their Summary Results PDF via the Electionware
# vendor share a consistent format that `ElectionwarePdfSource` parses
# with pdfplumber + regex (no API key needed). Many PA counties use
# Electionware per OpenElections' parser scripts (Berks, Chester, Centre,
# Bedford, Blair, Bradford, Carbon, Crawford, Elk, Franklin, etc.).

_BERKS_PDF_BASE = "https://www.berkspa.gov/getmedia/"
_CHESCO_PDF_BASE = "https://www.chesco.org/DocumentCenter/View/"
_NORCO_PDF_BASE = "https://www.norcopa.gov/corecode/uploads/document6/uploaded_pdfs/corecode/"
_CENTRE_PDF_BASE = "https://centrecountypa.gov/DocumentCenter/View/"
_LEBCO_PDF_BASE = "https://www.lebanoncountypa.gov/getmedia/"
_MERCER_PDF_BASE = "https://www.mercercountypa.gov/election/Election.Results/"
_NUMCO_PDF_BASE = "https://www.northumberlandcountypa.gov/htdocs/wp-content/uploads/documents/elections/"
_SCHU_PDF_BASE = "https://schuylkillcountypa.gov/Document_Center/Departments/Election%20Bureau/"
_LEHIGH_PDF_BASE = "https://www.lehighcounty.org/Portals/0/PDF/"

# 2021 Berks isn't here: their 2021 "Grand Totals" PDF is statewide-only
# (judges + ballot questions, 5 pages, no local contests), and the
# corresponding local-race data lives in the separate "By Precinct" PDF
# which would need a precinct-aware parser. 2023+ Berks publishes a single
# unified summary PDF that includes local contests at county totals.

ELECTIONWARE_PDF_SOURCES: list[ElectionSource] = [
    # Berks County (Reading + surrounding boroughs/townships)
    ElectionwarePdfSource(
        name="2023 Berks Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Berks County (Reading + boroughs + townships)",
        url=_BERKS_PDF_BASE + "bd209dc6-bd5c-431c-9072-1abf4e337e94/"
            "PE23-General-Summary-6-28.pdf",
    ),
    ElectionwarePdfSource(
        name="2025 Berks Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Berks County (Reading + boroughs + townships)",
        url=_BERKS_PDF_BASE + "d3291f1d-85d9-4ab8-9447-2cffa3f29b4d/"
            "Official-Summary-6-6-2025.pdf",
    ),
    # Chester County (West Chester + boroughs/townships). 126-page Summary
    # Results PDFs published at chesco.org/DocumentCenter.
    ElectionwarePdfSource(
        name="2021 Chester Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Chester County (West Chester + boroughs + townships)",
        url=_CHESCO_PDF_BASE + "63375",
    ),
    ElectionwarePdfSource(
        name="2023 Chester Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Chester County (West Chester + boroughs + townships)",
        url=_CHESCO_PDF_BASE + "72435",
    ),
    ElectionwarePdfSource(
        name="2025 Chester Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Chester County (West Chester + boroughs + townships)",
        url=_CHESCO_PDF_BASE + "80017",
    ),
    # Northampton County (Bethlehem + Easton + boroughs/townships). PDFs
    # under corecode/uploaded_pdfs/. The 2023 and 2025 filenames have spaces
    # in them, hence the %20 encoding.
    ElectionwarePdfSource(
        name="2021 Northampton Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Northampton County (Bethlehem + Easton + boroughs)",
        url=_NORCO_PDF_BASE + "pe2021_243.pdf",
    ),
    ElectionwarePdfSource(
        name="2023 Northampton Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Northampton County (Bethlehem + Easton + boroughs)",
        url=_NORCO_PDF_BASE + "Election%20Summary%20PE%202023_6.pdf",
    ),
    ElectionwarePdfSource(
        name="2025 Northampton Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Northampton County (Bethlehem + Easton + boroughs)",
        url=_NORCO_PDF_BASE + "PE25%20Summary%20Results_1804.pdf",
    ),
    # Centre County (State College + Bellefonte + townships). Only the
    # 2025 PDF ID was discovered cleanly; 2021/2023 land on year-specific
    # pages that we couldn't drill into in one pass.
    ElectionwarePdfSource(
        name="2025 Centre Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Centre County (State College + Bellefonte + townships)",
        url=_CENTRE_PDF_BASE + "31108/election-summary",
    ),
    # Lebanon County (City of Lebanon + boroughs + townships). 2023 is
    # skipped — the only available 2023 file (`LandA-2023-Municipal.pdf`)
    # is a Logic & Accuracy pre-election test run dated Sept 13 2023, not
    # the actual May 2023 primary results.
    ElectionwarePdfSource(
        name="2021 Lebanon Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Lebanon County (City of Lebanon + boroughs + townships)",
        url=_LEBCO_PDF_BASE + "8c6c067f-4cbe-4f62-b777-7819f1aba29e/"
            "2021-Municipal-Primary-Election-Summary-Results.pdf",
    ),
    ElectionwarePdfSource(
        name="2025 Lebanon Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Lebanon County (City of Lebanon + boroughs + townships)",
        url=_LEBCO_PDF_BASE + "f21ac1fe-65c2-4f9d-a5da-50d957ca0961/"
            "Summary-Results-V4_With_Provos.pdf",
    ),
    # Mercer County (Mercer + Sharon + Hermitage + townships). Clean
    # `Election.Results/<YEAR>/PRIMARY/SUMMARY.pdf` URL pattern.
    ElectionwarePdfSource(
        name="2021 Mercer Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Mercer County (Mercer + Sharon + Hermitage + townships)",
        url=_MERCER_PDF_BASE + "2021/PRIMARY/SUMMARY.pdf",
    ),
    ElectionwarePdfSource(
        name="2023 Mercer Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Mercer County (Mercer + Sharon + Hermitage + townships)",
        url=_MERCER_PDF_BASE + "2023/PRIMARY/SUMMARY.pdf",
    ),
    ElectionwarePdfSource(
        name="2025 Mercer Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Mercer County (Mercer + Sharon + Hermitage + townships)",
        url=_MERCER_PDF_BASE + "2025/PRIMARY/SUMMARY.pdf",
    ),
    # Northumberland County (Sunbury + Shamokin + boroughs/townships).
    # `documents/elections/<YYYY>_<MMDD>_official/overall.pdf` pattern;
    # 2021's filename is capitalized differently ("Overall-Official.pdf").
    ElectionwarePdfSource(
        name="2021 Northumberland Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Northumberland County (Sunbury + Shamokin + boroughs)",
        url=_NUMCO_PDF_BASE + "2021_0518_official/Overall-Official.pdf",
    ),
    ElectionwarePdfSource(
        name="2023 Northumberland Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Northumberland County (Sunbury + Shamokin + boroughs)",
        url=_NUMCO_PDF_BASE + "2023_0516_official/overall.pdf",
    ),
    ElectionwarePdfSource(
        name="2025 Northumberland Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Northumberland County (Sunbury + Shamokin + boroughs)",
        url=_NUMCO_PDF_BASE + "2025_0520_official/overall.pdf",
    ),
    # Northumberland 2025 GENERAL — replaces the OE-stitched version of
    # the same data, which has a known conflation bug (Mayor + Borough
    # Council candidates grouped under one office name). Excluded from
    # the OE 2025 stitch via `filename_exclude_substrs` above.
    ElectionwarePdfSource(
        name="2025 Northumberland General", year=2025, category="Generals",
        coverage_note="Northumberland County (Sunbury + Shamokin + boroughs)",
        url=_NUMCO_PDF_BASE + "2025_1104_official/overall.pdf",
    ),
    # Indiana County (Indiana + boroughs + townships). Only 2025 PDF URL
    # was readily discoverable; older years require drilling.
    ElectionwarePdfSource(
        name="2025 Indiana Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Indiana County (Indiana + boroughs + townships)",
        url="https://www.indianacountypa.gov/wp-content/uploads/"
            "May-20-2025-offical-results-summary-with-write-ins.pdf",
    ),
    # Schuylkill County (Pottsville + boroughs + townships). 2021 isn't
    # on the current results page — would require the archived results
    # site at co.schuylkill.pa.us (interactive viewer, no direct PDFs).
    ElectionwarePdfSource(
        name="2023 Schuylkill Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Schuylkill County (Pottsville + boroughs + townships)",
        url=_SCHU_PDF_BASE + "2023/Primary/2023%20Municipal%20Primary%20Official%20Results.pdf",
    ),
    ElectionwarePdfSource(
        name="2025 Schuylkill Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Schuylkill County (Pottsville + boroughs + townships)",
        url=_SCHU_PDF_BASE + "2025/Primary/Official%20Election%20Results.pdf",
    ),
    # Lehigh County (Allentown + boroughs + townships). 2021 primary
    # is the canonical PA RCV case — Tuerk won the 6-way DEM Allentown
    # mayoral by 122 votes over the incumbent. 2023 is live-portal only
    # on Lehigh's site (no direct PDF).
    # 2017 — the canonical pre-Tuerk Allentown DEM mayoral primary, where
    # then-incumbent Ed Pawlowski (under federal indictment) won an 8-way
    # race with 28.4% to challenger Ray O'Connell's 23.0%. Different PDF
    # vendor format than 2021+ Lehigh; parsed by LehighOldPdfSource.
    LehighOldPdfSource(
        name="2017 Lehigh Primary", year=2017, category="Primaries", is_primary=True,
        coverage_note="Lehigh County (Allentown + boroughs + townships)",
        url=_LEHIGH_PDF_BASE + "Voter/2017MPTotals.pdf",
    ),
    ElectionwarePdfSource(
        name="2021 Lehigh Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Lehigh County (Allentown + boroughs + townships)",
        url=_LEHIGH_PDF_BASE + "ElectionSummary12.pdf",
    ),
    ElectionwarePdfSource(
        name="2025 Lehigh Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Lehigh County (Allentown + boroughs + townships)",
        url=_LEHIGH_PDF_BASE + "Voter/LehighSummaryFirstCert.pdf",
    ),
    # Montgomery County (Norristown + Philly suburbs). Same Electionware
    # vendor as Berks/Chester but two layout differences `_parse_electionware_lines`
    # handles separately:
    #   1. Contest header has inline "(Vote for N)" instead of a standalone
    #      "Vote For N" line — matched by _ELECTIONWARE_INLINE_VF_RE.
    #   2. Candidate rows are "Name PARTY ED Mail Prov TOTAL" instead of
    #      Berks's "Name TOTAL ED Mail Prov" — matched by the new
    #      _ELECTIONWARE_CAND_PARTY_RE that takes the LAST number as total.
    # The 2025 URL is the cumulative summary on the live results webapp;
    # the "Unofficial" label notwithstanding, it's generated months after
    # election certification with final tallies.
    ElectionwarePdfSource(
        name="2021 Montgomery Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Montgomery County (Norristown + Philly suburbs)",
        url="https://www.montcopa.org/DocumentCenter/View/31614",
    ),
    ElectionwarePdfSource(
        name="2023 Montgomery Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Montgomery County (Norristown + Philly suburbs)",
        url="https://www.montcopa.org/DocumentCenter/View/39284",
    ),
    ElectionwarePdfSource(
        name="2025 Montgomery Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Montgomery County (Norristown + Philly suburbs)",
        url="https://webapp07.montcopa.org/election/"
            "2025UnofficialPrimaryElectionSummaryReport.pdf",
    ),
    # Erie County — pre-Clarity (2021/2023) primaries are hosted as PDFs on
    # the County Council site. Same Electionware vendor as Montgomery
    # (Name PARTY ED Mail Prov Total layout + inline "(Vote for N)"). Live
    # 2024+ primaries come from Clarity (see ClaritySummaryJsonSource above).
    ElectionwarePdfSource(
        name="2021 Erie Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Erie County (city + boroughs + townships)",
        url="https://www.eriecountycouncilpa.gov/uploads/modules/resources/"
            "520965_erie_county_2021_municipal_primary_official_results.pdf",
    ),
    ElectionwarePdfSource(
        name="2023 Erie Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Erie County (city + boroughs + townships)",
        url="https://www.eriecountycouncilpa.gov/uploads/modules/resources/"
            "688208_erie_county_2023_municipal_primary_official_results.pdf",
    ),
    ElectionwarePdfSource(
        name="2021 Erie General", year=2021, category="Generals", is_primary=False,
        coverage_note="Erie County (city + boroughs + townships)",
        url="https://www.eriecountycouncilpa.gov/uploads/modules/resources/"
            "520895_erie_county_2021_municipal_election_official_results.pdf",
    ),
    ElectionwarePdfSource(
        name="2023 Erie General", year=2023, category="Generals", is_primary=False,
        coverage_note="Erie County (city + boroughs + townships)",
        url="https://www.eriecountycouncilpa.gov/uploads/modules/resources/"
            "634675_erie_county_2023_municipal_election_official_results.pdf",
    ),
    # Washington County (suburban Pittsburgh) — Berks-style Electionware
    # layout (Name-Total-First, separate "Vote For N" line, party prefix
    # on contest name). Pre-2025 archive on cms.washingtoncopa.gov.
    ElectionwarePdfSource(
        name="2021 Washington Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Washington County (Washington + boroughs + townships)",
        url="https://cms.washingtoncopa.gov/uploads/"
            "2021_Primary_Official_Results_Election_Summary_3868b5c430.pdf",
    ),
    ElectionwarePdfSource(
        name="2023 Washington Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Washington County (Washington + boroughs + townships)",
        url="https://cms.washingtoncopa.gov/uploads/"
            "2023_Primary_Official_Results_Election_Summary_a63e8ce175.pdf",
    ),
    ElectionwarePdfSource(
        name="2025 Washington Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Washington County (Washington + boroughs + townships)",
        url="https://cms.washingtoncopa.gov/uploads/"
            "2025_Municipal_Primary_Election_Summary_Official_4424b742e7.pdf",
    ),
    # Lackawanna County (Scranton + Carbondale + townships) — same vendor as
    # Chester (Electionware with the optional VOTE % column). PDFs hosted on
    # cms8.revize.com under the county subpath. 2021 primary URL not yet
    # discovered.
    ElectionwarePdfSource(
        name="2023 Lackawanna Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Lackawanna County (Scranton + Carbondale + townships)",
        url="https://cms8.revize.com/revize/lackawanna/LackawannaPrimary2023Summary.pdf",
    ),
    ElectionwarePdfSource(
        name="2025 Lackawanna Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Lackawanna County (Scranton + Carbondale + townships)",
        url="https://cms8.revize.com/revize/lackawanna/Document_center/"
            "Certified%20Election%20Results/2025/25SUM.CERT.pdf",
    ),
    # Butler County (suburban Pittsburgh / Slippery Rock) — Berks-style
    # Electionware format.
    ElectionwarePdfSource(
        name="2023 Butler Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Butler County (Butler City + boroughs + townships)",
        url="https://www.butlercountypa.gov/DocumentCenter/View/6915/SUMMARY-PDF",
    ),
    ElectionwarePdfSource(
        name="2025 Butler Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Butler County (Butler City + boroughs + townships)",
        url="https://butlercountypa.gov/DocumentCenter/View/10551/SUMMARY-PDF",
    ),
    # Fayette County PDFs (pre-2025 primaries). Electionware format with
    # inline "(Vote for N)" headers, like Montgomery and Erie.
    ElectionwarePdfSource(
        name="2021 Fayette Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Fayette County (Uniontown + boroughs + townships)",
        url="https://www.fayettecountypa.org/DocumentCenter/View/9115",
    ),
    ElectionwarePdfSource(
        name="2023 Fayette Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Fayette County (Uniontown + boroughs + townships)",
        url="https://www.fayettecountypa.org/DocumentCenter/View/6591",
    ),
]


# Bucks County uses an "Election Source" vendor PDF with "Choice Total ED MI PR"
# columns instead of Electionware's "Name Party ED MI PR Total" layout. Rather
# than write a second vendor-specific parser for one county, we extract via
# Claude (LlmPdfSource) — the per-PDF API cost is a few dollars and cached after
# first run.
# LLM extraction is unavailable in this environment (no ANTHROPIC_API_KEY).
# LlmPdfSource still exists for future use but its registry stays empty.
LLM_PDF_SOURCES: list[ElectionSource] = []


# Bucks County (Doylestown + Philly suburbs) uses an "ElectionSource"-vendor
# certified results PDF that BucksPdfSource handles directly (no LLM needed).
_BUCKS_PDF_BASE = "https://www.buckscounty.gov/ArchiveCenter/ViewFile/Item/"

BUCKS_PDF_SOURCES: list[ElectionSource] = [
    BucksPdfSource(
        name="2021 Bucks Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Bucks County (Doylestown + Philly suburbs)",
        url=_BUCKS_PDF_BASE + "437",
    ),
    BucksPdfSource(
        name="2023 Bucks Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Bucks County (Doylestown + Philly suburbs)",
        url=_BUCKS_PDF_BASE + "518",
    ),
    BucksPdfSource(
        name="2025 Bucks Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Bucks County (Doylestown + Philly suburbs)",
        url=_BUCKS_PDF_BASE + "552",
    ),
]


# Lancaster County (570K, 7th-largest PA county) publishes pre-2025
# results as a hierarchical HTML site on vr.co.lancaster.pa.us. Walked
# by LancasterHtmlSource. 2025+ moved to a JS-rendered SPA whose PDF
# downloads require an in-browser session — deferred.
LANCASTER_HTML_SOURCES: list[ElectionSource] = [
    LancasterHtmlSource(
        name="2017 Lancaster Primary", year=2017, category="Primaries", is_primary=True,
        coverage_note="Lancaster County (Lancaster City + boroughs + townships)",
        year_path="May_16,_2017_-_Municipal_Primary",
        url="https://vr.co.lancaster.pa.us/ElectionReturns/"
            "May_16,_2017_-_Municipal_Primary/SelectParty.html",
    ),
    LancasterHtmlSource(
        name="2019 Lancaster Primary", year=2019, category="Primaries", is_primary=True,
        coverage_note="Lancaster County (Lancaster City + boroughs + townships)",
        year_path="May_21,_2019_-_Municipal_Primary",
        url="https://vr.co.lancaster.pa.us/ElectionReturns/"
            "May_21,_2019_-_Municipal_Primary/SelectParty.html",
    ),
    LancasterHtmlSource(
        name="2021 Lancaster Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Lancaster County (Lancaster City + boroughs + townships)",
        year_path="May_18,_2021_-_Municipal_Primary",
        url="https://vr.co.lancaster.pa.us/ElectionReturns/"
            "May_18,_2021_-_Municipal_Primary/SelectParty.html",
    ),
    LancasterHtmlSource(
        name="2023 Lancaster Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Lancaster County (Lancaster City + boroughs + townships)",
        year_path="May_16,_2023_-_Municipal_Primary",
        url="https://vr.co.lancaster.pa.us/ElectionReturns/"
            "May_16,_2023_-_Municipal_Primary/SelectParty.html",
    ),
]


# Lycoming County (Williamsport + boroughs + townships) uses a different PDF
# vendor — inline contest headers with party in parens, plus a VOTE % column.
# Parsed by `LycomingPdfSource`.
_LYCOMING_PDF_BASE = (
    "https://lycomingcountypa.gov/Documents/Government/Departments/"
    "Voter%20Services/PreviousResults/"
)

LYCOMING_PDF_SOURCES: list[ElectionSource] = [
    LycomingPdfSource(
        name="2021 Lycoming Primary", year=2021, category="Primaries", is_primary=True,
        coverage_note="Lycoming County (Williamsport + boroughs + townships)",
        url=_LYCOMING_PDF_BASE + "2021MPOfficial.pdf",
    ),
    LycomingPdfSource(
        name="2023 Lycoming Primary", year=2023, category="Primaries", is_primary=True,
        coverage_note="Lycoming County (Williamsport + boroughs + townships)",
        url=_LYCOMING_PDF_BASE + "2023MPOfficial.pdf",
    ),
    LycomingPdfSource(
        name="2025 Lycoming Primary", year=2025, category="Primaries", is_primary=True,
        coverage_note="Lycoming County (Williamsport + boroughs + townships)",
        url=_LYCOMING_PDF_BASE + "2025MPOfficial.pdf",
    ),
]


# Federal/statewide contests to exclude when running the Allegheny LOCAL
# workbook. Kept inclusive — both the modern "Justice of the X Court" naming
# and the older "Judge of the X Court" form, plus office variants.
_PA_STATE_FEDERAL_RACE_EXCLUDE = (
    r"justice of the (?:supreme|superior|commonwealth) court"
    r"|judge of the (?:superior|commonwealth|supreme) court"
    r"|(?:supreme|superior|commonwealth) court retention"  # retention questions
    r"|president of the united states|^president$"
    r"|u\.?s\.? senator|u\.?s\.? house|u\.?s\.? representative|congress"
    r"|governor|lieutenant governor|attorney general"
    r"|auditor general|state treasurer"
    r"|senator in the general assembly|representative in the general assembly"
    r"|state senator|state representative"
    # Ballot questions are not candidate races — they're up/down referendums
    # with YES/NO choices that our pipeline mistakes for a 2-3 candidate
    # field. Montgomery 2021 surfaced three of these (Constitutional
    # Amendments 1/2/3) at 50/50 ties.
    r"|constitutional amendment|ballot question|referendum"
    r"|home rule charter"
)


# --- Curated top-races workbook --------------------------------------------

# Office-name keywords that flag a race as "notable local" for the curated
# view (mayor, council, DA, sheriff, controllers, county exec, etc.).
_NOTABLE_LOCAL_OFFICE_RE = re.compile(
    r"\b(?:mayor|council|district attorney|sheriff|controller|chief executive"
    r"|county executive|county commissioner|judge of the court of common pleas"
    r"|magisterial district judge|tax collector|township supervisor)\b",
    re.IGNORECASE,
)


def _sheet_to_election_type(sheet_name: str, workbook: str) -> str:
    """Derive Primary vs General from the source sheet name / workbook.

    Pooled-by-category workbooks split into 'Primaries' / 'Generals' sheets,
    so the sheet name carries the type. Philly's workbook uses year-named
    sheets but is primary-only — fall back to the filename.
    """
    if sheet_name == 'Primaries':
        return 'Primary'
    if sheet_name == 'Generals':
        return 'General'
    if 'Primary' in workbook:
        return 'Primary'
    if 'General' in workbook:
        return 'General'
    return ''


def _race_summary(group: pd.DataFrame) -> dict:
    """Compress one race's rows into a single summary record.

    Sorts candidates by votes desc, captures top three candidate names + %,
    the gap between top two, total candidate count, and a `Candidates Inline`
    string with every candidate listed.
    """
    g = group.sort_values('Votes', ascending=False).reset_index(drop=True)
    leader = g.iloc[0] if len(g) >= 1 else None
    runner = g.iloc[1] if len(g) >= 2 else None
    third = g.iloc[2] if len(g) >= 3 else None
    inline = ', '.join(
        f"{r['Candidate']} {r['Percent']:.1f}%" for _, r in g.iterrows()
    )
    return {
        'Year': leader['Year'] if 'Year' in g.columns and leader is not None else '',
        'Election Type': leader.get('Election Type', '') if leader is not None else '',
        'Workbook': leader.get('Workbook', '') if leader is not None else '',
        'Coverage': leader.get('Coverage', '') if leader is not None else '',
        'Race_Name': leader['Race_Name'] if leader is not None else '',
        'Num_Candidates': len(g),
        'Leader': leader['Candidate'] if leader is not None else '',
        'Leader %': round(float(leader['Percent']), 2) if leader is not None else None,
        'Runner-Up': runner['Candidate'] if runner is not None else '',
        'Runner-Up %': round(float(runner['Percent']), 2) if runner is not None else None,
        'Third': third['Candidate'] if third is not None else '',
        'Third %': round(float(third['Percent']), 2) if third is not None else None,
        'Top-2 Gap': (
            round(float(leader['Percent']) - float(runner['Percent']), 2)
            if runner is not None else None
        ),
        'Candidates Inline': inline,
    }


def _read_all_non_majority_rows(workbooks: list[str]) -> pd.DataFrame:
    """Read every per-workbook output file and return a unified frame.

    Tags each row with its `Workbook` source, `Election Type`
    (Primary/General derived from sheet name + filename), and `Coverage`
    so the curated view can show which jurisdiction and contest cycle
    it came from. Drops blank separator rows.
    """
    frames = []
    for wb in workbooks:
        if not Path(wb).exists():
            print(f"  (skipping {wb} — not found)")
            continue
        sheets = pd.read_excel(wb, sheet_name=None)
        for sheet_name, df in sheets.items():
            if df.empty or 'Race_Name' not in df.columns:
                continue
            df = df[df['Race_Name'].astype(str) != ''].copy()
            df['Workbook'] = wb
            df['Sheet'] = sheet_name
            df['Election Type'] = _sheet_to_election_type(sheet_name, wb)
            # Per-source workbooks (Philly) don't have a Year column; default to
            # the sheet name (which is the year, e.g., "2007").
            if 'Year' not in df.columns:
                df['Year'] = sheet_name
            if 'Coverage' not in df.columns:
                df['Coverage'] = ''
            frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def make_top_races_workbook(
    workbooks: list[str],
    out_path: str = "Top_RCV_Races.xlsx",
) -> None:
    """Build a curated workbook from the per-source outputs.

    Sheets, each one row per race:
      1. **Closest two-way races** — sorted by gap between top two asc.
      2. **Crowded fields** — 4+ candidates with leader under 40%, sorted by
         leader % asc (most fragmented first).
      3. **Notable local offices** — mayor / council / DA / sheriff / etc.
      4. **All non-majority races** — every race from every workbook.
    """
    all_rows = _read_all_non_majority_rows(workbooks)
    if all_rows.empty:
        print("  (no rows found — workbooks haven't been generated yet)")
        return

    # One row per race summary
    summaries = []
    for _, group in all_rows.groupby(['Workbook', 'Sheet', 'Year', 'Race_Name'],
                                      sort=False):
        summaries.append(_race_summary(group))
    races = pd.DataFrame(summaries)

    # Sheet 1: closest two-way
    close = (
        races[races['Top-2 Gap'].notna() & (races['Top-2 Gap'] <= 5.0)]
        .sort_values('Top-2 Gap')
    )

    # Sheet 2: crowded fields
    crowded = (
        races[(races['Num_Candidates'] >= 4) & (races['Leader %'] < 40.0)]
        .sort_values('Leader %')
    )

    # Sheet 3: notable local offices
    notable = (
        races[races['Race_Name'].astype(str).str.contains(
            _NOTABLE_LOCAL_OFFICE_RE, na=False)]
        .sort_values(['Leader %', 'Num_Candidates'], ascending=[True, False])
    )

    # Sheet 4: everything
    full = races.sort_values(['Workbook', 'Year', 'Race_Name'])

    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        close.to_excel(writer, sheet_name='Closest two-way', index=False)
        crowded.to_excel(writer, sheet_name='Crowded fields', index=False)
        notable.to_excel(writer, sheet_name='Notable local offices', index=False)
        full.to_excel(writer, sheet_name='All non-majority races', index=False)


# --- Entry point ------------------------------------------------------------

if __name__ == "__main__":
    # --- Philly primaries ---
    # Include Mayor + District Council (the original RCV-relevant offices)
    # plus District Attorney and Controller (citywide single-winner races
    # that also surface in non-mayoral years — 2017 was the open-seat
    # DA primary that Larry Krasner won 38.2% in a 7-way race).
    include_all = False  # True -> all non-majority races; False -> the curated set
    if include_all:
        philly_out = "Philadelphia_Primary_AllRaces.xlsx"
        philly_race_pattern = None
    else:
        philly_out = "Philadelphia_Primary_Mayor_DistrictCouncil.xlsx"
        philly_race_pattern = r"mayor|district council|district attorney|controller"
    write_workbook(PHILLY_PRIMARY_SOURCES, philly_out, race_pattern=philly_race_pattern)
    print(f"Saved Philly workbook to '{philly_out}'")

    # --- OpenElections PA 2018+ ---
    # Exclude President: cross-county ticket fragmentation ("TRUMP" / "TRUMP /
    # VANCE" / "DONALD J TRUMP, PRESIDENT / JD VANCE, VICE-PRESIDENT" / etc.)
    # makes totals unreliable without an external candidate roster.
    pa_out = "Pennsylvania_NonMajority_2018plus.xlsx"
    write_workbook_pooled_by_category(
        OPENELECTIONS_PA_SOURCES, pa_out,
        race_exclude_pattern=r"president",
    )
    print(f"Saved PA workbook to '{pa_out}'")

    # --- Allegheny County local races (WPRDC, 2017-2025 odd-year cycles) ---
    ac_out = "Allegheny_NonMajority_Local.xlsx"
    write_workbook_pooled_by_category(
        ALLEGHENY_LOCAL_SOURCES, ac_out,
        race_exclude_pattern=_PA_STATE_FEDERAL_RACE_EXCLUDE,
    )
    print(f"Saved Allegheny workbook to '{ac_out}'")

    # --- 2025 PA all-counties local races (OpenElections county rollups) ---
    pa_local_out = "Pennsylvania_NonMajority_Local2025.xlsx"
    write_workbook_pooled_by_category(
        PA_LOCAL_2025_SOURCES, pa_local_out,
        race_exclude_pattern=_PA_STATE_FEDERAL_RACE_EXCLUDE,
    )
    print(f"Saved 2025 PA local workbook to '{pa_local_out}'")

    # --- Mid-sized counties: York/Delaware (Clarity), Electionware + Lycoming PDFs ---
    mid_out = "Pennsylvania_NonMajority_MidCounties.xlsx"
    write_workbook_pooled_by_category(
        CLARITY_PA_SOURCES + ELECTIONWARE_PDF_SOURCES + LYCOMING_PDF_SOURCES
            + BUCKS_PDF_SOURCES + LANCASTER_HTML_SOURCES + LLM_PDF_SOURCES,
        mid_out,
        race_exclude_pattern=_PA_STATE_FEDERAL_RACE_EXCLUDE,
    )
    print(f"Saved mid-counties workbook to '{mid_out}'")

    # --- Curated top-races view across every workbook above ---
    top_out = "Top_RCV_Races.xlsx"
    make_top_races_workbook(
        workbooks=[philly_out, pa_out, ac_out, pa_local_out, mid_out],
        out_path=top_out,
    )
    print(f"Saved curated top-races workbook to '{top_out}'")
