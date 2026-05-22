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
    'CONTEST TOTALS',    # Mercer / Northumberland: total ballots cast in the contest
    'TIMES BLANK VOTED', 'TIMES OVER VOTED', 'TIMES UNDER VOTED',  # Electionware stat rows
    'TOTAL',  # Lycoming per-contest sum row
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


def fuzzy_canonicalize_candidates(
    df: pd.DataFrame,
    *,
    ratio_threshold: float = 0.92,
    min_length: int = 6,
) -> pd.DataFrame:
    """Merge near-duplicate candidate names *within each race* (likely typos
    like `DOUGHHERTY` vs `DOUGHERTY`). Within a race, names are processed in
    vote-count order; the highest-vote spelling becomes canonical, and any
    later spelling within `ratio_threshold` is remapped to it.

    `min_length` (default 6) prevents short distinct names like `JOHN`/`JOAN`
    from being collapsed. The 0.92 threshold leaves gov/president ticket
    variants like `TRUMP` vs `TRUMP / VANCE` (~0.55) alone — fuzzy matching
    is the wrong tool for that problem.
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
            if len(name) < min_length:
                mapping[name] = name
                continue
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


def _parse_openelections_df(df: pd.DataFrame, *, is_primary: bool) -> pd.DataFrame:
    """Aggregate an OpenElections-format (2018+) tidy CSV into [Candidate, Race_Name, Votes].

    Expected columns: county, office, district, party, candidate, votes (extras
    like precinct/election_day/absentee/mail are ignored). Rows with no candidate
    are dropped — those are meta rows (Registered Voters, Ballots Cast, etc.).
    Race_Name is `<office> <district>` for generals; primaries prepend `<party>`
    so each party's contest becomes its own race. Candidate / office / party
    strings are upper-cased before grouping so case differences across counties
    don't split a single candidate's votes.
    """
    df = df.copy()
    df['_candidate'] = _normalize_candidate(df['candidate'])
    df = df[df['_candidate'] != '']
    df = df[~_is_non_candidate(df['_candidate'])].reset_index(drop=True)

    office = _normalize_text(df['office'])
    district = df['district'].fillna('').astype(str).str.strip()
    party_col = df['party'] if 'party' in df.columns else pd.Series('', index=df.index)
    party = _normalize_text(party_col)

    # Drop rows for office types that always require a district but where the
    # district column is blank — leaving them in would silently merge candidates
    # from many different districts into one bogus race.
    bad = office.isin(_NEEDS_DISTRICT_OFFICES) & (district == '')
    df = df[~bad].reset_index(drop=True)
    office = office[~bad].reset_index(drop=True)
    district = district[~bad].reset_index(drop=True)
    party = party[~bad].reset_index(drop=True)

    if is_primary:
        race_name = party + ' ' + office + ' ' + district
    else:
        race_name = office + ' ' + district
    race_name = race_name.str.replace(r'\s+', ' ', regex=True).str.strip()

    # Drop races that look like multi-seat — OE source data doesn't preserve
    # Vote For N, so without this filter School Director / County Commissioner
    # / at-large borough council races get treated as Vote For 1 and surface
    # as bogus non-majority RCV findings.
    multiseat = race_name.map(_is_likely_multiseat_oe_race)
    df = df[~multiseat].reset_index(drop=True)
    race_name = race_name[~multiseat].reset_index(drop=True)

    tidy = pd.DataFrame({
        'Candidate': df['_candidate'],
        'Race_Name': race_name,
        'Votes': pd.to_numeric(df['votes'], errors='coerce').fillna(0),
    })
    return tidy.groupby(['Race_Name', 'Candidate'], as_index=False)['Votes'].sum()


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
_OE_LIKELY_MULTISEAT_OFFICES_RE = re.compile(
    r"\b(?:school director|school board"
    r"|county commissioner|county council at[-\s]large"
    r"|borough council|township council|city council"
    r"|council at[-\s]large|council member city|council \(?at[-\s]large\)?)\b",
    re.IGNORECASE,
)
_DISTRICT_NUMBERED_RE = re.compile(r"\bdistrict\s+\d", re.IGNORECASE)


def _is_likely_multiseat_oe_race(race_name: str) -> bool:
    """True if this OE-sourced race looks like a multi-seat (Vote For N>1)
    contest we can't detect directly. District-numbered races (Vote For 1)
    are explicitly excluded from this check."""
    if _DISTRICT_NUMBERED_RE.search(race_name):
        return False
    return bool(_OE_LIKELY_MULTISEAT_OFFICES_RE.search(race_name))


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
    is_primary: bool = False
    expected_counties: int = 67

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
        return _parse_openelections_df(combined, is_primary=self.is_primary)

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
        return _filter_oe_listing(data, self.filename_substr, self.filename_suffix)


def _filter_oe_listing(
    listing: list[dict],
    substr: str,
    suffix: str,
) -> list[str]:
    """Filter a GitHub Contents API listing to file entries whose `name`
    contains `substr` AND ends with `suffix`. Returns sorted file names.
    Extracted from `StitchedOpenElectionsCountiesSource._list_matching_files`
    so the filter logic can be unit-tested without mocking HTTP."""
    return sorted(
        d['name'] for d in listing
        if d.get('type') == 'file'
           and substr in d.get('name', '')
           and d.get('name', '').endswith(suffix)
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

# Trailing "(DEM)" / "(REP)" / etc. parenthetical on Clarity contest names —
# duplicates the CAT field, so we strip it to get a clean contest name.
_CLARITY_PARTY_PAREN_RE = re.compile(
    r'\s*\((?:' + '|'.join(_PARTY_CODES) + r')\)\s*$', re.IGNORECASE
)


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
    """
    rows = []
    for contest in data:
        if int(contest.get('VF') or 1) > 1:
            continue  # multi-seat — skip
        contest_name = str(contest.get('C') or '').strip()
        contest_name = _CLARITY_PARTY_PAREN_RE.sub('', contest_name).strip()
        if not contest_name:
            continue
        cat = str(contest.get('CAT') or '').strip().upper()
        party_prefix = _CLARITY_CAT_TO_PARTY.get(cat, '')
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
        m = _ELECTIONWARE_VOTE_FOR_RE.match(lines[i])
        if not m:
            i += 1
            continue
        vote_for_n = int(m.group(1))

        # Walk backward (over blank lines and column-header lines) to find
        # the contest header — the most recent non-blank, non-header line.
        # "Column header" = a line whose tokens are all in the known set
        # (Election / TOTAL / Day / Mail / Provisional, in any whitespace
        # arrangement).
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
        if not contest_name:
            i += 1
            continue

        # Multi-seat: skip the whole contest.
        if vote_for_n > 1:
            i += 1
            continue

        # Scan forward for candidate rows; stop at the next "Vote For N".
        i += 1
        while i < len(lines):
            line = lines[i]
            if _ELECTIONWARE_VOTE_FOR_RE.match(line):
                break
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
        with urllib.request.urlopen(req, timeout=60, context=ctx) as resp:
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
        filename_suffix="__county.csv",
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
    include_all = False  # True -> all non-majority races; False -> mayor/council only
    if include_all:
        philly_out = "Philadelphia_Primary_AllRaces.xlsx"
        philly_race_pattern = None
    else:
        philly_out = "Philadelphia_Primary_Mayor_DistrictCouncil.xlsx"
        philly_race_pattern = r"mayor|district council"
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
        CLARITY_PA_SOURCES + ELECTIONWARE_PDF_SOURCES + LYCOMING_PDF_SOURCES, mid_out,
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
