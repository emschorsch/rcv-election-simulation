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
_ELECTIONWARE_HEADER_TOKENS = frozenset(
    {'Election', 'TOTAL', 'Day', 'Mail', 'Provisional'}
)
# Candidate row: indented name, two-or-more spaces, then 1-5 columns of numbers
# (with optional commas). The first captured number is TOTAL — we ignore the
# rest (Election Day / Mail / Provisional / Military). 2021 Berks PDFs have
# only a TOTAL column; 2023/2025 have four. Non-greedy name match plus the
# end-of-line anchor lets us locate the boundary correctly even for names
# containing commas like "Dante Santoni, Jr.".
_ELECTIONWARE_CAND_RE = re.compile(
    r'^\s+(.+?)\s{2,}([\d,]+)(?:\s+[\d,]+){0,4}\s*$'
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
            tokens = stripped.split()
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


def _extract_electionware_pdf(pdf_url: str) -> list[dict]:
    """Download an Electionware-format PDF and extract raw result rows.

    Uses pdfplumber's layout-preserving text extraction (no LLM, no API
    key). The PDF is fetched to a temp file because pdfplumber requires
    seekable input. Uses certifi's CA bundle explicitly so it works on
    systems whose default trust store doesn't accept the county sites'
    cert chains.
    """
    import ssl
    import tempfile
    import certifi  # lazy imports — only needed when this source is used
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
    return _parse_electionware_lines(all_lines)


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


# Berks County publishes PDF-only results going back to 2003. We extract
# them with pdfplumber + a regex state machine in `ElectionwarePdfSource`
# (no API key needed). The same Electionware format is used by many other
# PA counties, so adding more is one new source registry entry each.

_BERKS_PDF_BASE = "https://www.berkspa.gov/getmedia/"

# 2021 Berks isn't here: their 2021 "Grand Totals" PDF is statewide-only
# (judges + ballot questions, 5 pages, no local contests), and the
# corresponding local-race data lives in the separate "By Precinct" PDF
# which would need a precinct-aware parser. 2023+ Berks publishes a single
# unified summary PDF that includes local contests at county totals.

BERKS_PDF_SOURCES: list[ElectionSource] = [
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

    # --- Mid-sized counties: York + Delaware (Clarity JSON) + Berks (Electionware PDF) ---
    mid_out = "Pennsylvania_NonMajority_MidCounties.xlsx"
    write_workbook_pooled_by_category(
        CLARITY_PA_SOURCES + BERKS_PDF_SOURCES, mid_out,
        race_exclude_pattern=_PA_STATE_FEDERAL_RACE_EXCLUDE,
    )
    print(f"Saved mid-counties workbook to '{mid_out}'")
