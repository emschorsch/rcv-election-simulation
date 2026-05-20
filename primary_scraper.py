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

import json
import re
import urllib.error
import urllib.request
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
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
    'NOT ASSIGNED', 'WRITE-IN TOTALS', 'WRITE IN TOTALS',
})


def _normalize_candidate(s: pd.Series) -> pd.Series:
    """Canonicalize candidate names so cross-county variants collapse to one row.
    Upper-cases, strips periods (`L. WEISS` → `L WEISS`), collapses whitespace,
    and removes a trailing party code (`DAVE SUNDAY REP` → `DAVE SUNDAY`).

    Does NOT attempt to merge gov/lt-gov ticket variants like
    `SHAPIRO / DAVIS` vs `JOSH SHAPIRO` vs `JOSH SHAPIRO AUSTIN DAVIS`. Those
    would require an external roster and are flagged as a known limitation.
    """
    out = (
        s.fillna('').astype(str)
        .str.upper()
        .str.replace('.', '', regex=False)
        .str.replace(r'\s+', ' ', regex=True)
        .str.strip()
    )
    out = out.str.replace(_PARTY_SUFFIX_RE, '', regex=True).str.strip()
    return out


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
    df = df[~df['_candidate'].isin(_NON_CANDIDATE_NAMES)].reset_index(drop=True)

    office = _normalize_text(df['office'])
    district = df['district'].fillna('').astype(str).str.strip()
    party_col = df['party'] if 'party' in df.columns else pd.Series('', index=df.index)
    party = _normalize_text(party_col)

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
    files whose name contains `filename_substr`, then fetches each from
    `raw_base_url + filename` and concatenates. Populates `coverage_note` with
    the actual file count so partial-coverage years are visibly flagged.
    """
    listing_api_url: str = ""
    raw_base_url: str = ""
    filename_substr: str = ""
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
        return sorted(
            d['name'] for d in data
            if d.get('type') == 'file' and self.filename_substr in d.get('name', '')
        )


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
    threshold: float = 50.0,
    min_winner_votes: int = 100,
    min_candidate_percent: float = 1.0,
) -> None:
    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        for source in sources:
            print(f"Processing {source.name}...")
            tidy = source.fetch_tidy()
            tidy = add_percentages(tidy)
            tidy = filter_non_majority(tidy, threshold=threshold)
            tidy = filter_min_winner_votes(tidy, min_votes=min_winner_votes)
            tidy = filter_min_candidate_percent(tidy, min_percent=min_candidate_percent)
            tidy = filter_race_names(tidy, race_pattern)
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
    threshold: float = 50.0,
    min_winner_votes: int = 100,
    min_candidate_percent: float = 1.0,
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
        tidy = add_percentages(tidy)
        tidy = filter_non_majority(tidy, threshold=threshold)
        tidy = filter_min_winner_votes(tidy, min_votes=min_winner_votes)
        tidy = filter_min_candidate_percent(tidy, min_percent=min_candidate_percent)
        tidy = filter_race_names(tidy, race_pattern)
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
    pa_out = "Pennsylvania_NonMajority_2018plus.xlsx"
    write_workbook_pooled_by_category(OPENELECTIONS_PA_SOURCES, pa_out)
    print(f"Saved PA workbook to '{pa_out}'")
