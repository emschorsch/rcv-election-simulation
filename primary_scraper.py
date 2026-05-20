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

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd


# --- Source base class and subclasses ---------------------------------------

@dataclass(kw_only=True)
class ElectionSource(ABC):
    name: str  # used as the Excel sheet name
    url: str
    write_in_label: str = "write in"

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


# --- Pipeline ---------------------------------------------------------------

def add_percentages(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df['Votes'] = pd.to_numeric(df['Votes'], errors='coerce').fillna(0)
    df['Total_Votes_Race'] = df.groupby('Race_Name')['Votes'].transform('sum')
    df['Percent'] = df['Votes'] / df['Total_Votes_Race'] * 100
    return df


def filter_non_majority(df: pd.DataFrame, threshold: float = 50.0) -> pd.DataFrame:
    max_percent = df.groupby('Race_Name')['Percent'].transform('max')
    return df[max_percent < threshold].copy()


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
) -> None:
    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        for source in sources:
            print(f"Processing {source.name}...")
            tidy = source.fetch_tidy()
            tidy = add_percentages(tidy)
            tidy = filter_non_majority(tidy, threshold=threshold)
            tidy = filter_race_names(tidy, race_pattern)
            sheet = format_for_sheet(tidy)
            sheet.to_excel(writer, sheet_name=source.name, index=False)


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


# --- Entry point ------------------------------------------------------------

if __name__ == "__main__":
    include_all = False  # True -> all non-majority races; False -> mayor/council only

    if include_all:
        out_path = "Philadelphia_Primary_AllRaces.xlsx"
        race_pattern = None
    else:
        out_path = "Philadelphia_Primary_Mayor_DistrictCouncil.xlsx"
        race_pattern = r"mayor|district council"

    write_workbook(PHILLY_PRIMARY_SOURCES, out_path, race_pattern=race_pattern)
    print(f"All sources processed and saved to '{out_path}'")
