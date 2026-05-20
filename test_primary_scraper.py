"""Tests for the source-agnostic pipeline functions in primary_scraper.

Run with: `pytest test_primary_scraper.py`
"""
from types import SimpleNamespace

import pandas as pd
import pytest

from primary_scraper import (
    _parse_openelections_df,
    _wide_columns_to_tidy,
    add_percentages,
    filter_non_majority,
    filter_race_names,
    format_for_sheet,
    format_pooled_for_sheet,
)


# --- _wide_columns_to_tidy --------------------------------------------------

def test_wide_columns_candidate_first_sums_across_rows():
    df = pd.DataFrame({
        "WARD": [1, 2],
        "ALICE - MAYOR": [10, 5],
        "BOB - MAYOR": [3, 7],
    })
    out = _wide_columns_to_tidy(df, ["ALICE - MAYOR", "BOB - MAYOR"], candidate_first=True)
    expected = pd.DataFrame({
        "Candidate": ["ALICE", "BOB"],
        "Race_Name": ["MAYOR", "MAYOR"],
        "Votes": [15.0, 10.0],
    })
    pd.testing.assert_frame_equal(out, expected, check_dtype=False)


def test_wide_columns_candidate_first_multi_dash_race_joins_rest():
    df = pd.DataFrame({"ALICE - MAYOR - DEMOCRAT": [10, 5]})
    out = _wide_columns_to_tidy(df, ["ALICE - MAYOR - DEMOCRAT"], candidate_first=True)
    assert out["Candidate"].tolist() == ["ALICE"]
    assert out["Race_Name"].tolist() == ["MAYOR - DEMOCRAT"]


def test_wide_columns_race_first_two_parts():
    df = pd.DataFrame({"MAYOR - ALICE": [10, 5]})
    out = _wide_columns_to_tidy(df, ["MAYOR - ALICE"], candidate_first=False)
    assert out["Race_Name"].tolist() == ["MAYOR"]
    assert out["Candidate"].tolist() == ["ALICE"]


def test_wide_columns_race_first_multi_part_takes_first_two_as_race():
    # Mirrors the 2023 Philly shape, e.g. "COUNCIL - AT LARGE - JANE SMITH"
    df = pd.DataFrame({"COUNCIL - AT LARGE - JANE SMITH": [10, 5]})
    out = _wide_columns_to_tidy(df, ["COUNCIL - AT LARGE - JANE SMITH"], candidate_first=False)
    assert out["Race_Name"].tolist() == ["COUNCIL - AT LARGE"]
    assert out["Candidate"].tolist() == ["JANE SMITH"]


def test_wide_columns_strips_commas_from_numeric_strings():
    df = pd.DataFrame({"ALICE - MAYOR": ["1,000", "2,500"]})
    out = _wide_columns_to_tidy(df, ["ALICE - MAYOR"], candidate_first=True)
    assert out["Votes"].tolist() == [3500.0]


def test_wide_columns_no_dash_yields_unknown_candidate():
    df = pd.DataFrame({"TOTAL": [10, 5]})
    out = _wide_columns_to_tidy(df, ["TOTAL"], candidate_first=True)
    assert out["Candidate"].tolist() == ["Unknown"]
    assert out["Race_Name"].tolist() == ["TOTAL"]


def test_wide_columns_nan_coerced_to_zero():
    df = pd.DataFrame({"ALICE - MAYOR": [10, None]})
    out = _wide_columns_to_tidy(df, ["ALICE - MAYOR"], candidate_first=True)
    assert out["Votes"].tolist() == [10.0]


def test_wide_columns_does_not_mutate_input():
    df = pd.DataFrame({"ALICE - MAYOR": ["1,000"]})
    before = df.copy()
    _wide_columns_to_tidy(df, ["ALICE - MAYOR"], candidate_first=True)
    pd.testing.assert_frame_equal(df, before)


# --- add_percentages --------------------------------------------------------

def test_add_percentages_basic():
    df = pd.DataFrame({
        "Candidate": ["A", "B"],
        "Race_Name": ["X", "X"],
        "Votes": [30, 70],
    })
    out = add_percentages(df)
    assert out["Total_Votes_Race"].tolist() == [100, 100]
    assert out["Percent"].tolist() == [30.0, 70.0]


def test_add_percentages_groups_by_race():
    df = pd.DataFrame({
        "Candidate": ["A", "B", "C", "D"],
        "Race_Name": ["X", "X", "Y", "Y"],
        "Votes": [40, 60, 25, 75],
    })
    out = add_percentages(df)
    assert out["Total_Votes_Race"].tolist() == [100, 100, 100, 100]
    assert out["Percent"].tolist() == [40.0, 60.0, 25.0, 75.0]


def test_add_percentages_coerces_string_votes():
    df = pd.DataFrame({
        "Candidate": ["A"],
        "Race_Name": ["X"],
        "Votes": ["30"],
    })
    out = add_percentages(df)
    assert out["Votes"].tolist() == [30.0]
    assert out["Percent"].tolist() == [100.0]


# --- filter_non_majority ----------------------------------------------------

def test_filter_non_majority_keeps_split_race():
    df = add_percentages(pd.DataFrame({
        "Candidate": ["A", "B", "C"],
        "Race_Name": ["X", "X", "X"],
        "Votes": [40, 35, 25],
    }))
    out = filter_non_majority(df)
    assert set(out["Candidate"]) == {"A", "B", "C"}


def test_filter_non_majority_drops_majority_race():
    df = add_percentages(pd.DataFrame({
        "Candidate": ["A", "B"],
        "Race_Name": ["X", "X"],
        "Votes": [60, 40],
    }))
    out = filter_non_majority(df)
    assert len(out) == 0


def test_filter_non_majority_drops_only_majority_race_when_mixed():
    df = add_percentages(pd.DataFrame({
        "Candidate": ["A", "B", "C", "D", "E"],
        "Race_Name": ["X", "X", "X", "Y", "Y"],
        "Votes": [40, 35, 25, 80, 20],  # X: 40/35/25 (no majority); Y: 80/20 (majority)
    }))
    out = filter_non_majority(df)
    assert set(out["Race_Name"]) == {"X"}


def test_filter_non_majority_keeps_race_tied_at_threshold():
    # 50/50 tie: max percent == 50, which still counts as "no majority winner"
    # and is exactly the case RCV would need to resolve, so the race stays in.
    df = add_percentages(pd.DataFrame({
        "Candidate": ["A", "B"],
        "Race_Name": ["X", "X"],
        "Votes": [50, 50],
    }))
    assert len(filter_non_majority(df, threshold=50.0)) == 2


def test_filter_non_majority_drops_race_one_vote_over_threshold():
    df = add_percentages(pd.DataFrame({
        "Candidate": ["A", "B"],
        "Race_Name": ["X", "X"],
        "Votes": [51, 49],
    }))
    assert len(filter_non_majority(df, threshold=50.0)) == 0


def test_filter_non_majority_custom_threshold():
    df = add_percentages(pd.DataFrame({
        "Candidate": ["A", "B", "C"],
        "Race_Name": ["X", "X", "X"],
        "Votes": [45, 30, 25],  # max 45%
    }))
    assert len(filter_non_majority(df, threshold=40.0)) == 0
    assert len(filter_non_majority(df, threshold=45.0)) == 3  # boundary kept
    assert len(filter_non_majority(df, threshold=50.0)) == 3


# --- filter_race_names ------------------------------------------------------

def test_filter_race_names_none_is_noop():
    df = pd.DataFrame({"Race_Name": ["MAYOR", "DOG CATCHER"]})
    pd.testing.assert_frame_equal(filter_race_names(df, None), df)


def test_filter_race_names_case_insensitive():
    df = pd.DataFrame({"Race_Name": ["MAYOR - DEM", "Sheriff", "City Council"]})
    out = filter_race_names(df, r"mayor")
    assert out["Race_Name"].tolist() == ["MAYOR - DEM"]


def test_filter_race_names_regex_alternation():
    df = pd.DataFrame({"Race_Name": ["Mayor", "District Council 5", "Sheriff"]})
    out = filter_race_names(df, r"mayor|district council")
    assert set(out["Race_Name"]) == {"Mayor", "District Council 5"}


def test_filter_race_names_no_match_returns_empty():
    df = pd.DataFrame({"Race_Name": ["Sheriff", "Comptroller"]})
    out = filter_race_names(df, r"mayor")
    assert len(out) == 0


# --- format_for_sheet -------------------------------------------------------

def test_format_for_sheet_sorts_and_inserts_blank_between_races():
    df = pd.DataFrame({
        "Candidate": ["B", "A", "D", "C"],
        "Race_Name": ["X", "X", "Y", "Y"],
        "Votes": [20, 30, 10, 40],
        "Percent": [40.0, 60.0, 20.0, 80.0],
    })
    out = format_for_sheet(df)
    # X sorted by votes desc: A(30) then B(20); then blank; then Y: C(40) then D(10); then blank.
    assert out["Candidate"].tolist() == ["A", "B", "", "C", "D", ""]
    assert out["Race_Name"].tolist() == ["X", "X", "", "Y", "Y", ""]


def test_format_for_sheet_empty_input():
    df = pd.DataFrame(columns=["Candidate", "Race_Name", "Votes", "Percent"])
    out = format_for_sheet(df)
    assert len(out) == 0
    assert list(out.columns) == ["Candidate", "Race_Name", "Votes", "Percent"]


# --- _parse_openelections_df -----------------------------------------------

def _oe_df(rows):
    """Build an OpenElections-style DataFrame from a list of dicts."""
    cols = ['county', 'office', 'district', 'party', 'candidate', 'votes']
    return pd.DataFrame(rows, columns=cols).astype(str).replace('nan', '')


def test_oe_general_race_name_uses_office_only_when_no_district():
    df = _oe_df([
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': 'DEM', 'candidate': 'Biden', 'votes': '100'},
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': 'REP', 'candidate': 'Trump', 'votes': '200'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    # General: party is *not* part of race name, so Biden and Trump share a race.
    # Names are upper-cased by the parser to merge case-variant spellings.
    assert set(out['Race_Name']) == {'PRESIDENT'}
    assert set(out['Candidate']) == {'BIDEN', 'TRUMP'}


def test_oe_general_race_name_appends_district():
    df = _oe_df([
        {'county': 'Adams', 'office': 'U.S. House', 'district': '12',
         'party': 'DEM', 'candidate': 'Griffin', 'votes': '50'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert out['Race_Name'].tolist() == ['U.S. HOUSE 12']


def test_oe_primary_race_name_prepends_party():
    df = _oe_df([
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': 'DEM', 'candidate': 'Biden', 'votes': '100'},
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': 'REP', 'candidate': 'Trump', 'votes': '200'},
    ])
    out = _parse_openelections_df(df, is_primary=True)
    assert set(out['Race_Name']) == {'DEM PRESIDENT', 'REP PRESIDENT'}


def test_oe_drops_meta_rows_with_empty_candidate():
    df = _oe_df([
        {'county': 'Adams', 'office': 'Registered Voters', 'district': '',
         'party': '', 'candidate': '', 'votes': '500'},
        {'county': 'Adams', 'office': 'Ballots Cast', 'district': '',
         'party': '', 'candidate': '', 'votes': '400'},
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': 'DEM', 'candidate': 'Biden', 'votes': '100'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert out['Candidate'].tolist() == ['BIDEN']
    assert out['Race_Name'].tolist() == ['PRESIDENT']


def test_oe_sums_votes_across_counties():
    df = _oe_df([
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': 'DEM', 'candidate': 'Biden', 'votes': '100'},
        {'county': 'Allegheny', 'office': 'President', 'district': '',
         'party': 'DEM', 'candidate': 'Biden', 'votes': '5000'},
        {'county': 'Berks', 'office': 'President', 'district': '',
         'party': 'DEM', 'candidate': 'Biden', 'votes': '200'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert out['Votes'].tolist() == [5300.0]


def test_oe_strips_periods_and_party_suffix_from_candidate():
    # Real-world bug: some 2024 PA counties wrote candidates as "DAVE SUNDAY REP"
    # (party suffix from cross-filed ballot lines) and "RICHARD L. WEISS" vs
    # "RICHARD L WEISS". Normalization should collapse all variants.
    df = _oe_df([
        {'county': 'Adams', 'office': 'AG', 'district': '',
         'party': 'REP', 'candidate': 'DAVE SUNDAY REP', 'votes': '100'},
        {'county': 'Berks', 'office': 'AG', 'district': '',
         'party': 'REP', 'candidate': 'Dave Sunday', 'votes': '200'},
        {'county': 'Clinton', 'office': 'AG', 'district': '',
         'party': 'GRN', 'candidate': 'Richard L. Weiss', 'votes': '5'},
        {'county': 'Dauphin', 'office': 'AG', 'district': '',
         'party': 'GRN', 'candidate': 'RICHARD L WEISS GRN', 'votes': '7'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    by_cand = dict(zip(out['Candidate'], out['Votes']))
    assert by_cand == {'DAVE SUNDAY': 300.0, 'RICHARD L WEISS': 12.0}


def test_oe_drops_election_admin_pseudo_candidates():
    # OVER VOTES / UNDER VOTES / WRITE-IN TOTALS / NOT ASSIGNED inflate the
    # race denominator and falsely lower every real candidate's share, so they
    # are stripped before percentage calculation upstream.
    df = _oe_df([
        {'county': 'Adams', 'office': 'AG', 'district': '',
         'party': 'DEM', 'candidate': 'Real Person', 'votes': '100'},
        {'county': 'Adams', 'office': 'AG', 'district': '',
         'party': '', 'candidate': 'Over Votes', 'votes': '5'},
        {'county': 'Adams', 'office': 'AG', 'district': '',
         'party': '', 'candidate': 'Under Votes', 'votes': '10'},
        {'county': 'Adams', 'office': 'AG', 'district': '',
         'party': '', 'candidate': 'Write-in Totals', 'votes': '3'},
        {'county': 'Adams', 'office': 'AG', 'district': '',
         'party': '', 'candidate': 'Not Assigned', 'votes': '2'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert out['Candidate'].tolist() == ['REAL PERSON']


def test_oe_merges_case_variants_of_same_candidate():
    # Real-world bug: 2020 PA AG file had both "JOSH SHAPIRO" and "Josh Shapiro"
    # across counties, splitting his vote and falsely flagging the race as
    # non-majority. Upper-case normalization should merge them.
    df = _oe_df([
        {'county': 'Adams', 'office': 'Attorney General', 'district': '',
         'party': 'DEM', 'candidate': 'JOSH SHAPIRO', 'votes': '100'},
        {'county': 'Allegheny', 'office': 'Attorney General', 'district': '',
         'party': 'DEM', 'candidate': 'Josh Shapiro', 'votes': '200'},
        {'county': 'Berks', 'office': 'Attorney General', 'district': '',
         'party': 'DEM', 'candidate': '  Josh  Shapiro  ', 'votes': '50'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert out['Candidate'].tolist() == ['JOSH SHAPIRO']
    assert out['Votes'].tolist() == [350.0]


def test_oe_write_ins_survive_with_empty_party():
    df = _oe_df([
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': 'DEM', 'candidate': 'Biden', 'votes': '100'},
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': '', 'candidate': 'Write-ins', 'votes': '5'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert set(out['Candidate']) == {'BIDEN', 'WRITE-INS'}


def test_oe_votes_coerced_to_zero_when_missing():
    df = _oe_df([
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': 'DEM', 'candidate': 'Biden', 'votes': ''},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert out['Votes'].tolist() == [0.0]


def test_oe_missing_party_column_does_not_crash():
    df = pd.DataFrame([
        {'county': 'Adams', 'office': 'President', 'district': '',
         'candidate': 'Biden', 'votes': '100'},
    ]).astype(str)
    out = _parse_openelections_df(df, is_primary=False)
    assert out['Race_Name'].tolist() == ['PRESIDENT']


# --- format_pooled_for_sheet -----------------------------------------------

def test_format_pooled_emits_year_and_coverage_columns():
    source = SimpleNamespace(year=2020, coverage_note="67 of 67 counties")
    df = pd.DataFrame({
        "Candidate": ["A", "B"],
        "Race_Name": ["X", "X"],
        "Votes": [50, 50],
        "Percent": [50.0, 50.0],
    })
    out = format_pooled_for_sheet([(source, df)])
    assert list(out.columns) == ['Year', 'Race_Name', 'Candidate', 'Votes', 'Percent', 'Coverage']
    non_blank = out[out['Race_Name'] != '']
    assert set(non_blank['Year']) == {2020}
    assert set(non_blank['Coverage']) == {"67 of 67 counties"}


def test_format_pooled_sorts_by_year_then_race_then_votes_desc():
    s2020 = SimpleNamespace(year=2020, coverage_note="")
    s2018 = SimpleNamespace(year=2018, coverage_note="")
    df_2020 = pd.DataFrame({
        "Candidate": ["B", "A"], "Race_Name": ["Y", "Y"],
        "Votes": [20, 30], "Percent": [40.0, 60.0],
    })
    df_2018 = pd.DataFrame({
        "Candidate": ["D", "C"], "Race_Name": ["X", "X"],
        "Votes": [10, 40], "Percent": [20.0, 80.0],
    })
    out = format_pooled_for_sheet([(s2020, df_2020), (s2018, df_2018)])
    non_blank = out[out['Race_Name'] != '']
    # Expect: 2018 X (C 40, D 10), then 2020 Y (A 30, B 20).
    assert non_blank['Year'].tolist() == [2018, 2018, 2020, 2020]
    assert non_blank['Race_Name'].tolist() == ['X', 'X', 'Y', 'Y']
    assert non_blank['Candidate'].tolist() == ['C', 'D', 'A', 'B']


def test_format_pooled_inserts_blank_row_between_race_groups():
    source = SimpleNamespace(year=2020, coverage_note="")
    df = pd.DataFrame({
        "Candidate": ["A", "B", "C"],
        "Race_Name": ["X", "X", "Y"],
        "Votes": [30, 20, 10],
        "Percent": [60.0, 40.0, 100.0],
    })
    out = format_pooled_for_sheet([(source, df)])
    # Expect rows: X-A, X-B, blank, Y-C, blank
    assert out['Race_Name'].tolist() == ['X', 'X', '', 'Y', '']


def test_format_pooled_empty_input_returns_empty_with_correct_columns():
    out = format_pooled_for_sheet([])
    assert len(out) == 0
    assert list(out.columns) == ['Year', 'Race_Name', 'Candidate', 'Votes', 'Percent', 'Coverage']


def test_format_pooled_skips_sources_with_empty_dataframes():
    source_empty = SimpleNamespace(year=2018, coverage_note="x")
    source_data = SimpleNamespace(year=2020, coverage_note="y")
    empty_df = pd.DataFrame(columns=["Candidate", "Race_Name", "Votes", "Percent"])
    data_df = pd.DataFrame({
        "Candidate": ["A"], "Race_Name": ["X"], "Votes": [50], "Percent": [50.0],
    })
    out = format_pooled_for_sheet([(source_empty, empty_df), (source_data, data_df)])
    non_blank = out[out['Race_Name'] != '']
    assert non_blank['Year'].tolist() == [2020]
