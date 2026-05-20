"""Tests for the source-agnostic pipeline functions in primary_scraper.

Run with: `pytest test_primary_scraper.py`
"""
import pandas as pd
import pytest

from primary_scraper import (
    _wide_columns_to_tidy,
    add_percentages,
    filter_non_majority,
    filter_race_names,
    format_for_sheet,
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
