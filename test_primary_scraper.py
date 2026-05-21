"""Tests for the source-agnostic pipeline functions in primary_scraper.

Run with: `pytest test_primary_scraper.py`
"""
from types import SimpleNamespace

import pandas as pd
import pytest

from primary_scraper import (
    _filter_oe_listing,
    _parse_clarity_summary_json,
    _parse_electionware_lines,
    _parse_openelections_df,
    _parse_wprdc_summary_df,
    _rows_to_tidy,
    _strip_middle_initials,
    _wide_columns_to_tidy,
    add_percentages,
    filter_exclude_race_names,
    filter_min_candidate_percent,
    filter_min_leader_percent,
    filter_min_winner_votes,
    filter_non_majority,
    filter_race_names,
    format_for_sheet,
    format_pooled_for_sheet,
    fuzzy_canonicalize_candidates,
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


# --- filter_min_winner_votes -----------------------------------------------

def test_filter_min_winner_votes_keeps_race_with_large_winner():
    df = pd.DataFrame({
        "Candidate": ["A", "B"], "Race_Name": ["X", "X"], "Votes": [500, 400],
    })
    assert len(filter_min_winner_votes(df, min_votes=100)) == 2


def test_filter_min_winner_votes_drops_race_with_tiny_winner():
    df = pd.DataFrame({
        "Candidate": ["A", "B"], "Race_Name": ["X", "X"], "Votes": [40, 30],
    })
    assert len(filter_min_winner_votes(df, min_votes=100)) == 0


def test_filter_min_winner_votes_drops_only_tiny_races_when_mixed():
    df = pd.DataFrame({
        "Candidate": ["A", "B", "C", "D"],
        "Race_Name": ["BigRace", "BigRace", "TinyRace", "TinyRace"],
        "Votes": [500, 400, 40, 30],
    })
    out = filter_min_winner_votes(df, min_votes=100)
    assert set(out['Race_Name']) == {"BigRace"}


def test_filter_min_winner_votes_uses_max_not_total():
    # Two candidates at 60 each: total 120 but max is 60 — drop with min=100.
    df = pd.DataFrame({
        "Candidate": ["A", "B"], "Race_Name": ["X", "X"], "Votes": [60, 60],
    })
    assert len(filter_min_winner_votes(df, min_votes=100)) == 0


def test_filter_min_winner_votes_at_exact_threshold_kept():
    df = pd.DataFrame({
        "Candidate": ["A", "B"], "Race_Name": ["X", "X"], "Votes": [100, 50],
    })
    assert len(filter_min_winner_votes(df, min_votes=100)) == 2


# --- filter_min_leader_percent ---------------------------------------------

def test_filter_min_leader_percent_drops_write_in_chaos_race():
    # 20-candidate field where the "leader" got 6%: write-in only race, drop.
    df = pd.DataFrame({
        "Candidate": [f"C{i}" for i in range(20)],
        "Race_Name": ["X"] * 20,
        "Votes": [60] + [50] * 19,
        "Percent": [6.0] + [4.9474] * 19,
    })
    assert len(filter_min_leader_percent(df, min_percent=10.0)) == 0


def test_filter_min_leader_percent_keeps_legitimate_competitive_race():
    # 6-way primary with leader at 28% — real RCV-relevant race.
    df = pd.DataFrame({
        "Candidate": ["A", "B", "C", "D", "E", "F"],
        "Race_Name": ["X"] * 6,
        "Votes": [28, 22, 18, 15, 10, 7],
        "Percent": [28.0, 22.0, 18.0, 15.0, 10.0, 7.0],
    })
    out = filter_min_leader_percent(df, min_percent=10.0)
    assert len(out) == 6


def test_filter_min_leader_percent_only_drops_subthreshold_races_when_mixed():
    df = pd.DataFrame({
        "Candidate": ["A1", "A2", "B1", "B2"],
        "Race_Name": ["Real", "Real", "WriteInChaos", "WriteInChaos"],
        "Votes": [40, 35, 5, 4],
        "Percent": [40.0, 35.0, 5.0, 4.0],
    })
    out = filter_min_leader_percent(df, min_percent=10.0)
    assert set(out['Race_Name']) == {"Real"}


# --- filter_min_candidate_percent ------------------------------------------

def test_filter_min_candidate_percent_drops_long_tail_within_race():
    df = pd.DataFrame({
        "Candidate": ["A", "B", "C", "D"],
        "Race_Name": ["X", "X", "X", "X"],
        "Votes": [400, 300, 5, 1],
        "Percent": [55.0, 41.0, 0.7, 0.1],
    })
    out = filter_min_candidate_percent(df, min_percent=1.0)
    assert out['Candidate'].tolist() == ["A", "B"]


def test_filter_min_candidate_percent_keeps_exact_threshold():
    df = pd.DataFrame({
        "Candidate": ["A", "B"], "Race_Name": ["X", "X"],
        "Votes": [99, 1], "Percent": [99.0, 1.0],
    })
    assert len(filter_min_candidate_percent(df, min_percent=1.0)) == 2


def test_filter_min_candidate_percent_does_not_remove_winner():
    # The pruning is a display filter — it should never drop a top finisher
    # since their percent is always at least 1 / num_candidates.
    df = pd.DataFrame({
        "Candidate": ["A", "B"], "Race_Name": ["X", "X"],
        "Votes": [49, 47], "Percent": [49.0, 47.0],
    })
    out = filter_min_candidate_percent(df, min_percent=1.0)
    assert set(out['Candidate']) == {"A", "B"}


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


# --- filter_exclude_race_names ---------------------------------------------

def test_filter_exclude_race_names_drops_matching_races():
    df = pd.DataFrame({"Race_Name": ["PRESIDENT", "U.S. SENATE", "DEM PRESIDENT"]})
    out = filter_exclude_race_names(df, r"president")
    assert out['Race_Name'].tolist() == ["U.S. SENATE"]


def test_filter_exclude_race_names_none_is_noop():
    df = pd.DataFrame({"Race_Name": ["X", "Y"]})
    pd.testing.assert_frame_equal(filter_exclude_race_names(df, None), df)


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


# --- _strip_middle_initials ------------------------------------------------

def test_strip_middle_initials_drops_single_letter_middle_tokens():
    assert _strip_middle_initials("RYAN E MACKENZIE") == "RYAN MACKENZIE"
    assert _strip_middle_initials("CHRISTINA M HARTMAN") == "CHRISTINA HARTMAN"
    assert _strip_middle_initials("JOHN F K KENNEDY") == "JOHN KENNEDY"  # multiple


def test_strip_middle_initials_preserves_first_and_last_tokens():
    # Single-letter first tokens (like stage names "R KELLY") are preserved.
    assert _strip_middle_initials("R KELLY") == "R KELLY"
    # Multi-letter middle tokens (real names) are preserved.
    assert _strip_middle_initials("MARY ANN SMITH") == "MARY ANN SMITH"
    # Suffixes like JR/SR stay as the last token.
    assert _strip_middle_initials("WAYNE LANGERHOLC JR") == "WAYNE LANGERHOLC JR"


def test_strip_middle_initials_handles_short_names():
    assert _strip_middle_initials("ALICE") == "ALICE"
    assert _strip_middle_initials("ALICE SMITH") == "ALICE SMITH"


# --- _normalize_candidate (commas + middle initials) -----------------------

def test_oe_strips_comma_in_suffix():
    # WAYNE LANGERHOLC, JR / WAYNE LANGERHOLC JR — comma is the only difference.
    df = _oe_df([
        {'county': 'A', 'office': 'X', 'district': '',
         'party': 'R', 'candidate': 'Wayne Langerholc, Jr', 'votes': '100'},
        {'county': 'B', 'office': 'X', 'district': '',
         'party': 'R', 'candidate': 'Wayne Langerholc Jr', 'votes': '200'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert out['Candidate'].tolist() == ['WAYNE LANGERHOLC JR']
    assert out['Votes'].tolist() == [300.0]


def test_oe_drops_middle_initial_in_candidate():
    # RYAN E MACKENZIE / RYAN MACKENZIE — same person.
    df = _oe_df([
        {'county': 'A', 'office': 'X', 'district': '',
         'party': 'R', 'candidate': 'Ryan E Mackenzie', 'votes': '100'},
        {'county': 'B', 'office': 'X', 'district': '',
         'party': 'R', 'candidate': 'Ryan Mackenzie', 'votes': '200'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert out['Candidate'].tolist() == ['RYAN MACKENZIE']
    assert out['Votes'].tolist() == [300.0]


# --- fuzzy_canonicalize_candidates -----------------------------------------

def test_fuzzy_canonicalize_merges_typo_variants_into_higher_vote_spelling():
    # SHAUN DOUGHERTY (200 votes) and the typo SHAUN DOUGHHERTY (50 votes) —
    # ratio 0.97, well above the 0.92 default. Higher-vote spelling wins.
    df = pd.DataFrame({
        "Race_Name": ["X", "X"],
        "Candidate": ["SHAUN DOUGHERTY", "SHAUN DOUGHHERTY"],
        "Votes": [200.0, 50.0],
    })
    out = fuzzy_canonicalize_candidates(df)
    assert out['Candidate'].tolist() == ['SHAUN DOUGHERTY']
    assert out['Votes'].tolist() == [250.0]


def test_fuzzy_canonicalize_does_not_merge_distinct_short_names():
    # JOHN vs JOAN — ratio ~0.75 AND both are below the 6-char min_length.
    # Must not merge.
    df = pd.DataFrame({
        "Race_Name": ["X", "X"],
        "Candidate": ["JOHN", "JOAN"],
        "Votes": [100.0, 80.0],
    })
    out = fuzzy_canonicalize_candidates(df)
    assert set(out['Candidate']) == {"JOHN", "JOAN"}


def test_fuzzy_canonicalize_leaves_ticket_variants_alone():
    # PA gov/president ticket fragmentation: TRUMP vs TRUMP / VANCE.
    # Ratio ~0.55, well below the 0.92 default — fuzzy must not merge them.
    df = pd.DataFrame({
        "Race_Name": ["PRESIDENT", "PRESIDENT"],
        "Candidate": ["DONALD J TRUMP", "TRUMP / VANCE"],
        "Votes": [2_700_000.0, 250_000.0],
    })
    out = fuzzy_canonicalize_candidates(df)
    assert set(out['Candidate']) == {"DONALD J TRUMP", "TRUMP / VANCE"}


def test_fuzzy_canonicalize_only_merges_within_same_race():
    # Same near-duplicate name appearing in two different races stays separate.
    df = pd.DataFrame({
        "Race_Name": ["MAYOR", "MAYOR", "GOVERNOR", "GOVERNOR"],
        "Candidate": ["SHAUN DOUGHERTY", "SHAUN DOUGHHERTY",
                       "SHAUN DOUGHERTY", "SHAUN DOUGHHERTY"],
        "Votes": [200.0, 50.0, 300.0, 60.0],
    })
    out = fuzzy_canonicalize_candidates(df)
    # Both races collapse to canonical name, but the rows remain race-separated.
    by_race = dict(zip(out['Race_Name'], out['Votes']))
    assert by_race == {"MAYOR": 250.0, "GOVERNOR": 360.0}


def test_fuzzy_canonicalize_empty_input_returns_empty():
    df = pd.DataFrame(columns=['Race_Name', 'Candidate', 'Votes'])
    out = fuzzy_canonicalize_candidates(df)
    assert len(out) == 0


def test_oe_strips_periods_and_party_suffix_from_candidate():
    # Real-world bug: some 2024 PA counties wrote candidates as "DAVE SUNDAY REP"
    # (party suffix from cross-filed ballot lines) and "RICHARD L. WEISS" vs
    # "RICHARD L WEISS". Normalization collapses all variants — also dropping
    # the middle-initial "L" so all four rows for Weiss merge.
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
    assert by_cand == {'DAVE SUNDAY': 300.0, 'RICHARD WEISS': 12.0}


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


def test_oe_drops_aggregate_write_ins_entries():
    # Aggregate "Write-ins" rows are dropped because they're the sum of the
    # individual named write-in rows in the same file — leaving them in
    # double-counts and inflates the field size in write-in-heavy races.
    df = _oe_df([
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': 'DEM', 'candidate': 'Biden', 'votes': '100'},
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': '', 'candidate': 'Write-ins', 'votes': '5'},
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': '', 'candidate': 'Write-in', 'votes': '3'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert set(out['Candidate']) == {'BIDEN'}


def test_oe_votes_coerced_to_zero_when_missing():
    df = _oe_df([
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': 'DEM', 'candidate': 'Biden', 'votes': ''},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert out['Votes'].tolist() == [0.0]


def test_oe_drops_district_required_rows_when_district_is_blank():
    # 2024 Philadelphia State House rows have blank district in OpenElections;
    # leaving them in would merge 20+ unrelated reps into a single "STATE HOUSE"
    # race. The parser must drop them.
    df = _oe_df([
        {'county': 'Philly', 'office': 'State House', 'district': '',
         'party': 'DEM', 'candidate': 'Chris Rabb', 'votes': '500'},
        {'county': 'Philly', 'office': 'State House', 'district': '',
         'party': 'DEM', 'candidate': 'Ben Waxman', 'votes': '600'},
        {'county': 'Adams', 'office': 'State House', 'district': '76',
         'party': 'DEM', 'candidate': 'Joe Waltz', 'votes': '300'},
        {'county': 'Adams', 'office': 'U.S. House', 'district': '',
         'party': 'DEM', 'candidate': 'Some Person', 'votes': '999'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    # Only the districted State House 76 row survives.
    assert out['Race_Name'].tolist() == ['STATE HOUSE 76']
    assert out['Candidate'].tolist() == ['JOE WALTZ']


def test_oe_keeps_district_less_rows_for_statewide_offices():
    # President, US Senate, Governor, AG, etc. don't have districts — blank
    # district must NOT cause them to be dropped.
    df = _oe_df([
        {'county': 'Adams', 'office': 'President', 'district': '',
         'party': 'DEM', 'candidate': 'Biden', 'votes': '100'},
        {'county': 'Adams', 'office': 'U.S. Senate', 'district': '',
         'party': 'DEM', 'candidate': 'Casey', 'votes': '200'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert set(out['Race_Name']) == {'PRESIDENT', 'U.S. SENATE'}


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


# --- _parse_wprdc_summary_df -----------------------------------------------

def _wprdc_df(rows):
    cols = ['contest_name', 'choice_name', 'party_name', 'total_votes']
    return pd.DataFrame(rows, columns=cols).astype(str).replace('nan', '')


def test_wprdc_strips_vote_for_1_from_race_name():
    df = _wprdc_df([
        {'contest_name': 'District Attorney (Vote For 1)',
         'choice_name': 'Matt Dugan', 'party_name': 'DEM', 'total_votes': '100'},
    ])
    out = _parse_wprdc_summary_df(df, is_primary=False)
    assert out['Race_Name'].tolist() == ['District Attorney']
    assert out['Candidate'].tolist() == ['MATT DUGAN']


def test_wprdc_drops_multi_seat_vote_for_n_races():
    # "Vote For 2" is a multi-seat race; IRV majority doesn't apply.
    df = _wprdc_df([
        {'contest_name': 'County Council At-Large (Vote For 2)',
         'choice_name': 'Alice', 'party_name': 'DEM', 'total_votes': '100'},
        {'contest_name': 'County Council At-Large (Vote For 2)',
         'choice_name': 'Bob', 'party_name': 'DEM', 'total_votes': '90'},
        {'contest_name': 'Mayor (Vote For 1)',
         'choice_name': 'Carol', 'party_name': 'DEM', 'total_votes': '100'},
    ])
    out = _parse_wprdc_summary_df(df, is_primary=False)
    assert out['Race_Name'].tolist() == ['Mayor']


def test_wprdc_primary_contest_keeps_party_prefix():
    df = _wprdc_df([
        {'contest_name': 'DEM Mayor of Pittsburgh (Vote For 1)',
         'choice_name': 'Ed Gainey', 'party_name': 'DEM', 'total_votes': '100'},
        {'contest_name': 'REP Mayor of Pittsburgh (Vote For 1)',
         'choice_name': 'Tony Moreno', 'party_name': 'REP', 'total_votes': '50'},
    ])
    out = _parse_wprdc_summary_df(df, is_primary=True)
    assert set(out['Race_Name']) == {'DEM Mayor of Pittsburgh', 'REP Mayor of Pittsburgh'}


def test_wprdc_primary_prepends_party_when_missing_from_contest_name():
    # Older WPRDC files (2017, 2019) don't include the party in contest_name,
    # only in party_name. Without prepending, DEM and REP primaries for the
    # same office would merge into one bogus race.
    df = _wprdc_df([
        {'contest_name': 'Mayor Bellevue (Vote For 1)',
         'choice_name': 'Alice', 'party_name': 'DEM', 'total_votes': '100'},
        {'contest_name': 'Mayor Bellevue (Vote For 1)',
         'choice_name': 'Bob', 'party_name': 'REP', 'total_votes': '40'},
    ])
    out = _parse_wprdc_summary_df(df, is_primary=True)
    assert set(out['Race_Name']) == {'DEM Mayor Bellevue', 'REP Mayor Bellevue'}


def test_wprdc_drops_ballots_cast_and_registered_voters_meta_rows():
    df = _wprdc_df([
        {'contest_name': 'BALLOTS CAST - Nonpartisan (Vote For 0)',
         'choice_name': 'BALLOTS CAST', 'party_name': 'NON', 'total_votes': '50000'},
        {'contest_name': 'REGISTERED VOTERS - Nonpartisan (Vote For 0)',
         'choice_name': 'REGISTERED VOTERS', 'party_name': 'NON', 'total_votes': '100000'},
        {'contest_name': 'Mayor (Vote For 1)',
         'choice_name': 'Real Person', 'party_name': 'DEM', 'total_votes': '100'},
    ])
    out = _parse_wprdc_summary_df(df, is_primary=False)
    assert out['Race_Name'].tolist() == ['Mayor']
    assert out['Candidate'].tolist() == ['REAL PERSON']


def test_wprdc_applies_same_candidate_normalization_as_oe():
    # Periods, party-suffix, and middle-initial cleanups should all fire,
    # collapsing the variants into one row.
    df = _wprdc_df([
        {'contest_name': 'DEM Mayor (Vote For 1)',
         'choice_name': 'Wayne L. Langerholc, Jr', 'party_name': 'DEM', 'total_votes': '100'},
        {'contest_name': 'DEM Mayor (Vote For 1)',
         'choice_name': 'Wayne Langerholc Jr DEM', 'party_name': 'DEM', 'total_votes': '200'},
    ])
    out = _parse_wprdc_summary_df(df, is_primary=True)
    assert out['Candidate'].tolist() == ['WAYNE LANGERHOLC JR']
    assert out['Votes'].tolist() == [300.0]


# --- _filter_oe_listing ----------------------------------------------------

def _listing_entry(name):
    return {'name': name, 'type': 'file'}


def test_filter_oe_listing_substr_only_returns_all_matching_files():
    listing = [
        _listing_entry('20251104__pa__general__adams__county.csv'),
        _listing_entry('20251104__pa__general__adams__precinct.csv'),
        _listing_entry('20251104__pa__primary__adams__county.csv'),
    ]
    out = _filter_oe_listing(listing, substr='__general__', suffix='')
    assert out == [
        '20251104__pa__general__adams__county.csv',
        '20251104__pa__general__adams__precinct.csv',
    ]


def test_filter_oe_listing_suffix_disambiguates_county_vs_precinct():
    # Real 2025 case: dir has both __county.csv and __precinct.csv files.
    # Without suffix=__county.csv we'd double-count by stitching both.
    listing = [
        _listing_entry('20251104__pa__general__adams__county.csv'),
        _listing_entry('20251104__pa__general__adams__precinct.csv'),
        _listing_entry('20251104__pa__general__bucks__county.csv'),
        _listing_entry('20251104__pa__general__bucks__precinct.csv'),
    ]
    out = _filter_oe_listing(listing, substr='__general__', suffix='__county.csv')
    assert out == [
        '20251104__pa__general__adams__county.csv',
        '20251104__pa__general__bucks__county.csv',
    ]


def test_filter_oe_listing_ignores_non_file_entries():
    listing = [
        {'name': 'subdir', 'type': 'dir'},
        _listing_entry('20251104__pa__general__adams__county.csv'),
    ]
    out = _filter_oe_listing(listing, substr='__general__', suffix='__county.csv')
    assert out == ['20251104__pa__general__adams__county.csv']


def test_wprdc_aggregates_duplicate_candidate_rows_within_race():
    # Defensive: if a file somehow has two rows for the same (race, candidate),
    # they should sum rather than duplicate.
    df = _wprdc_df([
        {'contest_name': 'Mayor (Vote For 1)',
         'choice_name': 'Alice', 'party_name': 'DEM', 'total_votes': '60'},
        {'contest_name': 'Mayor (Vote For 1)',
         'choice_name': 'Alice', 'party_name': 'DEM', 'total_votes': '40'},
    ])
    out = _parse_wprdc_summary_df(df, is_primary=False)
    assert out['Candidate'].tolist() == ['ALICE']
    assert out['Votes'].tolist() == [100.0]


# --- _parse_clarity_summary_json -------------------------------------------

def _clarity_contest(cat, c, ch, v, p=None, vf=1):
    return {
        'CAT': cat, 'C': c, 'CH': ch, 'V': v, 'VF': vf,
        'P': p if p is not None else [''] * len(ch),
    }


def test_clarity_general_contest_keeps_bare_race_name():
    data = [_clarity_contest('Pennsylvania', 'Mayor of Media', ['Alice', 'Bob'], [100, 50])]
    out = _parse_clarity_summary_json(data, is_primary=False)
    assert out['Race_Name'].tolist() == ['Mayor of Media', 'Mayor of Media']
    assert set(out['Candidate']) == {'ALICE', 'BOB'}


def test_clarity_primary_prepends_party_from_cat():
    # CAT="Democratic" should become "DEM " prefix when is_primary=True.
    data = [_clarity_contest('Democratic', 'Council Member', ['Alice'], [100])]
    out = _parse_clarity_summary_json(data, is_primary=True)
    assert out['Race_Name'].tolist() == ['DEM Council Member']


def test_clarity_primary_does_not_double_prefix_when_contest_already_has_party():
    # Some Clarity files already include the party in the contest name itself
    # (the "(DEM)" parenthetical) AND in CAT. After stripping the paren we'd
    # double-prefix without the already-prefixed check.
    data = [_clarity_contest('Democratic', 'DEM Council Member', ['Alice'], [100])]
    out = _parse_clarity_summary_json(data, is_primary=True)
    assert out['Race_Name'].tolist() == ['DEM Council Member']


def test_clarity_drops_multi_seat_contests():
    data = [
        _clarity_contest('Democratic', 'Single Seat Race', ['Alice'], [100], vf=1),
        _clarity_contest('Democratic', 'Two Seat Race', ['Alice', 'Bob'], [100, 80], vf=2),
    ]
    out = _parse_clarity_summary_json(data, is_primary=True)
    assert set(out['Race_Name']) == {'DEM Single Seat Race'}


def test_clarity_strips_trailing_party_paren_from_contest_name():
    # Real Clarity output: contest names look like "Mayor of Media (DEM)" —
    # the parenthetical is a duplicate of CAT and clutters the race name.
    data = [_clarity_contest('Democratic', 'Mayor of Media (DEM)', ['Alice'], [100])]
    out = _parse_clarity_summary_json(data, is_primary=True)
    assert out['Race_Name'].tolist() == ['DEM Mayor of Media']


def test_clarity_drops_write_in_aggregate_choice():
    # Clarity often emits a "Write-In (Total)" choice as the aggregate write-in
    # count. After _normalize_candidate strips the paren, it becomes
    # "WRITE-IN TOTAL" — which isn't in our drop list. But the plain
    # "Write-Ins" / "Write-In" variants ARE dropped. Test the plain case here.
    data = [_clarity_contest('Democratic', 'Mayor', ['Alice', 'Write-In'], [100, 5])]
    out = _parse_clarity_summary_json(data, is_primary=True)
    assert out['Candidate'].tolist() == ['ALICE']


def test_clarity_empty_input_returns_empty_frame():
    out = _parse_clarity_summary_json([], is_primary=False)
    assert len(out) == 0
    assert list(out.columns) == ['Candidate', 'Race_Name', 'Votes']


# --- _rows_to_tidy -----------------------------------------------------

def test_rows_to_tidy_round_trip_general():
    rows = [
        {'contest_name': 'Mayor of Reading', 'candidate': 'Eddie Moran',
         'party': 'DEM', 'votes': 5000},
        {'contest_name': 'Mayor of Reading', 'candidate': 'Other Person',
         'party': 'REP', 'votes': 3000},
    ]
    out = _rows_to_tidy(rows, is_primary=False)
    assert set(out['Race_Name']) == {'Mayor of Reading'}
    assert set(out['Candidate']) == {'EDDIE MORAN', 'OTHER PERSON'}
    assert sorted(out['Votes'].tolist()) == [3000.0, 5000.0]


def test_rows_to_tidy_prepends_party_to_primary_when_missing():
    # LLM forgot to prefix the primary contest name. Should be added from
    # the `party` field — same safety net the WPRDC parser applies.
    rows = [
        {'contest_name': 'Mayor of Reading', 'candidate': 'Eddie Moran',
         'party': 'DEM', 'votes': 5000},
        {'contest_name': 'Mayor of Reading', 'candidate': 'Mary Smith',
         'party': 'REP', 'votes': 3000},
    ]
    out = _rows_to_tidy(rows, is_primary=True)
    assert set(out['Race_Name']) == {'DEM Mayor of Reading', 'REP Mayor of Reading'}


def test_rows_to_tidy_drops_meta_rows_if_extractor_included_them():
    rows = [
        {'contest_name': 'Mayor', 'candidate': 'Real Person',
         'party': 'DEM', 'votes': 100},
        {'contest_name': 'Mayor', 'candidate': 'Over Votes',
         'party': '', 'votes': 5},
        {'contest_name': 'Mayor', 'candidate': 'Write-Ins',
         'party': '', 'votes': 2},
    ]
    out = _rows_to_tidy(rows, is_primary=False)
    assert out['Candidate'].tolist() == ['REAL PERSON']


def test_rows_to_tidy_applies_candidate_normalization():
    # Same chain as OE: uppercase, periods stripped, middle initials dropped,
    # party suffix stripped. All variants below should collapse to one row.
    rows = [
        {'contest_name': 'Mayor', 'candidate': 'Ryan E. Mackenzie',
         'party': 'REP', 'votes': 100},
        {'contest_name': 'Mayor', 'candidate': 'RYAN MACKENZIE REP',
         'party': 'REP', 'votes': 50},
    ]
    out = _rows_to_tidy(rows, is_primary=False)
    assert out['Candidate'].tolist() == ['RYAN MACKENZIE']
    assert out['Votes'].tolist() == [150.0]


def test_rows_to_tidy_empty_input_returns_empty_frame():
    out = _rows_to_tidy([], is_primary=False)
    assert len(out) == 0
    assert list(out.columns) == ['Candidate', 'Race_Name', 'Votes']


# --- _parse_electionware_lines ---------------------------------------------

def test_electionware_parses_modern_4_column_format():
    # 2023/2025 Berks layout: TOTAL + Election Day + Mail + Provisional.
    lines = [
        "                                                                              ",
        "DEM Judge of the Court of Common Pleas",
        "Vote For 1",
        "                              Election",
        "                      TOTAL             Mail Provisional",
        "                              Day",
        "  Justin D. Bodor               13,589  7,838   5,718    33",
        "  Kurt Geishauser                7,743  4,233   3,487    23",
        "  Jill Scheidt                  14,698  8,454   6,199    45",
        "  Write-In Totals                  101      54     47     0",
        "   Not Assigned                    101      54     47     0",
    ]
    rows = _parse_electionware_lines(lines)
    by_cand = {r['candidate']: r['votes'] for r in rows
               if r['contest_name'] == 'DEM Judge of the Court of Common Pleas'}
    # All four real candidate rows captured with TOTAL only (other columns ignored).
    assert by_cand['Justin D. Bodor'] == 13589
    assert by_cand['Kurt Geishauser'] == 7743
    assert by_cand['Jill Scheidt'] == 14698
    # Write-In Totals / Not Assigned are emitted but the downstream
    # _rows_to_tidy + _is_non_candidate filter drops them. This test just
    # confirms the parser doesn't crash on them.
    assert 'Write-In Totals' in by_cand


def test_electionware_parses_legacy_1_column_format():
    # 2021 Berks layout: just TOTAL, no per-method breakdown.
    lines = [
        "DEM JUSTICE OF SUPREME COURT",
        "Vote For 1",
        "                       TOTAL",
        "  MARIA MCLAUGHLIN              21,453",
        "  Write-In Totals                 165",
    ]
    rows = _parse_electionware_lines(lines)
    by_cand = {r['candidate']: r['votes'] for r in rows
               if r['contest_name'] == 'DEM JUSTICE OF SUPREME COURT'}
    assert by_cand['MARIA MCLAUGHLIN'] == 21453


def test_electionware_skips_multi_seat_contests():
    lines = [
        "DEM Single Seat Race",
        "Vote For 1",
        "  Alice                       100  50  50  0",
        "DEM Multi Seat Race",
        "Vote For 2",
        "  Bob                         200  100  100  0",
        "  Carol                       150   75   75  0",
        "DEM Another Single Seat",
        "Vote For 1",
        "  Dave                         80   40   40  0",
    ]
    rows = _parse_electionware_lines(lines)
    contests = {r['contest_name'] for r in rows}
    assert contests == {'DEM Single Seat Race', 'DEM Another Single Seat'}


def test_electionware_handles_comma_suffixed_candidate_names():
    # "Santoni, Jr." has an internal comma — the non-greedy name match plus
    # 4-number suffix should still find the right boundary.
    lines = [
        "DEM County Commissioner",
        "Vote For 1",
        "  Dante Santoni, Jr.            12,027  6,963   5,027    37",
        "  Jess Royer                    10,857  5,930   4,901    26",
    ]
    rows = _parse_electionware_lines(lines)
    by_cand = {r['candidate']: r['votes'] for r in rows}
    assert by_cand['Dante Santoni, Jr.'] == 12027
    assert by_cand['Jess Royer'] == 10857


def test_electionware_skips_column_header_lines_when_finding_contest_title():
    # If the contest title is followed by column-header lines (Election, TOTAL,
    # Day, Mail, Provisional) before Vote For, those shouldn't be mistaken
    # for the title.
    lines = [
        "DEM Real Contest Title",
        "                      Election",
        "                  TOTAL    Day  Mail Provisional",
        "Vote For 1",
        "  Alice                          100   50  50  0",
    ]
    rows = _parse_electionware_lines(lines)
    assert rows == [
        {'contest_name': 'DEM Real Contest Title', 'candidate': 'Alice', 'votes': 100}
    ]


def test_electionware_walks_across_multiple_contests():
    lines = [
        "REP Office A",
        "Vote For 1",
        "  Alice                        100   50  50  0",
        "  Bob                          200  100  90  10",
        "  Write-In Totals                3    2   1  0",
        "REP Office B",
        "Vote For 1",
        "  Carol                        300  150 140 10",
    ]
    rows = _parse_electionware_lines(lines)
    by_contest = {}
    for r in rows:
        by_contest.setdefault(r['contest_name'], []).append(r['candidate'])
    assert by_contest['REP Office A'] == ['Alice', 'Bob', 'Write-In Totals']
    assert by_contest['REP Office B'] == ['Carol']


def test_electionware_empty_input_returns_empty():
    assert _parse_electionware_lines([]) == []
    assert _parse_electionware_lines(['', '', 'random noise', '']) == []
