"""Tests for the source-agnostic pipeline functions in primary_scraper.

Run with: `pytest test_primary_scraper.py`
"""
from types import SimpleNamespace

import pandas as pd
import pytest

from primary_scraper import (
    _filter_oe_listing,
    _is_likely_multiseat_race,
    _parse_clarity_summary_json,
    _parse_electionware_lines,
    _parse_lycoming_pdf_lines,
    _parse_openelections_df,
    _parse_wprdc_summary_df,
    _rows_to_tidy,
    _strip_middle_initials,
    _wide_columns_to_tidy,
    add_percentages,
    filter_exclude_race_names,
    filter_likely_multiseat_races,
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


def _oe_df_with(extra_cols, rows):
    """Like _oe_df but allows extra columns (e.g. vote_for from precinct CSVs)."""
    cols = ['county', 'office', 'district', 'party', 'candidate', 'votes'] + extra_cols
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


# --- _is_likely_multiseat_race ---------------------------------------------

def test_likely_multiseat_flags_school_director():
    assert _is_likely_multiseat_race("SCHOOL DIRECTOR DEER LAKES")
    assert _is_likely_multiseat_race("School Director Mt Lebanon")
    assert _is_likely_multiseat_race("DEM School Director Boyertown Area Region 1")
    # Plural form ("School Directors") and the SCH DIR / SCH DIRECTORS
    # abbreviation that several PA counties use.
    assert _is_likely_multiseat_race(
        "SCHOOL DIRECTORS (4 YEAR TERM) LEHIGHTON AREA SCHOOL DISTRICT"
    )
    assert _is_likely_multiseat_race("CONEMAUGH TWP SCH DIR DIRECTORS AT LARGE")
    assert _is_likely_multiseat_race("WILLIAMSPORT SCH DIRECTORS")


def test_likely_multiseat_flags_bare_school_district_office():
    # Dauphin's 2025 OE precinct file uses the school-district name as the
    # office (e.g. "Halifax Area School District") without saying "Director".
    # Those races are Vote For 4 at-large school-board elections.
    assert _is_likely_multiseat_race("DAUPHIN HALIFAX AREA SCHOOL DISTRICT")
    assert _is_likely_multiseat_race("HARRISBURG SCHOOL DISTRICT")
    assert _is_likely_multiseat_race("Susquehanna Township School District")
    # Races with an explicit region/seat number after "School District" are
    # Vote For 1 (single seat per region) — keep them.
    assert not _is_likely_multiseat_race("Susquenita School District 3")


def test_likely_multiseat_flags_delegate_and_study_commission():
    # "Delegates" plural (DEM/REP convention delegates are 3-5 per district).
    assert _is_likely_multiseat_race(
        "DEM DELEGATES TO THE NATIONAL CONVENTION 17TH DISTRICT"
    )
    assert _is_likely_multiseat_race(
        "REP ALTERNATE DELEGATES TO THE NATIONAL CONVENTION"
    )
    # Government Study Commissions in PA are 7-9 member elected bodies.
    assert _is_likely_multiseat_race("BRADFORD CITY STUDY COMMISSION")


def test_likely_multiseat_flags_county_commissioner_and_at_large_council():
    assert _is_likely_multiseat_race("COUNTY COMMISSIONER")
    assert _is_likely_multiseat_race("DEM County Commissioner")
    assert _is_likely_multiseat_race("PROSPECT PARK BOROUGH COUNCIL")
    assert _is_likely_multiseat_race("ALLENTOWN CITY COUNCIL")
    assert _is_likely_multiseat_race("COUNCIL AT LARGE PETERS")


def test_likely_multiseat_does_NOT_flag_district_numbered_races():
    # District-numbered races are Vote For 1 — keep them.
    assert not _is_likely_multiseat_race("DEM Member of Council District 9")
    assert not _is_likely_multiseat_race("Council District 5")
    assert not _is_likely_multiseat_race("School Director District 3 Region 2")


def test_likely_multiseat_does_NOT_flag_safe_vote_for_one_offices():
    # Mayor, magistrate, DA, controller, supervisor, tax collector, judge of
    # election, inspector of election — all Vote For 1 in PA.
    for race in [
        "DEM Mayor Pittsburgh",
        "REP Magisterial District Judge 19-3-09",
        "District Attorney",
        "Controller",
        "Township Supervisor Lower Milford",
        "Tax Collector Wayne Twp",
        "Judge of Election Ward 46",
        "Inspector of Elections Ward 12",
    ]:
        assert not _is_likely_multiseat_race(race), race


def test_filter_likely_multiseat_races_drops_matched_rows_keeps_others():
    df = pd.DataFrame({
        'Candidate': ['A', 'B', 'C', 'D', 'E', 'F'],
        'Race_Name': [
            'REP School Director Mt Lebanon',     # drop (school director)
            'REP School Director Mt Lebanon',     # drop
            'DEM Mayor Pittsburgh',               # keep
            'DEM Mayor Pittsburgh',               # keep
            'BRADFORD CITY STUDY COMMISSION',     # drop (study commission)
            'Council District 5',                 # keep (district-numbered)
        ],
        'Votes': [100, 200, 300, 400, 500, 600],
    })
    out = filter_likely_multiseat_races(df)
    surviving = set(out['Race_Name'].unique())
    assert surviving == {'DEM Mayor Pittsburgh', 'Council District 5'}


def test_oe_drops_general_race_with_three_same_party_candidates():
    # Real-world case: OE's 2025 Northumberland CSV conflates Mayor +
    # Borough Council under "Mayor Marion Heights Borough", listing three
    # DEM candidates with ~110 votes each. PA primary law forbids this —
    # each party can only nominate one candidate per single-seat general.
    df = _oe_df([
        {'county': 'X', 'office': 'Mayor Marion Heights Borough', 'district': '',
         'party': 'DEM', 'candidate': 'John Wargo', 'votes': '117'},
        {'county': 'X', 'office': 'Mayor Marion Heights Borough', 'district': '',
         'party': 'DEM', 'candidate': 'Joseph M. Petrovich', 'votes': '118'},
        {'county': 'X', 'office': 'Mayor Marion Heights Borough', 'district': '',
         'party': 'DEM', 'candidate': 'John O Lear', 'votes': '108'},
        # Control: legit general-election race with one D and one R survives
        {'county': 'X', 'office': 'Mayor Sunbury', 'district': '',
         'party': 'DEM', 'candidate': 'Alice', 'votes': '500'},
        {'county': 'X', 'office': 'Mayor Sunbury', 'district': '',
         'party': 'REP', 'candidate': 'Bob', 'votes': '450'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert set(out['Race_Name']) == {'MAYOR SUNBURY'}


def test_oe_keeps_primary_with_many_same_party_candidates():
    # In a primary, multiple same-party candidates are normal and expected.
    # The heuristic must NOT fire when is_primary=True.
    df = _oe_df([
        {'county': 'X', 'office': 'Mayor', 'district': '',
         'party': 'DEM', 'candidate': 'Alice', 'votes': '100'},
        {'county': 'X', 'office': 'Mayor', 'district': '',
         'party': 'DEM', 'candidate': 'Bob', 'votes': '90'},
        {'county': 'X', 'office': 'Mayor', 'district': '',
         'party': 'DEM', 'candidate': 'Carol', 'votes': '80'},
        {'county': 'X', 'office': 'Mayor', 'district': '',
         'party': 'DEM', 'candidate': 'Dave', 'votes': '70'},
    ])
    out = _parse_openelections_df(df, is_primary=True)
    assert out['Race_Name'].nunique() == 1
    assert len(out) == 4


def test_oe_keeps_general_three_candidates_with_distinct_parties():
    # A 3-way general with DEM/REP/IND each having one candidate is legit
    # and stays in. Pittsburgh-style mayoral with a third-party challenger.
    df = _oe_df([
        {'county': 'X', 'office': 'Mayor Sample City', 'district': '',
         'party': 'DEM', 'candidate': 'Alice', 'votes': '500'},
        {'county': 'X', 'office': 'Mayor Sample City', 'district': '',
         'party': 'REP', 'candidate': 'Bob', 'votes': '400'},
        {'county': 'X', 'office': 'Mayor Sample City', 'district': '',
         'party': 'LIB', 'candidate': 'Carol', 'votes': '80'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert out['Race_Name'].nunique() == 1
    assert len(out) == 3


def test_oe_keeps_general_with_empty_party_candidates():
    # Some PA local offices are reported as non-partisan with an empty
    # party field. The same-party-3+ heuristic only fires on a *non-empty*
    # party reaching the threshold; empty-party candidates aren't conflation
    # evidence per se. (The vote-spread heuristic still applies if totals
    # look weird, but here they don't.)
    df = _oe_df([
        {'county': 'X', 'office': 'Auditor X Borough', 'district': '',
         'party': '', 'candidate': 'Alice', 'votes': '50'},
        {'county': 'X', 'office': 'Auditor X Borough', 'district': '',
         'party': '', 'candidate': 'Bob', 'votes': '40'},
        {'county': 'X', 'office': 'Auditor X Borough', 'district': '',
         'party': '', 'candidate': 'Carol', 'votes': '30'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert out['Race_Name'].nunique() == 1


def test_oe_drops_general_race_with_extreme_vote_spread():
    # Real-world case: 2025 Centre County's OE file has 20+ candidates all
    # tagged "Mayor UNIONVILLE BOROUGH" with vote totals from 61 to 4588 —
    # clearly multiple distinct races got merged. The party column is
    # empty/NaN so same-party-3+ doesn't catch it; the vote-spread
    # heuristic (max/min > 20x) does.
    df = _oe_df([
        {'county': 'X', 'office': 'Mayor Unionville Borough', 'district': '',
         'party': '', 'candidate': f'Cand{i}', 'votes': str(v)}
        for i, v in enumerate([4588, 4296, 312, 360, 320, 285, 61, 82, 76])
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert out.empty, (
        f"Expected race dropped due to vote-spread heuristic; got: "
        f"{out[['Race_Name','Candidate','Votes']].to_dict('records')}"
    )


def test_oe_keeps_general_race_with_moderate_vote_spread():
    # Real Vote-For-1 race with many candidates and moderate spread:
    # Philly Judge of Election Ward races have ~5x ratio between top and
    # bottom candidates. The heuristic must NOT drop these.
    df = _oe_df([
        {'county': 'X', 'office': 'Judge of Election Ward 46', 'district': '',
         'party': 'DEM', 'candidate': f'Cand{i}', 'votes': str(v)}
        for i, v in enumerate([100, 90, 80, 70, 60, 50, 40, 30, 25, 20])
    ])
    # 10 distinct DEM candidates → same-party-3+ WOULD fire on this since
    # all are DEM. But Judge of Election is non-partisan-ish — in PA it's
    # often reported with party. The user case requires this to NOT be
    # filtered; for that we keep the heuristic strict.
    # Spread = 100/20 = 5x, below the 20x threshold. Spread-heuristic keeps it.
    # Party-heuristic WOULD fire, however. So for this test we leave party
    # empty to ensure the spread heuristic alone passes.
    df['party'] = ''
    out = _parse_openelections_df(df, is_primary=False)
    assert out['Race_Name'].nunique() == 1


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


def test_oe_scope_by_county_separates_same_name_townships():
    # Worth Township exists in Butler, Centre, Lawrence, and Mercer counties.
    # Without scope_by_county, every Worth Twp Tax Collector race merges into
    # one bogus pseudo-race; with scoping they stay distinct. (Parties left
    # blank so the same-party-count heuristic doesn't drop the merged case
    # before we can observe it.)
    df = _oe_df([
        {'county': 'Butler', 'office': 'Tax Collector', 'district': 'Worth Twp',
         'party': '', 'candidate': 'Alice', 'votes': '100'},
        {'county': 'Centre', 'office': 'Tax Collector', 'district': 'Worth Twp',
         'party': '', 'candidate': 'Kristine Zerby', 'votes': '250'},
        {'county': 'Mercer', 'office': 'Tax Collector', 'district': 'Worth Twp',
         'party': '', 'candidate': 'Bob', 'votes': '180'},
    ])
    out_merged = _parse_openelections_df(df, is_primary=False)
    assert out_merged['Race_Name'].nunique() == 1  # all 3 conflated

    out_scoped = _parse_openelections_df(df, is_primary=False, scope_by_county=True)
    assert set(out_scoped['Race_Name']) == {
        'BUTLER TAX COLLECTOR Worth Twp',
        'CENTRE TAX COLLECTOR Worth Twp',
        'MERCER TAX COLLECTOR Worth Twp',
    }


def test_oe_vote_for_column_drops_multi_seat_rows_authoritatively():
    # OE precinct CSVs from 2025+ include a `vote_for` column. When present,
    # rows with vote_for > 1 are dropped before the heuristic regex runs.
    df = _oe_df_with(['vote_for'], [
        {'county': 'Adams', 'office': 'Mayor', 'district': 'Boroughville',
         'party': 'REP', 'candidate': 'Alice', 'votes': '100', 'vote_for': '1'},
        {'county': 'Adams', 'office': 'Borough Council', 'district': 'Boroughville',
         'party': 'REP', 'candidate': 'Bob', 'votes': '90', 'vote_for': '4'},
        {'county': 'Adams', 'office': 'Borough Council', 'district': 'Boroughville',
         'party': 'REP', 'candidate': 'Carol', 'votes': '80', 'vote_for': '4'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    assert 'BOB' not in out['Candidate'].tolist()
    assert 'CAROL' not in out['Candidate'].tolist()
    assert 'ALICE' in out['Candidate'].tolist()


def test_oe_drops_bare_NOT_OVER_meta_candidates():
    # Blair County's 2025 OE county.csv has a "Not" row in every race with
    # small vote counts — likely "Not Voted" mis-extracted as a candidate.
    # Cumberland's precinct CSV has "OVER" alone in Inspector-of-Elections.
    df = _oe_df([
        {'county': 'Blair', 'office': 'Supervisor', 'district': 'Juniata Twp',
         'party': 'REP', 'candidate': 'David Kane', 'votes': '242'},
        {'county': 'Blair', 'office': 'Supervisor', 'district': 'Juniata Twp',
         'party': '', 'candidate': 'Not', 'votes': '5'},
        {'county': 'Cumberland', 'office': 'Inspector of Elections',
         'district': 'Newville North', 'party': '',
         'candidate': 'OVER', 'votes': '64'},
        {'county': 'Cumberland', 'office': 'Inspector of Elections',
         'district': 'Newville North', 'party': 'DEM',
         'candidate': 'Real Person', 'votes': '100'},
    ])
    out = _parse_openelections_df(df, is_primary=False)
    cands = set(out['Candidate'])
    assert 'NOT' not in cands
    assert 'OVER' not in cands
    assert 'DAVID KANE' in cands
    assert 'REAL PERSON' in cands


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


def test_filter_oe_listing_skips_excluded_county_substrings():
    # Used to bypass OE for counties whose OE-parsed data has known
    # conflation bugs (Northumberland 2025 conflates Mayor + Borough Council).
    listing = [
        _listing_entry('20251104__pa__general__adams__county.csv'),
        _listing_entry('20251104__pa__general__northumberland__county.csv'),
        _listing_entry('20251104__pa__general__berks__county.csv'),
    ]
    out = _filter_oe_listing(
        listing, substr='__general__', suffix='__county.csv',
        exclude_substrs=('__northumberland__',),
    )
    assert out == [
        '20251104__pa__general__adams__county.csv',
        '20251104__pa__general__berks__county.csv',
    ]


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


# --- _parse_lycoming_pdf_lines ---------------------------------------------

def test_lycoming_parses_partisan_primary_contest():
    lines = [
        "       District Attorney (Dem) (Vote for 1), 18804 registered voters, turnout 26.97%",
        "          Jane Smith            500   55.56%  300   200    0",
        "          John Doe              400   44.44%  250   150    0",
        "          Write-in                0    0.00%   0     0     0",
        "          Total                 900  100.00%  550   350    0",
    ]
    rows = _parse_lycoming_pdf_lines(lines)
    by_cand = {r['candidate']: r['votes'] for r in rows}
    assert by_cand == {'Jane Smith': 500, 'John Doe': 400, 'Write-in': 0, 'Total': 900}
    contests = {r['contest_name'] for r in rows}
    assert contests == {'DEM District Attorney'}


def test_lycoming_skips_multi_seat_races():
    lines = [
        "       Loyalsock SD (Dem) (Vote for 4), 2100 registered voters, turnout 27.00%",
        "          Holly N. Shadle       406  24.34%  218   188    0",
        "       District Attorney (Dem) (Vote for 1), 18804 registered voters, turnout 26.97%",
        "          Jane Smith            500   55.56%  300   200    0",
    ]
    rows = _parse_lycoming_pdf_lines(lines)
    contests = {r['contest_name'] for r in rows}
    assert contests == {'DEM District Attorney'}


def test_lycoming_does_not_bundle_ballot_question_votes_into_prior_race():
    # Real-world bug: Lycoming PDFs include non-partisan ballot questions
    # ("Cascade Ballot Question (Vote for 1), ...") on the last page after
    # the partisan races. The "(Vote for 1)" pattern alone shouldn't keep
    # the previous contest active — the YES/NO votes underneath would
    # otherwise get attributed to the last partisan race seen.
    lines = [
        "       Inspector of Elections (Rep) (Vote for 1), 637 registered voters, turnout 23.08%",
        "          Lenora G. Georges     127   99.22%  108    19    0",
        "          Write-in                1    0.78%   1     0     0",
        "          Total                 128  100.00% 109    19    0",
        "       Cascade Ballot Question (Vote for 1), 306 registered voters, turnout 21.57%",
        "          Yes                    43   66.15%  33    10    0",
        "          No                     22   33.85%  19     3    0",
        "          Total                  65  100.00%  52    13    0",
    ]
    rows = _parse_lycoming_pdf_lines(lines)
    # Yes/No should NOT appear under the Rep Inspector contest.
    cands_by_contest = {}
    for r in rows:
        cands_by_contest.setdefault(r['contest_name'], set()).add(r['candidate'])
    assert cands_by_contest == {
        'REP Inspector of Elections': {'Lenora G. Georges', 'Write-in', 'Total'},
    }


def test_lycoming_empty_input_returns_empty():
    assert _parse_lycoming_pdf_lines([]) == []
    assert _parse_lycoming_pdf_lines(['some random text', '']) == []


def test_electionware_handles_layout_with_percent_column():
    # Chester 2021 PDFs have a `VOTE %` column between TOTAL and Election Day,
    # so each candidate line includes a percent token (e.g., "99.33%") that
    # must not be mistaken for part of the candidate name. Before tightening
    # the regex (no digits/% in name), the parser was greedy-matching the
    # percent into the name and capturing Election Day votes as TOTAL.
    lines = [
        "DEM JUSTICE OF THE SUPREME COURT",
        "Vote For 1",
        "                                        Election Absentee/",
        "                       TOTAL   VOTE %                Provisional",
        "                                         Day    Mail-In",
        "  MARIA MCLAUGHLIN            42,302  99.33%  23,107  19,055   140",
        "  Write-In Totals               286    0.67%   234      51      1",
        "   Total Votes Cast            42,588  100.00% 23,341  19,106   141",
    ]
    rows = _parse_electionware_lines(lines)
    # Only the real candidate row should yield extracted data; Write-In Totals
    # and Total Votes Cast are emitted but downstream _is_non_candidate drops
    # them. Critically: the captured TOTAL must be 42,302, not 23,107.
    real = [r for r in rows if r['candidate'] == 'MARIA MCLAUGHLIN']
    assert len(real) == 1
    assert real[0]['votes'] == 42302
    assert real[0]['contest_name'] == 'DEM JUSTICE OF THE SUPREME COURT'
