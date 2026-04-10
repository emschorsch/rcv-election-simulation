# pip install requests pandas xlsxwriter openpyxl
import pandas as pd

# URLs for each year - from https://vote.phila.gov/resources-data/past-election-results/archived-data-sets/
year_urls = {
    2007: "https://vote.phila.gov/files/election-results/2007_Primary_WD.csv",
    2011: "https://vote.phila.gov/files/election-results/2011_Primary_WD.csv",
    2015: "https://vote.phila.gov/files/election-results/2015_PRIMARY/2015_PRIMARY_-_WARD_RESULTS_-_PUBLIC.xlsx",
    2019: "https://vote.phila.gov/files/election-results/2019_PRIMARY/2019_PRIMARY_RESULTS_BY_TYPE_BY_PRECINCT.xlsx",
    2023: "https://vote.phila.gov/media/2023_Primary_Results.xlsx"
}

sheet_name = 'Philadelphia_Primary_Mayor_DistrictCouncil.xlsx'
include_all = False # Include all races
if include_all:
    sheet_name = 'Philadelphia_Primary_AllRaces.xlsx'

with pd.ExcelWriter(sheet_name, engine='xlsxwriter') as writer:
    for year, url in year_urls.items():
        print(f"Processing {year}...")

        if year in [2007, 2011, 2023]:  # Wide CSVs
            if year == 2023:
                df = pd.read_excel(url, sheet_name='Totals')
                df.columns = df.columns.str.strip()  # remove leading/trailing whitespace and newlines
                df.columns = df.columns.str.replace('\n', ' - ', regex=False).str.strip()
                df.columns = df.columns.str.replace('AT-LARGE', 'AT LARGE', regex=False)
                df.columns = df.columns.str.replace('COUNCIL - ', 'COUNCIL', regex=False)
                df.columns = df.columns.str.replace('Write-in', 'write in', regex=False)
                start_idx = df.columns.get_loc('Council')  # drop columns before Council
                df_votes = df.iloc[:, start_idx:]
                vote_cols = df_votes.columns[1:]  # all columns after candidate column
            else:
                df = pd.read_csv(url)
                id_cols = ['WARD', 'DIVISION']
                vote_cols = [col for col in df.columns if col not in id_cols]

            # Remove commas and convert to numeric
            df[vote_cols] = df[vote_cols].replace({',': ''}, regex=True)
            df[vote_cols] = df[vote_cols].apply(pd.to_numeric, errors='coerce').fillna(0)

            totals = df[vote_cols].sum()
            rows = []
            for col, votes in totals.items():
                if '-' in col:
                    parts = [p.strip() for p in col.split('-')]
                    if year == 2023:
                        if len(parts) == 2:
                            race_name = parts[0]
                            candidate = parts[1]
                        else:
                            race_name = ' - '.join(parts[0:2])
                            candidate = ' - '.join(parts[2:])
                    else:
                        candidate = parts[0]
                        race_name = ' - '.join(parts[1:])
                    candidate = candidate.strip()
                    race_name = race_name.strip()
                else:
                    candidate = 'Unknown'
                    race_name = col.strip()
                rows.append({'Candidate': candidate, 'Race_Name': race_name, 'Votes': votes})
            tidy_df = pd.DataFrame(rows)

        elif year == 2015:  # CATEGORY + PARTY
            df = pd.read_excel(url)
            df['Race_Name'] = df['CATEGORY'].astype(str).str.strip() + " - " + df['PARTY'].astype(str).str.strip()
            df['Candidate'] = df['SELECTION'].astype(str).str.strip()
            df['Votes'] = pd.to_numeric(df['TOTAL'], errors='coerce').fillna(0)
            tidy_df = df.groupby(['Race_Name', 'Candidate'], as_index=False)['Votes'].sum()

        elif year == 2019:  # CATEGORY + SELECTION + VOTE COUNT
            df = pd.read_excel(url, sheet_name='2019 PRIMARY')
            df.columns = df.columns.str.strip()  # remove leading/trailing whitespace and newlines
            df['Race_Name'] = df['CATEGORY'].astype(str).str.strip()
            df['Candidate'] = df['SELECTION'].astype(str).str.strip()
            df['Votes'] = pd.to_numeric(df['VOTE COUNT'], errors='coerce').fillna(0)
            tidy_df = df.groupby(['Race_Name', 'Candidate'], as_index=False)['Votes'].sum()

        else:
            raise ValueError(f"Unknown year/file type for {year}")

        # Ensure Votes numeric
        tidy_df['Votes'] = pd.to_numeric(tidy_df['Votes'], errors='coerce').fillna(0)

        # Compute total votes per race
        tidy_df['Total_Votes_Race'] = tidy_df.groupby('Race_Name')['Votes'].transform('sum')

        # Percent calculation
        tidy_df['Percent'] = tidy_df['Votes'] / tidy_df['Total_Votes_Race'] * 100

        # Keep only races where the top candidate had <50%
        max_percent = tidy_df.groupby('Race_Name')['Percent'].transform('max')
        tidy_df = tidy_df[max_percent < 50]

        # Filter only Mayor or District Council races
        tidy_df['Race_Name'] = tidy_df['Race_Name'].str.strip()
        if not include_all:
            tidy_df = tidy_df[tidy_df['Race_Name'].str.lower().str.contains("mayor|district council")]

        # Count real candidates ignoring Write In
        real_candidate_counts = (
            tidy_df[tidy_df['Candidate'].str.lower() != 'write in']
            .groupby('Race_Name')['Candidate']
            .count()
            .reset_index()
            .rename(columns={'Candidate': 'Num_Real_Candidates'})
        )
        tidy_df = tidy_df.merge(real_candidate_counts, on='Race_Name', how='left')

        # Keep races with ≥3 candidates
        # tidy_df = tidy_df[tidy_df['Num_Real_Candidates'] >= 3]

        # Drop helper column
        tidy_df = tidy_df[['Candidate', 'Race_Name', 'Votes', 'Percent']]

        # Sort within each race
        tidy_df = tidy_df.sort_values(['Race_Name', 'Votes'], ascending=[True, False])

        # Add blank line between races
        final_rows = []
        for race, group in tidy_df.groupby('Race_Name', sort=True):
            final_rows.append(group)
            final_rows.append(pd.DataFrame([{'Candidate':'', 'Race_Name':'', 'Votes':'', 'Percent':''}]))

        if final_rows:
            final_df = pd.concat(final_rows, ignore_index=True)
        else:
            final_df = pd.DataFrame(columns=['Candidate', 'Race_Name', 'Votes', 'Percent'])

        # Write to Excel sheet
        final_df.to_excel(writer, sheet_name=str(year), index=False)

print("All years processed and saved to '." + sheet_name)
