import pandas as pd
from pathlib import Path


def main():

    xlsx = Path("demo/DiaryWorkedHours.xlsx")

    df = pd.read_excel(xlsx, sheet_name="Sheet1")

    df.columns = df.columns.str.strip()

    df_filtered = df[df['Mängd'] >= 0]

    total_mangd_per_person = (df_filtered.groupby(['IndividNamn', 'Dagboksdatum', 'Status'])['Mängd']
                              .sum()
                              .reset_index())

    reshaped_df = total_mangd_per_person.pivot(index=['IndividNamn', 'Dagboksdatum'],
                                               columns='Status',
                                               values='Mängd').reset_index()

    melted_df = reshaped_df.melt(id_vars=['Dagboksdatum', 'IndividNamn'], var_name='Metric')

    pivot_df = melted_df.pivot_table(index='Dagboksdatum', columns=['IndividNamn', 'Metric'], values='value')

    pivot_df = pivot_df.sort_index(axis=1, level=[0, 1])

    pivot_df.to_excel('demo/pivoted_dataframe.xlsx', merge_cells=True)


if __name__ == "__main__":
    main()
