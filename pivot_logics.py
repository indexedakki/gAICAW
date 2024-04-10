def pivot_logics(df):
    country_counts_df = df['Country'].value_counts()

    # Data for the table
    data = {
        'Age Group': ['Neonate (from 0 to ≤27 days)', 'Infant (from 28 days to 23 months)', 'Children (from 2 to 11 years)', 'Adolescent (from 12 to <18 years)', 'Adult (from 18 to <65 years)', 'Elderly (≥65 years)', 'Unknown*', 'Total'],
        'Current Reporting Interval (01 Oct 2013 to 30 Sep 2015)': [0, 0, 0, 0, 10, 6, 2, 18],
        'Cumulative Period (till 30 Sep 2015)': [0, 0, 0, 0, 10, 6, 2, 18]
    }
    df_age_group = pd.DataFrame(data)
    df_age_group.iloc[0,1] = df[df['Age (Group)'] == 'Neonate']
    df_age_group.iloc[1,1] = df[df['Age (Group)'] == 'Infant']
    df_age_group.iloc[2,1] = df[df['Age (Group)'] == 'Children']
    df_age_group.iloc[3,1] = df[df['Age (Group)'] == 'Adolescent']
    df_age_group.iloc[4,1] = df[df['Age (Group)'] == 'Adult']
    df_age_group.iloc[5,1] = df[df['Age (Group)'] == 'Elderly']
    df_age_group.iloc[6,1] = df[df['Age (Group)'] == '']

pivot_logics(df_creation_excel)