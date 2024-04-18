import pandas as pd
import numpy as np

def off_label_use_population(LL_path, Cognizant_CRS_path):
    df1 = pd.read_excel(LL_path, skiprows=7)
    df2 = pd.read_excel(Cognizant_CRS_path, sheet_name='Search Terms')

    index = df1.columns.get_loc('Indication(s)')
    df1.insert(loc=index+1, column='Off label use', value=0)
    df1.insert(loc=index+2, column='Population', value=0)

    df3 = pd.merge(df1, df2, left_on='Indication(s)', right_on="Approved, Related to Approve and Unknown indication", how='inner')
    df3_rest = df1[~df1['Indication(s)'].isin(df2["Approved, Related to Approve and Unknown indication"])]

    df3.loc[df1['Indication(s)'] == '1) Product used for unknown indication', 'Off label use'] = 'Unknown'
    df3.loc[df1['Indication(s)'] == '1) Hyperphosphataemia', 'Off label use'] = 'Approved'
    df3.loc[df1['Indication(s)'] == '1) Blood phosphorus increased', 'Off label use'] = 'Related to approved'

    merged_df = pd.concat([df3, df3_rest])

    # df3.loc[df1['Indication(s)'] == 0, 'Off label use'] = 'Off label use'
    merged_df['Off label use'] = merged_df['Off label use'].replace(0, 'Off label use')
    df4 = merged_df

    df4['Population'] = df4['Off label use'].apply(lambda x: 'Off label population' if x == 'Off label use' else 'Remaining population')
    df4['Population'] = df4['Population'].fillna('Remaining population')

    prr_df = df4[['MedDRA PT', 'Body system', 'Indication(s)', 'Off label use', 'Population']]
    prr_df.to_excel("C:\\Users\\2113196\\Downloads\\final.xlsx", index=False)

    prr_calculation(prr_df)

def prr_calculation(off_label_population_file):

    df1 = off_label_population_file
    # Count 'Off label population' and 'Remaining population' based on 'Body system' column
    off_label_count = df1[df1['Population'] == 'Off label population']['Body system'].value_counts()
    remaining_count = df1[df1['Population'] == 'Remaining population']['Body system'].value_counts()

    df2 = pd.DataFrame({'Off label population': off_label_count, 'Remaining population': remaining_count})

    index = df2.columns.get_loc('Off label population')
    df2.insert(loc=index+1, column='Proption of event in Off label population', value=0)
    df2.insert(loc=index+3, column='Proption of event in Remaining population', value=0)
    df2.insert(loc=index+4, column='PRR', value=0)

    off_label_sum = df2['Off label population'].sum()
    remaiining_population_sum = df2['Remaining population'].sum()

    df2['Proption of event in Off label population'] = (df2['Off label population'] / off_label_sum).round(2)
    df2['Proption of event in Remaining population'] = (df2['Remaining population'] / remaiining_population_sum).round(2)
    df2['PRR'] = (df2['Proption of event in Off label population'] / df2['Proption of event in Remaining population']).round(2)

    # Identify numeric columns
    numeric_cols = df2.select_dtypes(include=[np.number]).columns

    df2.loc['Total'] = df2[numeric_cols].sum()
    df2.fillna(0, inplace=True)

    df2.to_excel("C:\\Users\\2113196\\Downloads\\PRR.xlsx")

Cognizant_CRS_path = "C:\\Users\\2113196\\Downloads\\Cognizant_CRS.xlsx"
LL_reporting_path = "C:\\Users\\2113196\\Downloads\\LL_reporting.xlsx"
LL_cumulative_path = "C:\\Users\\2113196\\Downloads\\LLs_Cumulative Period (IBD to 30 Sep 2017)_Sample.xlsb"
off_label_use_population(LL_cumulative_path, Cognizant_CRS_path)
