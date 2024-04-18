import pandas as pd
import numpy as np
def special_population(file_path, special_population_file_path, final_population_file_path):
    df = pd.read_excel(file_path, skiprows=7)

    unique_age_count = df['Age (Group)'].value_counts().reset_index()
    unique_age_count.columns = ['Age (Group)', 'Count']
    df = unique_age_count
    df['Age (Group)'] = df['Age (Group)'].replace('', 'Unknown')

    # df.to_excel(special_population_file_path, index=False)
   
    try:
        Neonate = df[df['Age (Group)'] == 'Neonate']['Count'].values[0]
    except (IndexError) as error:
        Neonate = 0
    try:
        Infant = df[df['Age (Group)'] == 'Infant']['Count'].values[0]
    except (IndexError) as error:
        Infant = 0
    try:
        Children = df[df['Age (Group)'] == 'Children']['Count'].values[0]
    except (IndexError) as error:
        Children = 0
    try:
        Adolescent = df[df['Age (Group)'] == 'Adolescent']['Count'].values[0]
    except (IndexError) as error:
        Adolescent = 0
    try:
        Adult = df[df['Age (Group)'] == 'Adult']['Count'].values[0]
    except (IndexError) as error:
        Adult = 0
    try:
        Elderly = df[df['Age (Group)'] == 'Elderly']['Count'].values[0]
    except (IndexError) as error:
        Elderly = 0
    try:
        Unknown = df[df['Age (Group)'] == '']['Count'].values[0]
    except (IndexError) as error:
        Unknown = 0

    if file_path == LL_reporting_path:
        data = {
            'Age Group': ['Neonate (from 0 to ≤27 days)', 'Infant (from 28 days to 23 months)', 'Children (from 2 to 11 years)', 'Adolescent (from 12 to <18 years)', 'Adult (from 18 to <65 years)', 'Elderly (≥65 years)', 'Unknown*'],
            'Current Reporting Interval (01 Oct 2013 to 30 Sep 2015)': [Neonate, Infant, Children, Adolescent, Adult, Elderly, Unknown],
        }

        df1 = pd.DataFrame(data)
        df1.to_excel(final_population_file_path, index=False)

    if file_path == LL_cumulative_path:
        df = pd.read_excel(final_population_file_path)
        data  = {
                'Cumulative Period (till 30 Sep 2015)': [Neonate, Infant, Children, Adolescent, Adult, Elderly, Unknown],
            }
        df1 = pd.DataFrame(data)
        df3 = pd.concat([df, df1], axis=1)

        numeric_cols = df3.select_dtypes(include=[np.number]).columns
        df3.loc['Total'] = df3[numeric_cols].sum()
        df3.fillna('Total', inplace=True)


        df3.to_excel(final_population_file_path, index = False)
        

def run_functions(LL_reporting_path,LL_cumulative_path,final_population_file_path):
    special_population_file_path = 'C:\\Users\\2113196\\Downloads\\special_population_reporting.xlsx'
    special_population(LL_reporting_path, special_population_file_path, final_population_file_path)

    file_path = 'C:\\Users\\2113196\\Downloads\\LL_reporting.xlsx'
    special_population_file_path = 'C:\\Users\\2113196\\Downloads\\special_population_cumulative.xlsx'
    special_population(LL_cumulative_path, special_population_file_path, final_population_file_path)


LL_reporting_path = "C:\\Users\\2113196\\Downloads\\LLs_Reporting Interval (01 Oct 2015 to 30 Sep 2017).xlsx"
LL_cumulative_path = "C:\\Users\\2113196\\Downloads\\LLs_Cumulative Period (IBD to 30 Sep 2017)_Sample.xlsb"
final_population_file_path = 'C:\\Users\\2113196\\Downloads\\LL_special_population.xlsx'
run_functions(LL_reporting_path,LL_cumulative_path,final_population_file_path)