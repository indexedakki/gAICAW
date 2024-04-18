import pandas as pd
import numpy as np
def case_seriousness(LL_path, Cognizant_CRS_path, lookup_column, file_path_case_seriousness, line_listing_table_path, paragraph_path):
    df1 = pd.read_excel(LL_path, skiprows=7)
    df2 = pd.read_excel(Cognizant_CRS_path, sheet_name='Search Terms')

    df3 = pd.merge(df1, df2, left_on='MedDRA PT', right_on=lookup_column, how='inner')

    # df3.to_excel('C:\\Users\\2113196\\Downloads\\df3.xlsx', index=False)

    df4 = df3['Case Number'].unique()
    df4 = pd.DataFrame(df4, columns=['Case Number'])
    df5 = pd.merge(df4, df1, on='Case Number', how='inner')
    # Concatenate DataFrames vertically and reset index
    # df5.to_excel('C:\\Users\\2113196\\Downloads\\df5.xlsx', index=False)
    # df6 = pd.merge(df5, df3, left_on='Case Number', right_on='Case Number', how='outer')
    
    df_case_seriousness = df5[df5['CaseFlag'] == 1]

    file_path_case_seriousness = file_path_case_seriousness
    df5.to_excel(file_path_case_seriousness, index=False)

    df_creation_excel = df_case_seriousness

    create_excel_file(line_listing_table_path)

    report_types, medically_confirmed_cases, seriousness_events, target_cells = line_listing_table_logic()

    final_line_listing_table = process_dataframes_case_seriousness(df_creation_excel, line_listing_table_path, report_types, medically_confirmed_cases, seriousness_events, target_cells)

    # saving_line_listing_to_html(line_listing_table_path)

    updated_paragraph1, updated_paragraph2 = table_to_text_conversion(final_line_listing_table, df_creation_excel)
    AI_for_better_table_to_text(updated_paragraph1, updated_paragraph2, paragraph_path)


def case_seriousness_without_lookup(LL_path, age_type, file_path_case_seriousness, line_listing_table_path, paragraph_path):
    df1 = pd.read_excel(LL_path, skiprows=7)

    if age_type == 18:
        filtered_df = df1[df1['Age'] <= 18]
    elif age_type == 65:
        filtered_df = df1[df1['Age'] >= 65]

    filtered_df.to_excel(file_path_case_seriousness, index=False)

    df_case_seriousness = filtered_df[filtered_df['CaseFlag'] == 1]

    df_creation_excel = df_case_seriousness

    create_excel_file(line_listing_table_path)

    report_types, medically_confirmed_cases, seriousness_events, target_cells = line_listing_table_logic()

    final_line_listing_table = process_dataframes_case_seriousness(df_creation_excel, line_listing_table_path, report_types, medically_confirmed_cases, seriousness_events, target_cells)

    # saving_line_listing_to_html(line_listing_table_path)

    updated_paragraph1, updated_paragraph2 = table_to_text_conversion(final_line_listing_table, df_creation_excel)
    AI_for_better_table_to_text(updated_paragraph1, updated_paragraph2, paragraph_path)


def event_seriousness(LL_path, Cognizant_CRS_path, lookup_column, file_path_event_seriousness, line_listing_table_path, paragraph_path):
    df1 = pd.read_excel(LL_path, skiprows=7)
    df2 = pd.read_excel(Cognizant_CRS_path, sheet_name='Search Terms')

    df3 = pd.merge(df1, df2, left_on='MedDRA PT', right_on=lookup_column, how='inner')
    df3['Seriousness (Event)'] = df3.groupby('Case Number')['Seriousness (Event)'].transform(lambda x: 'Serious' if x.unique().size > 1 else x)

    df3.to_excel(file_path_event_seriousness, index=False)

    df3.drop_duplicates(subset='Case Number', keep='first', inplace=True)

    df_creation_excel = df3

    create_excel_file(line_listing_table_path)

    report_types, medically_confirmed_cases, seriousness_events, target_cells = line_listing_table_logic()

    final_line_listing_table = process_dataframes_case_seriousness(df_creation_excel, line_listing_table_path, report_types, medically_confirmed_cases, seriousness_events, target_cells)

    # saving_line_listing_to_html(line_listing_table_path)

    updated_paragraph1, updated_paragraph2 = table_to_text_conversion(final_line_listing_table, df_creation_excel)
    AI_for_better_table_to_text(updated_paragraph1, updated_paragraph2, paragraph_path)

def event_seriousness_without_lookup(LL_path, age_type, file_path_event_seriousness, line_listing_table_path, paragraph_path):
    df1 = pd.read_excel(LL_path, skiprows=7)

    if age_type == 18:
        filtered_df = df1[df1['Age'] <= 18]
    elif age_type == 65:
        filtered_df = df1[df1['Age'] >= 65]
    
    df_event_seriousness = filtered_df.to_excel(file_path_event_seriousness, index=False)

    df_creation_excel = pd.read_excel(file_path_event_seriousness)

    create_excel_file(line_listing_table_path)

    report_types, medically_confirmed_cases, seriousness_events, target_cells = line_listing_table_logic()

    final_line_listing_table = process_dataframes_case_seriousness(df_creation_excel, line_listing_table_path, report_types, medically_confirmed_cases, seriousness_events, target_cells)

    # saving_line_listing_to_html(line_listing_table_path)

    updated_paragraph1, updated_paragraph2 = table_to_text_conversion(final_line_listing_table, df_creation_excel)
    AI_for_better_table_to_text(updated_paragraph1, updated_paragraph2, paragraph_path)

def off_label_use_population(LL_path, Cognizant_CRS_path, lookup_column, file_path_event_seriousness, line_listing_table_path, paragraph_path):
    df1 = pd.read_excel(LL_path, skiprows=7)
    df2 = pd.read_excel(Cognizant_CRS_path, sheet_name='Search Terms')

    index = df1.columns.get_loc('Indication(s)')
    df1.insert(loc=index+1, column='Off label use', value=0)
    df1.insert(loc=index+2, column='Population', value=0)

    df3 = pd.merge(df1, df2, left_on='Indication(s)', right_on=lookup_column, how='inner')
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

    df5 = df4[df4['Population'] == 'Off label population']

    df_creation_excel = df5

    create_excel_file(line_listing_table_path)

    report_types, medically_confirmed_cases, seriousness_events, target_cells = line_listing_table_logic()

    final_line_listing_table = process_dataframes_case_seriousness(df_creation_excel, line_listing_table_path, report_types, medically_confirmed_cases, seriousness_events, target_cells)

    # saving_line_listing_to_html(line_listing_table_path)

    updated_paragraph1, updated_paragraph2 = table_to_text_conversion(final_line_listing_table, df_creation_excel)
    AI_for_better_table_to_text(updated_paragraph1, updated_paragraph2, paragraph_path)

    prr_df = df4[['MedDRA PT', 'Body system', 'Indication(s)', 'Off label use', 'Population']]

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


def create_excel_file(line_listing_table_path):
    df = pd.DataFrame(index=range(10), columns=range(6))
    # Fill NaN values with empty strings
    df.fillna('', inplace=True) 

    df.iat[0, 1] = 'Serious' 
    df.iat[0, 3] = 'Not-Serious'
    df.iat[1, 0] = 'Report Type'

    # report_type_unique = df_creation_excel['Report Type'].unique()
    # for report_types_i in range (len(report_type_unique)):
    #     df.iat[report_types_i+2,0] = report_type_unique[report_types_i]
    df.iat[2, 0] = 'Clinical Trials'
    df.iat[3, 0] = 'Spontaneous Report'
    df.iat[4, 0] = 'Literature'
    df.iat[5, 0] = 'Post market survey'
    df.iat[6, 0] = 'Total'

    df.iat[1, 1] = 'HCP'
    df.iat[1, 2] = 'Non HCP'
    df.iat[1, 3] = 'HCP'
    df.iat[1, 4] = 'Non HCP'
    df.iat[1, 5] = 'Total'

    df.to_excel(line_listing_table_path, index=False)


def process_dataframes_event_seriousness(df1, file2_path, report_types, medically_confirmed_cases, seriousness_events, target_cells):
    df2 = pd.read_excel(file2_path)

    # Define a function to filter df1 and update df2
    def update_df2(report_type, medically_confirmed_case, seriousness_event, target_cell):
        df_filtered = df1[(df1['Report Type'].isin(report_type)) & 
                          (df1['Medically confirmed case'].isin(medically_confirmed_case)) & 
                          (df1['Seriousness (Event)'].isin(seriousness_event))]
        df2.iloc[target_cell[0], target_cell[1]] = len(df_filtered)
    
    # Apply the function for each set of conditions and target cell
    for report_type, medically_confirmed_case, seriousness_event, target_cell in zip(report_types, medically_confirmed_cases, seriousness_events, target_cells):
        update_df2(report_type, medically_confirmed_case, seriousness_event, target_cell)
    
    # Totalling logic
    # Find sum of columns specified
    # Get the number of rows
    num_rows = df2.shape[0]

    serious_hcp_sum = df2[1].iloc[2:].sum()
    df2.iloc[num_rows-1,1] = serious_hcp_sum
    serious_nonhcp_sum = df2[2].iloc[2:].sum()
    df2.iloc[num_rows-1,2] = serious_nonhcp_sum
    nonserious_hcp_sum = df2[3].iloc[2:].sum()
    df2.iloc[num_rows-1,3] = nonserious_hcp_sum
    nonserious_nonhcp_sum = df2[4].iloc[2:].sum()
    df2.iloc[num_rows-1,4] = nonserious_nonhcp_sum

    # Sum up the specific row till the end, excluding the specified column
    for i in range(num_rows-2):
        sum_result = df2.iloc[i+2].drop(0).sum()
        df2.iloc[i+2,5] = sum_result
    df2 = df2.fillna(0)
    df2.to_excel(file2_path, index=False)

    return df2

def process_dataframes_case_seriousness(df1, file2_path, report_types, medically_confirmed_cases, seriousness_events, target_cells):
    df2 = pd.read_excel(file2_path)

    def update_df2(report_type, medically_confirmed_case, seriousness_event, target_cell):
        df_filtered = df1[(df1['Report Type'].isin(report_type)) & 
                          (df1["""Medically confirmed
case"""].isin(medically_confirmed_case)) &
                          (df1['''Case 
seriousness'''].isin(seriousness_event))]
        df2.iloc[target_cell[0], target_cell[1]] = len(df_filtered)
    
    for report_type, medically_confirmed_case, seriousness_event, target_cell in zip(report_types, medically_confirmed_cases, seriousness_events, target_cells):
        update_df2(report_type, medically_confirmed_case, seriousness_event, target_cell)

    num_rows = df2.shape[0]

    serious_hcp_sum = df2[1].iloc[2:].sum()
    df2.iloc[num_rows-1,1] = serious_hcp_sum
    serious_nonhcp_sum = df2[2].iloc[2:].sum()
    df2.iloc[num_rows-1,2] = serious_nonhcp_sum
    nonserious_hcp_sum = df2[3].iloc[2:].sum()
    df2.iloc[num_rows-1,3] = nonserious_hcp_sum
    nonserious_nonhcp_sum = df2[4].iloc[2:].sum()
    df2.iloc[num_rows-1,4] = nonserious_nonhcp_sum

    for i in range(num_rows-2):
        sum_result = df2.iloc[i+2].drop(0).sum()
        df2.iloc[i+2,5] = sum_result

    df2 = df2.fillna(0)
    df2.to_excel(file2_path, index=False)

    return df2

def line_listing_table_logic():
    # report_type_unique = df_creation_excel['Report Type'].unique()
    report_type_unique = ['Clinical trial', 'Spontaneous Report', 'Literature', 'Post-marketing Surveillance']

    # Incase one more report_types come i.e 5, generally it's for 4*4 matrix, increase multiplication to make it 5*4
    report_types = [item for report_type in report_type_unique for item in [[report_type]] * 4] 

    medically_confirmed_cases_unique = ['Yes', 'No']
    medically_confirmed_cases_not_list = medically_confirmed_cases_unique * 10
    medically_confirmed_cases = [[element] for element in medically_confirmed_cases_not_list]

    seriousness_events_unique = ['Serious', 'Serious', 'Not Serious', 'Not Serious']
    seriousness_events_not_list = seriousness_events_unique * 5
    seriousness_events = [[element] for element in seriousness_events_not_list]
    
    target_cells = [(i, j) for i in range(2, 7) for j in range(1, 5)]

    return report_types, medically_confirmed_cases, seriousness_events, target_cells

from spire.xls import *
from spire.xls.common import *
def saving_line_listing_to_html(line_listing_table_path):
    workbook = Workbook()
    # Load an Excel file
    workbook.LoadFromFile(line_listing_table_path)
    workbook.SaveToHtml("C:\\Users\\2113196\\Downloads\\ExcelToHTML.html")
    workbook.Dispose()

def table_to_text_conversion(df, df2):
    # Get the number of rows
    i = df.shape[0]-1

    placeholder_1 = df.iloc[i, 5]

    report_type_1 = df.iloc[2, 0]
    placeholder_2 = df.iloc[2, 5]
    report_type_2 = df.iloc[3, 0]
    placeholder_3 = df.iloc[3, 5]
    report_type_3 = df.iloc[4, 0]
    placeholder_4 = df.iloc[4, 5]
    report_type_4 = df.iloc[5, 0]
    placeholder_5 = df.iloc[4, 5]

    placeholder_6 = df.iloc[i, 1] + df.iloc[i, 2]
    placeholder_7 = df.iloc[i, 3] + df.iloc[i, 4]
    placeholder_8 = df.iloc[i, 1] + df.iloc[i, 3]
    placeholder_9 = df.iloc[i, 2] + df.iloc[i, 4]

    updated_paragraph1 = f"""During the reporting interval, a total of {placeholder_1} cases were retrieved based on 
    search criteria mentioned in Appendix 2. Of these {placeholder_1} cases, {placeholder_2} were retrieved from 
    {report_type_1}, {placeholder_3} were {report_type_2} reports, {placeholder_4} were {report_type_3} reports, 
    {placeholder_5} were {report_type_4} reports. Of the {placeholder_1} cases, {placeholder_6} 
    were serious cases and {placeholder_7} were non-serious, and further {placeholder_8} were reported by an HCP 
    while {placeholder_9} was reported by non-HCP."""
    
    
    try:
        placeholder_1 = df.iloc[i, 5]
        placeholder_2 = df2['Age'].count()
        placeholder_3 = int(df2['Age'].min())
        placeholder_4 = int(df2['Age'].max())
        placeholder_5 = int(df2['Age'].mean())
        placeholder_6 = int(df2['Age'].median())
        placeholder_7 = df2['Age'].count()
        placeholder_8 = df2[df2['Gender'] == 'Male']['Gender'].count()
        placeholder_9 = df2[df2['Gender'] == 'Female']['Gender'].count()
    except (ValueError, TypeError) as error:
        pass

    updated_paragraph2 = f"""Of these {placeholder_1} cases, age was reported in {placeholder_2} cases which ranged 
    from {placeholder_3} years to {placeholder_4} years, with mean age of {placeholder_5} years and median age of 
    {placeholder_6} years. Gender of patients was reported in {placeholder_7} cases, {placeholder_8} were males 
    while {placeholder_9} were female patients. No cases were reported with a fatal outcome."""

    return updated_paragraph1, updated_paragraph2

def AI_for_better_table_to_text(updated_paragraph1, updated_paragraph2, paragraph_path):
    import os
    from openai import AzureOpenAI

    from dotenv import load_dotenv, find_dotenv
    load_dotenv(find_dotenv(), override=True)

    api_key=os.getenv("OPENAI_API_KEY")
    api_version=os.getenv("OPENAI_API_VERSION")
    azure_endpoint=os.getenv("OPENAI_API_BASE")

    client = AzureOpenAI(
    api_key=api_key,  
    api_version=api_version,
    azure_endpoint=azure_endpoint 
    )
    
    conversation = [{"role": "system", "content": "You are a helpful assistant."}]

    user_input = f"""{updated_paragraph1}
        
        Amend this paragraph in a better way and remove those reports where 0 cases were found
        
        """  

    conversation.append({"role": "user", "content": user_input})

    response = client.chat.completions.create(
        model="LMM_OPENAI",
        messages=conversation
    )

    conversation.append({"role": "assistant", "content": response.choices[0].message.content})

    final_paragraph = str(response.choices[0].message.content)

    with open(paragraph_path, 'w') as text_file:
        text_file.write(final_paragraph)

    print("\n" + response.choices[0].message.content + "\n")

def run_functions(Cognizant_CRS_path, LL_reporting_path, LL_cumulative_path, lookup_column, safety_topic):

    safety_topic = "pediatric"
    type = "case"
    paragraph_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_{type}_seriousness_reporting'
    file_path_case_seriousness = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_{type}_seriousness_reporting.xlsx'
    line_listing_table_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_{type}_distribution_reporting.xlsx'
    case_seriousness_without_lookup(LL_reporting_path, 18, file_path_case_seriousness, line_listing_table_path, paragraph_path)

    safety_topic = "geriatric"
    paragraph_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_{type}_seriousness_reporting'
    file_path_case_seriousness = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_{type}_seriousness_reporting.xlsx'
    line_listing_table_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_{type}_distribution_reporting.xlsx'
    case_seriousness_without_lookup(LL_reporting_path, 65, file_path_case_seriousness, line_listing_table_path, paragraph_path)


    # paragraph_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_event_seriousness_reporting'
    # file_path_case_seriousness = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_event_seriousness_reporting.xlsx'
    # line_listing_table_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_event_distribution_reporting.xlsx'
    # event_seriousness(LL_reporting_path, Cognizant_CRS_path, lookup_column, file_path_case_seriousness, line_listing_table_path, paragraph_path)

    # paragraph_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_event_seriousness_cumulative'
    # file_path_case_seriousness = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_event_seriousness_cumulative.xlsx'
    # line_listing_table_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_event_distribution_cumulative.xlsx'
    # event_seriousness(LL_cumulative_path, Cognizant_CRS_path, lookup_column, file_path_case_seriousness, line_listing_table_path, paragraph_path)

    # safety_topic = "drugOverdose"
    # lookup_column = """Overdose: cases reporting PTs ‘Accidental overdose’, ‘Accidental poisoning’, ‘Chemical poisoning’, ‘Exposure to toxic agent’, ‘Intentional overdose’, overdose’, Poisoning’, ‘Poisoning deliberate’, ‘Prescribed overdose’, ‘Systemic toxicity’ and ‘Toxicity to various agents’. """
    # paragraph_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_{type}_seriousness_reporting'
    # file_path_case_seriousness = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_{type}_seriousness_reporting.xlsx'
    # line_listing_table_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_{type}_distribution_reporting.xlsx'
    # case_seriousness(LL_reporting_path, Cognizant_CRS_path, lookup_column, file_path_case_seriousness, line_listing_table_path, paragraph_path)

    # safety_topic = "drugMisuse"
    # lookup_column = """Drug abuse/ Misuse: SMQ (narrow) ‘Drug abuse and dependence’ and PTs: ‘Dependence’, ‘Drug diversion’, ‘Intentional product use issue’, ‘Narcotic bowel syndrome’, ‘ Prescription from tampering’, ‘ Prescription drug used without prescription’, ‘Performance enhancing product use’, ‘Product use in unapproved therapeutic environment’, ‘Substance-induced mood disorder’, ‘Substance-induced psychotic disorder’. """
    # paragraph_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_{type}_seriousness_cumulative'
    # file_path_case_seriousness = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_{type}_seriousness_cumulative.xlsx'
    # line_listing_table_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_{type}_distribution_cumulative.xlsx'
    # case_seriousness(LL_cumulative_path, Cognizant_CRS_path, lookup_column, file_path_case_seriousness, line_listing_table_path, paragraph_path)

    # safety_topic = "offLabel"
    # lookup_column = """Approved, Related to Approve and Unknown indication"""
    # paragraph_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_case_seriousness_cumulative'
    # file_path_case_seriousness = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_case_seriousness_cumulative.xlsx'
    # line_listing_table_path = f'C:\\Users\\2113196\\Downloads\\{safety_topic}_case_distribution_cumulative.xlsx'
    # off_label_use_population(LL_cumulative_path, Cognizant_CRS_path, lookup_column, file_path_case_seriousness, line_listing_table_path, paragraph_path)

Cognizant_CRS_path = "C:\\Users\\2113196\\Downloads\\Cognizant_CRS.xlsx"
LL_reporting_path = "C:\\Users\\2113196\\Downloads\\Reporting interval_Sample  2 Line listing (1).xlsb"
LL_cumulative_path = "C:\\Users\\2113196\\Downloads\\Cumulative period_Sample 2 Line listing (1).xlsb"
safety_topic = "intestinal_perforation"
lookup_column = """SMQ (narrow) Gastrointestinal perforation, ulceration, haemorrhage of obstruction"""
run_functions(Cognizant_CRS_path, LL_reporting_path, LL_cumulative_path, lookup_column, safety_topic)
