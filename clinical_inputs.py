import pandas as pd
import numpy as np
def clinical_input_table(file_path, clinical_input_table_path):
    df = pd.read_excel(file_path, skiprows=4)
    try:
        grouped_df = df.groupby('Formulation')['Patients treated'].sum().unstack().fillna(0)
    except ValueError:
        grouped_df = df.groupby(['Formulation', 'Status'])['Patients treated'].sum().unstack().fillna(0)

    df2 = grouped_df.reset_index()

    # Identify numeric columns
    numeric_cols = df2.select_dtypes(include=[np.number]).columns

    df2.loc['Total'] = df2[numeric_cols].sum()
    df2.fillna('Total', inplace=True)

    df2.to_excel(clinical_input_table_path, index= False)

file_path = 'C:\\Users\\2113196\\Downloads\\Clinical inputs_Section 5.1.xlsx'
clinical_input_table_path = 'C:\\Users\\2113196\\Downloads\\clinical_input_table.xlsx'
clinical_input_table(file_path, clinical_input_table_path)