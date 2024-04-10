import pandas as pd
import numpy as np
def exposure_calculation_sales_excel_creation(sales_file_path, sheet_name, sales_file_path_filled):
    df1 = pd.read_excel(sales_file_path, sheet_name=sheet_name, skiprows=4)
    df2 = pd.read_excel(sales_file_path, sheet_name='Region')

    df1.insert(df1.columns.get_loc('Product name') + 1, 'Formulation', '')
    df1['Product name'] = df1['Product name'].astype(str)
    for index, row in df1.iterrows():
        product_name = row['Product name']
        if 'FCT' in product_name:
            df1.at[index, 'Formulation'] = 'Film coated tablet'
        elif 'POS' in product_name:
            df1.at[index, 'Formulation'] = 'Powder for oral suspension'

    df1.insert(df1.columns.get_loc('Country') + 1, 'Region', '')
    df1['Region'] = df1['Country'].map(df2.set_index('Country')['Region'])

    df1.insert(df1.columns.get_loc('Strength') + 1, 'Strength value', '')
    for index, row in df1.iterrows():
        strength = str(row['Strength'])
        if 'MG' in strength:
            df1.at[index, 'Strength value'] = float(strength.strip('MG'))
        elif 'G' in strength:
            strength_value = float(strength.strip('G')) * 1000
            df1.at[index, 'Strength value'] = strength_value

    df1['Total mg'] = ''
    df1['Total mg'] = df1['Strength value'] * df1['Pack size'] * df1['Quantity sold']
    total_mg_sum = df1['Total mg'].sum()
    df1['Total mg'] = pd.to_numeric(df1['Total mg'], errors='coerce')

    df1.to_excel(sales_file_path_filled, index=False)

def exposure_calucaltion_by_condition(sales_file_path, exposure_table_filled, condition):
    df1 = pd.read_excel(sales_file_path)
    df3 = df1.groupby(condition)['Total mg'].sum()
    df3 = df3.to_frame().reset_index()

    user_choice = '2'

    DDD = 6400
    if user_choice == '1':
        df3['PTD'] = None
        df3['PTD'] = (df3['Total mg'] / DDD).round(2)
    if user_choice == '2':
        df3['PTD'] = None
        df3['PTD'] = (df3['Total mg'] / (DDD * 365)).round(2)
    if user_choice == '3':
        courses = input("Enter input for PTD formula: ")
        df3['PTD'] = None
        df3['PTD'] = (df3['Total mg'] / courses).round(2)

    if sales_file_path == "C:\\Users\\2113196\\Downloads\\loratadine_reporting_exposure_sales.xlsx":
        df3 = df3.rename(columns={df3.columns[1]: 'Current reporting interval Amount Sold'})
        df3 = df3.rename(columns={df3.columns[2]: 'Current Estimated exposure'})
    
    if sales_file_path == "C:\\Users\\2113196\\Downloads\\loratadine_cumulative_exposure_sales.xlsx":
        df3 = df3.rename(columns={df3.columns[1]: 'Cumulative reporting interval Amount Sold'})
        df3 = df3.rename(columns={df3.columns[2]: 'Cumulative Estimated exposure'})
    df3.to_excel(exposure_table_filled)


def final_exposure_table(file1, file2, final_file_path, condition):
    df1 = pd.read_excel(file1)   
    df2 = pd.read_excel(file2)

    df2 = df2.drop(df2.columns[0], axis=1)
    df2 = df2.drop(condition, axis=1)

    df = pd.concat([df1, df2], axis=1)
    df = df.drop(df.columns[0], axis=1)

    # Identify numeric columns
    numeric_cols = df.select_dtypes(include=[np.number]).columns

    df.loc['Total'] = df[numeric_cols].sum()
    df.fillna('Total', inplace=True)

    df.to_excel(final_file_path, index=False)

# sales_file_path = 'C:\\Users\\2113196\\Downloads\\SEVELAMER reporting and cumulative period sales_current.xlsx'
# sheet_name = 'Reporting interval'
# sales_file_path_filled = "C:\\Users\\2113196\\Downloads\\loratadine_reporting_exposure_sales.xlsx"
# exposure_calculation_sales_excel_creation(sales_file_path, sheet_name, sales_file_path_filled)

# sales_file_path = 'C:\\Users\\2113196\\Downloads\\SEVELAMER reporting and cumulative period sales_current.xlsx'
# sheet_name = 'Cumulative period'
# sales_file_path_filled = "C:\\Users\\2113196\\Downloads\\loratadine_cumulative_exposure_sales.xlsx"
# exposure_calculation_sales_excel_creation(sales_file_path, sheet_name, sales_file_path_filled)

# sales_file_path = "C:\\Users\\2113196\\Downloads\\loratadine_reporting_exposure_sales.xlsx"
# exposure_table_filled = "C:\\Users\\2113196\\Downloads\\loratadine_reporting_exposure_table.xlsx"
# condition = 'Formulation'
# exposure_calucaltion_by_condition(sales_file_path, exposure_table_filled, condition)

# sales_file_path = "C:\\Users\\2113196\\Downloads\\loratadine_cumulative_exposure_sales.xlsx"
# exposure_table_filled = "C:\\Users\\2113196\\Downloads\\loratadine_cumulative_exposure_table.xlsx"
# exposure_calucaltion_by_condition(sales_file_path, exposure_table_filled, condition)

# file1 = "C:\\Users\\2113196\\Downloads\\loratadine_reporting_exposure_table.xlsx"
# file2 = "C:\\Users\\2113196\\Downloads\\loratadine_cumulative_exposure_table.xlsx"
# final_file_path = "C:\\Users\\2113196\\Downloads\\loratadine_exposure_table_formulation.xlsx"
# final_exposure_table(file1, file2, final_file_path, condition)

sales_file_path = "C:\\Users\\2113196\\Downloads\\loratadine_reporting_exposure_sales.xlsx"
exposure_table_filled = "C:\\Users\\2113196\\Downloads\\loratadine_reporting_exposure_table.xlsx"
condition = 'Region'
exposure_calucaltion_by_condition(sales_file_path, exposure_table_filled, condition)

sales_file_path = "C:\\Users\\2113196\\Downloads\\loratadine_cumulative_exposure_sales.xlsx"
exposure_table_filled = "C:\\Users\\2113196\\Downloads\\loratadine_cumulative_exposure_table.xlsx"
exposure_calucaltion_by_condition(sales_file_path, exposure_table_filled, condition)

file1 = "C:\\Users\\2113196\\Downloads\\loratadine_reporting_exposure_table.xlsx"
file2 = "C:\\Users\\2113196\\Downloads\\loratadine_cumulative_exposure_table.xlsx"
final_file_path = "C:\\Users\\2113196\\Downloads\\loratadine_exposure_table_region.xlsx"
final_exposure_table(file1, file2, final_file_path, condition)

sales_file_path = "C:\\Users\\2113196\\Downloads\\loratadine_reporting_exposure_sales.xlsx"
exposure_table_filled = "C:\\Users\\2113196\\Downloads\\loratadine_reporting_exposure_table.xlsx"
condition = 'Country'
exposure_calucaltion_by_condition(sales_file_path, exposure_table_filled, condition)

sales_file_path = "C:\\Users\\2113196\\Downloads\\loratadine_cumulative_exposure_sales.xlsx"
exposure_table_filled = "C:\\Users\\2113196\\Downloads\\loratadine_cumulative_exposure_table.xlsx"
exposure_calucaltion_by_condition(sales_file_path, exposure_table_filled, condition)

file1 = "C:\\Users\\2113196\\Downloads\\loratadine_reporting_exposure_table.xlsx"
file2 = "C:\\Users\\2113196\\Downloads\\loratadine_cumulative_exposure_table.xlsx"
final_file_path = "C:\\Users\\2113196\\Downloads\\loratadine_exposure_table_country.xlsx"
final_exposure_table(file1, file2, final_file_path, condition)