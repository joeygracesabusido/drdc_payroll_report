import pandas as pd
from fuzzywuzzy import fuzz

class ExcellConnection():
    """This class is for connection purposes"""
    def employee_list():
        file_path = 'DRDC EMPLOYEE LIST.xlsx'
        
        df_sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')

        # print(df_sheet2)

        return df_sheet2

    def payroll_list():
        file_path = 'DRDC-FEBRUARY-2023.xlsx'
        
        df_sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')



        # print(df_sheet2)

        return df_sheet2


    def test_computation():

        employee_list = ExcellConnection.employee_list()
        payroll_comp = ExcellConnection.payroll_list()

        def find_similarity(row):
            if pd.notna(row['DEPARTMENT']):
                return row['DEPARTMENT']
            else:
                best_match = max(employee_list['NAME'], key=lambda name: fuzz.partial_ratio(str(row['Name']), str(name)))

                if fuzz.partial_ratio(str(row['Name']), str(best_match)) > 80:
                    if 'DEPARTMENT' in payroll_comp.columns:
                        matches = payroll_comp.loc[payroll_comp['Name'] == best_match, 'DEPARTMENT']
                        if not matches.empty:
                            return matches.values[0]
                    return 'Not Found'
        
        merged_df = pd.merge(employee_list, payroll_comp, how='left', left_on='NAME', right_on='Name')
        merged_df['DEPARTMENT'] = merged_df.apply(find_similarity, axis=1)

        print(merged_df.columns)

        selected_columns = ['Name', 'BOOKS', 'DEPARTMENT','Total_Gross']

        # Group by 'DEPARTMENT', 'Full Name', and 'BOOKS', and sum the 'Amount'
        # grouped_df = merged_df.groupby(['DEPARTMENT'])[['Total_Gross','SSS_Employee_Remt']].sum().reset_index()
        grouped_df = merged_df.groupby(['DEPARTMENT'])[['Total_Gross', 'SSS_Employee_Remt',
                                                    'SSS_Employer_share']].sum().reset_index()

        pd.set_option('display.max_rows', None)

        grouped_df.to_excel('payroll_gross.xlsx', index=False)
        print(merged_df[selected_columns])

ExcellConnection.test_computation()





