import pandas as pd
from fuzzywuzzy import fuzz
from prettytable import PrettyTable
import subprocess


@staticmethod
def transaction():
    
    """This function is for selection of transactions"""
    
   
    

    TransactionList = [
        
               {"Code": '2000',"Transaction":'All Computation'},
               {"Code": '2001',"Transaction":'GCE Computation'},
               {"Code": '2002',"Transaction":'EP Computation'},
               {"Code": '2003',"Transaction":'WP Computation'},
               {"Code": '2004',"Transaction":'QH2 Computation'},
           
           
        ]
    

    menu = PrettyTable()
    menu.field_names=['Code','Transactions']
        
    
    for x in TransactionList:      
        menu.add_row([
            x['Code'],
            x['Transaction'],
          
        ])
    print(menu)

    ans = input('Please enter code for your Desire transaction: ')

    if ans == '2000':
        return ExcellConnection.all_computation()

    elif ans == '2001':
        return ExcellConnection.GCE_computation()

    elif ans == '2002':
        return ExcellConnection.EP_computation()
    elif ans == '2003':
        return ExcellConnection.WP_computation()
    elif ans == '2004':
        return ExcellConnection.QH2_computation()

    elif ans == 'x' or ans =='X':
        exit()


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
    

    @staticmethod
    def all_computation():

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

       
        # grouped_df = merged_df.groupby(['DEPARTMENT'])[['Total_Gross','SSS_Employee_Remt']].sum().reset_index()
        grouped_df = merged_df.groupby(['DEPARTMENT','BOOKS'])[['Total_Gross', 'SSS_Employee_Remt',
                                                    'SSS_Employer_share']].sum().reset_index()
        
        # Assuming 'BOOKS' is a column in the grouped DataFrame
        # filtered_df = grouped_df[grouped_df['BOOKS'] == 'GENERAL COMMON EXPENSE']

        pd.set_option('display.max_rows', None)

        grouped_df.to_excel('payroll_gross.xlsx', index=False)

        # Open the generated Excel file using subprocess
        subprocess.run(['xdg-open', 'payroll_gross.xlsx'])

        # Open the generated Excel file
            # startfile("payroll.xlsx")
        
        transaction()

    @staticmethod
    def GCE_computation():

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

       
        # grouped_df = merged_df.groupby(['DEPARTMENT'])[['Total_Gross','SSS_Employee_Remt']].sum().reset_index()
        grouped_df = merged_df.groupby(['DEPARTMENT','BOOKS'])[['Total_Gross', 'SSS_Employee_Remt',
                                                    'SSS_Employer_share']].sum().reset_index()
        
        # Assuming 'BOOKS' is a column in the grouped DataFrame
        filtered_df = grouped_df[grouped_df['BOOKS'] == 'GENERAL COMMON EXPENSE']

        pd.set_option('display.max_rows', None)

        filtered_df.to_excel('payroll_gross.xlsx', index=False)

        # Open the generated Excel file using subprocess
        subprocess.run(['xdg-open', 'payroll_gross.xlsx'])

        # Open the generated Excel file
            # startfile("payroll.xlsx")
        
        transaction()

    @staticmethod
    def EP_computation():

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

       
        # grouped_df = merged_df.groupby(['DEPARTMENT'])[['Total_Gross','SSS_Employee_Remt']].sum().reset_index()
        grouped_df = merged_df.groupby(['DEPARTMENT','BOOKS'])[['Total_Gross', 'SSS_Employee_Remt',
                                                    'SSS_Employer_share']].sum().reset_index()
        
        # Assuming 'BOOKS' is a column in the grouped DataFrame
        filtered_df = grouped_df[grouped_df['BOOKS'] == 'ELISTON PLACE']

        pd.set_option('display.max_rows', None)

        filtered_df.to_excel('payroll_gross.xlsx', index=False)

        # Open the generated Excel file using subprocess
        subprocess.run(['xdg-open', 'payroll_gross.xlsx'])

        # Open the generated Excel file
            # startfile("payroll.xlsx")
        
        transaction()


    @staticmethod
    def WP_computation():

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

       
        # grouped_df = merged_df.groupby(['DEPARTMENT'])[['Total_Gross','SSS_Employee_Remt']].sum().reset_index()
        grouped_df = merged_df.groupby(['DEPARTMENT','BOOKS'])[['Total_Gross', 'SSS_Employee_Remt',
                                                    'SSS_Employer_share']].sum().reset_index()
        
        # Assuming 'BOOKS' is a column in the grouped DataFrame
        filtered_df = grouped_df[grouped_df['BOOKS'] == 'WELLINGTON PLACE 6-12']

        pd.set_option('display.max_rows', None)

        filtered_df.to_excel('payroll_gross.xlsx', index=False)

        # Open the generated Excel file using subprocess
        subprocess.run(['xdg-open', 'payroll_gross.xlsx']) 

        # Open the generated Excel file
            # startfile("payroll.xlsx")
        
        transaction()


    @staticmethod
    def QH2_computation():

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

       
        # grouped_df = merged_df.groupby(['DEPARTMENT'])[['Total_Gross','SSS_Employee_Remt']].sum().reset_index()
        grouped_df = merged_df.groupby(['DEPARTMENT','BOOKS'])[['Total_Gross', 'SSS_Employee_Remt',
                                                    'SSS_Employer_share']].sum().reset_index()
        
        # Assuming 'BOOKS' is a column in the grouped DataFrame
        filtered_df = grouped_df[grouped_df['BOOKS'] == 'QH2']

        pd.set_option('display.max_rows', None)

        filtered_df.to_excel('payroll_gross.xlsx', index=False)

        # Open the generated Excel file using subprocess
        subprocess.run(['xdg-open', 'payroll_gross.xlsx']) 

        # Open the generated Excel file
            # startfile("payroll.xlsx")
        
        transaction()



    


transaction()





