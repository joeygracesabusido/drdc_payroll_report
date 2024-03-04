import pandas as pd
from fuzzywuzzy import fuzz
from prettytable import PrettyTable
import subprocess
from os import startfile


@staticmethod
def transaction():
    
    """This function is for selection of transactions"""
    
   
    

    TransactionList = [
               {"Code": '1000',"Transaction":'Checking Payroll'},
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
    if ans == '1000':
        return ExcellConnection.checking_all_payroll()

    elif ans == '2000':
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
        
        df_sheet2 = pd.read_excel(file_path, sheet_name='Sheet3')

        # print(df_sheet2)

        return df_sheet2

    def payroll_list():
        file_path = 'DRDC-FEBRUARY-2023.xlsx'
        
        df_sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')



        # print(df_sheet2)

        return df_sheet2
    
    @staticmethod
    def checking_all_payroll():
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

        

        selected_columns = ['Name', 'BOOKS', 'DEPARTMENT','Total_Gross']

        # print(merged_df[selected_columns])

        

        merged_df[selected_columns].to_excel('payroll_gross_all.xlsx', index=False)

        open_excel_file = input('Do you want to open the Excel File: ').lower()

        if open_excel_file == 'yes':
            # subprocess.run(['xdg-open', 'payroll_gross_all.xlsx'])
            # Open the generated Excel file
            startfile("payroll_gross_all.xlsx")
        
        transaction()

    

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
        merged_df['TOTAL SSS'] = merged_df['SSS_Employee_Remt'] + merged_df['SSS_Employer_share'] + merged_df['EC']
        merged_df['TOTAL PHIC'] = merged_df['PHIC_Employee'] + merged_df['PHIC_Rmployer_Share']
        merged_df['TOTAL PHIC'] = merged_df['PHIC_Employee'] + merged_df['PHIC_Rmployer_Share']
        merged_df['TOTAL HDMF'] = merged_df['HDMF_CONTRIBUTION_employee'] + merged_df['HDMF_CONTRIBUTION_employer']


        grouped_df = merged_df.groupby(['DEPARTMENT','BOOKS'])[['Total_Gross', 'SSS_Employee_Remt',
                                            'SSS_Employer_share','EC','TOTAL SSS',
                                            'SSS_Loan','SSS_Calamity Loan','PHIC_Employee',
                                            'PHIC_Rmployer_Share','TOTAL PHIC',
                                            'HDMF_CONTRIBUTION_employee','HDMF_CONTRIBUTION_employer',
                                            'TOTAL HDMF','HDMF_LOAN','HDMF_CALAMITY',
                                            'W_TAX_2024','CASH_ADVANCE','Personal_Loan_(MA)',
                                            '13th_Month_Pay_over_Payment','Ad 13 Month Pay',
                                            'Return_loan_sss_loan','Regular_Allowance',
                                            'Holiday_RDOT_Pay','meal','Developmental','Add_Others Adjustment','Net_Pay']].sum().reset_index()
        
        # Assuming 'BOOKS' is a column in the grouped DataFrame
        # filtered_df = grouped_df[grouped_df['BOOKS'] == 'GENERAL COMMON EXPENSE']

        pd.set_option('display.max_rows', None)

        grouped_df.to_excel('payroll_gross.xlsx', index=False)

        # Open the generated Excel file using subprocess
        # subprocess.run(['xdg-open', 'payroll_gross.xlsx'])


        open_excel_file = input('Do you want to open the Excel File: ').lower()

        if open_excel_file == 'yes':
            # subprocess.run(['xdg-open', 'payroll_gross.xlsx'])

            # Open the generated Excel file
            startfile("payroll_gross.xlsx")
        
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
        merged_df['TOTAL SSS'] = merged_df['SSS_Employee_Remt'] + merged_df['SSS_Employer_share'] + merged_df['EC']
        merged_df['TOTAL PHIC'] = merged_df['PHIC_Employee'] + merged_df['PHIC_Rmployer_Share']
        merged_df['TOTAL PHIC'] = merged_df['PHIC_Employee'] + merged_df['PHIC_Rmployer_Share']
        merged_df['TOTAL HDMF'] = merged_df['HDMF_CONTRIBUTION_employee'] + merged_df['HDMF_CONTRIBUTION_employer']
        grouped_df = merged_df.groupby(['DEPARTMENT','BOOKS'])[['Total_Gross', 'SSS_Employee_Remt',
                                            'SSS_Employer_share','EC','TOTAL SSS',
                                            'SSS_Loan','SSS_Calamity Loan','PHIC_Employee',
                                            'PHIC_Rmployer_Share','TOTAL PHIC',
                                            'HDMF_CONTRIBUTION_employee','HDMF_CONTRIBUTION_employer',
                                            'TOTAL HDMF','HDMF_LOAN','HDMF_CALAMITY',
                                            'W_TAX_2024','CASH_ADVANCE','Personal_Loan_(MA)',
                                            '13th_Month_Pay_over_Payment','Ad 13 Month Pay',
                                            'Return_loan_sss_loan','Regular_Allowance',
                                            'Holiday_RDOT_Pay','meal','Developmental','Add_Others Adjustment','Net_Pay']].sum().reset_index()
        
        # Create a new column 'TOTAL SHARES' by summing 'SSS_Employer_share' and 'EC'
        grouped_df['TOTAL SHARES'] = grouped_df['SSS_Employer_share'] + grouped_df['EC']
        
      
        
        # Assuming 'BOOKS' is a column in the grouped DataFrame
        filtered_df = grouped_df[grouped_df['BOOKS'] == 'GENERAL COMMON EXPENSE']

        

        pd.set_option('display.max_rows', None)


        departments = ['ACCOUNTING DEPARTMENT', 'ADMIN DEPARTMENT', 'EMD DEPARTMENT', 'ENGINEERING DEPARTMENT - ANTIPOLO',
                    'ENGINEERING DEPARTMENT - CAVITE', 'FINANCE DEPARTMENT', 'HR DEPARTMENT', 'LEGAL DEPARTMENT',
                    'OFFICE OF THE PRESIDENT', 'PERMITS & LICENSES DEPARTMENT', 'PLANNING & DESIGN DEPARTMENT',
                    'SALES & LOAN DOCUMENTATION', 'TREASURY DEPARTMENT']

        salary_dfs = []
        sss_dfs = []
        phic_dfs = []
        hdmf_dfs = []
        

        for department in departments:
            # Calculate total gross
            total_gross = filtered_df.loc[grouped_df['DEPARTMENT'] == department, 'Total_Gross'].sum()
            
            # Calculate total shares (assuming 'TOTAL SHARES' is the sum of 'SSS_Employer_share' and 'EC')
            total_gross_ss = filtered_df.loc[grouped_df['DEPARTMENT'] == department, 'TOTAL SHARES'].sum()
            
            # Calculate total PHIC (assuming 'PHIC_Rmployer_Share' is the column for PHIC)
            total_gross_phic = filtered_df.loc[grouped_df['DEPARTMENT'] == department, 'PHIC_Rmployer_Share'].sum()
            
            # Calculate total PHIC (assuming 'PHIC_Rmployer_Share' is the column for PHIC)
            total_gross_hdmf = filtered_df.loc[grouped_df['DEPARTMENT'] == department, 'HDMF_CONTRIBUTION_employee'].sum()

            

            # Create DataFrames for SALARIES & WAGES
            salary_df = pd.DataFrame({'DEPARTMENT': [f'SALARIES & WAGES - {department}'], 'BOOKS': [total_gross]})
            salary_dfs.append(salary_df)

            # Create DataFrames for SSS, MEDICARE & ECC CONTRIBUTIONS
            ss_df = pd.DataFrame({'DEPARTMENT': [f'SSS, MEDICARE & ECC CONTRIBUTIONS - {department}'], 'BOOKS': [total_gross_ss]})
            sss_dfs.append(ss_df)

            # Create DataFrames for PHILHEALTH CONTRIBUTIONS
            phic_df = pd.DataFrame({'DEPARTMENT': [f'PHILHEALTH CONTRIBUTIONS - {department}'], 'BOOKS': [total_gross_phic]})
            phic_dfs.append(phic_df)

            # Create DataFrames for HDMF CONTRIBUTIONS
            hdmf_df = pd.DataFrame({'DEPARTMENT': [f'PAG-IBIG CONTRIBUTIONS - {department}'], 'BOOKS': [total_gross_hdmf]})
            hdmf_dfs.append(hdmf_df)

           
        # Concatenate all SALARIES & WAGES DataFrames into a single DataFrame
        salary_df = pd.concat(salary_dfs, ignore_index=True)

        # Concatenate all SSS, MEDICARE & ECC CONTRIBUTIONS DataFrames into a single DataFrame
        sss_df = pd.concat(sss_dfs, ignore_index=True)

        # Concatenate all PHILHEALTH CONTRIBUTIONS DataFrames into a single DataFrame
        phic_df = pd.concat(phic_dfs, ignore_index=True)

        hdmf_df = pd.concat(hdmf_dfs, ignore_index=True)

        

        
        

        # Concatenate the new rows to the existing DataFrame
        filtered_df = pd.concat([filtered_df,salary_df, sss_df, phic_df,
                                 hdmf_df,], ignore_index=True)
        
        # Calculate the total sum of Credit
        total_13th_month_add = filtered_df['Ad 13 Month Pay'].sum()
        total_13th_month_less = filtered_df['13th_Month_Pay_over_Payment'].sum()
        total_withholding_taxes_payable = filtered_df['W_TAX_2024'].sum()
        sss_total_payable = filtered_df['TOTAL SSS'].sum()
        sss_total_calamity = filtered_df['SSS_Calamity Loan'].sum()
        phic_contri_payable = filtered_df['TOTAL PHIC'].sum()
        hdmf_contri_payable =  filtered_df['TOTAL HDMF'].sum()   
        hdmf_loan_payable =  filtered_df['HDMF_LOAN'].sum()     
        advances_to_personel =  filtered_df['CASH_ADVANCE']
        
        print(advances_to_personel)

        # Create a new DataFrame Credit
        total_13th_month_add_df = pd.DataFrame({'DEPARTMENT': ['13th MONTH - ADD'], 'BOOKS': [total_13th_month_add]})
        withholding_tax_df = pd.DataFrame({'DEPARTMENT': ['WITHOLDING TAXES PAYABLE- COMPENSATION'], 'Total_Gross': [total_withholding_taxes_payable]})
        total_13th_month_less_df = pd.DataFrame({'DEPARTMENT': ['13th MONTH - LESS'], 'Total_Gross': [total_13th_month_less]})
        total_sss_remittance_df = pd.DataFrame({'DEPARTMENT': ['SSS/MEDICARE/ECC PAYABLE'], 'Total_Gross': [sss_total_payable]})
        total_sss_total_calamitiy_df = pd.DataFrame({'DEPARTMENT': ['SSS CALAMITY LOAN PAYABLE'], 'Total_Gross': [sss_total_calamity]})
        total_phic_contri_payable_df = pd.DataFrame({'DEPARTMENT': ['PHILHEALTH CONTRIBUTIONS PAYABLE'], 'Total_Gross': [phic_contri_payable]})
        total_hdmf_contri_payable_df = pd.DataFrame({'DEPARTMENT': ['PAG-IBIG CONTRIBUTIONS PAYABLE'], 'Total_Gross': [hdmf_contri_payable]})
        total_hdmf_loan_df = pd.DataFrame({'DEPARTMENT': ['PAG-IBIG SALARY LOAN PAYABLE'], 'Total_Gross': [hdmf_loan_payable]})

        total_hdmf_loan_df = pd.DataFrame({'DEPARTMENT': ['ADVANCES'], 'Total_Gross': [advances_to_personel]})
        
        
        # Concatenate the new row to the existing DataFrame
        filtered_df = pd.concat([filtered_df, total_13th_month_add_df,
                                 total_13th_month_less_df, withholding_tax_df,
                                 total_sss_remittance_df,total_sss_total_calamitiy_df,
                                 total_phic_contri_payable_df,total_hdmf_contri_payable_df,
                                 total_hdmf_loan_df,total_hdmf_loan_df], ignore_index=True)

        # Save to Excel file
        filtered_df.to_excel('payroll_gross.xlsx', index=False)

        ans2 = input('Do you want to Open Excel file?: ').lower()

        if ans2 == 'yes':
            # Open the generated Excel file
            startfile("payroll_gross.xlsx")
            
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
        merged_df['TOTAL SSS'] = merged_df['SSS_Employee_Remt'] + merged_df['SSS_Employer_share'] + merged_df['EC']
        merged_df['TOTAL PHIC'] = merged_df['PHIC_Employee'] + merged_df['PHIC_Rmployer_Share']
        merged_df['TOTAL PHIC'] = merged_df['PHIC_Employee'] + merged_df['PHIC_Rmployer_Share']
        merged_df['TOTAL HDMF'] = merged_df['HDMF_CONTRIBUTION_employee'] + merged_df['HDMF_CONTRIBUTION_employer']
        grouped_df = merged_df.groupby(['DEPARTMENT','BOOKS'])[['Total_Gross', 'SSS_Employee_Remt',
                                            'SSS_Employer_share','EC','TOTAL SSS',
                                            'SSS_Loan','SSS_Calamity Loan','PHIC_Employee',
                                            'PHIC_Rmployer_Share','TOTAL PHIC',
                                            'HDMF_CONTRIBUTION_employee','HDMF_CONTRIBUTION_employer',
                                            'TOTAL HDMF','HDMF_LOAN','HDMF_CALAMITY',
                                            'W_TAX_2024','CASH_ADVANCE','Personal_Loan_(MA)',
                                            '13th_Month_Pay_over_Payment','Ad 13 Month Pay',
                                            'Return_loan_sss_loan','Regular_Allowance',
                                            'Holiday_RDOT_Pay','meal','Developmental','Add_Others Adjustment','Net_Pay']].sum().reset_index()
        
        # Assuming 'BOOKS' is a column in the grouped DataFrame
        filtered_df = grouped_df[grouped_df['BOOKS'] == 'ELISTON PLACE']

        pd.set_option('display.max_rows', None)

        filtered_df.to_excel('payroll_gross.xlsx', index=False)

        # Open the generated Excel file using subprocess
        # subprocess.run(['xdg-open', 'payroll_gross.xlsx'])

        # Open the generated Excel file
        startfile("payroll_gross.xlsx")
        
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
        merged_df['TOTAL SSS'] = merged_df['SSS_Employee_Remt'] + merged_df['SSS_Employer_share'] + merged_df['EC']
        merged_df['TOTAL PHIC'] = merged_df['PHIC_Employee'] + merged_df['PHIC_Rmployer_Share']
        merged_df['TOTAL PHIC'] = merged_df['PHIC_Employee'] + merged_df['PHIC_Rmployer_Share']
        merged_df['TOTAL HDMF'] = merged_df['HDMF_CONTRIBUTION_employee'] + merged_df['HDMF_CONTRIBUTION_employer']
        grouped_df = merged_df.groupby(['DEPARTMENT','BOOKS'])[['Total_Gross', 'SSS_Employee_Remt',
                                            'SSS_Employer_share','EC','TOTAL SSS',
                                            'SSS_Loan','SSS_Calamity Loan','PHIC_Employee',
                                            'PHIC_Rmployer_Share','TOTAL PHIC',
                                            'HDMF_CONTRIBUTION_employee','HDMF_CONTRIBUTION_employer',
                                            'TOTAL HDMF','HDMF_LOAN','HDMF_CALAMITY',
                                            'W_TAX_2024','CASH_ADVANCE','Personal_Loan_(MA)',
                                            '13th_Month_Pay_over_Payment','Ad 13 Month Pay',
                                            'Return_loan_sss_loan','Regular_Allowance',
                                            'Holiday_RDOT_Pay','meal','Developmental','Add_Others Adjustment','Net_Pay']].sum().reset_index()
        
        # Assuming 'BOOKS' is a column in the grouped DataFrame
        filtered_df = grouped_df[grouped_df['BOOKS'] == 'WELLINGTON PLACE 6-12']

        pd.set_option('display.max_rows', None)

        filtered_df.to_excel('payroll_gross.xlsx', index=False)

        # Open the generated Excel file using subprocess
        # subprocess.run(['xdg-open', 'payroll_gross.xlsx']) 

        # Open the generated Excel file
        startfile("payroll_gross.xlsx")
        
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
        merged_df['TOTAL SSS'] = merged_df['SSS_Employee_Remt'] + merged_df['SSS_Employer_share'] + merged_df['EC']
        merged_df['TOTAL PHIC'] = merged_df['PHIC_Employee'] + merged_df['PHIC_Rmployer_Share']
        merged_df['TOTAL PHIC'] = merged_df['PHIC_Employee'] + merged_df['PHIC_Rmployer_Share']
        merged_df['TOTAL HDMF'] = merged_df['HDMF_CONTRIBUTION_employee'] + merged_df['HDMF_CONTRIBUTION_employer']
        grouped_df = merged_df.groupby(['DEPARTMENT','BOOKS'])[['Total_Gross', 'SSS_Employee_Remt',
                                            'SSS_Employer_share','EC','TOTAL SSS',
                                            'SSS_Loan','SSS_Calamity Loan','PHIC_Employee',
                                            'PHIC_Rmployer_Share','TOTAL PHIC',
                                            'HDMF_CONTRIBUTION_employee','HDMF_CONTRIBUTION_employer',
                                            'TOTAL HDMF','HDMF_LOAN','HDMF_CALAMITY',
                                            'W_TAX_2024','CASH_ADVANCE','Personal_Loan_(MA)',
                                            '13th_Month_Pay_over_Payment','Ad 13 Month Pay',
                                            'Return_loan_sss_loan','Regular_Allowance',
                                            'Holiday_RDOT_Pay','meal','Developmental','Add_Others Adjustment','Net_Pay']].sum().reset_index()
        
        # Assuming 'BOOKS' is a column in the grouped DataFrame
        filtered_df = grouped_df[grouped_df['BOOKS'] == 'QH2']

        pd.set_option('display.max_rows', None)

        filtered_df.to_excel('payroll_gross.xlsx', index=False)

        # Open the generated Excel file using subprocess
        # subprocess.run(['xdg-open', 'payroll_gross.xlsx']) 

        # Open the generated Excel file for windows
        startfile("payroll_gross.xlsx")
        
        transaction()



    


transaction()





