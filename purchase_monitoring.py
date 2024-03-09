import pandas as pd
from fuzzywuzzy import fuzz
from prettytable import PrettyTable
# import subprocess
from os import startfile

import os


@staticmethod
def transaction():
    
    """This function is for selection of transactions"""
    
   
    

    TransactionList = [
               {"Code": '4000',"Transaction":'Checking Payroll'},
               {"Code": '4001',"Transaction":'Expense Account'},
           
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
    if ans == '4000':
        return PmConnection.pm_conn()
    elif ans == '4001':
        return PmConnection.select_expense_account()

    elif ans == 'x' or ans =='X':
        exit()


class PmConnection():
    """This class is for connection purposes"""
    def pm_conn():
        folder_path = 'purchase_monitoring'
        file_name = 'January 2024 PM-Jerome.xlsx'
        file_path = os.path.join(folder_path, file_name)
        sheet_name = 'PURCHASE-MONITORING'
        pm_df = pd.read_excel(file_path, sheet_name=sheet_name)

        pd.set_option('display.max_rows', None)

        # print(pm_df)

        return pm_df
    
    def ewt_connection():
        folder_path = 'purchase_monitoring'
        file_name = 'January 2024 PM-Jerome.xlsx'
        file_path = os.path.join(folder_path, file_name)
        sheet_name = 'GL-WITHOLDING-TAX'
        pm_df = pd.read_excel(file_path, sheet_name=sheet_name)

        pd.set_option('display.max_rows', None)

        # print(pm_df)

        return pm_df
    
    
    

      

    def select_expense_account() -> None:

        pm_df = PmConnection.pm_conn()
        ewt_df = PmConnection.ewt_connection()

        column_list = ['VOUCHER NO.', 'TOTAL SALES', '12% VAT', 'NET OF VAT', 'EXPENSE ACCOUNT','W/TAX']

        column_list_ewt = ['Transaction No', 'Credit']
        # print(pm_df[column_list])
        # print(ewt_df[column_list_ewt])

        # Assuming 'Transaction No' and 'Credit' are common columns between pm_df and ewt_df
        merged_df = pd.merge(pm_df, ewt_df, how='left', left_on='VOUCHER NO.', right_on='Transaction No')

        # Update 'W/TAX' column based on 'Credit' column
        merged_df['W/TAX'] = merged_df['Credit']

        # Fill NaN values in 'W/TAX' with 0
        merged_df['W/TAX'].fillna(0, inplace=True)

        # Drop the redundant 'Transaction No' and 'Credit' columns if needed
        merged_df.drop(['Transaction No', 'Credit'], axis=1, inplace=True)

        # Print the updated dataframe
        # print(merged_df[column_list])

        # Define the list of desired 'EXPENSE ACCOUNT' values
        desired_accounts = ['INPUT TAX SERVICES', 'DEFERRED INPUT TAX', 'INPUT TAX GOODS']

        # Add a new column '12% VAT' and set it to 0 only for rows in desired_accounts
        # pm_df['12% VAT'] = pm_df.apply(lambda row: round(row['NET OF VAT'] * 0.12, 2) if row['EXPENSE ACCOUNT'] not in desired_accounts else 0, axis=1)
        # pm_df['12% VAT'] = pm_df.apply(lambda row: round(row['NET OF VAT'] * 0.12, 2) if 'VAT Exempt' not in row['EXPENSE ACCOUNT'] and row['EXPENSE ACCOUNT'] not in desired_accounts else 0, axis=1)
        # pm_df['12% VAT'] = pm_df.apply(lambda row: round(row['NET OF VAT'] * 0.12, 2) if isinstance(row['EXPENSE ACCOUNT'], str) and 'VAT Exempt' not in row['EXPENSE ACCOUNT'] or 'DELIVERY' not in row['EXPENSE ACCOUNT'] and row['EXPENSE ACCOUNT'] not in desired_accounts else 0, axis=1)
        merged_df['12% VAT'] = merged_df.apply(lambda row: round(row['NET OF VAT'] * 0.12, 2) if pd.notna(row['EXPENSE ACCOUNT']) and isinstance(row['EXPENSE ACCOUNT'], str)
                                        and ('VAT Exempt' not in row['EXPENSE ACCOUNT']) and row['EXPENSE ACCOUNT'] 
                                        not in desired_accounts else 0, axis=1)

        # Add a new column 'TOTAL SALES' and set it to 0 only for rows in desired_accounts
        merged_df['TOTAL SALES'] = merged_df.apply(lambda row: round(row['NET OF VAT'] + row['12% VAT'], 2) if row['EXPENSE ACCOUNT'] not in desired_accounts else 0, axis=1)

        # Use boolean indexing to filter rows based on 'EXPENSE ACCOUNT' and select specified columns
        filtered_df = merged_df[merged_df['EXPENSE ACCOUNT'].isin(desired_accounts)][column_list]

        # Extract 'VOUCHER NO' from the filtered DataFrame
        desired_voucher_numbers = filtered_df['VOUCHER NO.'].tolist()

        # Use boolean indexing to select rows with the same 'VOUCHER NO' from the original DataFrame
        additional_rows = merged_df[merged_df['VOUCHER NO.'].isin(desired_voucher_numbers)]

        additional_rows.loc[additional_rows['VOUCHER NO.'].str.contains('JV', case=False, na=False), '12% VAT'] = 0


       
        # Print only the specified columns from the resulting DataFrame
        additional_rows[column_list] = additional_rows[column_list].round({'TOTAL SALES': 2, '12% VAT': 2})
        print(additional_rows[column_list])

         # Save to Excel file in the 'purchase_monitoring' folder
        excel_file_path = os.path.join('purchase_monitoring', 'pm_report.xlsx')
        additional_rows[~additional_rows['EXPENSE ACCOUNT'].isin(desired_accounts)][column_list].to_excel(excel_file_path, index=False)
        # additional_rows[column_list].to_excel(excel_file_path, index=False)

        ans2 = input('Do you want to Open Excel file?: ').lower()

        if ans2 == 'yes':
            # Open the generated Excel file using the full path
            os.startfile(excel_file_path)

        

        transaction()


transaction()