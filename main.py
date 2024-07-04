import pandas as pd
from fuzzywuzzy import fuzz
from prettytable import PrettyTable

from os import startfile

@staticmethod
def transaction():
    
    """This function is for selection of transactions"""
    
   
    

    TransactionList = [
               {"Code": '1000',"Transaction":'Get Total'},
               {"Code": '1002',"Transaction":'Get Total Per Name'},
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
        return get_total_per_book_per_department()
    elif ans == '1002':
        return get_total_per_book_per_department_per_name()

    elif ans == 'x' or ans =='X':
        exit()



def  create_per_department():
    # Load data from Excel sheets
    file_path = 'DRDC EMPLOYEE LIST.xlsx'
    df_sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')
    df_sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')

    # Function to find similarity using fuzzywuzzy
    def find_similarity(row):
        if pd.notna(row['DEPARTMENT']):
            return row['DEPARTMENT']
        else:
            # Iterate through 'NAME' in Sheet2 and find the best match
            best_match = max(df_sheet2['NAME'], key=lambda name: fuzz.partial_ratio(str(row['LAST NAME']), str(name)))
            
            # If similarity is above a certain threshold (adjust as needed), consider it a match
            if fuzz.partial_ratio(str(row['LAST NAME']), str(best_match)) > 80:
                return df_sheet2.loc[df_sheet2['NAME'] == best_match, 'DEPARTMENT'].values[0]
            else:
                return 'Not Found'

    # Apply the similarity function to each row
    merged_df = pd.merge(df_sheet1, df_sheet2, how='left', left_on='LAST NAME', right_on='NAME')
    merged_df['DEPARTMENT'] = merged_df.apply(find_similarity, axis=1)

    # Save the result back to Sheet1 or a new Excel file
    merged_df.to_excel('initial_membership.xlsx', index=False)

    print(merged_df)

def  create_per_department2():
    # Load data from Excel sheets
    file_path = 'DRDC EMPLOYEE LIST.xlsx'
    df_sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')
    df_sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')

    # Function to find similarity using fuzzywuzzy
    def find_similarity(row):
        if pd.notna(row['DEPARTMENT']):
            return row['DEPARTMENT']
        else:
            # Iterate through 'NAME' in Sheet2 and find the best match
            best_match = max(df_sheet2['NAME'], key=lambda name: fuzz.partial_ratio(str(row['LAST NAME']), str(name)))
            
            # If similarity is above a certain threshold (adjust as needed), consider it a match
            if fuzz.partial_ratio(str(row['LAST NAME']), str(best_match)) > 80:
                # return df_sheet2.loc[df_sheet2['NAME'] == best_match, 'DEPARTMENT'].values[0]
                matches = df_sheet2.loc[df_sheet2['NAME'] == best_match, 'DEPARTMENT']
                if not matches.empty:
                    return matches.values[0]
            return 'Not Found'
        
    def find_similarity2(row):
        if pd.notna(row['BOOKS']):
            return row['BOOKS']
        else:
            # Iterate through 'NAME' in Sheet2 and find the best match
            best_match = max(df_sheet2['NAME'], key=lambda name: fuzz.partial_ratio(str(row['LAST NAME']), str(name)))
            
            # If similarity is above a certain threshold (adjust as needed), consider it a match
            if fuzz.partial_ratio(str(row['LAST NAME']), str(best_match)) > 80:
                # return df_sheet2.loc[df_sheet2['NAME'] == best_match, 'DEPARTMENT'].values[0]
                matches = df_sheet2.loc[df_sheet2['NAME'] == best_match, 'BOOKS']
                if not matches.empty:
                    return matches.values[0]
            return 'Not Found'

    # Apply the similarity function to each row
    merged_df = pd.merge(df_sheet1, df_sheet2, how='left', left_on='LAST NAME', right_on='NAME')
    # merged_df['BOOKS'] = merged_df.apply(find_similarity, axis=1)
    # print(merged_df)
    merged_df['DEPARTMENT'] = merged_df.apply(find_similarity, axis=1)
    merged_df['BOOKS'] = merged_df.apply(find_similarity2, axis=1)
    
    
    print(merged_df)
    grouped_df = merged_df.groupby(['LAST NAME','FIRST NAME','BOOKS','DEPARTMENT'])['AMOUNT'].sum().reset_index()
    grouped_df['INPUT TAX'] = grouped_df['AMOUNT'] * .12

    group_by_department = merged_df.groupby(['DEPARTMENT'])['AMOUNT'].sum().reset_index()

    # Save the result back to Sheet1 or a new Excel file
    grouped_df.to_excel('initial_membership2.xlsx', index=False)
    group_by_department.to_excel('group_by_department_initial_membership.xlsx', index=False)

    print(grouped_df)

def get_total_per_department():
    # Load data from Excel sheets
    file_path = 'LIST-AVEGA-CLAIMS1.xlsx'
    df_sheet1 = pd.read_excel(file_path, sheet_name='claims')
    df_sheet3 = pd.read_excel(file_path, sheet_name='EMPLOYEE-LIST')

    # Function to find similarity using fuzzywuzzy
    def find_similarity(row):
        if pd.notna(row['DEPARTMENT']):
            return row['DEPARTMENT']
        else:
            # Iterate through 'NAME' in Sheet2 and find the best match
            best_match = max(df_sheet3['NAME'], key=lambda name: fuzz.partial_ratio(str(row['Full Name']), str(name)))
            
            # If similarity is above a certain threshold (adjust as needed), consider it a match
            if fuzz.partial_ratio(str(row['Full Name']), str(best_match)) > 80:
                matches = df_sheet3.loc[df_sheet3['NAME'] == best_match, 'DEPARTMENT']
                if not matches.empty:
                    return matches.values[0]
            return 'Not Found'

    # Apply the similarity function to each row
    merged_df = pd.merge(df_sheet1, df_sheet3, how='left', left_on='Full Name', right_on='NAME')
    merged_df['DEPARTMENT'] = merged_df.apply(find_similarity, axis=1)

    # Group by 'DEPARTMENT', 'Full Name', and 'BOOKS', and sum the 'Amount'
    grouped_df = merged_df.groupby(['DEPARTMENT', 'Full Name', 'BOOKS'])['Amount'].sum().reset_index()
    grouped_df['INPUT TAX'] = grouped_df['Amount'] * .12

    

    print(grouped_df)

    grouped_df.to_excel('output_file.xlsx', index=False)

    startfile("output_file.xlsx")

    transaction()
    

    # Calculate total amount per department
    # total_per_department = merged_df.groupby('DEPARTMENT','Full Name','BOOKS')['Amount'].sum().reset_index()

    # Add a new column for input tax (12% of the amount)
    # total_per_department['Input Tax'] = total_per_department['Amount'] * 0.12

    # print(total_per_department)

    # # Save the result back to Sheet1 or a new Excel file
    # merged_df.to_excel('output_file.xlsx', index=False)

    # # Save total per department with input tax to a new Excel file
    # total_per_department.to_excel('total_per_department.xlsx', index=False)

    # print(total_per_department)

    # Pivot the DataFrame to create separate columns for each book
    # pivoted_df = total_per_department.pivot(index='DEPARTMENT', columns='BOOKS', values='Amount').reset_index()

    # # Save the result back to Sheet1 or a new Excel file
    # merged_df.to_excel('output_file.xlsx', index=False)

    # # Save total per department with input tax and separate columns for each book to a new Excel file
    # pivoted_df.to_excel('total_per_department_with_books.xlsx', index=False)

    # print(pivoted_df)

def get_total_per_department2():
    # Load data from Excel sheets
    # Load data from Excel sheets
    file_path = 'LIST-AVEGA-CLAIMS1.xlsx'
    df_sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')
    df_sheet3 = pd.read_excel(file_path, sheet_name='Sheet3')

    # Function to find similarity using fuzzywuzzy
    def find_similarity(row):
        if pd.notna(row['DEPARTMENT']):
            return row['DEPARTMENT']
        else:
            # Iterate through 'NAME' in Sheet2 and find the best match
            best_match = max(df_sheet3['NAME'], key=lambda name: fuzz.partial_ratio(str(row['Full Name']), str(name)))
            
            # If similarity is above a certain threshold (adjust as needed), consider it a match
            if fuzz.partial_ratio(str(row['Full Name']), str(best_match)) > 80:
                matches = df_sheet3.loc[df_sheet3['NAME'] == best_match, 'DEPARTMENT']
                if not matches.empty:
                    return matches.values[0]
            return 'Not Found'

    # Apply the similarity function to each row
    merged_df = pd.merge(df_sheet2, df_sheet3, how='left', left_on='Full Name', right_on='NAME')
    merged_df['DEPARTMENT'] = merged_df.apply(find_similarity, axis=1)

    # Group by 'DEPARTMENT', 'Full Name', and 'BOOKS', and sum the 'Amount'
    grouped_df = merged_df.groupby(['DEPARTMENT', 'Full Name', 'BOOKS'])['Amount'].sum().reset_index()
    grouped_df['INPUT TAX'] = grouped_df['Amount'] * .12

    print(grouped_df)

    grouped_df.to_excel('avega2.xlsx', index=False)

def get_total_per_book_per_department():
    # Load data from Excel sheets
    file_path = 'LIST-AVEGA-CLAIMS1.xlsx'
    df_sheet1 = pd.read_excel(file_path, sheet_name='claims')
    df_sheet3 = pd.read_excel(file_path, sheet_name='EMPLOYEE-LIST')

    # Function to find similarity using fuzzywuzzy
    def find_similarity(row):
        if pd.notna(row['DEPARTMENT']):
            return row['DEPARTMENT']
        else:
            # Iterate through 'NAME' in Sheet2 and find the best match
            best_match = max(df_sheet3['NAME'], key=lambda name: fuzz.partial_ratio(str(row['Full Name']), str(name)))
            
            # If similarity is above a certain threshold (adjust as needed), consider it a match
            if fuzz.partial_ratio(str(row['Full Name']), str(best_match)) > 80:
                matches = df_sheet3.loc[df_sheet3['NAME'] == best_match, 'DEPARTMENT']
                if not matches.empty:
                    return matches.values[0]
            return 'Not Found'

    # Apply the similarity function to each row
    merged_df = pd.merge(df_sheet1, df_sheet3, how='left', left_on='Full Name', right_on='NAME')
    merged_df['DEPARTMENT'] = merged_df.apply(find_similarity, axis=1)

    # Group by 'DEPARTMENT', 'BOOKS', and sum the 'Amount'
    grouped_df = merged_df.groupby(['DEPARTMENT', 'BOOKS'])['Amount'].sum().reset_index()
    grouped_df['INPUT TAX'] = grouped_df['Amount'] * 0.12

    # Sort the DataFrame by 'BOOKS' column
    grouped_df = grouped_df.sort_values(by='BOOKS')
    
    print(grouped_df)

    grouped_df.to_excel('output_file.xlsx', index=False)

 
    startfile("output_file.xlsx")


def get_total_per_book_per_department_per_name():
    
    # Load data from Excel sheets
    file_path = 'LIST-AVEGA-CLAIMS1.xlsx'
    df_sheet1 = pd.read_excel(file_path, sheet_name='claims')
    df_sheet3 = pd.read_excel(file_path, sheet_name='EMPLOYEE-LIST')

    # Function to find similarity using fuzzywuzzy
    def find_similarity(row):
        if pd.notna(row['DEPARTMENT']):
            return row['DEPARTMENT']
        else:
            # Iterate through 'NAME' in Sheet2 and find the best match
            best_match = max(df_sheet3['NAME'], key=lambda name: fuzz.partial_ratio(str(row['Full Name']), str(name)))
            
            # If similarity is above a certain threshold (adjust as needed), consider it a match
            if fuzz.partial_ratio(str(row['Full Name']), str(best_match)) > 80:
                matches = df_sheet3.loc[df_sheet3['NAME'] == best_match, 'DEPARTMENT']
                if not matches.empty:
                    return matches.values[0]
            return 'Not Found'

    # Apply the similarity function to each row
    merged_df = pd.merge(df_sheet1, df_sheet3, how='left', left_on='Full Name', right_on='NAME')
    merged_df['DEPARTMENT'] = merged_df.apply(find_similarity, axis=1)

    # Group by 'Full Name', 'DEPARTMENT', and 'BOOKS' and sum the 'Amount'
    grouped_df = merged_df.groupby(['Full Name', 'DEPARTMENT', 'BOOKS']).agg({'Amount': 'sum'}).reset_index()
    grouped_df['INPUT TAX'] = grouped_df['Amount'] * 0.12

    # Sort the DataFrame by 'Full Name'
    grouped_df = grouped_df.sort_values(by='Full Name')
    
    print(grouped_df)

    grouped_df.to_excel('output_file2.xlsx', index=False)

    startfile("output_file2.xlsx")
    


# get_total_per_department2()
# create_per_department()
transaction()