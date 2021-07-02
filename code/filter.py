import pandas as pd
from prompt_toolkit import prompt
from xlsxwriter import Workbook

#prompt user to enter source file and target file path
source_file = prompt('Please enter source file path:\n')
print('')
target_file = prompt('Please enter target file path:\n')

#import target build and retrieve source
source_table = pd.read_excel (source_file)
target_table = pd.read_excel (target_file)

#input DateFrames from target and source excel files
source_data = pd.DataFrame(source_table)
target_data = pd.DataFrame(target_table)
i = x = 0

#tables title of TVAs who have/haven't updated to the correct version
col_correct = ['Index','Updated TVAs']
col_incorrect = ['Index','Not updated TVAs']

#output dataframes for correct/incorrlect TVAs
result_df_correct = pd.DataFrame(index=None, columns=col_correct)
result_df_incorrect = pd.DataFrame(index=None, columns=col_incorrect)
writer = pd.ExcelWriter('results.xlsx')

#iterate each column of target table
for columnName, columnValue in target_data.iteritems():
    print('')
    print(columnName)
    #for chiclet column
    if columnName == 'chiclet':
        for index, row in source_data.iterrows():
            if str(row[columnName]).find(str(columnValue.values[0])) != -1:
                result_df_add = pd.DataFrame([[i+1,str(row['TVA'])]], columns=col_correct)
                result_df_correct=result_df_correct.append(result_df_add)
                i = i+1
            elif str(row[columnName]) != '0-0':
                result_df_add = pd.DataFrame([[x+1,str(row['TVA'])]], columns=col_incorrect)
                result_df_incorrect=result_df_incorrect.append(result_df_add)
                x = x+1
    #for other columns
    else:
        for index, row in source_data.iterrows():
            if row[columnName] == columnValue.values[0]:
                result_df_add = pd.DataFrame([[i+1,str(row['TVA'])]], columns=col_correct)
                result_df_correct=result_df_correct.append(result_df_add)
                i = i+1
            elif row[columnName] != 0:
                result_df_add = pd.DataFrame([[x+1,str(row['TVA'])]], columns=col_incorrect)
                result_df_incorrect=result_df_incorrect.append(result_df_add)
                x = x+1
    
    if i!=0: result_df_correct.to_excel(writer, sheet_name=columnName, startrow=0, startcol=0, index=False)
    if x!=0: result_df_incorrect.to_excel(writer, sheet_name=columnName, startrow=0, startcol=3, index=False)
    


    #ignore empty tables and print not empty tables
    if i != 0:print(result_df_correct)
    print('')
    if x != 0:print(result_df_incorrect)
    
    #print out relavent values
    print('total tester:', i+x)
    print('correct = ', i)
    print('incorrect = ', x)

    #clear dataframes and counters
    result_df_correct = pd.DataFrame(columns=col_correct)
    result_df_incorrect = pd.DataFrame(columns=col_incorrect)
    x = i = 0
    print('')
    print('')


writer.close()
