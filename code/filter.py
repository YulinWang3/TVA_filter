import PySimpleGUI as sg
import pandas as pd
from xlsxwriter import Workbook

#GUI window design
sg.theme('Light Blue 2')
layout = [[sg.Text('Browse your source and target files')],
          [sg.Text('Source File', size=(8, 1)), sg.Input(key='source'), sg.FileBrowse()],
          [sg.Text('Target File', size=(8, 1)), sg.Input(key='target'), sg.FileBrowse()],
          [sg.Text('Results', size=(8, 1)), sg.Input(key='results'), sg.FolderBrowse()],
          [sg.Submit('Filter Them'), sg.Cancel('Close Program')]]
window = sg.Window('Filter', layout)

#GUI window
while True:
    event, values = window.read()
    if event in (sg.WINDOW_CLOSED, "Close Program"):break
    elif event == 'Filter Them':
        file_1, file_2, results = values['source'], values['target'], values['results']
        if bool(file_1) == False or bool(file_2) == False:break
       
        #import target and source files and input DataFrames
        source_table = pd.read_csv (file_1)
        target_table = pd.read_excel (file_2)
        source_data = pd.DataFrame(source_table)
        target_data = pd.DataFrame(target_table)
        i = x = 0

        #Create output dataframe tables
        col_correct = ['Index','Updated TVAs']
        col_incorrect = ['Index','Not updated TVAs']
        result_df_correct = pd.DataFrame(index=None, columns=col_correct)
        result_df_incorrect = pd.DataFrame(index=None, columns=col_incorrect)
        writer = pd.ExcelWriter(results + '/results.xlsx')

        #iterate each column of target table
        for columnName, columnValue in target_data.iteritems():
            print('')
            print(columnName)
           
            #for chiclet column - allows user to check version and PI separately
            if columnName == 'chiclet-version' or columnName == 'chiclet-PI':
                for index, row in source_data.iterrows():
                    if str(row['chiclet']).find(str(columnValue.values[0])) != -1:
                        result_df_add = pd.DataFrame([[i+1,str(row['TVA'])]], columns=col_correct)
                        result_df_correct=result_df_correct.append(result_df_add)
                        i = i+1
                    elif str(row['chiclet']) != '0-0':
                        result_df_add = pd.DataFrame([[x+1,str(row['TVA'])]], columns=col_incorrect)
                        result_df_incorrect=result_df_incorrect.append(result_df_add)
                        x = x+1
            #for other columns - only version checking needed
            else:
                for index, row in source_data.iterrows():
                    if row[columnName] == columnValue.values[0]:
                        result_df_add = pd.DataFrame([[i+1,str(row['TVA'])]], columns=col_correct)
                        result_df_correct=result_df_correct.append(result_df_add)
                        i = i+1
                    elif row[columnName] != '0':
                        result_df_add = pd.DataFrame([[x+1,str(row['TVA'])]], columns=col_incorrect)
                        result_df_incorrect=result_df_incorrect.append(result_df_add)
                        x = x+1
            
            #export tables to a excel
            if i!=0: result_df_correct.to_excel(writer, sheet_name=columnName, startrow=0, startcol=0, index=False)
            if x!=0: result_df_incorrect.to_excel(writer, sheet_name=columnName, startrow=0, startcol=3, index=False)
            
            #print out relavent infomation in terminal
            if i != 0:print(result_df_correct)
            print('')
            if x != 0:print(result_df_incorrect)
            print('total tester:', i+x)
            print('correct = ', i)
            print('incorrect = ', x)

            #clear dataframes and counters
            result_df_correct = pd.DataFrame(columns=col_correct)
            result_df_incorrect = pd.DataFrame(columns=col_incorrect)
            x = i = 0
            print('')
            print('')

        writer.close()#close excel writer
window.close()#close GUI window