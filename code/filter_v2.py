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
        writer = pd.ExcelWriter(results + '/results.xlsx')
        result_df_column = pd.DataFrame(index=None)
        i = x = z = 0

        #Iterate TVA column
        for index, row in source_data.iterrows():
            result_df_column=result_df_column.append(pd.DataFrame([row['TVA']], columns=['TVA']))
        result_df_column.to_excel(writer, sheet_name='results', startcol=z, index=False)
        result_df_column = pd.DataFrame(index=None)#clear result_df_column
        z = z + 1 #column counter
         
        #iterate columns within target file
        for columnName, columnValue in target_data.iteritems():
            
            if columnName == 'chiclet-version' or columnName == 'chiclet-PI':
                for index, row in source_data.iterrows():
                    if str(row['chiclet']).find(str(columnValue.values[0])) != -1:result_df_add = pd.DataFrame(['YES'], columns=[columnName]);i = i + 1
                    elif str(row['chiclet']) == '0-0':result_df_add = pd.DataFrame([0], columns=[columnName])
                    else:result_df_add = pd.DataFrame(['NO'], columns=[columnName]);x=x+1
                    result_df_column=result_df_column.append(result_df_add)
            else:
                for index, row in source_data.iterrows():
                    if row[columnName] == columnValue.values[0]:result_df_add = pd.DataFrame(['YES'], columns=[columnName]);i = i + 1
                    elif str(row[columnName]) == '0':result_df_add = pd.DataFrame([0], columns=[columnName])
                    else:result_df_add = pd.DataFrame(['NO'], columns=[columnName]);x = x + 1
                    result_df_column=result_df_column.append(result_df_add)
            #add correct/incorrect/total/correct percentage to the end of each column
            result_df_column=result_df_column.append(pd.DataFrame(['Total = ', i+x], columns=[columnName]))
            result_df_column=result_df_column.append(pd.DataFrame(['Correct = ', i], columns=[columnName]))
            result_df_column=result_df_column.append(pd.DataFrame(['Incorrect = ', x], columns=[columnName]))
            result_df_column=result_df_column.append(pd.DataFrame(['% = ', round((i/(i+x))*100)], columns=[columnName]))
            result_df_column.to_excel(writer, sheet_name='results', startcol=z, index=False)
            result_df_column = pd.DataFrame(index=None)#clear result_df_column
            z = z + 1 #count the column number and write the next column on the right of this one
            i = x = 0

        #iterate Playback Hours column
        for index, row in source_data.iterrows():
            result_df_column=result_df_column.append(pd.DataFrame([row['Playback Hours']], columns=['Playback Hours']))
        result_df_column.to_excel(writer, sheet_name='results', startcol=z, index=False)

        writer.close()#close excel writer

        print('Successfully output results')
window.close()#close GUI window