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

#function to compare
def compare(platform, target):
    #for non-chiclet-PI
    if '.' in str(platform): 
        version_split = platform.split(".")
        target_split = str(target).split(".")
        length = len(version_split)
        
        if int(version_split[0]) > int(target_split[0]): return platform
        elif int(version_split[0]) < int(target_split[0]): return target 
        else: 
            if int(version_split[1]) > int(target_split[1]): return platform
            elif int(version_split[1]) < int(target_split[1]): return target
            else:
                if int(version_split[2]) > int(target_split[2]): return platform
                elif int(version_split[2]) < int(target_split[2]): return target
                else:
                    if length == 3: return platform
                    else:
                        if int(version_split[3]) >= int(target_split[3]): return platform
                        elif int(version_split[3]) < int(target_split[3]): return target
    #for chiclet-PI
    else: 
        if int(platform) >= target: return platform
        else: return target
       
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

        for index, row in source_data.iterrows():#iterate TVA column
            result_df_column=result_df_column.append(pd.DataFrame([row['TVA']], columns=['TVA']))
        result_df_column.to_excel(writer, sheet_name='results', startcol=z, index=False)
        result_df_column = pd.DataFrame(index=None)#clear result_df_column
        z = z + 1 #column counter
         
        for columnName, columnValue in target_data.iteritems():#iterate columns within target file
        
            for index, row in source_data.iterrows():#iterate rows within source fule
                if "-" in columnName:#if columnName == 'chiclet-version' or columnName == 'chiclet-PI', also split 0-0 to 0 only
                    split = row['chiclet'].split("-")
                    if columnName == 'chiclet-version': row[columnName] = split[0]; 
                    else: row[columnName] = split[1]

                if str(row[columnName]) == '0':result_df_add = pd.DataFrame([' '], columns=[columnName])
                elif compare(row[columnName], columnValue.values[0]) == row[columnName]:result_df_add = pd.DataFrame([1], columns=[columnName]);i=i+1
                else:result_df_add = pd.DataFrame([0], columns=[columnName]);x=x+1
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


