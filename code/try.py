import pandas as pd
import PySimpleGUI as sg
from xlsxwriter import Workbook

columns_correct = ['Index','Updated TVAs']
columns_incorrect = ['Index','Not updated TVAs']
resultDataFrame_correct = pd.DataFrame(columns=columns_correct)
resultDataFrame_incorrect = pd.DataFrame(columns=columns_incorrect)


resultDataFrame_addition = pd.DataFrame([[5,9]], columns=columns_correct, index=[1])
resultDataFrame_correct = resultDataFrame_correct.append(resultDataFrame_addition)

resultDataFrame_addition = pd.DataFrame([[6,10]], columns=columns_correct, index=[2])
resultDataFrame_correct = resultDataFrame_correct.append(resultDataFrame_addition)

resultDataFrame_addition = pd.DataFrame([[7,15]], columns=columns_correct, index=[3])
resultDataFrame_correct = resultDataFrame_correct.append(resultDataFrame_addition)


resultDataFrame_addition = pd.DataFrame([[6,98]], columns=columns_incorrect, index=[1])
resultDataFrame_incorrect = resultDataFrame_incorrect.append(resultDataFrame_addition)

resultDataFrame_addition = pd.DataFrame([[7,99786867868768686868]], columns=columns_incorrect, index=[2])
resultDataFrame_incorrect = resultDataFrame_incorrect.append(resultDataFrame_addition)

resultDataFrame_addition = pd.DataFrame([[8,888768768767868767868]], columns=columns_incorrect, index=[3])
resultDataFrame_incorrect = resultDataFrame_incorrect.append(resultDataFrame_addition)



layout = [
    [
        sg.Text("Please enter your source file path here:"),
        sg.In(size=(25,1), enable_events=True, key='-FOLDER-'),
        sg.FolderBrowse(),
    ]
]

sg.Window(title='filter', layout=[[]])


resultDataFrame_correct.name="correct"
resultDataFrame_correct.name="incorrect"

writer = pd.ExcelWriter('/Users/wangyulin/Desktop/results.xlsx')
# workbook=writer.book
# worksheet=workbook.add_worksheet('Sheet1')
# writer.sheets['Sheet1'] = worksheet

#worksheet.write_string(0,0,resultDataFrame_correct.name)
resultDataFrame_correct.to_excel(writer, sheet_name='Sheet1', startrow=0, startcol=0, index=False)

# worksheet.write_string(resultDataFrame_correct.shape[0]+4,0,resultDataFrame_incorrect.name)
resultDataFrame_incorrect.to_excel(writer, sheet_name='Sheet1', startrow=0, startcol=3, index=False)
writer.close()

print(resultDataFrame_correct)
print(resultDataFrame_incorrect)
print('')


