# TVA_filter
* A filter that checks if TVA update Fibe trial update to the correct version.
* A prompt window requires user to input a source excel file, a target excel file and a path where the place to generate result file
* Source file contains a table of all account number(TVAs) with different app verion on multiple platforms
* Target file contains a list of platforms with their correct app version
* Note: Target file cannot check a platform that does exist on Source file
* This program will generate a result excel file with all filtered TVAs for different platforms

# Requirement
* Python
* pyopenxl
* pandas
* PySimpleGUI 
* xlsxwriter
