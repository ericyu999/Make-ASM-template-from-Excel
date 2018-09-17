# Make-ASM-template-from-Excel
create ASM worksheet from invpack excel, using python openpyxl

07/09/2018 main function is working now, for the next step will be implement below Function

#TODO, file can be saved in another location instead of python script folder. Done 12/09/2018
#TODO, create one file to input path and file name, separate the main script in another file, by doing this, can teach other colleagues to use it. Done 12/09/2018
#TODO, distribute on colleagues PC


installation procedure
1. download python
2. pip install openpyxl
3. pip install pywin32
4. edit WorksheetMaker.bat, change the file path
5. edit MakingAsmWorksheet.py, change the template file path
6. move WorksheetMaker.bat to Python script folder
7. win + R, type worksheetmaker
8. follow the prompt and enjoy!


This Branch is an experiment to replace WIN32 module with Python standard open and close. 
see if it solves the ASM error. Guess the problem lies when openpyxl save the file, the cursor stays at the end of file.
