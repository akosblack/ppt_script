# PPT Script
A script that replace placeholders in a ppt* file form a xls* file.

How its working:
  - From the input directory checking if there is any xls and a ppt file available
  - From the xlx file its gets the data (one row is one data)  
  - Then replace all defined keywords to the data (the first keyword to the first data)
  
## Input
  The excel input file should be a one sheet file with the following structure:
  | keyword1 | data1 |
  | -------- | ----- |
  | keyword2 | data2 |
  | keyword3 | data3 |

  etc.

  The ppt input file should be a pptx file with the following structure:
  - Where you want to replace the data you should write the keyword in. The keyword should be the same as in the excel file. Like this: "<\keyword1>"

## Output
  The output file will be in the output directory. The name of the file will be the same as the input file name but with a timestamp at the end. The new file should now contain the replaced data.

## How to create an executeable file from project:
  - Install pyinstaller with command
  ```sh
  pip install pyinstaller
  ```
  - Run the following command in the project directory:
  ```sh
  pyinstaller --onefile --icon=icon.ico script.py
  ```
  - The executeable file will be in the dist folder

