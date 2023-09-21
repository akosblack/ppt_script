# ppt_script
A script that replace placeholders in a ppt* file form a xls* file.

How its working:

  - From the input directory checking if there is any xls and a ppt file available
  - From the xlx file its gets the data (one row is one data)  
  - Then replace all defined keywords to the data (the first keyword to the first data)

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

