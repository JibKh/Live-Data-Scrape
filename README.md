# Live-Data-Scrape
Scrape live company data from Screener.in. You can add how many every companies you want and how many every columns you would like to extract. The columns for all the companies will be extracted onto an Excel sheet at increments you enter.

### Update ###
Due to website updates:
1) This code will only work for maximum 15 companies at a time.
2) Adding companies function may not work as expected.

### How to run exe ###
Simply run the exe file and input the details prompted.
The Excel file will be output in the 'output' folder.

### How to run without exe ###
Simply open the main.py in your choice of code editting software and run the code.
You will need to do:
1) pip install selenium
2) pip install XlsxWriter

### How to edit ###
In the 'files' folder, you can open the txt files and add or remove columns or companies.

### Make your own exe file ###
If you updated the main.py and would like to make your own exe file:
1) pip install pyinstaller
2) In the directory of the main.py open cmd. 
   pyinstaller --onefile main.py
3) Open dist folder and copy all the file dependencies the main.py has.

### Warnings ###
The txt files cannot contain any empty lines at the start, end or middle.<br />
The adding companies of the code may cause error due to the website. Please manually add websites if it throws an error.<br />
Do not open the Excel file while the code is running. Please open it during its sleep time. Or copy it and open the copy file while it is running.
