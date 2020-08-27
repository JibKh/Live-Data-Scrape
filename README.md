# Live-Data-Scrape
Scrape live company data from Screener.in. You can add how many every companies you want and how many every columns you would like to extract. The columns for all the companies will be extracted onto an Excel sheet at increments you enter.

### How to run exe ###
Simply run the exe file and input the details prompted.
The Excel file will be output in the 'output' folder.

### How to run without exe ###
Simply open the main.py in your choice of code editting software and run the code.
You will need to do:
pip install selenium
pip install XlsxWriter

### How to edit ###
In the 'files' folder, you can open the txt files and add or remove columns or companies.

### Warnings ###
The txt files cannot contain any empty lines at the start, end or middle.<br />
The adding companies of the code may cause error due to the website. Please manually add websites if it throws an error.<br />
Do not open the Excel file while the code is running. Please open it during its sleep time. Or copy it and open the copy file while it is running.
