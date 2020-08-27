from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter
from datetime import datetime

chromedriver = "drivers/chromedriver.exe"
chrome_options = Options()
chrome_options.add_argument('--headless')

# ==== USER DEFAULT ====
workbookName = "scraped"
email = ""
password = ""
time = 20
sheetName = "Results"
numberOfColumnsRetrieve = 15

class Website:
    def __init__(self, username, password, companyNames):
        self.time = time
        self.sheetName = sheetName
        self.numberOfColumnsRetrieve = numberOfColumnsRetrieve
        self.noIter = 0
        self.iter = 1

        # Initialize
        self.username = username
        self.password = password
        self.companyNames = companyNames
        # Create worksheet
        self.writeTitleColumn = 3
        self.writeDataColumn = 3
        # Start Time
        self.startTime = datetime.now()
        print("Starting Time:", self.startTime)
        
        # Get all columns
        self.readColumn = 0 # For what column we are at
        self.columnNames = []
        self.getColumns() # To fill column names

        # Get Drivers
        self.driver = webdriver.Chrome(chromedriver, options=chrome_options)
        # self.driver = webdriver.Chrome(chromedriver)
        self.driver.get("https://www.screener.in/login/")

        # Values to be reinitialized when loop is re run
        self.reinitValues = [numberOfColumnsRetrieve, self.writeTitleColumn, self.writeDataColumn, self.readColumn]

        # Login and start
        self.login()

    # Gets all the column names from text file and puts it in self.columnNames
    def getColumns(self):
        file = open("files/columns.txt", "r")
        self.columnNames = file.read().split("\n")
        self.noIter = int((len(self.columnNames) / self.numberOfColumnsRetrieve)+1)

    # Sets up the self.worksheet
    def setupWorksheet(self):
        column = 3
        # Export data
        for i in self.columnNames:
            self.worksheet.write(0, column, i)
            column += 1

    def login(self):
        # Wait for load
        self.waitForLoad("/html/body/main/div/form/div[1]/input")

        # Input
        self.driver.find_element_by_xpath('/html/body/main/div/form/div[1]/input').send_keys(self.username)
        self.driver.find_element_by_xpath('/html/body/main/div/form/div[2]/input').send_keys(self.password)
        
        # Login Button
        self.driver.find_element_by_xpath('/html/body/main/div/form/button').click()

        self.getToWatchlist()

    def getToWatchlist(self):
        # Click watchlist
        self.waitForLoad("/html/body/nav/div/div/div[3]/a[2]")
        self.driver.find_element_by_xpath("/html/body/nav/div/div/div[3]/a[2]").click()

        # Begin adding companies and extracting
        self.company()

    # Adds and extract the information of the whole company
    def company(self):
        self.addCompany()

        while 1:
            workbook = xlsxwriter.Workbook('output/' + workbookName + '.xlsx')
            self.worksheet = workbook.add_worksheet(self.sheetName)
            self.extractInfo()
            workbook.close()
            
            # End Time
            print("Time:", datetime.now())
            print("Duration:", datetime.now() - self.startTime)

            # Reinit
            self.reinit()

            # Sleep until next iteration
            print("\nSleeping for:", self.time, "seconds\n")
            sleep(self.time)
            self.startTime = datetime.now()
            print("Starting Time:", self.startTime)

    def addCompany(self):
        # Check if login is correct. If not then loop input.
        while 1:
            print("\nLogging in...")
            try:
                self.waitForLoad("/html/body/main/div/div[1]/div[2]/a")
                print("Login successful")
                break
            except:
                self.driver.get("https://www.screener.in/login/")
                print("Failed to login. Please enter details again")
                self.email = input("Email: ")
                self.password = input("Password: ")
                self.login()
                print("Please wait")
        
        print("\nAdding Companies...")
        self.driver.find_element_by_xpath("/html/body/main/div/div[1]/div[2]/a").click()

        # Wait for load
        self.waitForLoad("/html/body/main/div/div[1]/input")

        for name in self.companyNames:
            try: # Try to find if it is already added
                test = self.driver.find_element_by_class_name("items").text.lower()
                if name.lower() in test:
                    continue
                else:
                    fail = 10/0
            except: # Add if not
                # Enter company name
                self.driver.find_element_by_xpath("/html/body/main/div/div[1]/input").clear()
                sleep(0.05)
                self.driver.find_element_by_xpath("/html/body/main/div/div[1]/input").send_keys(name)

                # Wait for dropdown
                for i in range(0,1):
                    try:
                        self.waitForLoad("/html/body/main/div/div[1]/ul/li")
                        self.driver.find_element_by_xpath("/html/body/main/div/div[1]/ul/li").click()
                        break
                    except:
                        self.driver.find_element_by_xpath("/html/body/main/div/div[1]/input").clear()
                        self.driver.find_element_by_xpath("/html/body/main/div/div[1]/input").send_keys(name[0:-1])
                        sleep(0.1)
                        if i == 148:
                            self.driver.get("https://www.screener.in/user/stocks/?next=/watchlist/613221/")
                            try:
                                self.waitForLoad("/html/body/main/div/div[1]/ul/li")
                                self.driver.find_element_by_xpath("/html/body/main/div/div[1]/input").clear()
                                self.driver.find_element_by_xpath("/html/body/main/div/div[1]/input").send_keys(name)
                            except:
                                print("Unable to add company after 200 tries:", name)
                                print("Please add manually and run code again")
                                break
                        elif i == 0:
                            print("Unable to add company after 200 tries:", name)
                            print("Please add manually and run code again")

        # Click Done
        button = self.driver.find_element_by_xpath("/html/body/main/div/div[2]/a[1]")
        self.driver.execute_script("arguments[0].click();", button)
        print("Done Adding")

    def extractInfo(self):
        print("\nExtracting Info...")
        while True:
            print("         Iteration:", self.iter, "/",self.noIter)
            self.iter += 1
            # Click edit columns
            # self.waitForLoad("/html/body/main/div/div[2]/div[1]/a")
            # self.driver.find_element_by_xpath("/html/body/main/div/div[2]/div[1]/a").click()
            while 1:
                try:
                    self.waitForLoad("/html/body/main/div/div[2]/div[1]/a")
                    self.driver.find_element_by_xpath("/html/body/main/div/div[2]/div[1]/a").click()
                    break
                except:
                    print("Unable to edit columns, retrying..")
                    self.driver.get("https://www.screener.in/watchlist/")
                    sleep(1)

            # Remove all columns
            self.removeColumns()

            # Extract Columns
            if (len(self.columnNames) - self.readColumn) < self.numberOfColumnsRetrieve:
                self.numberOfColumnsRetrieve = len(self.columnNames) - self.readColumn
                if self.numberOfColumnsRetrieve != 0:
                    self.extractColumns(self.numberOfColumnsRetrieve, True)
                break
            self.extractColumns(self.numberOfColumnsRetrieve, False)
        print("Done Extracting")

    def removeColumns(self):
        try:
            self.waitForLoad("/html/body/main/div/form/ul")
        except:
            self.driver.get("https://www.screener.in/user/columns/?next=/watchlist/")
            
        while True:
            try:
                self.driver.find_elements_by_class_name("icon-cancel-thin")[0].click()
                sleep(0.2)
            except:
                break

    def extractColumns(self, x, last):
        print("                 Extracting Columns")
        columnName = ""
        # Find x columns and add them
        for i in range(0, x):
            # Try Except for incase we go out of list range
            try:
                columnName = self.columnNames[i + self.readColumn]
            except:
                break 
            # Fill form
            self.driver.find_element_by_xpath("/html/body/main/div/form/div[2]/div/div[1]/div/input").clear()
            self.driver.find_element_by_xpath("/html/body/main/div/form/div[2]/div/div[1]/div/input").send_keys(columnName)
            sleep(0.2)

            # Find the exact one to check
            for i in range(0,150):
                try:
                    search = "//input[@value='" + columnName + "']"
                    self.driver.find_element_by_xpath(search).click()
                    break
                except:
                    if i == 149:
                        print("Unable to find columns:", columnName)
        
        # Increment
        self.readColumn = self.readColumn + x

        # Save columns and return to table
        try:
            button = self.driver.find_element_by_xpath("/html/body/main/div/form/div[1]/div[1]/button")
            self.driver.execute_script("arguments[0].click();", button)
        except:
            print("Unable to find Save Column Button")
        # self.driver.find_element_by_xpath("/html/body/main/div/form/div[1]/div[1]/button").click()
        
        print("                         Writing Columns")
        # All Titles
        titles = self.driver.find_elements_by_tag_name("th")

        # Export Titles
        for i in titles[3:]:
            self.worksheet.write(0, self.writeTitleColumn, i.text)
            self.writeTitleColumn += 1

        # All Value Data
        data = self.driver.find_elements_by_tag_name("td")

        # Split Data by number of rows for each company
        a = 0
        splitData = []
        for i in range (a, (x+3)*len(self.companyNames), x+3):
            a = i
            splitData.append(data[a : a+x+3])

        # Export data
        for i, value in enumerate(splitData):
            for j in value[3:]:
                try:
                    self.worksheet.write(i+1, self.writeDataColumn, float(j.text))
                except:
                    self.worksheet.write(i+1, self.writeDataColumn, j.text)
                self.writeDataColumn += 1
            self.writeDataColumn = self.writeDataColumn - x
        self.writeDataColumn += x

        # If its the last iteration, add the first 3 columns
        if last:
            # Title
            for i in range(0,3):
                self.worksheet.write(0, i, titles[i].text)
            # Values
            for i in range(len(self.companyNames)):
                for j in range(0, 3):
                    try:
                        try:
                            self.worksheet.write(i+1, j, float(splitData[i][j].text))
                        except:
                            self.worksheet.write(i+1, j, splitData[i][j].text)
                    except:
                        print("Error printing first 3 columns. Continuing.")
        print("                         Writing Done")
        print("                 Done Extracting")
    
    def reinit(self):
        self.numberOfColumnsRetrieve = self.reinitValues[0]
        self.writeTitleColumn = self.reinitValues[1]
        self.writeDataColumn = self.reinitValues[2]
        self.readColumn = self.reinitValues[3]
        self.iter = 1

    def waitForLoad(self, xpath):
        Wait = WebDriverWait(self.driver, 10)       
        Wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))

def main():

    # User Input
    print("Please input the initial details:\n")
    global workbookName
    workbookName = input("Workbook Name. For ex: Results:\n-> ")
    global sheetName
    sheetName = input("\nName of the sheet within the workbook. For example: Sheet1:\n-> ")
    global email
    email = input("\nEmail:\n-> ")
    global password
    password = input("\nPassword:\n-> ")
    global time
    time = int(input("\nTime Interval. How often would you like the data to be retrieved.\nIn Seconds:\n-> "))
    print("")

    # Open Company File
    file = open("files/companies.txt", "r")
    companyNames = file.read().split("\n")

    # Start Bot
    Website(email, password, companyNames)
    
    #workbook.close()

main()