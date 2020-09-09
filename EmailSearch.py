from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
import time
import csv

# These cookies must be taken from your browser before using this script. Download the following browser extension: https://chrome.google.com/webstore/detail/editthiscookie/fngmhnnpilhplaeedifhccceomclgfbg (optional) to read your browser cookies.
# Log into your instance of Salesforce and use the extension to find out the values for the 9 cookies listed below. All required values and some domain names are marked as #####. Try looking for the cookies with the same names.
cookies = [{"domain": "#####", "hostOnly": "true", "httpOnly": False, "name": "sid_Client",
             "path": "/", "secure": True, "session": "true", "storeId": "0", "value": "#####",
             "id": 9},
           {"domain": "#####", "hostOnly": "true", "httpOnly": True, "name": "sid",
             "path": "/", "secure": True, "session": "true", "storeId": "0",
             "value": "#####",
             "id": 8},
            {"domain": "#####", "expirationDate": 1597770033.128076, "hostOnly": "true",
             "httpOnly": False, "name": "sfdc-stream", "path": "/", "secure": True, "session": "false", "storeId": "0",
             "value": "#####", "id": 7},
            {"domain": "#####", "expirationDate": 1597770033.128183, "hostOnly": "true",
             "httpOnly": False, "name": "force-stream", "path": "/", "secure": True, "session": "false", "storeId": "0",
             "value": "#####", "id": 6},
            {"domain": "#####", "expirationDate": 1597770033.128132, "hostOnly": "true",
             "httpOnly": False, "name": "force-proxy-stream", "path": "/", "secure": True, "session": "false",
             "storeId": "0",
             "value": "#####", "id": 5},
            {"domain": "#####", "hostOnly": "true", "httpOnly": False, "name": "clientSrc",
             "path": "/", "secure": True, "session": "true", "storeId": "0", "value": "#####", "id": 4},
            {"domain": ".force.com", "hostOnly": "false", "httpOnly": False, "name": "inst", "path": "/",
            "secure": True, "session": "true", "storeId": "0", "value": "#####", "id": 3},
            {"domain": ".force.com", "expirationDate": 1629295233.12803, "hostOnly": "false", "httpOnly": False,
             "name": "BrowserId_sec", "path": "/", "secure": True, "session": "false", "storeId": "0",
             "value": "#####", "id": 2},
            {"domain": ".force.com", "expirationDate": 1629295233.127926, "hostOnly": "false", "httpOnly": False,
             "name": "BrowserId", "path": "/", "secure": True, "session": "false", "storeId": "0",
             "value": "#####", "id": 1}]

# READ BEFORE EXECUTION
# This small script is made for one specific application: Turning CSV lists of accounts into lists of emails of contacts that are associated with those accounts.
# The script  will automate your browser keystrokes and can run in the background, once the necessary cookies are inserted to login to Salesforce (see above)
# This is performed by opening your company's Salesforce organization at the account overview level with the Requests library and opening a CSV file that you need to provide in the same directory as this script.
# The csv should contain 2 columns without headings. One for Account names and one for the emails, which this script will find. You provide all the account names on the left most column and save the csv.
# The script then searches for each account in the account tab search field, finds the account name's contact record, pulls the first found email of a contacts associated with this account, and saves it in the csv.

def Main():
    # The csv file with your account data should be named "data.csv"
    filename = "data.csv"

    with open(filename, "r") as csvfile:
        reader = csv.DictReader(csvfile)

        with open("email_data.csv", "w", newline="") as target_file:
            fieldnames = ['AccountId', 'Email']
            writer = csv.DictWriter(target_file, fieldnames=fieldnames)
            writer.writeheader()

            options = Options()
            options.headless = True
            driver = webdriver.Chrome(options=options)
            driver.get("https://google.com")

            # Adding tasty cookies
            for cookie in cookies:
                driver.add_cookie(cookie)

            time.sleep(2)

            driver.get("https://#####.lightning.force.com/lightning/o/Account/list?filterName=Recent")

            # Wait for Search bar to load
            WebDriverWait(driver, timeout=1).until(lambda d: d.find_element_by_css_selector("input.slds-input"))


            for row in reader:

                if row["AccountId"] == "":
                    continue

                time.sleep(1) # Wait 1 second explicitly for general load

                # Finding search bar
                searchBar = driver.find_element(By.CSS_SELECTOR, "input.slds-input")
                searchBar.click()
                searchBar.send_keys(row["AccountId"] + Keys.RETURN)

                # Wait for emails to be loaded
                try:
                    WebDriverWait(driver, timeout=10).until(EC.presence_of_element_located((By.CLASS_NAME,"emailuiFormattedEmail")))
                except:
                    # Deleting data from search bar if email not found and continue with the next record
                    searchBar.click()
                    searchBar.send_keys(Keys.CONTROL + "a")
                    searchBar.send_keys(Keys.DELETE)
                    continue

                # Find the first contact's email and copy it to the "email" variable
                try:
                    email = driver.find_element(By.CLASS_NAME, "emailuiFormattedEmail").text
                    print(email)
                except:
                    time.sleep(1)
                    email = driver.find_element(By.CLASS_NAME, "emailuiFormattedEmail").text
                    print(email)


                writer.writerow({
                    "AccountId": row["AccountId"],
                    "Email": email
                })

                # Deleting data from search bar
                searchBar.click()
                searchBar.send_keys(Keys.CONTROL + "a")
                searchBar.send_keys(Keys.DELETE)

            driver.quit()



if ("__main__" == __name__):
    Main()