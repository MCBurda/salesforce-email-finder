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
cookies = [
{
    "domain": ".force.com",
    "expirationDate": 1621676401.339031,
    "hostOnly": False,
    "httpOnly": False,
    "name": "_ga",
    "path": "/",
    "secure": False,
    "session": False,
    "storeId": "0",
    "value": "GA1.2.906777506.1595527784",
    "id": 1
},
{
    "domain": ".force.com",
    "expirationDate": 1613471645.780588,
    "hostOnly": False,
    "httpOnly": False,
    "name": "BrowserId",
    "path": "/",
    "secure": False,
    "session": False,
    "storeId": "0",
    "value": "aZKTbOLQEeqVRbs2Ldz3Gw",
    "id": 2
},
{
    "domain": ".force.com",
    "expirationDate": 1613471645.780809,
    "hostOnly": False,
    "httpOnly": False,
    "name": "BrowserId_sec",
    "path": "/",
    "secure": True,
    "session": False,
    "storeId": "0",
    "value": "aZKTbOLQEeqVRbs2Ldz3Gw",
    "id": 3
},
{
    "domain": ".force.com",
    "hostOnly": False,
    "httpOnly": False,
    "name": "inst",
    "path": "/",
    "secure": True,
    "session": True,
    "storeId": "0",
    "value": "APP_5I",
    "id": 4
},
{
    "domain": "XXX.lightning.force.com",
    "hostOnly": True,
    "httpOnly": False,
    "name": "clientSrc",
    "path": "/",
    "secure": True,
    "session": True,
    "storeId": "0",
    "value": "81.161.133.81",
    "id": 5
},
{
    "domain": "XXX.lightning.force.com",
    "expirationDate": 1606652904.628993,
    "hostOnly": True,
    "httpOnly": False,
    "name": "force-proxy-stream",
    "path": "/",
    "secure": True,
    "session": False,
    "storeId": "0",
    "value": "!mUS/mIzRHKXLLq2smU9OHObjWe25sbrd64B7NBAWiFFG2UAzmj84vT6bu9O7xIAlWPcm6CapNV7W5Jo=",
    "id": 6
},
{
    "domain": "XXX.lightning.force.com",
    "expirationDate": 1606652904.629051,
    "hostOnly": True,
    "httpOnly": False,
    "name": "force-stream",
    "path": "/",
    "secure": True,
    "session": False,
    "storeId": "0",
    "value": "!ZkbnQzo5lEMAzplZVR8FTdvJzFPr6bP0cSqXrtbvuZ7E9mW0nvugFJGN85CGLr4GKc4an3NC+YWTCA==",
    "id": 7
},
{
    "domain": "XXX.lightning.force.com",
    "expirationDate": 1606652904.628905,
    "hostOnly": True,
    "httpOnly": False,
    "name": "sfdc-stream",
    "path": "/",
    "secure": True,
    "session": False,
    "storeId": "0",
    "value": "!ZkbnQzo5lEMAzplZVR8FTdvJzFPr6bP0cSqXrtbvuZ7E9mW0nvugFJGN85CGLr4GKc4an3NC+YWTCA==",
    "id": 8
},
{
    "domain": "XXX.lightning.force.com",
    "hostOnly": True,
    "httpOnly": True,
    "name": "sid",
    "path": "/",
    "secure": True,
    "session": True,
    "storeId": "0",
    "value": "00D20000000BU0P!AR8AQNzdR5_V8sWg.kvga_QGzMjvNVdpt21DMyhNFnQGZFfB.fVmN7IbXa1EuLPLxCvmT_nzShtwXl4ySuTIr8vhrhHEpP2l",
    "id": 9
},
{
    "domain": "XXX.lightning.force.com",
    "hostOnly": True,
    "httpOnly": False,
    "name": "sid_Client",
    "path": "/",
    "secure": True,
    "session": True,
    "storeId": "0",
    "value": "N000009ZuIY0000000BU0P",
    "id": 10
}
]

# READ BEFORE EXECUTION
# This small script is made for one specific application: Turning CSV lists of accounts into lists of emails of contacts that are associated with those accounts.
# The script  will automate your browser keystrokes and can run in the background, once the necessary cookies are inserted to login to Salesforce (see above)
# This is performed by opening your company's Salesforce organization at the account overview level with the Requests library and opening a CSV file that you need to provide in the same directory as this script.
# The csv should contain 2 columns without headings. One for Account names and one for the emails, which this script will find. You provide all the account names on the left most column and save the csv.
# The script then searches for each account in the account tab search field, finds the account name's contact record, pulls the first found email of a contacts associated with this account, and saves it in the csv.
# After the last update, the script also find the first and last name of each contact, in order to identify them.

def Main():
    # The csv file with your account data should be named "data.csv"
    filename = "data.csv"
    fieldnames = ['AccountId', 'Email', "First_name", "Last_name"]

    with open(filename, "r") as csvfile:
        reader = csv.DictReader(csvfile, fieldnames=fieldnames)

        with open("email_data.csv", "w", newline="") as target_file:
            writer = csv.DictWriter(target_file, fieldnames=fieldnames)
            writer.writeheader()

            options = Options()
            options.headless = False
            driver = webdriver.Chrome(options=options)
            driver.get("https://duck.com")

            # Adding tasty cookies
            for cookie in cookies:
                driver.add_cookie(cookie)

            time.sleep(2)

            driver.get("https://XXX.lightning.force.com/lightning/o/Account/list?filterName=Recent")

            # Wait for Search bar to load
            WebDriverWait(driver, timeout=1).until(lambda d: d.find_element_by_css_selector("div.forceSearchInputEntitySelector"))
            time.sleep(5)

            ObjectSelect = driver.find_element(By.CLASS_NAME, "forceSearchInputEntitySelector")
            ObjectSelect.click()
            ContactSelect = driver.find_element_by_xpath("//span[@title='Contacts']")
            ContactSelect.click()


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
                    time.sleep(1)
                    email = driver.find_element(By.CLASS_NAME, "emailuiFormattedEmail").text
                    print(email)
                except:
                    # Deleting data from search bar if email not found and continue with the next record
                    searchBar.click()
                    searchBar.send_keys(Keys.CONTROL + "a")
                    searchBar.send_keys(Keys.DELETE)
                    continue

                try:
                    time.sleep(1)
                    name = driver.find_element(By.CSS_SELECTOR, "div.hideRowNumberColumn div div div.actionBarPlugin div table.forceRecordLayout tbody tr th.slds-cell-edit span a").get_attribute("title")
                    print(name)
                except:
                    # Deleting data from search bar if email not found and continue with the next record
                    searchBar.click()
                    searchBar.send_keys(Keys.CONTROL + "a")
                    searchBar.send_keys(Keys.DELETE)
                    continue

                if len(name.split()) < 2: # If the name of the contact consists of one word, we add "Unknown" as the last name
                    name = name + " Unknown"

                writer.writerow({
                    "AccountId": row["AccountId"],
                    "Email": email,
                    "First_name": name.split()[0],  # The first name is the first part of the found contact name
                    "Last_name": name.split()[1]  # The last name is the second part.
                })

                # Deleting data from search bar
                searchBar.click()
                searchBar.send_keys(Keys.CONTROL + "a")
                searchBar.send_keys(Keys.DELETE)

            driver.quit()



if ("__main__" == __name__):
    Main()
