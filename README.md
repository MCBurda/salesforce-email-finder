# Salesforce Email Finder

**READ BEFORE EXECUTION:**
This small script is made for one specific application: Turning CSV lists of accounts into lists of emails of contacts that are associated with those accounts.

The script  will automate your browser keystrokes and can run in the background, once the necessary cookies are inserted to login to Salesforce (see "cookies" variable in file). This is performed by opening your company's Salesforce organization at the account overview level with the Requests library and opening a CSV file that you need to provide in the same directory as this script. The csv should contain 2 columns without headings. One for Account names and one for the emails, which this script will find. You provide all the account names on the left most column and save the csv. The script then searches for each account in the account tab search field, finds the account name's contact record, pulls the first found email of a contacts associated with this account, and saves it in the csv.

**Why is this useful?**

Many intersting Salesforce reports can only be generated on **Account** basis or **Account and Opportunity** basis. These reports don't allow the user to include Contact emails in the report, which can be downloaded as a CSV and uploaded to other marketing systems (e.g. Hubspot). This script solves this problem by finding of email addresses linked to CSVs of account names, which can be downloaded from a report.

