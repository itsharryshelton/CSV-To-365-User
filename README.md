# Automatic 365 User Creation via CSV

This was something I needed for a work job, so putting here if anyone needs it. 

--------------------------------------------------------

Currently, this will take a CSV File with the below columns:

_"Preferred First Name"_

_"Preferred Last Name"_

_"Preferred Email (subject to availability)"_

_"Employee Type"_ (This one gets converted to Department in 365)

_"Job Title"_

Convert it to a clean CSV file for 365, then enter it into 365 for user creation - it does not do any license changes, as wasn't required here.
It will export a CSV file of passwords for each user, with reset on next sign on set.

---------------------------------------------------------

### Current Limitation:

Does not automatically select the CSV file it makes to import into 365, so it will ask you again for this.

**The way it'll work through:**

1. Ask you for the source CSV
2. Ask for save location for the clean CSV
3. Ask you to select the clean CSV to import to 365
4. Ask for the save location for the password list

----------------------------------------------------------

# Prerequisites

Latest Version of Python Installed to run the .py file

Script uses these modules, so you will need to be sure they are installed: csv, tkinter, pandas, requests, msal, random, string

You will need to edit the script to make it work with your 365 - **you will need to make a Entra App Registration, and add the Secret ID, Tenant ID and Application ID on line 106, 107 & 108**

![image](https://github.com/user-attachments/assets/bfe15dbb-9464-47f2-ace6-3c1b9e03df88)


----------------------------------------------------------

## Entra App API Settings:

You will need to add:

Microsoft Graph:

User.ReadWrite - Delegated

User.ReadWrite.All - Application

![image](https://github.com/user-attachments/assets/d43e11ad-03fe-43da-8a04-8c271ff3ce61)
