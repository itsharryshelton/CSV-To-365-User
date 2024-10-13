# 🔣 Automatic 365 User Creation via CSV
### Latest Version: v1.2

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

_Looking to have this pull this info from a TXT so you can customize the bad CSV file easier than editing the code on the next version maybe_

---------------------------------------------------------

# 📖 The way it'll work through:

GUI Box will open, select the function you want to run, either cleaning CSV or just 365 making.
Terminal Box within the GUI so you can see the output as it goes on.

![image](https://github.com/user-attachments/assets/2d867f9a-403c-4848-a23b-dd07b747393e)


----------------------------------------------------------

# 🧠 Prerequisites

Latest Version of Python Installed to run the .py file

Script uses these modules, so you will need to be sure they are installed: csv, tkinter, pandas, requests, msal, random, string

You will need to download the AzureAPIDetails.txt and place it within the same folder as the .py, update this with your Azure API Details

![image](https://github.com/user-attachments/assets/d397b6f9-223d-49d4-8154-b63b77aa6499)



----------------------------------------------------------

## 💡 Entra App API Settings:

You will need to add:

Microsoft Graph:

User.ReadWrite - Delegated

User.ReadWrite.All - Application

![image](https://github.com/user-attachments/assets/d43e11ad-03fe-43da-8a04-8c271ff3ce61)


----------------------------------------------------------

## ⚡Python Modules used

You will need to make sure that the below modules are installed, and are callable.

**Standard Python Modules**

_csv_

_os_

_random_

_string_

_sys_

_threading_

_io_


**Third-party imports**

_pandas_

_requests_

_msal_

_customtkinter_


**Tkinter - CustomTkinter is for main GUI, Tkinter is for File Calling**

_tkinter_
