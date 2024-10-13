# -*- coding: utf-8 -*-
"""
Copyright (c) 2024 - Harry Shelton
All rights reserved.

This project is licensed under the MIT License. You may obtain a copy of the License at
https://opensource.org/licenses/MIT

You are free to use, modify, and distribute this software under the terms of the MIT License.
This software is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the MIT License for more details.

Author: Harry Shelton
Date: 13th October 2024
Repository: https://github.com/itsharryshelton/CSV-To-365-User

CSV to 365 v1.2
"""


import csv
import os
import random
import string
import sys
import threading
import io

#Third-party imports - getting any errors with the 365 side or main GUI interface, make sure these are installed :D
import pandas as pd
import requests
import msal
import customtkinter as ctk

#Tkinter
import tkinter as tk
from tkinter import filedialog, messagebox





# ++++++++++++++++++ CSV CLEAN UP FUNCTIONS ++++++++++++++++++

#Create the new CSV Function
def create_csv_file(input_file, output_file):
    new_headers = ['UserPrincipalName', 'FirstName', 'LastName', 'DisplayName', 'Job Title', 'Department', 'Country']

    with open(input_file, mode='r', newline='', encoding='utf-8') as infile, \
         open(output_file, mode='w', newline='', encoding='utf-8') as outfile:
        
        reader = csv.DictReader(infile)
        writer = csv.writer(outfile)
        
        writer.writerow(new_headers)

        for row in reader:
            UserPrincipalName = row.get('Preferred Email (subject to availability)', '').replace(" ", "")
            FirstName = row.get('Preferred First Name', '')
            LastName = row.get('Preferred Last Name', '')
            display_name = f"{FirstName} {LastName}"
            job_title = row.get('Job Title', '')
            department = row.get('Employee Type', '')
            country = row.get('Country', '')

            #This starts the writing bit
            writer.writerow([UserPrincipalName, FirstName, LastName, display_name, job_title, department, country])

    print(f"CSV file '{output_file}' created successfully with transformed data.")

#Get CSV File for Cleanup
def get_file_locations():
    root = tk.Tk()
    root.withdraw()

    input_file = filedialog.askopenfilename(title="Select Input CSV File", filetypes=[("CSV Files", "*.csv")])
    output_file = filedialog.asksaveasfilename(title="Save New CSV File As", defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])

    return input_file, output_file











# ++++++++++++++++++ 365 BIT STARTS HERE ++++++++++++++++++

def generate_random_password(length=8):
    if length < 8:
        raise ValueError("Password length should be at least 8 characters.")
    
    lower = string.ascii_lowercase
    upper = string.ascii_uppercase
    digits = string.digits
    special = string.punctuation

    password = [
        random.choice(lower),
        random.choice(upper),
        random.choice(digits),
        random.choice(special),
    ]

    all_characters = lower + upper + digits + special
    password += random.choices(all_characters, k=length - 4)
    random.shuffle(password)

    return ''.join(password)

#Get the existing 365 users here to check for duplicates
def get_existing_users(access_token):
    existing_users = []
    url = 'https://graph.microsoft.com/v1.0/users'
    headers = {'Authorization': f'Bearer {access_token}'}

    while url:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            users = response.json()
            existing_users.extend(user['userPrincipalName'] for user in users.get('value', []))
            url = users.get('@odata.nextLink')
        else:
            print(f"Error retrieving users: {response.text}")
            break

    return existing_users

#Function to create users in 365
def create_users(csv_path, save_path):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    api_details_path = os.path.join(script_dir, 'AzureAPIDetails.txt')
    
    #Get the API Details from the AzureAPIDetails.txt (THIS NEEDS TO BE SAME ROOT AS THE .PY :D !)
    try:
        with open(api_details_path, 'r') as f:
            lines = f.readlines()
            client_id = lines[0].strip()
            client_secret = lines[1].strip()
            tenant_id = lines[2].strip()
    except FileNotFoundError:
        print("Error: AzureAPIDetails.txt file not found.")
        return
    except IndexError:
        print("Error: AzureAPIDetails.txt file is missing some required details. Enter Client ID, Client Secret and Tenant ID as one line each, no extra commas or speech marks")
        return

    authority_url = f'https://login.microsoftonline.com/{tenant_id}'
    scope = ['https://graph.microsoft.com/.default']

    app = msal.ConfidentialClientApplication(client_id, authority=authority_url, client_credential=client_secret)
    token_response = app.acquire_token_for_client(scopes=scope)

    #Check for 365 token
    if 'access_token' not in token_response:
        print("Failed to obtain access token")
        exit()

    access_token = token_response['access_token']
    existing_users = get_existing_users(access_token)
    users = None
    
    #Read the cleaned CSV file
    try:
        users = pd.read_csv(csv_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read CSV file: {str(e)}")
        return

    #Check all required columns were made
    required_columns = ['DisplayName', 'UserPrincipalName', 'FirstName', 'LastName']
    missing_columns = [col for col in required_columns if col not in users.columns]

    if missing_columns:
        messagebox.showerror("Error", f"CSV file is missing the following columns: {', '.join(missing_columns)}")
        return

    #Remove duplicates from listbased on existing users
    users = users[~users['UserPrincipalName'].isin(existing_users)]
    users['UserPrincipalName'] = users['UserPrincipalName'].fillna('')
    users = users[users['UserPrincipalName'] != '']  #Remove rows where UserPrincipalName is empty so no errors happen

    if users.empty:
        messagebox.showinfo("Info", "No new users to create. All emails are duplicates or invalid.")
        return

    email_password_list = []

    for index, user in users.iterrows():
        if isinstance(user['UserPrincipalName'], str) and user['UserPrincipalName']:
            password = generate_random_password()
            user_data = {
                "accountEnabled": True,
                "displayName": user['DisplayName'],
                "mailNickname": user['UserPrincipalName'].split('@')[0],
                "userPrincipalName": user['UserPrincipalName'],
                "givenName": user['FirstName'],
                "surname": user['LastName'],
                "passwordProfile": {
                    "forceChangePasswordNextSignIn": True,
                    "password": password
                }
            }

            response = requests.post(
                'https://graph.microsoft.com/v1.0/users',
                headers={'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'},
                json=user_data
            )

            if response.status_code == 201:
                print(f"Created user: {user['DisplayName']}")
                email_password_list.append(f"{user['UserPrincipalName']},{password}")
            else:
                print(f"Failed to create user: {user['DisplayName']} - {response.text}")
        else:
            print(f"Invalid UserPrincipalName for user: {user['DisplayName']}")

    #This part will save the usernames into a list for you
    with open(save_path, 'w') as file:
        file.write("Email,Password\n")
        for entry in email_password_list:
            file.write(f"{entry}\n")

    #The End
    messagebox.showinfo("Success", f"User creation completed. Passwords saved to {save_path}.")


#Function to select CSV file and save location for 365 user bit
def select_file():
    csv_path = filedialog.askopenfilename(title="Select the New CSV File", filetypes=[("CSV files", "*.csv")])
    if csv_path:
        save_path = filedialog.asksaveasfilename(defaultextension=".csv", title="Save Passwords List As", filetypes=[("CSV files", "*.csv")])
        if save_path:
            create_users(csv_path, save_path)

app = ctk.CTk()
app.geometry("400x200")
app.title("Create Microsoft 365 Users")

#Buttons for File Dialogue for 365 selection bit
select_button = ctk.CTkButton(app, text="Select the new CSV and Password Location", command=select_file)
select_button.pack(pady=40)









# ++++++++++++++++++ APPLICATION GUI + APP LOOP starts here ++++++++++++++++++

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")
stop_flag = threading.Event()

def on_select_csv_click():
    input_file, output_file = get_file_locations()
    if input_file and output_file:
        create_csv_file(input_file, output_file)
    else:
        print("No files selected.")

def on_select_365_click():
    input_file, output_file = get_file_locations()
    if input_file and output_file:
        create_users(input_file, output_file)
    else:
        print("No files selected.")

#Function to start the process
def start_program():
    stop_flag.clear()
    threading.Thread(target=run_creation_process).start()

def run_creation_process():
    try:
        input_file, output_file = get_file_locations()
        create_csv_file(input_file, output_file)
        print("Process finished successfully.")
    except Exception as e:
        print(f"Error occurred: {e}")

#Button for CSV Clean
start_button = ctk.CTkButton(master=app, text="Clean CSV File - Step 1", command=start_program)
start_button.pack(pady=20)

#Button for 365 creation
start_button = ctk.CTkButton(master=app, text="Create 365 Users - Step 2", command=start_program)
start_button.pack(pady=20)

#Main GUI window bit
app = ctk.CTk()
app.geometry("400x400")
app.title("User Management")

#Terminal Box in the GUI
output_textbox = ctk.CTkTextbox(app, width=400, height=150)
output_textbox.pack(pady=20)

class TextRedirector(io.StringIO):
    def __init__(self, textbox):
        super().__init__()
        self.textbox = textbox

    def write(self, message):
        self.textbox.insert("end", message)
        self.textbox.see("end")
        self.textbox.update() 
redirector = TextRedirector(output_textbox)
sys.stdout = redirector

make_new_starters_button = ctk.CTkButton(app, text="Clean CSV File - Step 1", 
                                         command=on_select_csv_click, 
                                         width=150,
                                         height=40,
                                         corner_radius=10)
make_new_starters_button.pack(pady=20)

select_button = ctk.CTkButton(app, text="Create 365 Users - Step 2", 
                              command=on_select_365_click, 
                              width=150, 
                              height=40, 
                              corner_radius=10)
select_button.pack(pady=20)

#APP Loop starts here
app.mainloop()
