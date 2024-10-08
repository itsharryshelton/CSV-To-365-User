import csv
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import requests
import msal
import random
import string

#CSV Function
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
            display_name = f"{FirstName} {LastName}"  #This combines First and Last name
            job_title = row.get('Job Title', '')
            department = row.get('Employee Type', '')
            country = ''  #If you want to add logic to get country, modify here

            #This starts the writing bit
            writer.writerow([UserPrincipalName, FirstName, LastName, display_name, job_title, department, country])

    print(f"CSV file '{output_file}' created successfully with transformed data.")

#Function for GUI
def get_file_locations():
    root = tk.Tk()
    root.withdraw()

    #Ask for the input file
    input_file = filedialog.askopenfilename(title="Select Input CSV File", filetypes=[("CSV Files", "*.csv")])
    
    #Ask for the output file location
    output_file = filedialog.asksaveasfilename(title="Save New CSV File As", defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])

    return input_file, output_file

#File locations and usage
input_file, output_file = get_file_locations()
create_csv_file(input_file, output_file)

#365 BIT STARTS HERE

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

    #Shuffle Shuffle
    random.shuffle(password)

    return ''.join(password)

#Get existing 365 users
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
    # Azure AD app details
    client_id = 'EDIT ME'
    client_secret = 'EDIT ME'
    tenant_id = 'EDIT ME'
    authority_url = f'https://login.microsoftonline.com/{tenant_id}'
    scope = ['https://graph.microsoft.com/.default']

    # Authenticate
    app = msal.ConfidentialClientApplication(client_id, authority=authority_url, client_credential=client_secret)
    token_response = app.acquire_token_for_client(scopes=scope)

    if 'access_token' not in token_response:
        print("Failed to obtain access token")
        exit()

    access_token = token_response['access_token']

    #Get existing users for checking
    existing_users = get_existing_users(access_token)
    users = None
    
    try:
        users = pd.read_csv(csv_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read CSV file: {str(e)}")
        return

    #Check for required columns
    required_columns = ['DisplayName', 'UserPrincipalName', 'FirstName', 'LastName']
    missing_columns = [col for col in required_columns if col not in users.columns]

    if missing_columns:
        messagebox.showerror("Error", f"CSV file is missing the following columns: {', '.join(missing_columns)}")
        return

    #Remove duplicates based on existing users
    users = users[~users['UserPrincipalName'].isin(existing_users)]
    
    #Handle missing or NaN values in UserPrincipalName
    users['UserPrincipalName'] = users['UserPrincipalName'].fillna('')
    users = users[users['UserPrincipalName'] != '']  # Remove rows where UserPrincipalName is empty

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

    try:
        #This part will save the usernames into a list for you
        with open(save_path, 'w') as file:
            file.write("Email,Password\n")
            for entry in email_password_list:
                file.write(f"{entry}\n")

        #The End
        messagebox.showinfo("Success", f"User creation completed. Passwords saved to {save_path}.")
    finally:
        root.destroy()
        root.quit()

#GUI to select CSV file and save location
def select_file():
    csv_path = filedialog.askopenfilename(title="Select the New CSV File", filetypes=[("CSV files", "*.csv")])
    if csv_path:
        save_path = filedialog.asksaveasfilename(defaultextension=".csv", title="Save Passwords List As", filetypes=[("CSV files", "*.csv")])
        if save_path:
            create_users(csv_path, save_path)

#Setup the GUI for 365 side
root = tk.Tk()
root.title("Create Microsoft 365 Users")

select_button = tk.Button(root, text="Select the new CSV and Password Location", command=select_file)
select_button.pack(pady=40)

root.mainloop()
