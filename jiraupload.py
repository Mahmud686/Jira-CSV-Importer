#Mahmudur Rahman
#mahmudurrahman298@gmail.com
# Copyright © 2023 Mahmudur Rahman. All rights reserved.

import tkinter as tk
from tkinter import messagebox
import requests
import csv
import sys
import os

# Define a dictionary to store the epic keys
epic_key_dict = {}

script_directory = os.path.dirname(os.path.abspath(__file__))



selected_project_key = None
def get_full_path(filename):
    if getattr(sys, 'frozen', False):
        # Running as a compiled .exe, use sys._MEIPASS
        return os.path.join(sys._MEIPASS, filename)
    else:
        # Running as a script, use relative path
        script_directory = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(script_directory, filename)


def check_and_initialize_token():
    global email, api_token
    token_file_path = get_full_path("token.txt")
    if os.path.exists(token_file_path):
        with open(token_file_path, "r") as txt_file:
            email = txt_file.readline().strip()
            api_token = txt_file.readline().strip()
    else:
        setup_token()


def setup_token():
    def save_token():
        global email, api_token
        email = email_entry.get()
        api_token = api_token_entry.get()

        if not email or not api_token:
            messagebox.showerror("Error", "Email and API Token are required.")
            return

        with open(get_full_path("token.txt"), "w") as txt_file:
            txt_file.write(email + "\n" + api_token)

        token_window.destroy()

    token_window = tk.Tk()
    token_window.title("Token Setup")
    token_window.geometry("500x350")

    email_label = tk.Label(token_window, text="Email:", font=("Helvetica", 12))
    email_label.pack(pady=10)

    email_entry = tk.Entry(token_window, font=("Helvetica", 12))
    email_entry.pack()

    api_token_label = tk.Label(token_window, text="API Token:", font=("Helvetica", 12))
    api_token_label.pack(pady=10)

    api_token_entry = tk.Entry(token_window, font=("Helvetica", 12))
    api_token_entry.pack()

    save_button = tk.Button(token_window, text="Save", command=save_token, font=("Helvetica", 12))
    save_button.pack(pady=10)

    token_window.mainloop()

def check_project_key(project_key):
    try:
        with open(get_full_path("project_key.txt"), "r") as txt_file:
            stored_project_key = txt_file.read().strip()
            return project_key == stored_project_key
    except FileNotFoundError:
        return False

def check_duplicate_epic(email, api_token, project_key, epic_name):
    url = "https://your-jira-instance.atlassian.net/rest/api/3/search"              #write your jira instance
    headers = {
        "Accept": "application json",
        "Content-Type": "application/json"
    }

    jql_query = f'project = "{project_key}" AND issuetype = Epic AND "Epic Name" ~ "{epic_name}"'

    payload = {
        "jql": jql_query,
        "fields": ["key"],
    }

    response = requests.post(url, json=payload, headers=headers, auth=(email, api_token))

    if response.status_code == 200:
        epic_data = response.json()
        if epic_data.get("issues"):
            return True
    return False

def get_or_create_epic_key(email, api_token, project_key, epic_name):
    # Check Epic already exists in the dictionary
    epic_key = epic_key_dict.get(epic_name)
    if epic_key:
        print(f"Epic Key for '{epic_name}' found in epic_key_dict: {epic_key}")
        return epic_key

    # Check Epic with name already existing in Jira
    existing_epic_key = get_epic_key_by_name(email, api_token, project_key, epic_name)
    if existing_epic_key:
        print(f"Epic with the name '{epic_name}' already exists in Jira. Using existing Epic Key: {existing_epic_key}")
        epic_key_dict[epic_name] = existing_epic_key
        return existing_epic_key

    # Create the Epic if it doesn't exists
    epic_key = create_epic(email, api_token, project_key, epic_name)
    if epic_key:
        epic_key_dict[epic_name] = epic_key
        print(f"New Epic Key for '{epic_name}' added to epic_key_dict: {epic_key}")
    return epic_key


def get_epic_key_by_name(email, api_token, project_key, epic_name):
    url = "https://your-jira-instance.atlassian.net/rest/api/3/search"                      #write your jira instance
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    jql_query = f'project = "{project_key}" AND issuetype = Epic AND summary ~ "{epic_name}"'

    payload = {
        "jql": jql_query,
        "fields": ["key"],
    }

    response = requests.post(url, json=payload, headers=headers, auth=(email, api_token))

    if response.status_code == 200:
        epic_data = response.json()
        if epic_data.get("issues"):
            return epic_data["issues"][0]["key"]
    return None

def create_epic(email, api_token, project_key, epic_name):
    url = "https://your-jira-instance.atlassian.net/rest/api/2/issue"                   #write your jira instance
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    payload = {
        "fields": {
            "project": {"key": project_key},
            "issuetype": {"name": "Epic"},
            "summary": epic_name,
        }
    }

    response = requests.post(url, json=payload, headers=headers, auth=(email, api_token))

    if response.status_code == 201:
        epic_data = response.json()
        return epic_data["key"]
    return None

def process_csv(project_key, email, api_token, update_option):
    url = "https://your-jira-instance.atlassian.net/rest/api/2/issue"                   #write your jira instance
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    
    #epics_processed = 0
    issues_processed = 0
    total_time = 0

     
    #response = None  # Initialize response to None

    try:
        with open(get_full_path("CSVFilePath.txt"), "r") as txt_file:
            csv_file_path = txt_file.read().strip()
    except FileNotFoundError:
        print("CSVFilePath.txt not found. Make sure the VBA macro has created it.")
        return

    with open(csv_file_path, "r", encoding="utf-8-sig") as csvfile:
        csv_reader = csv.DictReader(csvfile, delimiter=";")
        for row in csv_reader:
            vorgangs_id = row['Task-ID']
            issue_type = row["Issue Type"]

            # Check if the issue_type is "epic" and skip it
            if issue_type.lower() == "epic":
                continue

            # Process other fields for non-epic rows
            summary = f"{row['Summary']}- {vorgangs_id}"
            description = row["Description"]
            epic_link = row["Epic Link"]
            original_estimate = int(row["Original Estimate (in seconds)"]) if row["Original Estimate (in seconds)"].isdigit() else 0
            original_estimate /=3600
            remaining_estimate = int(row["Remaining Estimate (in seconds)"]) if row["Remaining Estimate (in seconds)"].isdigit() else 0
            remaining_estimate /=3600
            components = row["Components"]

            if issue_type == "Epic":
                # Epic link is not applicable for epics
                epic_link = None
            else:
                # Check if the Epic Link is not empty, and get or create the Epic Key
                if epic_link:
                    if not check_duplicate_epic(email, api_token, project_key, epic_link):
                        epic_key = get_or_create_epic_key(email, api_token, project_key, epic_link)
                        if epic_key:
                            epic_link = epic_key
                    else:
                        epic_link = None  # No Epic Link provided in CSV

            # Build a JQL query to search for an issue with the same summary (Task-ID)
            jql_query = f'project = "{project_key}" AND summary ~ "{vorgangs_id}"'
            search_url = f"https://your-jira-instance.atlassian.net/rest/api/2/search?jql={jql_query}"          #write your jira instance

            search_response = requests.get(search_url, headers=headers, auth=(email, api_token))

            if search_response.status_code != 200:
                print(search_response.text)
                continue

            try:
                search_data = search_response.json()

                if search_data['issues']:
                    existing_issue_id = search_data['issues'][0]['key']

                    if update_option == "1":
                        # Build the URL for the existing issue
                        update_url = f"https://your-jira-instance.atlassian.net/rest/api/2/issue/{existing_issue_id}"       #write your jira instance

                        # Define the fields you want to update
                        update_fields = {
                            "summary": summary,
                            "description": description,
                            "timetracking": {
                                "originalEstimate": original_estimate,
                                "remainingEstimate": remaining_estimate
                            },
                            "components": [{"name": components}],
                            "customfield_10014": epic_link  # Epic Link
                        }

                        update_payload = {
                            "fields": update_fields
                        }

                        update_response = requests.put(update_url, json=update_payload, headers=headers,
                                                       auth=(email, api_token))

                        if update_response.status_code == 204:
                            print(f"Issue {existing_issue_id} updated successfully.")
                        else:
                            print(f"Failed to update issue {existing_issue_id}: {update_response.text}")
                    elif update_option == "2":
                        user_input = input(f"issue with key: {existing_issue_id} exists. Do you want to update this issue? (yes/no): ")
                        if user_input.lower() == "y" or user_input.lower() == "yes":
                            # Build the URL for the existing issue
                            update_url = f"https://your-jira-instance.atlassian.net/rest/api/2/issue/{existing_issue_id}"          #write your jira instance

                            # Define the fields you want to update
                            update_fields = {
                                "summary": summary,
                                "description": description,
                                "timetracking": {
                                    "originalEstimate": original_estimate,
                                    "remainingEstimate": remaining_estimate
                                },
                                "components": [{"name": components}],
                                "customfield_10014": epic_link  # Epic Link
                            }

                            update_payload = {
                                "fields": update_fields
                            }

                            update_response = requests.put(update_url, json=update_payload, headers=headers,
                                                           auth=(email, api_token))

                            if update_response.status_code == 204:
                                print(f"Issue {existing_issue_id} updated successfully.")
                            else:
                                print(f"Failed to update issue {existing_issue_id}: {update_response.text}")
                else:
                    # Create a new issue
                    payload = {
                        "fields": {
                            "project": {"key": project_key},
                            "issuetype": {"name": issue_type},
                            "summary": summary,
                            "timetracking": {
                                "originalEstimate": original_estimate,
                                "remainingEstimate": remaining_estimate
                            },
                            "description": description,
                            "components": [{"name": components}],
                            "customfield_10014": epic_link  # Epic Link
                        }
                    }

                    response = requests.post(url, json=payload, headers=headers, auth=(email, api_token))
                    if response.status_code == 201:
                        print(f"New issue created: {response.json()['key']}")
                    else:
                        print(f"Failed to create a new issue: {response.text}")
            except Exception as e:
                print(f"An error occurred while processing Task-ID {vorgangs_id}: {str(e)}")
                continue
            if issue_type == "Epic":
                epics_processed += 1

            else:
                issues_processed += 1
                total_time += original_estimate 

    # Create a summary message
    summary_message = f"Processing completed!\n"
    #summary_message += f"Number of Epic Processed: {epics_processed}\n"
    summary_message += f"Number of Issues Processed: {issues_processed}\n"
    summary_message += f"Total Time (hours): {total_time:.2f}"

    # Display the summary message in a message box
    messagebox.showinfo("Processing Summary", summary_message)


def project_selected():
    selected_project = project_var.get()
    if selected_project:
        project_key = selected_project.split(" - ")[0]
        stored_project_key = check_project_key(project_key)
        if stored_project_key:
            # Create a new tkinter window
            option_window = tk.Toplevel(window)
            option_window.title("Update Options")
            option_window.geometry("400x150")
            
            # Add a label and text field
            label = tk.Label(option_window, text="Choose an update option:", font=("Helvetica", 12))
            label.pack(pady=10)

            update_option = tk.IntVar()
            option1_button = tk.Radiobutton(option_window, text="Update All Issues", variable=update_option, value=1)
            option2_button = tk.Radiobutton(option_window, text="Update One by One", variable=update_option, value=2)
            option1_button.pack()
            option2_button.pack()

            def update_issues():
                if update_option.get() == 1:
                    # Update All Issues
                    option_window.destroy()
                    process_csv(project_key, email, api_token, "1")
                    window.destroy()
                elif update_option.get() == 2:
                    # Update One by One
                    option_window.destroy()
                    process_csv(project_key, email, api_token, "2")
                    window.destroy()
                    
            update_button = tk.Button(option_window, text="Update", command=update_issues, font=("Helvetica", 12))
            update_button.pack(pady=10)
        else:
            recent_project_key = get_recent_project_key()
            error_message = f"Error: Recent Project Key is {recent_project_key}. Check your Selected Project or Project Key."
            messagebox.showerror("Error", error_message)

def get_recent_project_key():
    file_path = get_full_path("project_key.txt")
    if os.path.exists(file_path):
        with open(file_path, "r") as txt_file:
            stored_project_key = txt_file.read().strip()
            return stored_project_key
    else:
        print("project_key.txt file not found")
        return None  # Or a default value if needed

def get_components(email, api_token, project_key):
    url = f"https://your-jira-instance.atlassian.net/rest/api/3/project/{project_key}/components"           #write your jira instance
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers, auth=(email, api_token))

    if response.status_code == 200:
        component_data = response.json()
        component_names = [component["name"] for component in component_data]
        return component_names
    else:
        return []

def show_components():
    selected_project = project_var.get()
    if selected_project:
        project_key = selected_project.split(" - ")[0]
        project_name = selected_project.split(" - ")[1]

        components = get_components(email, api_token, project_key)

        if components:
            # Create a new window to display components
            components_window = tk.Toplevel(window)
            components_window.title("Components in the Project")
            components_window.geometry("400x700")

            project_name_label = tk.Label(components_window, text=f"Selected Project: {project_key}", font=("Helvetica", 12))
            project_name_label.pack(pady=10, anchor="w")

            components_label = tk.Label(components_window, text="Components in the Project:", font=("Helvetica", 12))
            components_label.pack(anchor="w")

            for component in components:
                component_label = tk.Label(components_window, text=component, font=("Helvetica", 12))
                component_label.pack(anchor="w")

        else:
            messagebox.showinfo("No Components", "No components found in the selected project.")
            
if __name__ == "__main__":
    check_and_initialize_token()
    default_project_key = get_recent_project_key()

    def get_project_info(email, api_token):
        url = "https://your-jira-instance.atlassian.net/rest/api/3/project"             #write your jira instance
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json"
        }

        response = requests.get(url, headers=headers, auth=(email, api_token))

        if response.status_code == 200:
            project_data = response.json()
            project_info = [(f"{project['key']} - {project['name']}") for project in project_data]
            return project_info
        else:
            return []


    window = tk.Tk()
    window.title("Jira Issue Uploader")
    window.geometry("800x350")

    frame = tk.Frame(window)
    frame.pack(pady=20)

    project_info = get_project_info(email, api_token)

    project_var = tk.StringVar()
    project_label = tk.Label(frame, text="Select Project:", font=("Helvetica", 12))
    project_label.grid(row=0, column=0, padx=10, pady=5)

    if project_info:
        # Set the default project based on the content of "project_key.txt"
        default_project = next((proj for proj in project_info if proj.startswith(default_project_key)) if default_project_key else project_info[0])
        project_var.set(default_project)

    project_dropdown = tk.OptionMenu(frame, project_var, *project_info)
    project_dropdown.grid(row=0, column=1)

    error_label = tk.Label(frame, text="", font=("Helvetica", 12), fg="red")
    error_label.grid(row=2, column=0, columnspan=2)

    submit_button = tk.Button(frame, text="Submit", command=project_selected, font=("Helvetica", 12))
    submit_button.grid(row=1, column=0, columnspan=2, pady=20)
    
    show_components_button = tk.Button(frame, text="Show Components", command=show_components, font=("Helvetica", 12))
    show_components_button.grid(row=3, column=0, columnspan=2, pady=20)

window.mainloop()