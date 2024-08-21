import msal
import requests
from datetime import datetime
from urllib.parse import urlparse, parse_qs
import json
import os
import html2text
import re

########################################################################################################################
# INSTRUCTIONS
# 1. Create an app registration in Azure
#     1. Supported Account Type: 'Personal Microsoft accounts only'
#     2. Redirect URIs: 
#         1. Public client/native http://localhost
# 2. Copy the 'Client ID' to APP_REGISTRATION_CLIENT_ID
# 3. Click on the 'CreateApplication ID URI' link then create the URI with the default value it that it generates
# 4. Click on the 'Authentication' tab then toggle on 'Allow public client flows'
# 5. Click on the 'API Permissions' tab then add permissions then add the following permissions:
#     1. Task.Read
# 6. Click on the 'Manifest' tab then modify the value of "signInAudience" to be "AzureADandPersonalMicrosoftAccount"
# 7. Click on the 'Expose an API' tab then add Task.Read to 'Authorized client applications'
########################################################################################################################

SAVE_AS_MARKDOWN = True
SAVE_ATTACHMENTS = False
GRAPH_API_TOKEN_CACHE = "graph_api_token_cache.bin"
GRAPH_API_URL = "https://graph.microsoft.com/beta/me/todo/lists/"
APP_REGISTRATION_CLIENT_ID = "SET THIS VALUE"
REDIRECT_URI = "http://localhost"
SCOPES = ["Tasks.Read"]

def get_or_create_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(GRAPH_API_TOKEN_CACHE):
        with open(GRAPH_API_TOKEN_CACHE, "r") as cache_file:
            cache.deserialize(cache_file.read())
    return cache

def save_cache(cache):
    with open(GRAPH_API_TOKEN_CACHE, "w") as cache_file:
        cache_file.write(cache.serialize())

def clean_markdown(content):
    # When the HTML is converted to markdown it isn't perfectly 'clean'
    # This method will 'clean' the converted markdown by removing unnecessary characters, spaces, and newlines

    # Remove standalone **, __, or _ lines
    content = re.sub(r'^\s*(\*\*|__)\s*$', '', content, flags=re.MULTILINE)
    content = re.sub(r'^\s*(\*\*|_)\s*$', '', content, flags=re.MULTILINE)
    # Remove lines that are just underscores followed by spaces
    content = re.sub(r'^\s*_{2,}\s*$', '', content, flags=re.MULTILINE)
    # Remove extra newlines (more than 2 consecutive)
    content = re.sub(r'\n{3,}', '\n\n', content)
    # Remove spaces at the end of lines
    content = re.sub(r' +$', '', content, flags=re.MULTILINE)
    return content.strip()

def sanitize_filename(filename):
    # Remove or replace characters that are invalid in filenames
    return re.sub(r'[<>:"/\\|?*]', '_', filename)

# Get or create the token cache
token_cache = get_or_create_cache()

# Create the client
app = msal.PublicClientApplication(
    APP_REGISTRATION_CLIENT_ID,
    token_cache=token_cache
)

# Try to get a token from the cache
accounts = app.get_accounts()
if accounts:
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if result:
        print("Token found in cache")
        access_token = result['access_token']
    else:
        print("No suitable token found in cache. Acquiring new token...")
        result = app.acquire_token_interactive(SCOPES)
        if "access_token" in result:
            access_token = result['access_token']
        else:
            print(result.get("error"))
            print(result.get("error_description"))
            exit()
else:
    print("No accounts found. Acquiring new token...")
    result = app.acquire_token_interactive(SCOPES)
    if "access_token" in result:
        access_token = result['access_token']
    else:
        print(result.get("error"))
        print(result.get("error_description"))
        exit()

# Save the token cache
save_cache(token_cache)

if not result:
    # No suitable token exists in cache. Let's get a new one from AAD.
    flow = app.initiate_auth_code_flow(scopes=SCOPES, redirect_uri=REDIRECT_URI)
    
    print("Please go to this URL and sign in:")
    print(flow["auth_uri"])
    
    auth_response = input("After signing in, paste the full URL of the page you were redirected to: ")
    
    try:
        # Parse the URL to extract the query parameters
        parsed_url = urlparse(auth_response)
        query_params = parse_qs(parsed_url.query)
        
        # Create a dictionary with the parsed parameters
        auth_response_dict = {
            'code': query_params.get('code', [None])[0],
            'state': query_params.get('state', [None])[0]
        }
        
        result = app.acquire_token_by_auth_code_flow(flow, auth_response_dict)
        print("Token acquisition result:")
        print(json.dumps(result, indent=2))
    except Exception as e:
        print(f"Error acquiring token: {str(e)}")
        if hasattr(e, 'args') and len(e.args) > 0:
            print("Full error details:")
            print(json.dumps(e.args[0], indent=2))
        exit()

if "access_token" not in result:
    print(f"Error: {result.get('error')}")
    print(f"Error description: {result.get('error_description')}")
    exit()

# Get all task lists
headers = {
    'Authorization': f'Bearer {access_token}',
    'Prefer': 'outlook.body-content-type="html"'
}
lists_response = requests.get(f"{GRAPH_API_URL}delta", headers=headers)
lists = lists_response.json().get('value', [])

if SAVE_AS_MARKDOWN:
    # Create HTML to Markdown converter
    h = html2text.HTML2Text()
    h.ignore_links = False
    h.body_width = 0
    h.ul_item_mark = '-'  # Use - for unordered lists instead of *

    # Open the file in write mode
    for task_list in lists:
            # Create a filename with list name and current date and time
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            list_name = sanitize_filename(task_list['displayName'])
            filename = f"{list_name}_{current_time}.md"

            # Get tasks for this list
            tasks_response = requests.get(f"{GRAPH_API_URL}{task_list['id']}/tasks", headers=headers)
            tasks = tasks_response.json().get('value', [])

            # Skip empty lists
            if not tasks:
                continue

            with open(filename, 'w', encoding='utf-8') as file:
                # List name
                file.write(f"# List: {task_list['displayName']}\n\n\n") # (ID: {task_list['id']})
                
                # Tasks
                is_first_task = True
                for task in tasks:
                    # New line between tasks
                    if not is_first_task:
                        file.write("\n\n")
                    else:
                        is_first_task = False
                    
                    # Title
                    file.write(f"## Task: {task['title']}\n")

                    # Status
                    file.write(f"### Status: {'Completed' if task['status'] == 'completed' else 'Not Completed'}\n")
                    
                    # Due Date
                    if task.get('dueDateTime'):
                        file.write(f"### Due: {task['dueDateTime']['dateTime']}\n")

                    # Reminder
                    if task.get('reminderDateTime'):
                        file.write(f"### Reminder: {task['reminderDateTime']['dateTime']}\n")
                    
                    # Attachments
                    attachments_response = requests.get(f"{GRAPH_API_URL}{task_list['id']}/tasks/{task['id']}/attachments", headers=headers)
                    attachments = attachments_response.json().get('value', [])
                    if attachments:
                        file.write("### Attachments:\n")
                        for attachment in attachments:
                            file.write(f"- {attachment['name']} (Size: {attachment['size']} bytes)\n")
                            
                            # Download Attachments
                            if SAVE_ATTACHMENTS and attachment['@odata.type'] == '#microsoft.graph.taskFileAttachment':
                                attachment_content_response = requests.get(f"{GRAPH_API_URL}{task_list['id']}/tasks/{task['id']}/attachments/{attachment['id']}/$value", headers=headers)
                                if attachment_content_response.status_code == 200:
                                    attachment_filename = f"attachment_{attachment['name']}"
                                    with open(attachment_filename, 'wb') as attachment_file:
                                        attachment_file.write(attachment_content_response.content)
                                    file.write(f"  - Saved attachment: {attachment_filename}\n")
                    
                    # Body Content
                    if task.get('body'):
                        content_type = task['body']['contentType']
                        content = task['body']['content']
                        markdown_content = clean_markdown(h.handle(content))
                        if markdown_content and markdown_content.strip():
                            # file.write(f"### Content Type: {content_type}\n")
                            file.write("### Content:")
                            
                            if content_type.lower() == 'html':
                                # Convert HTML to Markdown
                                file.write(markdown_content)
                            else:
                                file.write(content)
                            file.write("\n")
            
            print(f"Tasks for list '{task_list['displayName']}' have been exported to {filename}")

else:
    # Write the task lists to file
    for task_list in lists:
        # Get tasks for this list
        tasks_response = requests.get(f"{GRAPH_API_URL}{task_list['id']}/tasks", headers=headers)
        tasks = tasks_response.json().get('value', [])

        # Skip empty lists
        if not tasks:
            continue

        # Create a filename with list name and current date and time
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        list_name = sanitize_filename(task_list['displayName'])
        filename = f"{list_name}_{current_time}.txt"

        with open(filename, 'w', encoding='utf-8') as file:
            # List name
            file.write(f"List: {task_list['displayName']}\n\n\n") # (ID: {task_list['id']})\n")
            
            # Tasks
            is_first_task = True
            for task in tasks:
                # Line breaks between tasks
                if not is_first_task:
                    file.write("\n\n")
                else:
                    is_first_task = False

                # Title
                file.write(f"  Task: {task['title']}\n")

                # Status
                file.write(f"  Status: {'Completed' if task['status'] == 'completed' else 'Not Completed'}\n")
                
                # Due date
                if task.get('dueDateTime'):
                    file.write(f"  Due: {task['dueDateTime']['dateTime']}\n")

                # Reminder
                if task.get('reminderDateTime'):
                    file.write(f"  Reminder: {task['reminderDateTime']['dateTime']}\n")
                
                # Attachments
                attachments_response = requests.get(f"{GRAPH_API_URL}{task_list['id']}/tasks/{task['id']}/attachments", headers=headers)
                attachments = attachments_response.json().get('value', [])
                if attachments:
                    file.write("  Attachments:\n")
                    for attachment in attachments:
                        file.write(f"    - {attachment['name']} (Size: {attachment['size']} bytes)\n")
                        
                        # Download Attachments
                        if SAVE_ATTACHMENTS and attachment['@odata.type'] == '#microsoft.graph.taskFileAttachment':
                            attachment_content_response = requests.get(f"{GRAPH_API_URL}{task_list['id']}/tasks/{task['id']}/attachments/{attachment['id']}/$value", headers=headers)
                            if attachment_content_response.status_code == 200:
                                attachment_filename = f"attachment_{attachment['name']}"
                                with open(attachment_filename, 'wb') as attachment_file:
                                    attachment_file.write(attachment_content_response.content)
                                file.write(f"      Saved attachment: {attachment_filename}\n")
                
                # Body Content
                if task.get('body'):
                    content_type = task['body']['contentType']
                    content = task['body']['content']
                    #file.write(f"  Content Type: {content_type}\n")
                    file.write(f"  Content: {content}\n")
        
        print(f"Tasks for list '{task_list['displayName']}' have been exported to {filename}")
