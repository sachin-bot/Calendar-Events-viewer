import re
import requests
import tkinter as tk
from tkinter import messagebox
from tkinter.ttk import Button, Treeview
import datetime
import duration
import webbrowser


client_id = "Your Client ID Here"
client_secret = "Your Client secret here"
tenant_id = "Your Tenanat ID here"
GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/'


def get_access_token():
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "resource": "https://graph.microsoft.com"
    }
    response = requests.post(url, data=data)
    return response.json().get("access_token")



def delete_event(selected_item):
    user_email = email_entry.get ( )
    headers = {
        'Authorization': f'Bearer {get_access_token ( )}'
    }
    response = requests.get (GRAPH_ENDPOINT + f"users?$filter=mail eq '{user_email}'", headers=headers)
    user_data = response.json ( )
    if 'value' in user_data:
        user = user_data['value'][0]
        user_id = user['id']
        if selected_item:
            item_values = events_table.item(selected_item, "values")
            event_id = item_values[-1]  # Assuming event ID is stored at the last index
            DelReq1 = requests.delete (GRAPH_ENDPOINT + f"users/{user_id}/events/{event_id}", headers=headers)
            print(DelReq1.text)
            events_table.delete(selected_item)



def fetch_calendar_events():
    user_email = email_entry.get ( )
    start_time = start_time_entry.get()
    end_time = end_time_entry.get()

    headers = {
        'Authorization': f'Bearer {get_access_token ( )}'
    }

    response = requests.get (GRAPH_ENDPOINT + f"users?$filter=mail eq '{user_email}'", headers=headers)
    user_data = response.json ( )

    if 'value' in user_data:
        user = user_data['value'][0]
        user_id = user['id']
        resp = requests.get (GRAPH_ENDPOINT + f"users/{user_id}/events?$filter=start/dateTime ge '{start_time}' and end/dateTime le '{end_time}'", headers=headers)
        while resp:
            #resp = requests.get (GRAPH_ENDPOINT + resp, headers=headers)
            if resp.status_code==200:
                data=resp.json()
                if "value" in data:
                    events_data = data["value"]
                    #events_table.delete (*events_table.get_children ( ))  # Clear existing data
                    for event in events_data:
                        body_preview = event.get ('bodyPreview', '')
                        meeting_id_match = re.search (r"Meeting ID: (.+)$", body_preview, re.MULTILINE)
                        meeting_id = meeting_id_match.group (1) if meeting_id_match else 'N/A'
                        events_table.insert ("", "end", values=(
                        event['subject'],
                        meeting_id,
                        event['webLink'],
                        event['start']['dateTime'],
                        event['end']['dateTime'],
                        event.get ('location', {}).get ('displayName', 'Not specified'),
                        event.get('organizer',{}).get('emailAddress','name'),
                        event['id']
                    ))
                next_page = data.get('@odata.nextLink')
                if next_page:
                    resp = requests.get (next_page, headers=headers)
                else:
                    break
            else:
                break

def show_event_details(event):
    selected_item = events_table.selection()
    if selected_item:
        values = events_table.item(selected_item, "values")
        event_details = f"Subject: {values[0]}\nMeeting ID: {values[1]}\n"\
                        f"Start Time: {values[3]}\nEnd Time: {values[4]}\nLocation: {values[5]}\n"\
                        f"Organizer: {values[6]}"
        messagebox.showinfo("Event Details", event_details)

def show_context_menu(event):
    selected_item = events_table.identify_row(event.y)
    if selected_item:
        context_menu = tk.Menu(root, tearoff=0)
        context_menu.add_command(label="View Details", command=lambda: show_event_details(selected_item))
        context_menu.add_command(label="Delete", command=lambda: delete_event(selected_item))
        context_menu.tk_popup(event.x_root, event.y_root)


root = tk.Tk ( )
root.title ("Calendar Event Viewer")

# Create labels, entry, and button for email input


email_entry = tk.Entry (root)
email_entry.pack ( )

fetch_button = Button (root, text="Fetch Events", command=fetch_calendar_events)
fetch_button.pack ( )

start_time_label = tk.Label(root, text="Start Time (UTC):")
start_time_label.pack()

start_time_entry = tk.Entry(root)  # Use Entry widget for time input
start_time_entry.pack()

end_time_label = tk.Label(root, text="End Time (UTC):")
end_time_label.pack()

end_time_entry = tk.Entry(root)  # Use Entry widget for time input
end_time_entry.pack()


events_table = Treeview(root, columns=("Subject", "Meeting ID", "Join Link", "Start Time", "End Time", "Location","organizer"),
                         show="headings")
#events_table.heading("Join", text="Join")
#events_table.column("Join", width=100, anchor="center")

events_table.heading("organizer",text="Organizer")
events_table.heading ("Subject", text="Subject")
events_table.heading ("Meeting ID", text="Meeting ID")
events_table.heading ("Join Link", text="Join Link")
events_table.heading ("Start Time", text="Start Time")
events_table.heading ("End Time", text="End Time")
events_table.heading ("Location", text="Location")
events_table.heading ("organizer", text="organizer")
events_table.pack ( )

connect_button = Button(root, text="Connect", command=connect_to_selected_meeting)
connect_button.pack()

delete_button = Button(root, text="Delete", command=lambda: delete_event(events_table.selection()))
delete_button.pack()

events_table.bind("<Double-1>", show_event_details)
events_table.bind("<Button-3>", show_context_menu)

root.mainloop ( )
