import pandas as pd
import tkinter as tk
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name(r'party-checker-2dd96bd51289.json', scope)

client = gspread.authorize(credentials)
sheet_url = 'https://docs.google.com/spreadsheets/d/1cuUC00cy8Lf42c9CaMw-MsraMXZEN8timPpGzlOnu5c/edit#gid=1499291877'
sheet = client.open_by_url(sheet_url)

worksheet = sheet.worksheet('Form Responses 1') 

rsvp = worksheet.get_all_values()

def update_row(row_to_edit):
    background_color = {"red": 0.0, "green": 0.5, "blue": 1.0}
    requests = [
    {
        "updateCells": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": row_to_edit - 1,
                "endRowIndex": row_to_edit,
            },
            "fields": "userEnteredFormat.backgroundColor",
            "rows": [{"values": [{"userEnteredFormat": {"backgroundColor": background_color}}]}],
        }
    }
    ]

    body = {"requests": requests}
    response = worksheet.spreadsheet.batch_update(body)


def search():
    
    found = False
    first_name = entry_first_name.get().lower()
    last_name = entry_last_name.get().lower()

    for index, row in enumerate(rsvp, start=1):
        firstName = row[1].lower()
        lastName=row[2].lower()

        if first_name == firstName and last_name == lastName:
            found = True
            update_row(index)
            break
    
    if found:
        submitted_label.config(text=f"{first_name} {last_name} has RSVP!", fg="green")
    else: 
        submitted_label.config(text=f"{first_name} {last_name} has not RSVP!", fg="red")

    entry_first_name.delete(0, tk.END)
    entry_last_name.delete(0, tk.END)

window = tk.Tk()
window.title("Sigma Nu Party checker")
window.geometry("1000x750")


label_first_name = tk.Label(window, text="First Name:")
label_first_name.pack(pady=10)

entry_first_name = tk.Entry(window, width=30)
entry_first_name.pack(pady=10)

label_last_name = tk.Label(window, text="Last Name:")
label_last_name.pack(pady=10)

entry_last_name = tk.Entry(window, width=30) 
entry_last_name.pack(pady=10)

submit_button = tk.Button(window, text="Submit", command=search)
submit_button.pack(pady=10)

submitted_label = tk.Label(window, text="", font=("Helvetica", 16))
submitted_label.pack()


window.mainloop()

search('eric', 'king')
