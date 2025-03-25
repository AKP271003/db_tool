import tkinter as tk
from tkinter import filedialog
import requests

def select_file():
    file_path = filedialog.askopenfilename(title="Select SQL Queries File", filetypes=[("Text Files", "*.txt")])
    entry_file.delete(0, tk.END)
    entry_file.insert(tk.END, file_path)

def execute_queries():
    file_path = entry_file.get()
    case_number = entry_case_number.get()
    if file_path and case_number:
        with open(file_path, 'rb') as file:
            response = requests.post('http://localhost:5000/execute_queries', files={'file': file}, data={'case_number': case_number})
            if response.status_code == 200:
                save_path = filedialog.asksaveasfilename(title="Save Results", defaultextension=".zip", filetypes=[("Zip Files", "*.zip")])
                if save_path:
                    with open(save_path, 'wb') as f:
                        f.write(response.content)
                    label_status.config(text="Results saved successfully!", fg="green")
                else:
                    label_status.config(text="Save operation cancelled.", fg="blue")
            else:
                label_status.config(text="Failed to download results.", fg="red")
    else:
        label_status.config(text="Please select a file and enter case number.", fg="red")

root = tk.Tk()
root.title("SQL Query Executor")

label_file = tk.Label(root, text="SQL Queries File:")
label_file.grid(row=0, column=0, padx=5, pady=5)

entry_file = tk.Entry(root, width=50)
entry_file.grid(row=0, column=1, padx=5, pady=5)

button_browse = tk.Button(root, text="Browse", command=select_file)
button_browse.grid(row=0, column=2, padx=5, pady=5)

label_case_number = tk.Label(root, text="Case Number:")
label_case_number.grid(row=1, column=0, padx=5, pady=5)

entry_case_number = tk.Entry(root)
entry_case_number.grid(row=1, column=1, padx=5, pady=5)

button_execute = tk.Button(root, text="Execute Queries", command=execute_queries)
button_execute.grid(row=2, column=1, padx=5, pady=5)

label_status = tk.Label(root, text="")
label_status.grid(row=3, column=1, padx=5, pady=5)

root.mainloop()
