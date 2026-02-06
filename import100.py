import json
import os
import pyodbc
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import chardet
import urllib
from sqlalchemy import create_engine, types as sqltypes
import logging
import threading
import re

# Configure logging
logging.basicConfig(filename='import_logfile.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Import the config.json file
try:
    with open('config.json') as f:
        config = json.load(f)
except FileNotFoundError:
    error_message = "Error: The file 'config.json' does not exist. Please make sure the file is in the correct location."
    print(error_message)
    root = tk.Tk()
    root.withdraw()
    tk.messagebox.showerror("File Not Found", error_message)
    exit()
except json.JSONDecodeError:
    error_message = "Error: The file 'config.json' is not properly formatted. Please check the file for JSON syntax errors."
    print(error_message)
    root = tk.Tk()
    root.withdraw()
    tk.messagebox.showerror("Invalid JSON Format", error_message)
    exit()

# Create the GUI window
window = tk.Tk()
window.title("Syscon Import Wizard")
window.geometry("400x500")
window.resizable(True, True)

# Add padding for all widgets
padding = {'padx': 10, 'pady': 10}

# Configure dynamic sizing
window.grid_columnconfigure(0, weight=0)
window.grid_columnconfigure(1, weight=1)

# Server Selection Section
server_label_header = tk.Label(window, text="Server Selection", font=('Helvetica', 11, 'bold'))
server_label_header.grid(row=0, column=0, columnspan=2, **padding)

server_label = tk.Label(window, text="Server:")
server_label.grid(row=1, column=0, sticky='e', **padding)
server_combobox = ttk.Combobox(window, values=config["Server"], state='normal')
server_combobox.grid(row=1, column=1, sticky='w', **padding)
server_combobox.set("Select or Enter Server")

variant_label = tk.Label(window, text="Variant:")
variant_label.grid(row=2, column=0, sticky='e', **padding)
variant_values = [property["Variant"] for property in config["Properties"]]
variant_combobox = ttk.Combobox(window, values=variant_values)
variant_combobox.grid(row=2, column=1, sticky='w', **padding)
variant_combobox.set(variant_values[0])

# Login Method Selection via Combobox
login_label_header = tk.Label(window, text="Login Method", font=('Helvetica', 11, 'bold'))
login_label_header.grid(row=3, column=0, columnspan=2, **padding)

login_method = tk.StringVar()
login_combobox = ttk.Combobox(window, textvariable=login_method, state='readonly')
login_options = {
    "Windows": "windows",
    "SQL Login": "sql",
    "Entra": "entra"
}
login_combobox['values'] = list(login_options.keys())
login_combobox.grid(row=4, column=1, sticky='w', **padding)
login_combobox.set("Windows")

login_label = tk.Label(window, text="Login Type:")
login_label.grid(row=4, column=0, sticky='e', **padding)

# Login Credentials Section
user_label = tk.Label(window, text="Username:")
user_label.grid(row=5, column=0, sticky='e', **padding)
user_entry = tk.Entry(window, state='disabled')
user_entry.grid(row=5, column=1, sticky='w', **padding)

password_label = tk.Label(window, text="Password:")
password_label.grid(row=6, column=0, sticky='e', **padding)
password_entry = tk.Entry(window, show="*", state='disabled')
password_entry.grid(row=6, column=1, sticky='w', **padding)

# Enable/Disable Credentials Based on Login Type

def update_login_fields(event=None):
    selected_display = login_method.get()
    method = login_options.get(selected_display, "windows")
    if method == "windows":
        user_entry.configure(state='disabled')
        password_entry.configure(state='disabled')
    elif method == "sql":
        user_entry.configure(state='normal')
        password_entry.configure(state='normal')
    elif method == "entra":
        user_entry.configure(state='normal')
        password_entry.configure(state='disabled')

login_combobox.bind("<<ComboboxSelected>>", update_login_fields)

# Initial login field state (moved below function definition)
update_login_fields()

# Execute button and message
execute_button = tk.Button(window, text="Import File", command=lambda: threading.Thread(target=execute).start())
execute_button.grid(row=7, column=1, columnspan=1, sticky='w', **padding)

# Checkbox for executing stored procedures
execute_sp_var = tk.BooleanVar(value=True)
execute_sp_checkbox = tk.Checkbutton(window, text="Execute SP",variable=execute_sp_var)
execute_sp_checkbox.grid(row=7, column=0, sticky='w', columnspan=1, padx=10)

message = tk.Label(window,wraplength=250, justify="left")
message.grid(row=8, column=1, columnspan=1, sticky='w')

# # Progress Bar
# progress = ttk.Progressbar(window, orient='horizontal', mode='determinate')
# progress.grid(row=8, column=0, columnspan=2, **padding)

def execute():
    message.config(text='')

    selected_server = server_combobox.get()
    selected_variant = variant_combobox.get()
    if not selected_variant:
        message.config(text="Please select a variant.", fg='red')
        return
    try:
        selected_property = next(property for property in config["Properties"] if property["Variant"] == selected_variant)
    except StopIteration:
        message.config(text=f"Variant '{selected_variant}' not found in config.", fg='red')
        return

    server = selected_server
    database = selected_property["database"]
    schema = selected_property["schema"]
    table = selected_property["table"]
    sp = selected_property["SP"]

    username = user_entry.get().strip()
    password = password_entry.get().strip()
    login_type = login_options.get(login_method.get(), "windows")

    file_path = filedialog.askopenfilename()
    if not file_path:
        message.config(text="No file selected", fg='red')
        return

    file_extension = os.path.splitext(file_path)[1]
    sheet_config = selected_property.get("sheet", 0)
    sheet_name_or_index = sheet_config if isinstance(sheet_config, str) else int(sheet_config) - 1

    try:
        if login_type == "windows":
            conn_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;'
        elif login_type == "sql":
            conn_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        elif login_type == "entra":
            conn_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};Authentication=ActiveDirectoryInteractive;UID={username}'
        else:
            raise ValueError("Unknown login method")

        conn = pyodbc.connect(conn_string)
        params = urllib.parse.quote_plus(conn_string)
        engine = create_engine(f"mssql+pyodbc:///?odbc_connect=%s" % params)

        if file_extension == '.csv':
            with open(file_path, 'rb') as f:
                result = chardet.detect(f.read())
            encoding = result['encoding']
            with open(file_path, encoding=encoding) as f:
                first_line = f.readline()
            if ';' in first_line:
                delimiter = ';'
            elif ',' in first_line:
                delimiter = ','
            else:
                delimiter = '\t'
            data = pd.read_csv(file_path, sep=delimiter, encoding=encoding, dtype=str)
        elif file_extension in ['.xlsx', '.xls', '.xlsm']:
            data = pd.read_excel(file_path, sheet_name=sheet_name_or_index, dtype=str)
        else:
            raise ValueError('Unsupported file format')

        message.config(text='processing...', fg='green')

        def clean_number(val):
            if not isinstance(val, str):
                return val
            val_stripped = val.strip().replace("'", "")
            if re.match(r'^\d{1,3}(\.\d{3})*,\d{1,2}$', val_stripped):
                return val_stripped.replace('.', '').replace(',', '.')
            elif re.match(r'^\d{1,3}(,\d{3})*\.\d{1,2}$', val_stripped):
                return val_stripped.replace(',', '')
            return val

        data = data.applymap(clean_number)
        dtype_dict = {col: sqltypes.VARCHAR(4000) for col in data.columns}
        data.to_sql(table, engine, schema=schema, if_exists='replace', index=False, dtype=dtype_dict)

        cursor = conn.cursor()
        if execute_sp_var.get() and sp:
            cursor.execute('EXEC ' + sp)

        conn.commit()
        conn.close()

        success_message = 'Success'
        message.config(text=success_message, fg='green')
        logging.info(success_message)
    except Exception as e:
        error_message = str(e)
        message.config(text=error_message, fg='red')
        logging.error(error_message)


window.mainloop()
