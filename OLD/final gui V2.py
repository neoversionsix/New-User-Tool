import os
import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog, Text
from tkinter.scrolledtext import ScrolledText


# CCL Code
ccl_code = [
    '; USER RE-ACTIVATION SCRIPT',
    'update into prsnl p',
    'set p.end_effective_dt_tm = cnvtdatetime("31-DEC-2100")',
    ', p.updt_dt_tm = cnvtdatetime(curdate,curtime3)',
    ', p.updt_id = reqinfo->updt_id',
    ', p.updt_cnt = p.updt_cnt + 1',
    'where p.username = "SWAPME123"',
    ''
]

# CCL Credential Code
ccl_cred_code = [
    '; CREDENTIAL MOVE SCRIPT ------------------------------------------'
    , 'update into credential cred'
    , 'set cred.prsnl_id = (select person_id from prsnl where username = "SWAPME123")'
    , ', cred.credential_cd = (select code_value from code_value where code_set = 29600 and display = "SWAP_TO_REAL_CREDENTIAL")'
    , ', cred.credential_type_cd = 686580 ; License from code set 254874'
    , ', cred.beg_effective_dt_tm = cnvtdatetime(curdate,curtime3)'
    , ', cred.active_ind = 1'
    , ', cred.active_status_dt_tm = cnvtdatetime(curdate,curtime3)'
    , ', cred.active_status_cd = 188 ; Active from code set 48'
    , ', cred.updt_dt_tm = cnvtdatetime(curdate,curtime3)'
    , ', cred.updt_id = reqinfo->updt_id'
    , ', cred.updt_cnt = cred.updt_cnt + 1'
    , 'where cred.credential_id  = ('
    , 'select min(credential_id)'
    , 'from credential'
    , 'where prsnl_id = 13876656 ; Credential Box user in prod or cert'
    , ')'
    , 'and not exists ('
    , 'select 1'
    , 'from credential'
    , 'where prsnl_id = (select person_id from prsnl where username = "SWAPME123")'
    , 'and credential_cd = (select code_value from code_value where code_set = 29600 and display = "SWAP_TO_REAL_CREDENTIAL")'
    , 'and active_ind = 1'
    , ')'
    , ''
    , '; DIRECTORY IND SCRIPT'
    , 'update into ea_user eau'
    , 'set'
    , 'eau.directory_ind = 1'
    , ',eau.updt_dt_tm = cnvtdatetime(curdate,curtime3)'
    , ',eau.updt_id = reqinfo->updt_id'
    , ',eau.updt_cnt = eau.updt_cnt + 1'
    , 'where eau.username = "SWAPME123"'
    , ';---------------------------------------------------------------------'
]

# Tkinter root window
root = tk.Tk()
root.geometry("400x800")  # size of the window

# Button widget
button = tk.Button(root, text="Open Excel File")
button.grid(row=0, column=0, sticky='w', pady=10)

# Text widget for re-activation script
code_text = ScrolledText(root, wrap='word')  # Wrap text at WORD level
code_text.grid(row=0, column=1, sticky='nsew')  # Fill the grid cell

# Text widget for credential move script
code_cred_text = ScrolledText(root, wrap='word')  # Wrap text at WORD level
code_cred_text.grid(row=0, column=2, sticky='nsew')  # Fill the grid cell

root.grid_rowconfigure(0, weight=1)  # Row 0 expands with window
root.grid_columnconfigure(1, weight=1)  # Column 1 (code_text) expands with window
root.grid_columnconfigure(2, weight=1)  # Column 2 (code_cred_text) expands with window


def generate_code():
    filename = filedialog.askopenfilename(initialdir="/", title="Select File",
                                          filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
    if filename:  # if a file was selected
        input_data = pd.read_excel(filename, dtype='str' )
        code_text.delete('1.0', tk.END)  # Clear text widget for re-activation script
        code_cred_text.delete('1.0', tk.END)  # Clear text widget for credential move script
        for index, row in input_data.iterrows():
            # Columns that has the usernames and credentials to put in the code
            to_switch_1 = row['USERNAME'].upper()
            to_switch_2 = str(row['CREDENTIAL'])

            # Generate re-activation code
            for a_row in ccl_code:
                new_row = a_row.replace('SWAPME123', to_switch_1)
                code_text.insert(tk.END, new_row + '\n')

            # Only generate credential move code where a credential is filled out
            if to_switch_2 != 'nan':
                # Generate credential move code
                for a_row in ccl_cred_code:
                    new_row = a_row.replace('SWAPME123', to_switch_1)
                    new_row_2 = new_row.replace('SWAP_TO_REAL_CREDENTIAL', to_switch_2)
                    code_cred_text.insert(tk.END, new_row_2 + '\n')

button['command'] = generate_code

root.mainloop()
