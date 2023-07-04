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

# Tkinter root window
root = tk.Tk()
root.geometry("800x600")  # size of the window

# Text widget
code_text = ScrolledText(root, wrap='word')  # Wrap text at WORD level
code_text.pack(fill='both', expand=True)  # Fill the entire window

def generate_code():
    filename = filedialog.askopenfilename(initialdir="/", title="Select File",
                                          filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
    if filename:  # if a file was selected
        input_data = pd.read_excel(filename, dtype='str' )
        code_text.delete('1.0', tk.END)  # Clear text widget
        for index, row in input_data.iterrows():
            # Column that has the usernames to put in the code
            to_switch = row['USERNAME'].upper()
            # Generate code
            for a_row in ccl_code:
                # REPLACE SWAPME123 with the username in each row of the code slab
                new_row = a_row.replace('SWAPME123', to_switch)
                code_text.insert(tk.END, new_row + '\n')

# Button widget
button = tk.Button(root, text="Open Excel File", command=generate_code)
button.pack(pady=10)

root.mainloop()
