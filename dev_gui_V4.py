import os
import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog, Text
from tkinter.scrolledtext import ScrolledText
import tkinter.messagebox as messagebox


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

def create_df_cm(input_data):
    # Define header row
    header_row = ['*Last Name', '*First Name', 'Middle Name', 'Username', 'External Id'
                  , 'External Id Alias Pool', 'Name Full Formatted', 'Title', 'Suffix'
                  , 'Position', 'Begin Date+Time', '*End Date+Time', 'Physician Ind'
                  , 'SSN', 'SSN Pool', 'Birthdate', 'Sex', 'VIP', 'Active Ind'
                  , 'Primary Assigned Location', 'Email', '*Prsnl Alias Type', '*Prsnl Alias'
                  , '*Prsnl Alias Pool', 'Prsnl Alias Active Ind', 'Prsnl_Alias_End_Dt', '*Org Name'
                  , 'Org Confid Level', 'Org_End_Dt', '*Address Type', '*Address Type Seq'
                  , 'Address Street 1', 'Address Street 2', 'Address Street 3', 'Address Street 4'
                  , 'City', 'County', 'State or Prov', 'Country', 'Zip Code', 'Contact', 'Comment'
                  , 'District Health UK', 'Primary Care UK', 'Address_Delete_Ind'
                  , 'Org Address Reltn Ind', 'Org Addr Name', 'Org Addr Type', 'Org Addr Sequence'
                  , '*Phone Type', '*Phone Type Seq', 'Phone Number', 'Phone Extension'
                  , 'Phone Format', 'Phone Description', 'Phone Contact', 'Phone Call Instruction'
                  , 'Phone_Delete_Ind', 'Org Phone Reln Ind', 'Org Phone Name', 'Org Phone Type'
                  , 'Org Phone Seq', '*Location Type', '*Location Name', 'Location_Delete_Ind'
                  , '*Org Group Type', '*Org Group Name', 'Org_Group_Delete_Ind', '*Prsnl Group Type'
                  , '*Prsnl Group Name', '*Prsnl Group Class', 'Prsnl_Group_Delete_Ind'
                  , '*Clinical Service Display', 'Clinical Service Default'
                  , 'Clinical Service Org Name', 'Clin_Serv_Delete_Ind'
                ]

    end_date = "31/12/2100 00:00"
    alias_pool = 'External ID'
    today = datetime.date.today()
    d1 = today.strftime("%d/%m/%Y")
    begin_date = d1 + ' 00:00'
    act_ind = '1'
    org_g_type = 'SECURITY'
    org_g_name = 'Western Health'
    
    # iterate over input data
    data_dicts = []
    for index, row in input_data.iterrows():
        data_dict = {}
        a_user = row['USERNAME'].upper()
        a_credential = str(row['CREDENTIAL'])
        if a_credential == 'nan':
            a_credential = ''
        a_fname = row['FIRST']
        a_lname = row['LAST']
        a_fullname = a_lname + ', ' + a_fname + ' - ' + a_credential
        a_position = row['POSITION']
        a_extid = 'WHS' + a_user
        physician_ind = '0'
        if a_position in ['Medical Officer', 'Medical Officer P1', 'Medical Officer P2']:
            physician_ind = '1'
        data_dict.update({'Username': a_user, 'External Id': a_extid, 'External Id Alias Pool': alias_pool, '*Last Name': a_lname, '*First Name': a_fname, 'Name Full Formatted': a_fullname, 'Position': a_position, 'Begin Date+Time': begin_date, '*End Date+Time': end_date, 'Physician Ind': physician_ind, 'Active Ind': act_ind, '*Org Group Type': org_g_type, '*Org Group Name': org_g_name})
        data_dicts.append(data_dict)
    df_cm = pd.DataFrame(data_dicts, columns=header_row)  # create dataframe from the list of dictionaries
    return df_cm



def save_as_csv():
    global input_data_global, button_close
    if input_data_global is not None:
        df_cm = create_df_cm(input_data_global)
        
        # Create a DataFrame containing only the header row
        df_header = pd.DataFrame(columns=df_cm.columns)
        df_header.loc[0] = df_cm.columns

        # Concatenate the header DataFrame with the original DataFrame twice
        df_cm = pd.concat([df_header, df_cm]).reset_index(drop=True)
        
        filename = filedialog.asksaveasfilename(defaultextension='.csv')
        df_cm.to_csv(filename, index=False)

        # Show confirmation message
        messagebox.showinfo("Success", "File saved successfully")

        # Create Close button
        button_close = tk.Button(root, text="Close", command=root.destroy)
        button_close.grid(row=2, column=0, sticky='w', pady=10)
    else:
        print("No input data to save")


# Tkinter root window
root = tk.Tk()
root.state('zoomed')  # Maximize the window
#root.geometry("400x800")  # size of the window

# Button widget
button_open = tk.Button(root, text="Open Excel File")
button_open.grid(row=0, column=0, sticky='w', pady=10)

button_save = tk.Button(root, text="Save Content Manager file to Upload as CSV", command=save_as_csv)
button_save.grid(row=1, column=0, sticky='w', pady=10)

# Text widget for re-activation script
code_text = ScrolledText(root, wrap='word')  # Wrap text at WORD level
code_text.grid(row=0, column=1, sticky='nsew')  # Fill the grid cell

# Text widget for credential move script
code_cred_text = ScrolledText(root, wrap='word')  # Wrap text at WORD level
code_cred_text.grid(row=0, column=2, sticky='nsew')  # Fill the grid cell

root.grid_rowconfigure(0, weight=1)  # Row 0 expands with window
root.grid_columnconfigure(1, weight=1)  # Column 1 (code_text) expands with window
root.grid_columnconfigure(2, weight=1)  # Column 2 (code_cred_text) expands with window


# global variable to hold the last read data
input_data_global = None

def generate_code():
    global input_data_global
    filename = filedialog.askopenfilename(initialdir="/", title="Select File",
                                          filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
    if filename:  # if a file was selected
        input_data = pd.read_excel(filename, dtype='str' )
        input_data_global = input_data
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

button_open['command'] = generate_code

root.mainloop()