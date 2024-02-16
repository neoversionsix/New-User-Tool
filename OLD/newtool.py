#import libraries
import os
import datetime
import pandas as pd



# READ EXCEL FILE
input_data = pd.read_excel(input.xlsx, dtype= 'str' )

# creat
for index, row in input_data.iterrows():
    # Column that has the usernames to put in the code
    to_switch = row['USERNAME'].upper()
    #Write to file
    for a_row in ccl_activate_code:
        # REPLACE SWAPME123 with the username in each row of the code slab
        new_row = a_row.replace('SWAPME123', to_switch)
        f = open(outputfilename, "a")
        f.write(new_row + '\n')
f.close()

# CCL Code
ccl_activate_code = [
'; USER RE-ACTIVATION SCRIPT',
'update into prsnl p',
'set p.end_effective_dt_tm = cnvtdatetime("31-DEC-2100")',
', p.updt_dt_tm = cnvtdatetime(curdate,curtime3)',
', p.updt_id = reqinfo->updt_id',
', p.updt_cnt = p.updt_cnt + 1',
'where p.username = "SWAPME123"',
''
]


#endregion

# Create Credential Code file
outfilp = r'C:\Users\\'
outfilp = outfilp + username + '\\'
outputfilename = '-CREDENTIAL-CODE'
datetime_str = str(datetime.datetime.now())
datetime_str = datetime_str.replace('.', '_')
datetime_str = datetime_str.replace(':', '-')
outputfilename = outfilp + datetime_str + outputfilename + filetype
outputfilename = str(outputfilename)

# CCL Code for Credentials and AD/LDAP Sync
#region
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

# READ CSV FILE
input_data = pd.read_excel(path, sheet_name= 'DATA', dtype= 'str' )

# WRITE CODE TO TXT FILE
for index, row in input_data.iterrows():
    # Column that has the usernames to put in the code
    to_switch_1 = str(row['USERNAME'].upper())
    to_switch_2 = str(row['CREDENTIAL'])
    #Write to file
    for a_row in ccl_cred_code:
        # Only generate code where a credential is filled out
        if to_switch_2 != 'nan':
            # REPLACE SWAPME123 with the username in each row of the code slab
            new_row = a_row.replace('SWAPME123', to_switch_1)
            new_row_2 = new_row.replace('SWAP_TO_REAL_CREDENTIAL', to_switch_2)
            f = open(outputfilename, "a")
            f.write(new_row_2 + '\n')
f.close()
#endregion


# Create csv for uploading to content manager
#region
# df = pd.read_excel(r'INPUT_TEMPLATES\Personnel_template.xlsx', header=[1], dtype= 'str')
header_row = ['*Last Name', '*First Name', 'Middle Name', 'Username', 'External Id', 'External Id Alias Pool', 'Name Full Formatted', 'Title', 'Suffix', 'Position', 'Begin Date+Time', '*End Date+Time', 'Physician Ind', 'SSN', 'SSN Pool', 'Birthdate', 'Sex', 'VIP', 'Active Ind', 'Primary Assigned Location', 'Email', '*Prsnl Alias Type', '*Prsnl Alias', '*Prsnl Alias Pool', 'Prsnl Alias Active Ind', 'Prsnl_Alias_End_Dt', '*Org Name', 'Org Confid Level', 'Org_End_Dt', '*Address Type', '*Address Type Seq', 'Address Street 1', 'Address Street 2', 'Address Street 3', 'Address Street 4', 'City', 'County', 'State or Prov', 'Country', 'Zip Code', 'Contact', 'Comment', 'District Health UK', 'Primary Care UK', 'Address_Delete_Ind', 'Org Address Reltn Ind', 'Org Addr Name', 'Org Addr Type', 'Org Addr Sequence', '*Phone Type', '*Phone Type Seq', 'Phone Number', 'Phone Extension', 'Phone Format', 'Phone Description', 'Phone Contact', 'Phone Call Instruction', 'Phone_Delete_Ind', 'Org Phone Reln Ind', 'Org Phone Name', 'Org Phone Type', 'Org Phone Seq', '*Location Type', '*Location Name', 'Location_Delete_Ind', '*Org Group Type', '*Org Group Name', 'Org_Group_Delete_Ind', '*Prsnl Group Type', '*Prsnl Group Name', '*Prsnl Group Class', 'Prsnl_Group_Delete_Ind', '*Clinical Service Display', 'Clinical Service Default', 'Clinical Service Org Name', 'Clin_Serv_Delete_Ind']
df_cm = pd.DataFrame(data=[header_row], columns=header_row)

input_data = pd.read_excel(path, sheet_name= 'DATA', dtype= 'str' )

end_date = r"30/12/2100"
alias_pool = 'External ID'
today = datetime.date.today()
d1 = today.strftime("%d/%m/%Y")
begin_date = d1 + ' 00:00'
end_date = '31/12/2100 00:00'
act_ind = '1'
org_g_type = 'SECURITY'
org_g_name = 'Western Health'


#Content Manager Dataframe
row_number = 1
for index, row in input_data.iterrows():
    # Column that has the usernames to put in the code
    a_user = row['USERNAME'].upper()
    a_credential = str(row['CREDENTIAL'])
    # Set credential to blank rather than 'nan' if it's not filled out
    if a_credential == 'nan':
        a_credential = ''
    a_fname = row['FIRST']
    a_lname = row['LAST']
    a_fullname = a_lname + ', ' + a_fname + ' - ' + a_credential
    a_position = row['POSITION']
    a_extid = 'WHS' + a_user
    physician_ind = '0'
    if a_position == 'Medical Officer':
        physician_ind = '1'
    if a_position == 'Medical Officer P1':
        physician_ind = '1'
    if a_position == 'Medical Officer P2':
        physician_ind = '1'
    # Edit Sheet
    df_cm.loc[row_number,'Username'] = a_user
    df_cm.loc[row_number,'External Id'] = a_extid
    df_cm.loc[row_number,'External Id Alias Pool'] = alias_pool
    df_cm.loc[row_number,'*Last Name'] = a_lname
    df_cm.loc[row_number,'*First Name'] = a_fname
    df_cm.loc[row_number,'Name Full Formatted'] = a_fullname
    df_cm.loc[row_number,'Position'] = a_position
    df_cm.loc[row_number,'Begin Date+Time'] = begin_date
    df_cm.loc[row_number,'*End Date+Time'] = end_date
    df_cm.loc[row_number,'Physician Ind'] = physician_ind
    df_cm.loc[row_number,'Active Ind'] = act_ind
    df_cm.loc[row_number,'*Org Group Type'] = org_g_type
    df_cm.loc[row_number,'*Org Group Name'] = org_g_name
    row_number+=1

outfilp = r'C:\Users\\'
outfilp = outfilp + username + '\\'
outputfilename = '-CONTENTMGR'
datetime_str = str(datetime.datetime.now())
datetime_str = datetime_str.replace('.', '_')
datetime_str = datetime_str.replace(':', '-')
outputfilename = outfilp + datetime_str + outputfilename + '.csv'
outputfilename = str(outputfilename)
# Add a blank row
df_cm.to_csv(outputfilename, index=False)
#endregion