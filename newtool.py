# Username is used for finding your folder eg. c/users/whittlj2
username = 'whittlj2'

# Path for input data
path = 'input.xlsx'

# Output filetype for CCL and authview code
filetype = '.txt'

#import libraries
import os
import datetime
import pandas as pd

# Create Activate Code file
#region
outfilp = r'C:\Users\\'
outfilp = outfilp + username + '\\'
outputfilename = '-ACTIVATE-CODE'
datetime_str = str(datetime.datetime.now())
datetime_str = datetime_str.replace('.', '_')
datetime_str = datetime_str.replace(':', '-')
outputfilename = outfilp + datetime_str + outputfilename + filetype
outputfilename = str(outputfilename)


# READ EXCEL FILE
input_data = pd.read_excel(path, dtype= 'str' )

# WRITE CODE TO TXT FILE
for index, row in input_data.iterrows():
    # Column that has the usernames to put in the code
    to_switch = row['USERNAME'].upper()
    #Write to file
    for a_row in ccl_code:
        # REPLACE SWAPME123 with the username in each row of the code slab
        new_row = a_row.replace('SWAPME123', to_switch)
        f = open(outputfilename, "a")
        f.write(new_row + '\n')
f.close()

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
df = pd.read_excel(r'INPUT_TEMPLATES\Personnel_template.xlsx', header=[1], dtype= 'str')
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

row_number = 0
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
    df.loc[row_number,'Username'] = a_user
    df.loc[row_number,'External Id'] = a_extid
    df.loc[row_number,'External Id Alias Pool'] = alias_pool
    df.loc[row_number,'*Last Name'] = a_lname
    df.loc[row_number,'*First Name'] = a_fname
    df.loc[row_number,'Name Full Formatted'] = a_fullname
    df.loc[row_number,'Position'] = a_position
    df.loc[row_number,'Begin Date+Time'] = begin_date
    df.loc[row_number,'*End Date+Time'] = end_date
    df.loc[row_number,'Physician Ind'] = physician_ind
    df.loc[row_number,'Active Ind'] = act_ind
    df.loc[row_number,'*Org Group Type'] = org_g_type
    df.loc[row_number,'*Org Group Name'] = org_g_name
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
df.loc[-1] = df.columns.values
df.index = df.index + 1  # shifting index
df = df.sort_index()  # sorting by index
df.to_csv(outputfilename, index=False)
#endregion