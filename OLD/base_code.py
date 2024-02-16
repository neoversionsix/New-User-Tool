#import libraries
import os
import datetime
import pandas as pd

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

# READ CSV FILE
path = "input.xlsx"
input_data = pd.read_excel(path, dtype= 'str' )

# Create Activate Code file
outputfilename = "ccl_activate_code.txt"

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