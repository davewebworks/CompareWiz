# Excel comparison script
# All IDs and other data in File1 will be matched and compaired to those in File2 and a report will be generated
# Order and structure of the two excel files may be different as long as both files will contain "ID" column ("ID" column names can differ in the two files)
# D.A. 6/27 v1
# Install:
# install python: https://www.python.org/downloads/windows/
# run "cmd", then:
# pip install pandas
# pip install xlrd
# pip install openpyxl
# run "IDLE" and from it open this script
# edit excel_diff_settings.json adjusting it to any comparison or report settings you prefer
# column NAMES have to be copied to excel_diff_settings.json EXACTLY, pay attention even to spaces
# In other fields spaces will be taken care of

# DON"T EDIT BELOW THIS POINT ----------------------------------------------------------------
# excel_diff_settings.json is the only user-edited file
# log file will be created (name also specified in excel_diff_settings.json, default name excel_log.txt)

import pandas as pd
import json
import datetime

#Reading two Excel Sheets

with open("excel_diff_settings.json") as json_file:
    data=json.load(json_file)

file1=data['excel1']['filename']
sheet1=data['excel1']['sheet_name']

file2=data['excel2']['filename']
sheet2=data['excel2']['sheet_name']


df1 = pd.read_excel(file1,sheet_name=sheet1, keep_default_na=False)
df2 = pd.read_excel(file2,sheet_name=sheet2, keep_default_na=False)
df1=df1.dropna()
df2=df2.dropna()
#print("DF1: ", df1)
#print("DF2: ", df2)

FieldsToMatchDict= data['excel1']["FieldsToMatch"]
FieldsToMatchKeys=list(FieldsToMatchDict.keys())
FieldsToMatchKeys.append('Student ID')
now = datetime.datetime.now()

# log dictionary (json format, to be saved in the log file)
log = {
    "Description":"All IDs and other data in File1 will be matched and compaired to those in File2",
    "FieldsToMatch": FieldsToMatchKeys,
    "File1": file1,
    "Sheet1": sheet1,
    "File2": file1,
    "Sheet2": sheet1,
    "log created": now.strftime("%Y-%m-%d %H:%M:%S"),
    "version": "v1",
    "mismatches": 0,
    "mismatched_entries": []
    }

print("All IDs and other data in File1 will be matched and compaired to those in File2")
print("Settings file: excel_diff_settings.json")
print("FieldsToMatch :", FieldsToMatchKeys)
print("File1 :", file1)
print("Sheet1 :", sheet1)
print("File2 :", file2)
print("Sheet2 :", sheet2)
print("log created :", now.strftime("%Y-%m-%d %H:%M:%S"))

    

table1_multientries_duplicates=[]
table2_multientries_duplicates=[]

table1_othertableissues_duplicates=[]
table2_othertableissues_duplicates=[]

mismatches=0
# ID column in 
file1_id_col=data['excel1']["ID_field"]
file2_id_col=data['excel2']["ID_field"]

#print("ID column in file1: ", file1_id_col)

for i in df1[file1_id_col]: # over each 'ID' column cell in file 1
    # print(i)
    #if i == '' or i== nan :
    #    continue
    #   print("never gets here")
    id_matches_num= len(df1[df1[file1_id_col] == i].index)
    if(id_matches_num > 1) and (i not in table1_multientries_duplicates):
      table1_multientries_duplicates.append(int(i))  
      message= 'ID '+str(int(i))+' found ['+str(id_matches_num)+'] times in table1'
      print(message)
      mismatches += 1
      new_mismatch={}
      new_mismatch['Mismatch #'] = mismatches
      new_mismatch['ID'] = int(i)
      new_mismatch['error_type'] = message
      log["mismatched_entries"].append(new_mismatch)

    # current row in file 1
    row1=df1[df1[file1_id_col] == i].index[0]

    #print("Searching student with ID ",i," from  file 1 in file 2") # Student ID from file1


 
    #print("Search the row containing ID ",i," in 2nd file")
    id_matches_num= len(df2[df2[file2_id_col] == i].index)
    if(id_matches_num > 1) and (i not in table1_othertableissues_duplicates):
      table1_othertableissues_duplicates.append(int(i))

      message= 'ID '+str(int(i))+' from table1 found ['+str(id_matches_num)+'] times in table2'
      print(message)
      mismatches += 1
      new_mismatch={}
      new_mismatch['Mismatch #'] = mismatches
      new_mismatch['ID'] = int(i)
      new_mismatch['error_type'] = message
      log["mismatched_entries"].append(new_mismatch)
      
    elif (id_matches_num == 0) and (i not in table1_othertableissues_duplicates):
      table1_othertableissues_duplicates.append(int(i))  
      message= 'ID '+str(int(i))+' from table1 NOT found in table2'
      print(message)
      mismatches += 1
      new_mismatch={}
      new_mismatch['Mismatch #'] = mismatches
      new_mismatch['ID'] = int(i)
      new_mismatch['error_type'] = message
      log["mismatched_entries"].append(new_mismatch)

    elif (id_matches_num == 0):
        continue

    else:

     # row in file 2 with that ID   
     row2=df2[df2[file2_id_col] == i].index[0]
    
     #print("Found row with ID ",i, ": ",row2)
     #print("Print the found row in file2: ", df2.iloc[row2,:])

     # In the found row in file2 match all the fields under "FieldsToMatch" 
     for key in data['excel1']["FieldsToMatch"]:
         # column names in file1 and file2 with the field to match
         row1_item_column_name = data['excel1']["FieldsToMatch"][key]
         row2_item_column_name = data['excel2']["FieldsToMatch"][key]
         #print("field in 2nd file",row2_item_column_name)
         #print("field in 1st file",row1_item_column_name)
         cell1=df1.iloc[row1][row1_item_column_name]
         cell2=df2.iloc[row2][row2_item_column_name]
         #print("Cell in file1: ",cell1)
         #print("Cell in file2: ",cell2)
         if (str(cell1).strip() != str(cell2).strip()):
            message='Element '+row1_item_column_name+ ' mismatch: '+str(cell1)+' vs '+str(cell2)+ ' for ID: '+str(int(i))
            print(message)
            mismatches += 1
            new_mismatch={}
            new_mismatch['Mismatch #'] = mismatches
            new_mismatch['ID'] = int(i)
            new_mismatch['error_type'] = message
            log["mismatched_entries"].append(new_mismatch)
            
         #else:
         #   print("Element ",row1_item_column_name, " matches in both files")


for i in df2[file2_id_col]: # over each 'ID' column cell in file 2
    row2=df2[df2[file2_id_col] == i].index[0]
    row2_item_column_name = data['excel2']["FieldsToMatch"]["LastName_col"]
    cell2=df2.iloc[row2][row2_item_column_name]
    #print("Search the row containing ID ",i," in 2nd file")
    id_matches_num= len(df1[df1[file1_id_col] == i].index)
    if(id_matches_num == 0 and (data['reports']["report_ID_infile2_but_not_infile1"] == 'yes')  and (i not in table2_othertableissues_duplicates) ):
      table2_othertableissues_duplicates.append(int(i))
      message='ID '+str(int(i))+' from table2 NOT found in table1'
      print(message)
      mismatches += 1
      new_mismatch={}
      new_mismatch['Mismatch #'] = mismatches
      new_mismatch['ID'] = int(i)
      new_mismatch['error_type'] = message
      log["mismatched_entries"].append(new_mismatch)


log['mismatches']=mismatches

            
filename = data["reports"]["logfile_name"]         

json_object = json.dumps(log, indent = 4)
with open(filename, "w") as outfile:
    outfile.write(json_object)


