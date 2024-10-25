#   The purpose of this code is take information from 
#   Intune, ServiceDesk+, SpaceIQ, and Ninja and merge it into 
#   one and make a report out of the Excel file that will be 
#   created
#
#   Inputs:
#   Sheet_Name: a string that contains the file name of a 
#   .xlsx file that has the four sheets of information 

# pandas is a package that contains tools for Excel file
# processing, it's used extensively throughout this code
import pandas as pd

Sheet_Name = str(input("Enter the name of the Excel file: ")) #get the file name from the user

# In this next block of code, the excel file gets processed
# using pandas and then each sheet gets assigned to a variable
Inventory_Spreadsheet = pd.ExcelFile(Sheet_Name) 
SD_info = Inventory_Spreadsheet.parse("Sheet1")
SpaceIQ_info = Inventory_Spreadsheet.parse("Sheet2")
Intune_info = Inventory_Spreadsheet.parse("Sheet3")
Ninja_info = Inventory_Spreadsheet.parse("Sheet4")
#PowerBI_info = Inventory_Spreadsheet.parse("Sheet5")

SpaceIQ = [] # this list is a list of lists that stores the SpaceIQ information of each user

# This for loop goes through every row in the SpaceIQ spreadsheet
# and puts the user's email and Space code into a list and then adds
# it to the SpaceIQ
for index, row in SpaceIQ_info.iterrows():
    if type(row["Employee Email"]) == str:
        SpaceIQ.append([row["Space Code"], row["Employee Email"]])
        
# This nested for loop is responsible for creating two columns in the 
# final spreadsheet, the SpaceIQ column and the SD-SpaceIQ column as well
new_SpaceIQ = []
SD_SpaceIQ = []

for index, row in SD_info.iterrows():
    temp = ""
    for i in range(len(SpaceIQ)):
        if row["User Email"].lower() == SpaceIQ[i][1].lower():
            temp = SpaceIQ[i][0]
            
    if temp != "":
        new_SpaceIQ.append(temp)
        if row["Location"] == "Remote Workstation":
            SD_SpaceIQ.append("N/A")
        elif temp != row["Location"]:
            SD_SpaceIQ.append("n")
            #print(row["Location"])
            #print(temp)
        else:
            SD_SpaceIQ.append("y")
    else:
        new_SpaceIQ.append("")
        if row["State"] == "Disposed" or row["State"] == "In Repair" or row["State"] == "In Store":
            SD_SpaceIQ.append("N/A")
        else:
            SD_SpaceIQ.append("N/A")

# This for loop goes through each row in the Intune spreadsheet and adds the
# Intune device name and the user's email address
Intune = []
for index, row in Intune_info.iterrows():
    Intune.append([row["Device name"], row["Primary user email address"]])

# This nested for loop goes through each workstation and matches the Intune
# device name and will eventually add the Intune email as a seperate column
Intune_Emails = []
for index, row in SD_info.iterrows():
    temp = ""
    for i in range(len(Intune)):
        if row["Workstation"] == Intune[i][0]:
            temp = Intune[i][1]
    
    if temp == "":
        Intune_Emails.append("")
    else:
        Intune_Emails.append(temp)

# This for loop gathers the information from the Ninja spreadsheet and saves it
# as a list      
Ninja = []
for index, row in Ninja_info.iterrows():
    temp = ""
    if type(row["Last LoggedIn User"]) == str:
        temp = row["Last LoggedIn User"]
        temp = temp[8:]
        length_of_temp = len(temp)
        if temp[length_of_temp-2:] == "-c" or temp[length_of_temp-2:] == "-C":
            temp = temp[:length_of_temp-2]
        #print(temp)
        #print(temp1[length_of_temp1-2:])
    #print(temp1)
    Ninja.append([row["SystemName"], temp])

# This nested for loop adds the Ninja information into a list for the final Excel sheet
Ninja_Col = []
for index, row in SD_info.iterrows():
    temp = ""
    temp_Email = ""
    if type(row["User Email"]) == str:
        temp_Email = row["User Email"]
        temp_Email = temp_Email.replace("@cgf.com", "")
    for i in range(len(Ninja)):
        if row["Workstation"] == Ninja[i][0]:
            temp = Ninja[i][1]
    
    Ninja_Col.append(temp)

# This adds a ticket number column in the final Excel sheet 
Ticket_Number = []
for index, row in SD_info.iterrows():
    Ticket_Number.append("")

# This block of code offically adds new columns in the final Excel sheet by using 
# lists that were made   
SD_info["Intune-User-Email"] = Intune_Emails
SD_info["SpaceIQ-Seat"] = new_SpaceIQ
SD_info["Ninja Last LoggedIn User"] = Ninja_Col
SD_info["SD-SpaceIQ Location Correct?"] = SD_SpaceIQ

# This implements the logic that the User's computer in SD and the last logged in user
# on Ninja are the same
Emails_Check = [] 
for index, row in SD_info.iterrows():
    temp_Email = row["User Email"]
    if "@" in temp_Email:
        temp_Email = temp_Email.replace("@cgf.com", "")
    if temp_Email == row["Ninja Last LoggedIn User"].lower():
        Emails_Check.append("y")
    elif row["Ninja Last LoggedIn User"] == "":
        Emails_Check.append("N/A")
    else:
        Emails_Check.append("n")

SD_info["SD User and Ninja User the same?"] = Emails_Check


#PowerBI = []
#for index, row in PowerBI_info.iterrows():
    #temp = row["Name"]
    #if temp != "Total":
        #PowerBI.append(temp)

Intune_Devices = []
for index, row in Intune_info.iterrows():
    Intune_Devices.append(row["Device name"])

Ninja_Devices = []
for index, row in Ninja_info.iterrows():
    Ninja_Devices.append(row["Display Name"])

# The purpose of the flag is to mark and rows that needs to be investigated to see if 
# it's correct, if any of the two comparison columns have a "N", a flag will be displayed
# in a column
Flag = []
Flag_Reason = []
for index, row in SD_info.iterrows():
    temp = ""
    Flag_Reason_temp = []
    if row["SD User and Ninja User the same?"] == "n":
        temp = "FLAG"
        Flag_Reason_temp.append("Assigned user to this workstation may be incorrect")
    if row["SD-SpaceIQ Location Correct?"] == "n":
        temp = "FLAG"
        Flag_Reason_temp.append("Location of workstation may be incorrect")
    Flag.append(temp)
    #for i in range(len(PowerBI)):
        #if PowerBI[i] == row["Workstation"]:
            #if row["State"] == "In Use" or row["State"] == "For Loan" or row["State"] == "For Guest":
                #print()
            #else:
                #temp = "FLAG"
                #Flag_Reason_temp.append("PowerBI (Invalid state detected)")
    temp_string = ""
    for i in range (len(Flag_Reason_temp)):
        if temp_string == "":
            temp_string += Flag_Reason_temp[i]
        else:
            temp_string += ", " + Flag_Reason_temp[i] 
    Flag_Reason.append(temp_string)



# Adds the flag list as a column in the final excel sheet       
SD_info["Flag"] = Flag
SD_info["Flag Reason"] = Flag_Reason
SD_info["Ticket Number"] = Ticket_Number
        
# This creates the final excel spreadsheet, this will be created in the same directory as where the .exe is
SD_info.to_excel("Result Spreadsheet.xlsx", index=False)

print("Document created successfully")
