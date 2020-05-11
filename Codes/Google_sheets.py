
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import serial


ser1 = serial.Serial('COM6', 9600)
while True:
    scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    code = str(ser1.read(8))
    
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)

    client = gspread.authorize(creds)
    
    print (code)
    sheet = client.open("iot").sheet1  # Open the spreadhseet
    count = 0
    data = sheet.get_all_records()  # Get a list of all records
    length = len(data)

    row = sheet.row_values(2)  # Get a specific row
    col = sheet.col_values(2)  # Get a specific column
    #cell = sheet.cell(1,2).value  # Get the value of a specific cell
    for x in range(1,length+1):
        y = x+1
        cell = str(sheet.cell(y,1).value)
        #print(celll)
        if cell == code:
            name = sheet.cell(y,2).value
            num = sheet.cell(y,3).value
            print ("Name : " + name)
            print ("passport number : " + num)
            count =1;
            print("")
            break
    if count != 1:
        print('No such passport found')
        count =0
        print ("Do you wnt to add new passport?? y/n")
        inp = input()
        if inp== 'y':
            sheet.update_cell(length + 2,1, code)
            print("Enter the name of the person")
            n = input()
            sheet.update_cell(length + 2,2, n)
            print("Enter your passport number")
            p = input()
            sheet.update_cell(length + 2,3, p)
            print("database updated successfully")
            print("")
        else:
            print ("okay no problem")
            print("")
           
   

