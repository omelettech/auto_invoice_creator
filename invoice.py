import gspread
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import load_workbook
import shutil
import os

def createinvoice(date, name, address, phn, arttype, invoicenumber,price,email):
    global invoice
    invoice["B3"] = "Date: " + str(date)[0:10]
    invoice["B5"] = "Name: " + name
    invoice["B6"]= "Email: "+email
    invoice["B7"] = "phone number: " + phn
    invoice["B8"] = "address: " + address
    invoice["B10"] = arttype
    invoice["B2"] = "INVOICE# " + str(invoicenumber)
    invoice["C10"]=price
    if address == "" and phn == "":
        invoice["C14"] = 0
    else:
        invoice["C14"] = 460
    return


original = r'C:\Users\Butt_crack\PycharmProjects\invoice\Invoice_template.xlsx'
target = r'C:\Users\Butt_crack\PycharmProjects\invoice\Invoice_template_copy.xlsx'
shutil.copyfile(original, target)
workbook = load_workbook(filename=target)
invoice = workbook.active

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\Butt_crack\PycharmProjects\invoice\invoic-974bbb116225.json", scope)
client = gspread.authorize(creds)
sheet = client.open("Respondents details").sheet1

number_of_rows = len(sheet.get_all_records())
address = ["" for c in range(number_of_rows+1)]
phn = ["" for c in range(number_of_rows+1)]
date = sheet.col_values(1)
name = sheet.col_values(3)
contact = sheet.col_values(4)

for i in sheet.col_values(8):
    address[sheet.col_values(8).index(i)] = i
for i in sheet.col_values(9):
    phn[sheet.col_values(9).index(i)] = i

arttype = sheet.col_values(10)
email = sheet.col_values(11)


# creating a new xl file
with open(r"C:\Users\Butt_crack\PycharmProjects\invoice\invoice_counter.txt", "r") as bitch:
    invoice_no = int(bitch.readline())



data = sheet.get_all_records()


print("___________________________________")
for i in range(1,number_of_rows+1):
    print(i,"|",name[i],"|", email[i],"|", date[i][0:10])
print("___________________________________")
print("\n")


entry = int(input("enter row number "))
price = int(input("price: "))

filename = name[entry] + ".xlsx"
directory= "C:/Users/Butt_crack/Google Drive/Digital drawings/commissions/"+name[entry]+"/"
if not os.path.exists(directory):
    os.makedirs(directory)

createinvoice(date[entry], name[entry], address[entry], phn[entry], arttype[entry], invoice_no,price, email[entry])


workbook.save(filename=directory+filename)
print("File saved successfully")
with open(r"C:\Users\Butt_crack\PycharmProjects\invoice\invoice_counter.txt", "w") as bitch:
    invoice_no = bitch.write(str(invoice_no+1))