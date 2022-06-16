# openpyxl = a python library to read/write Excel
import openpyxl 
from openpyxl import Workbook
from openpyxl.styles import Font

# open the excel file
path = "input.xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active



# email_dict = {customer_id: [[primary email address], [non-primary email]]}
email_dict = {}

for i in range(2, sheet_obj.max_row):
	cell_1 = sheet_obj.cell(row = i, column = 1)
	cell_2 = sheet_obj.cell(row = i, column = 2)
	cell_3 = sheet_obj.cell(row = i, column = 3)

	customer_id = cell_1.value
	email = cell_2.value
	primary = cell_3.value

	if customer_id not in email_dict:
		email_dict[customer_id] = [[], []]

	if primary == 'yes':
			email_dict[customer_id][0].append(email)
	else:
			email_dict[customer_id][1].append(email)



# sort alphabetically each list of email address
for email in email_dict.values():
  email[0].sort() # sort the list of primary email addresses
  email[1].sort() # sort the list of non-primary email addresses

for customer_id, email in email_dict.items():
  email_dict[customer_id] = email[0] + email[1]



# make string of email summary
# e.g. Primary_email_1[;Primary_email_2 to N][;Non_Primary_email_1 to N][;+ X more]

data = (("CustomerID", "EmailSummary"),)

for customer_id, email_list in email_dict.items():
	email_summary = email_list[0]
	for i in range(1, len(email_list)):
		if 64 - len(email_summary) >= len(email_list[i]):
			email_summary += ';' + email_list[i]
		else:
			email_summary += ';+ ' + str(len(email_list[i:])) + ' more'
			break
		
	data += (customer_id, email_summary),



# Writing to Spreadsheets
workbook = Workbook()
workbook.save(filename="email_summary.xlsx")

wb = openpyxl.load_workbook("email_summary.xlsx")
sheet = wb.active

for row in data:
	sheet.append(row)

sheet['A1'].font = Font(bold=True)
sheet['B1'].font = Font(bold=True)
wb.save("output.xlsx")

