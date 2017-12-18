import xlsxwriter as xl
import sys
import datetime
now = datetime.datetime.now()

def readFile():
	'''
	Reading file named rates.txt: Contains per quantity rates and tax of that item.
	'''
	file_rates = open("rates.txt","a+")
	'''
	'''
	rates = []
	lines = file_rates.readlines()
	for i in lines:
		rates.append(i.split(" "))
	file_rates.close()

	return rates

def fetchDetails(item, rates):
	for i in rates:
		if i[0] == item:
			return i
	print "Item not found."	
	return []
try:
	workbook = xl.Workbook(sys.argv[1]+".xlsx")
except IndexError:
	print "Filename not found.\n"
	exit()

worksheet = workbook.add_worksheet()
row = 0
col = 0

'''
Formating of a file
'''
bold = workbook.add_format({'bold': True, 'text_wrap':'true'})
money_format = workbook.add_format({'num_format': '##,##,##,##,##0', 'text_wrap':'true'})
bold_border = workbook.add_format({'bold':True})
side_border = workbook.add_format({'text_wrap':'true'})
left_border = workbook.add_format({'text_wrap':'true'})
bottom_border = workbook.add_format({'text_wrap':'true'})
left_bottom_border = workbook.add_format({'text_wrap':'true'})
left_border_no_wrap = workbook.add_format({'bold':'true'})
top_border = workbook.add_format()
top_right_border = workbook.add_format()

'''
The company proposing the quotation
'''
quoteBy = 'MICROLINK TECHNOCRATES P.LTD.\nSHOP D, SUPER APPARTMENT\nSANDHKUVA NAVSARI.\nCONTACT- 2491440   / 9377512300'
worksheet.write(row, col, quoteBy, left_border_no_wrap) #Takes col, col + 1 and col + 2 
worksheet.write(row, col + 1, "", top_border)
worksheet.write(row, col + 2, "", top_border)
worksheet.write(row, col + 3, "Quote No:", left_border_no_wrap) #Takes col + 3, col + 4
worksheet.write(row, col + 4, "", top_border)
worksheet.write(row, col + 5, "Dated:\n" + now.strftime("%d-%m-%Y"), left_border_no_wrap) #Takes col + 5, col + 6

worksheet.write(row, col + 6, "", top_right_border)
left_border_no_wrap.set_top()
left_border_no_wrap.set_align('top')
left_border_no_wrap.set_left()
top_border.set_top()
top_right_border.set_top()
top_right_border.set_right()
worksheet.set_row(row, 60)# Conversion 60

row += 1	
'''
'''

'''
Defining the heading of the tables sheets
'''
worksheet.write(row, col, "Sr.\nNo.", bold_border)
worksheet.write(row, col + 1, "Description of Goods", bold_border)
worksheet.write(row, col + 2, "HSN\nCode", bold_border)
worksheet.write(row, col + 3, "Quantity", bold_border)
worksheet.write(row, col + 4, "Rates\n(in Rs)", bold_border)
worksheet.write(row, col + 5, "Per", bold_border)
worksheet.write(row, col + 6, "Amt", bold_border)
worksheet.set_row(row, 30) # Some conversion 30 = 1.06 cm
'''
'''

row += 1
nextLine = "Y"
'''
Input request (space separated input is requested)
'''
print "Input formate: Item name <space>  Quantity"
'''
'''
rates = readFile()
serialNumber = 1
totalAmount = 0
tax = 0
while nextLine == "Y" or nextLine == "y":
	item_quantity = raw_input()
	querry = item_quantity.split(" ") # [0]: itemName, [1]: Quantity
	result = fetchDetails(querry[0], rates) # [0]: itemName, [1]: HSN number, [2]: rates(Rs) [3]: per [4]: %tax
	if len(result) != 0:
		print result
		worksheet.write(row, col, serialNumber, side_border)
		worksheet.write(row, col + 1, result[0], side_border)
		worksheet.write(row, col + 2, result[1], side_border)
		worksheet.write(row, col + 3, querry[1], side_border)
		worksheet.write_number(row, col + 4, float(result[2]), money_format)
		money_format.set_left()
		money_format.set_right()
		money_format.set_bottom()
		worksheet.write(row, col + 5, float(result[3]), side_border)
		tempAmt = (float(result[2]) / float(result[3])) * float(querry[1])
		# print tempAmt
		worksheet.write_number(row, col + 6, float(tempAmt), money_format)
		money_format.set_left()
		money_format.set_right()
		money_format.set_bottom()	
		totalAmount += tempAmt
		tax += (tempAmt * float(result[4])) / 100
		row += 1
		serialNumber += 1
	nextLine = raw_input("Press Y/y to make a new entry. ")

worksheet.write(row, col,"", left_border)
worksheet.write(row, col + 5, "Total\n(Base cost):", bold_border)
worksheet.set_row(row, 30)
worksheet.write_number(row, col + 6, totalAmount, money_format)
money_format.set_border()

row += 1

worksheet.write(row, col,"", left_border)
worksheet.write(row, col + 5, "GST:", bold_border)
worksheet.write_number(row, col + 6, tax, money_format)
money_format.set_border()

row += 1

worksheet.write(row, col, "", left_bottom_border)
worksheet.write(row, col + 1, "", bottom_border)
worksheet.write(row, col + 2, "", bottom_border)
worksheet.write(row, col + 3, "", bottom_border)
worksheet.write(row, col + 4, "", bottom_border)
worksheet.write(row, col + 5, "Total\n(With GST):", bold_border)
worksheet.set_row(row, 30)
worksheet.write_number(row, col + 6, totalAmount + tax, money_format)

bold_border.set_border()
side_border.set_left()
side_border.set_right()
side_border.set_bottom()
left_border.set_left()
left_bottom_border.set_left()
left_bottom_border.set_bottom()
bottom_border.set_bottom()
money_format.set_border()
column_widths = [1.10, 5.58, 1.85, 2.32, 1.83, 2.83, 2.67]
for i in range(0, 6):
	worksheet.set_column(i, i, 3.7 * column_widths[i]) #Some coversion factor 3.7
workbook.close()
