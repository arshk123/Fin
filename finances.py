import csv
import xlwt
from datetime import datetime
from dateutil.parser import parse


f = open ("FORMAT.csv")

#Write to Excel Setup
wb = xlwt.Workbook()
rd = wb.add_sheet('FINANCE')
style1 = xlwt.easyxf('font: name Times New Roman')
style = xlwt.XFStyle()
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
# pattern.pattern_fore_colour = xlwt.Style.colour_map['red']
style.pattern = pattern

finance_csv = csv.reader(f)
dates_wr = []
locations = []
amounts = []
payers = []
#Initialize other payers following same format 'Name' : (float(initial amount paid))
reg_payer = {'Arsh' : (float(0.0)), 'Niket' : (float(0.0)), 'Vivek' : (float(0.0)), 'Narain' : (float(0.0)) }
net_bal = 0.0

#Setup initial Headers
rd.write(0,0, "Date")
rd.write(0,1, "Amount Paid")
rd.write(0,2, "Where?")
rd.write(0,3, "Payer")
rd.write(0,4, "Total Balance")
i = 1;

success = True
first = True
#parse CSV
for row in finance_csv:
		#Checks for empty row or first row
		if first:
			# print row
			first = False
		elif row[1] == '':
			print "empty row"
		else:
			# print row
			locations.append(row[1])
			dates_wr.append(row[0])
			row[2] = row[2].replace("$", "")
			amounts.append(float(row[2]))
			payers.append(row[3])
			found = False

			for k in reg_payer.keys():
				# print k
				if row[3] == k:
					# print "yes"
					found = True
					reg_payer[k] += amounts[len(amounts)-1]
			# bool found_return_val = false

			if found == False:
				success = False
				print "Error: Could not find this user in the reg_payer dict, please check spelling, noting that all users are case sensitive"
				break
			# expenses[amounts[len(amounts) - 1]] = {dates_wr[len(dates_wr) - 1], locations[len(locations)-1]}
			#mon_s = amounts[len(amounts) - 1] + mon_s
			# print "successfully found payer in registered list"
			#first element of our list
			net_bal = net_bal + amounts[len(amounts)-1]

			# rd.write(i, 1, amounts[len(amounts) - 1], style)
			rd.write(i, 1, amounts[len(amounts) - 1], style1)
			rd.write(i, 0, dates_wr[len(dates_wr) - 1], style1)
			rd.write(i, 2, locations[len(locations) - 1], style1)
			rd.write(i, 3, payers[len(payers) - 1], style1)
			rd.write(i, 4, net_bal, style1)
			i = i + 1
print "Total spent", (net_bal), "which amounts to", (net_bal/4), "per person"
for k in reg_payer.keys():
	print k, "paid this much:", reg_payer[k], "and we each owe him", (reg_payer[k]/len(reg_payer))
if success:
	print "Success!"
else:
	print "Failure"

wb.save('finance.xls')
