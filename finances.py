import csv
import xlwt
from datetime import datetime
from dateutil.parser import parse


def get_data(file_name):
	""" open file name for input data"""
	f = open ("FORMAT.csv")
	finance_csv = csv.reader(f)
	return finance_csv

def setup():
	"""Function to setup excel sheet for writing to"""
	#Write to Excel Setup
	wb = xlwt.Workbook()
	rd = wb.add_sheet('FINANCE')
	style1 = xlwt.easyxf('font: name Times New Roman')
	style = xlwt.XFStyle()
	pattern = xlwt.Pattern()
	pattern.pattern = xlwt.Pattern.SOLID_PATTERN
	# pattern.pattern_fore_colour = xlwt.Style.colour_map['red']
	style.pattern = pattern

	#Setup initial Headers
	rd.write(0,0, "Date")
	rd.write(0,1, "Amount Paid")
	rd.write(0,2, "Where?")
	rd.write(0,3, "Payer")
	rd.write(0,4, "Total Balance")
	return [wb, rd]

def calculate(rd, wb, finance_csv):
	""" Function to calculate and determine spending split """
	dates_wr = []
	locations = []
	amounts = []
	payers = []
	#Initialize other payers following same format 'Name' : (float(initial amount paid))
	reg_payer = {'Arsh' : (float(0.0)), 'Niket' : (float(0.0)), 'Vivek' : (float(0.0)), 'Narain' : (float(0.0)) }
	net_bal = 0.0


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
				locations.append(row[1])
				dates_wr.append(row[0])
				row[2] = row[2].replace("$", "")
				amounts.append(float(row[2]))
				payers.append(row[3])
				found = False

				for k in reg_payer.keys():
					if row[3] == k:
						found = True
						reg_payer[k] += amounts[len(amounts)-1]

				if found == False:
					success = False
					print "Error: Could not find this user in the reg_payer dict, please check spelling, noting that all users are case sensitive"
					break

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

def main():
	csv = get_data('FORMAT.csv')
	output = setup()
	calculate(rd=output[1], wb=output[0], finance_csv=csv)

if __name__ == '__main__':
	main()
