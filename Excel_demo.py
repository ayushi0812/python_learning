import xlrd #read excelsheet
import xlwt #write excelsheet

book = xlrd.open_workbook("F:\Machinelearning\Machine_learning\Excelsheet_demo.xlsx")  #give full path

sheet = book.sheets()[0]
#sheet = book.sheet_by_name("Sheet1")
#sheet = book.sheet_by_index(0)

r = sheet.row(0)
c = sheet.col_values(0)
print(sheet.row_values(1)[2])

print(r)
print(c)

data = []
for i in xrange(sheet.nrows):
	data.append(sheet.row_values(i))
	
print(data)

#--------------------------------------------------------------------------------

# def main():
	# """
		# purpose: abc
		# author:abc
		# exception:abc
	# """
	
	# book = xlwt.Workbook()
	# sheet1 = book.add_sheet("Pysheet1")
	
	# cols = ["A", "B", "C", "D", "E"]
	# txt = "Row {0}, Col {1}"
	
	# for num in range(5):
		# row = sheet1.row(num)
		# for index, col in enumerate(cols):
			# print("col {col}".format(col = col))
			# value = txt.format(num+1,col)
			# print("value {value}".format(value=value))
			# row.write(index,value)
			
	# book.save("test_excel.xls")
	
# #------------------------
# if __name__== "__main__":
	# main()
	
		