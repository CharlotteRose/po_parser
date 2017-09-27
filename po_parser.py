# _author_ scooties
#
# PLAN :
# accept xls filename as input
# parse the sheet for occurrences of quantity (this will need to be a dynamic search - ie not sensitive to casing)
#
# !!!!!!!!! NEED TO VERIFY THAT FILES WILL NOT BE MULTI SHEETED !!!!!!!!!!!!!!

import xlrd     # import module for use
import margins
workbook = xlrd.open_workbook('/home/scoots/Desktop/automation_files/REDH160114B-5.xls')
sheets = workbook.nsheets # get the number of sheets in the work book
print (sheets) # see the number of sheets
print (workbook.sheet_by_index(0))
print (workbook.sheet_by_name("Sheet1"))
print ("all", workbook.sheet_names())

current_sheet = workbook.sheet_by_index(0)
print (current_sheet)

# IMPORTANT - this module operates on the assumption of a zero indexed r,c system
print (current_sheet.cell(8,2).value)
print (current_sheet.cell(10,2).value)

# get data from indiviual cells:
#sheet.cell


#print (workbook.sheet_by_index(sheets))
#print ("prior")
#while sheets > 0:
 #   print (sheets)
  #  print (workbook.sheet_by_index(sheets))
  #  print (sheets)
  #  sheets-1
  #  print (sheets)
