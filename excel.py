from xlrd import open_workbook
from xlutils.copy import copy

def saveData(order, quote, author):
	workbook = open_workbook('quotes.xls')
	workbookWrite = copy(workbook)
	worksheetQuotes = workbookWrite.get_sheet(order.worksheetIndex)
	worksheetQuotes.write(order.totalQuotes,0,quote)
	worksheetQuotes.write(order.totalQuotes,1,author)
	workbookWrite.save('quotes.xls')
	order.totalQuotes += 1

def getTotalQuotes(order,query):
	workbook = open_workbook('quotes.xls')
	worksheet = workbook.sheet_by_name(query)
	try:
		while worksheet.cell(order.totalQuotes,0).value:
			order.totalQuotes += 1
	except:
		print("Total Quotes saved: " + str(order.totalQuotes))

def getQuoteWorksheet(order,query):
	workbook = open_workbook('quotes.xls')
	for worksheet in workbook.sheet_names():
		if str(worksheet) == query:
			return query
		order.worksheetIndex += 1
	workbookWrite = copy(workbook)
	workbookWrite.add_sheet(query)
	worksheetNew = workbookWrite.get_sheet(order.worksheetIndex)
	worksheetNew.write(0,0,'Quote')
	worksheetNew.write(0,1,'Author')
	workbookWrite.save('quotes.xls')
	return query