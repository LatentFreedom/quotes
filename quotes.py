import time

from selenium import webdriver
from selenium.webdriver.common.keys import Keys

from xlrd import open_workbook
from xlutils.copy import copy


class QuoteBot:

	def __init__(self):
		self.url = "https://www.brainyquote.com/search_results.html?q="
		self.pageURL = "&pg="
		self.query = None
		self.totalQuotes = 0
		self.driver = None
		self.currentPage = 1
		self.pageCount = None
		self.worksheet = None
		self.worksheetIndex = 0

	def setURL(self):
		self.url = self.url + self.query

	def setPageURL(self):
		self.pageURL = "https://www.brainyquote.com/search_results.html?q=" + self.query + self.pageURL

	def setPageCount(self):
		time.sleep(1)
		pagination = self.driver.find_element_by_class_name('pagination')
		pageLinks = pagination.find_elements_by_tag_name('li')
		self.pageCount = int(pageLinks[len(pageLinks)-2].text)
		print("Total pages: " + str(self.pageCount))

	def getQuoteWorksheet(self):
		# import Excel workbook to write quotes to
		workbook = open_workbook('quotes.xls')
		self.worksheet = self.query
		for worksheet in workbook.sheet_names():
			if str(worksheet) == self.query:
				return True
			self.worksheetIndex += 1
		workbookWrite = copy(workbook)
		workbookWrite.add_sheet(self.query)
		worksheetNew = workbookWrite.get_sheet(self.worksheetIndex)
		worksheetNew.write(0,0,'Quote')
		worksheetNew.write(0,1,'Author')
		workbookWrite.save('quotes.xls')

	def getTotalQuotes(self):
		# import Excel workbook to read from
		workbook = open_workbook('quotes.xls')
		worksheet = workbook.sheet_by_name(self.query)
		try:
			while worksheet.cell(self.totalQuotes,0).value:
				self.totalQuotes += 1
		except:
			print("Total Quotes saved: " + str(self.totalQuotes))

	def getPageQuotes(self):
		quotes = self.driver.find_elements_by_class_name('bqQuoteLink')
		authors = self.driver.find_elements_by_class_name('bq-aut')
		for i in range(1, len(quotes)):
			quote = quotes[i-1].text
			author = authors[i-1].text
			saveData(self,quote,author)
			print("{0:3}. {1:20} \n\t -{2:20}".format(self.totalQuotes,quote[:40],author))
			self.totalQuotes += 1

	def nextPage(self):
		self.currentPage += 1
		self.driver.get(self.pageURL+str(self.currentPage))


def startDriver(d):
	if d == 'f':
		return webdriver.Firefox()
	else:
		return webdriver.PhantomJS()

def saveData(order, quote, author):
	# import Excel workbook to write quotes to
	workbook = open_workbook('quotes.xls')
	workbookWrite = copy(workbook)
	worksheetQuotes = workbookWrite.get_sheet(order.worksheetIndex)
	worksheetQuotes.write(order.totalQuotes,0,quote)
	worksheetQuotes.write(order.totalQuotes,1,author)
	workbookWrite.save('quotes.xls')

def createOrder(query):
	order = QuoteBot()
	order.query = query
	order.setURL()
	order.setPageURL()
	order.getQuoteWorksheet()
	order.getTotalQuotes()
	return order

def run(order):
	order.driver = startDriver('p')
	order.driver.get(order.url)
	#order.setPageCount()
	for i in range(0,order.pageCount):
		order.getPageQuotes()
		order.nextPage()
	order.driver.quit()

def main():
	query = raw_input("Query: ")
	order = createOrder(query)
	run(order)


if __name__ == "__main__":
	main()