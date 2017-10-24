#### System
import sys
import time
#### Selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
### Local
import excel

def startDriver(d):
	driver = webdriver.Firefox() if d == 'f' else webdriver.PhantomJS()
	return driver

def scrollDown(driver):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

class QuoteOrder:

	def __init__(self):
		self.url = "https://www.brainyquote.com/search_results.html?q="
		self.totalQuotes = 0
		self.driver = None
		self.worksheet = None
		self.worksheetIndex = 0

	def setURL(self,query):
		self.url = self.url + query

	def getQuotes(self):
		loopcatch = 0
		index = 0

		quote_container = self.driver.find_element_by_id('quotesList')
		quote_list = quote_container.find_elements_by_class_name('m-brick')

		while len(quote_list) > index and loopcatch < 4:
			loopcatch = loopcatch + 1 if len(quote_list) == index else 0

			# Save quotes on page from last saved
			for quote in range(index,len(quote_list)):

				quote = quote_list[index].find_element_by_class_name('b-qt').text
				author = quote_list[index].find_element_by_class_name('bq-aut').text

				print("{0:3}. {1:20} \n\t -{2:20}".format(index,quote[:40],author))
				excel.saveData(self,quote,author)
				index += 1

			# Check for more quotes
			scrollDown(self.driver)
			time.sleep(2)

			quote_container = self.driver.find_element_by_id('quotesList')
			quote_list = quote_container.find_elements_by_class_name('m-brick')

def createOrder(query):
	order = QuoteOrder()
	order.setURL(query)
	order.worksheet = excel.getQuoteWorksheet(order,query)
	excel.getTotalQuotes(order,query)
	return order

def run(order):
	order.driver = startDriver('f')
	order.driver.get(order.url)
	order.getQuotes()
	order.driver.quit()

def main():
	# python quotes.py <query>
	query = sys.argv[1]
	order = createOrder(query)
	run(order)


if __name__ == "__main__":
	main()
