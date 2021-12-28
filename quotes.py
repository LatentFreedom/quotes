# System
import sys
import time
# Selenium
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
# Local
import excel

def startDriver(headless=False):
	if headless:
		fireFoxOptions = Options()
		fireFoxOptions.headless = True
		return webdriver.Firefox(options=fireFoxOptions)
	else:
		return webdriver.Firefox()

def scrollDown(driver):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

class QuoteOrder:

	def __init__(self):
		self.url = "https://www.brainyquote.com/search_results?q="
		self.totalQuotes = 0
		self.driver = None
		self.worksheet = None
		self.worksheetIndex = 0

	def setURL(self,query):
		self.url = self.url + query

	def getQuotes(self):
		total_pages = 1
		current_page = 1

		# Get total pages
		try:
			pagination = self.driver.find_element(By.CLASS_NAME, 'pagination')
			page_links = pagination.find_elements(By.CLASS_NAME, 'page-link')
			total_pages = int(page_links[len(page_links)-2].text)
		except:
			pass


		while current_page < total_pages:

			self.driver.get(self.url+'&pg='+str(current_page))
			time.sleep(1)
			quote_container = self.driver.find_element(By.ID, 'quotesList')
			quote_list = quote_container.find_elements(By.CLASS_NAME, 'grid-item')

			# Save quotes on page from last saved
			for i in range(0,len(quote_list)):
				if 'm-ad-brick' in quote_list[i].get_attribute('class').split():
					# Ad found
					continue

				quote = quote_list[i].find_element(By.CLASS_NAME, 'b-qt').text
				author = quote_list[i].find_element(By.CLASS_NAME, 'bq-aut').text

				# print("{0:3}. {1:20} \n\t -{2:20}".format(i,quote[:100],author))
				excel.saveData(self,quote,author)

			# Go to next page
			current_page += 1
			

def createOrder(query):
	order = QuoteOrder()
	order.setURL(query)
	order.worksheet = excel.getQuoteWorksheet(order,query)
	excel.getTotalQuotes(order,query)
	return order

def run(order):
	order.driver = startDriver(True)
	order.driver.get(order.url)
	order.getQuotes()
	order.driver.quit()

def main():
	if len(sys.argv) == 2:
		query = sys.argv[1]
	else:
		print("Usage: python quotes.py <query>")
		return
	order = createOrder(query)
	run(order)

if __name__ == "__main__":
	main()
