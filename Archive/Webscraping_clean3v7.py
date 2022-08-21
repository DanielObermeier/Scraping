# Before executing this script, the .xls-File needs to be prepared using the following
# two macros:

# VBA Code to transform Hyperlinks into Python-readable URL's:
# 
# Sub ExtractHL()
#     Dim HL As Hyperlink
#     For Each HL In ActiveSheet.Hyperlinks
#         HL.Range.Offset(0, 0).Value = HL.Address
#     Next
# End Sub
# 
# Sub TruncateURLS()
# 
# Dim cval, url, remove As String
# Dim lrow, i, rev As Integer
# 
# lrow = ActiveSheet.Cells(Rows.Count, "C").End(xlUp).Row
# 
# For i = 2 To lrow
#         cval = ActiveSheet.Cells(i, "C").Value
#         remove = "?utm_source=crunchbase&utm_medium=export&utm_campaign=odm_csv"
#         rev = InStrRev(cval, remove)
#         If rev > 0 Then
#             url = Replace(cval, remove, "")
#             ActiveSheet.Cells(i, "C").Value = url
#         End If
# Next
# 
# End Sub

# Also: Shuffle the order of the data points in the .xls-File so as to minimize the chance of
# being detected by Crunchbase Webmasters ;-) (e.g. sort by columns K, P, A in A-Z ascending
# order)

# Header
import os
import re
import sys
import concurrent
from concurrent import futures
from xlrd import open_workbook

from fake_useragent import UserAgent
import time

from random import randint
import random

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from datetime import datetime

import bs4

import xlwt

import subprocess

#for downloading proxy list
import urllib.request

# Definitions

opts = Options()
detection = "something about your browser made us think you were a bot. There are a few reasons this might happen:"
blocked = "Access To Website Blocked"
distil="distil_ident_block"
xlsfile = 'Python_input1032_1499.xls'
proxylist = []
proxylist_path = "proxylist.txt"
	
# script = '''set vpn_name to "'LRZ VPN'"

# tell application "System Events"
	# set rc to do shell script "scutil --nc status " & vpn_name
	# if rc starts with "Connected" then
		# do shell script "scutil --nc stop " & vpn_name
	# else
		# do shell script "scutil --nc start " & vpn_name
	# end if
# end tell
# '''

# Functions

def readexcel():
		
	if os.path.isfile(proxylist_path):
		with open(proxylist_path, "r") as text_file:
			for line in text_file:
				proxylist.append(str.rstrip(line,"\n"))
	
	else:

		try:
			proxylist_source=urllib.request.urlopen("http://www.rosinstrument.com/proxy/plab100.xml").read()

		except:
			print('Error when fetching proxy list from web')
			sys.exit()
			
		proxylist_sourcesoup = bs4.BeautifulSoup(proxylist_source, "html.parser")
		proxylist_source = proxylist_sourcesoup.find_all("title")
		
		for eachProxy in proxylist_source[2:]:
			proxylist.append(eachProxy.text)
		
		
		
		with open(proxylist_path, "w") as text_file:
			for proxy in proxylist:
				print(proxy, file=text_file)
				
	
	

	if __name__ == '__main__':
		print("init excelmanager")
		myRealExcelMgr = ExcelThreadManager(xlsfile)
		print("init excelmanager done")
		request(myRealExcelMgr)


def request(myExcelMgr):
# TO DO: implement mode switch for raw source


	currn = myExcelMgr.getStartRow()
	trn = myExcelMgr.getLastRow()
	trn = trn+1
	
	executor = concurrent.futures.ThreadPoolExecutor(10)
	futures = [executor.submit(getContents, i, myExcelMgr) for i in range(currn, trn+1)]
	concurrent.futures.wait(futures)

	
	# for i in range(currn, trn):
		# print ("----")
		# print ("Working on row " + str(i) + " of " + str(trn))
		# getContents(i, myExcelMgr)

def getContents(i, myExcelMgr):

	cb_timestamp="."
	cb_status="."
	cb_totalfunding="."
	lfd="."
	cb_headquarter="."
	cb_founders="."
	cb_description="."
	cb_statusWhen="."
	cb_statusBy="."
	cb_noOfEmployees="."
	cb_noOfAcquisitions="."
	cb_noOfProducts="."
	cb_dateFounded="."
	cb_initialinvestment_date="."
	cb_initialinvestment_amount="."
	cb_initialinvestment_round="."
	cb_initialinvestment_leadinvestor="."
	cb_news_visible="."
	cb_categories ="."
	text = ""
	cb_investors ="."
	cb_investmentstatus = "."
	cb_initialinvestment_round = "."
	cb_investmentstatus_round = "."
	cb_noOfInvestments = "."
	investmentstatus_amount = "."
				
	if (len(proxylist) == 0):
		print("No proxies left to use")
		sys.exit()
		
		
	wt = randint(1,3)
	col = myExcelMgr.getValue(2,i)
	print(col)
	
	ua = "user-agent=" + UserAgent().chrome
	opts.add_argument(ua)
	
	#pick random proxy-server
	# TO DO: check proxy availability, kill from list if unavailable
	
	thisProxy = random.choice(proxylist)
	print("Setting proxy to " + thisProxy)
	opts.add_argument('--proxy-server=' + thisProxy)
	opts.add_argument('--headless')
	
	#adblock
	opts.add_argument('load-extension=C:/Users/Tobias Hlavka/AppData/Local/Google/Chrome/User Data/Default/Extensions/cjpalhdlnbpafiamejdnhcphjbkeiagm/1.13.8_0/')


	try:
		browser = webdriver.Chrome("C:/Users/Tobias Hlavka/AppData/Local/Programs/Python/Python36/chromedriver/chromedriver.exe", chrome_options=opts)
	
	except:
		print('Webdriver failure at row ' + str(i) + ' - aborting')
		browser.close()
		myExcelMgr.writeData()
		sys.exit()
	
	if (col != 'cb_URL') and (col != '.'):

		try:
			browser.get(col)
			print("Got some content")
			
		except SystemExit:
			print("Exception: SystemExit - But Why?")
			sys.exit()
		
		except:
			print('Could not get page - retrying...')
			browser.close()
			getContents(i, excel)
						
		time.sleep(1)
		try:
			text = browser.page_source
		except:
			print('Could not get source - retrying...')
			browser.close()
			getContents(i, excel)
			
		cbsourcesoup = bs4.BeautifulSoup(text, "html.parser")

		if (detection in text) or (blocked in text) or (distil in text):
			print("Script detected by Crunchbase at line " + str(i)) 
			print("Removing proxy from list and retrying...")
			proxylist.remove(thisProxy)
			print("Number of proxies left: " + str(len(proxylist)))
			
			with open(proxylist_path, "w") as text_file:
				for proxy in proxylist:
					print(proxy, file=text_file)
			
			browser.close()
			time.sleep(wt)
			print("Starting over with same request for line " + str(i))
			getContents(i, excel)
			 
		elif ("ERR_PROXY_CONNECTION_FAILED" in text) or ("ERR_TUNNEL_CONNECTION_FAILED" in text) or ("(92)" in text) or ("ERR_EMPTY_RESPONSE" in text) or ("ERR_CONNECTION_TIMED_OUT" in text) or (len(text) < 200):
			print("Proxy down - removing from list: " + thisProxy)
			proxylist.remove(thisProxy)
			print("Number of proxies left: " + str(len(proxylist)))
			with open(proxylist_path, "w") as text_file:
				for proxy in proxylist:
					print(proxy, file=text_file)
		
			browser.close()
			time.sleep(wt)
			print("Starting over with same request for line " + str(i))
			getContents(i, myExcelMgr)
		
		elif ("crunchbase.com/assets/404" in text):
			print("Invalid URL in Excel on line " + str(i) +", skipping...")
			#TO DO: SET ALL VARIABLES TO X
			cb_status="x"
			cb_totalfunding="x"
			lfd="x"
			
			browser.close()
			myExcelMgr.setValue(3,i,cb_status)
			myExcelMgr.setValue(4,i,cb_totalfunding)
			myExcelMgr.setValue(5,i,lfd)
			
		else:
			print("Extracting contents of "+ col)
			#set timestamp
			cb_timestamp = str(datetime.now().date().strftime("%Y/%m/%d") + " " + datetime.now().time().strftime("%H:%M:%S"))
			print(cb_timestamp)
			
			#save source to disk
			filename_timestamp = datetime.now().date().strftime("%Y%m%d")
			with open("cb_source/"+filename_timestamp+"_"+str.split(col,"/")[len(str.split(col,"/"))-1]+".txt", "w", encoding="utf-8") as text_file:
				print(text, file=text_file)
		
			try:
				cb_status = str(cbsourcesoup.find("dt",text="Status").parent.findNext("dd").contents[0])
			except AttributeError:
				cb_status = "."
			
			if 'closed' in cb_status:
				cb_status = "Closed"
			elif 'Acquired' in cb_status:
				cb_status = "Acquired"
			print("cb_status")
			print(cb_status)
			
			try:
				cb_totalfunding = str.strip(cbsourcesoup.find('span', {'class':'collection-count'}).next_element.next_element)
			except:
				cb_totalfunding = "."
			
			if '$' not in cb_totalfunding:
				cb_totalfunding = "."
			elif ' /' in cb_totalfunding:
				cb_totalfunding_new = cb_totalfunding.replace(' /', '', 1)
				cb_totalfunding = cb_totalfunding_new
			elif '-' in cb_totalfunding:
				cb_totalfunding_new = cb_totalfunding.replace('-', '', 1)
				cb_totalfunding = cb_totalfunding_new
			
			if cb_totalfunding == ".":
				try:
					cb_totalfunding = str(cbsourcesoup.find('span', {'class':'funding_amount'}).next_element.next_element.next_element)
				except:
					cb_totalfunding = "."	
			
			
			print("cb_totalfunding")
			print(cb_totalfunding)
			
			
			try:
				lfd = str(cbsourcesoup.find('span', {'class':'funding-type'}).next_element.next_element.next_element.next_element.next_element.next_element)
				new_lfd = lfd.lstrip()
				lfd = new_lfd
				
			except AttributeError:
				lfd = "."
				
			print("lfd")
			print(lfd)
			
			
			# find cb_headquarter
			try:
				cb_headquarter = cbsourcesoup.find('dt', text="Headquarters:").find_next('a').text
			
			except AttributeError:
				cb_headquarter = "."
				
			print ("HQ:")
			print (cb_headquarter)
		
			
			#find cb_founders
			try:
				cb_founders = cbsourcesoup.find('dt', text="Founders:").find_next('dd').text
				
			except AttributeError:
				cb_founders = "."
				
			print ("Founders:")
			print (cb_founders)
			
			#find cb_categories
			try:
				cb_categories = cbsourcesoup.find('dt', text="Categories:").find_next('dd').text
				
			except AttributeError:
				cb_categories = "."
				
			print ("Categories:")
			print (cb_categories)
			
			
			#find cb_description
			try:
				cb_description = cbsourcesoup.find('dt', text="Description:").find_next('dd').text
				
			except AttributeError:
				cb_description = "."
				
			print ("Description:")
			print (cb_description)
			
			
			#find cb_statusWhen
			try:
				cb_statusWhen = str.strip(cbsourcesoup.find("dt",text="Status").parent.find_next("dd").contents[7])
				
			except:
				cb_statusWhen = "."
				
			print ("cb_statusWhen:")
			print (cb_statusWhen)
			
			
			#find cb_statusBy
			try:
				cb_statusBy = cbsourcesoup.find("dt",text="Status").parent.find_next("dd").contents[4].text
				
			except:
				cb_statusBy = "."
				
			print ("cb_statusBy:")
			print (cb_statusBy)
			
			
			#find cb_noOfEmployees
			#TO DO: rstrip " | "
			# TO if contains "crunchbase" dann nur linkes wort nehmen
			try:
				cb_noOfEmployees = cbsourcesoup.find('dt', text="Employees:").find_next('dd').contents[0].text
				
				cb_noOfEmployees.replace(" in Crunchbase", "")
				
			except AttributeError:
				cb_noOfEmployees = "."
				
			print ("cb_noOfEmployees:")
			print (cb_noOfEmployees)
			
			
			#find cb_noOfAcquisitions
			try:
				cb_noOfAcquisitions = str(cbsourcesoup.find("dt",text="Acquisitions").parent.findNext("a").text).replace(" Acquisitions","")
				
			except:
				cb_noOfAcquisitions = "."
				
			print ("cb_noOfAcquisitions:")
			print (cb_noOfAcquisitions)
			
			
			#find cb_noOfProducts
			#TO DO: Strip nbsp and ()
			try:
				cb_noOfProducts = str.strip(str.strip(str.strip(cbsourcesoup.find('div', {'class':'base info-tab products'}).findNext("span").text),"("),")")
				
			except AttributeError:
				cb_noOfProducts = "."
				
			print ("cb_noOfProducts:")
			print (cb_noOfProducts)
			
			#find cb_dateFounded
			try:
				cb_dateFounded = cbsourcesoup.find('dt', text="Founded:").find_next('dd').text
				
			except AttributeError:
				cb_dateFounded = "."
				
			print ("cb_dateFounded:")
			print (cb_dateFounded)
			
			#find cb_public
			#TO DO: Split by Date and Symbol
			try:
				cb_public = cbsourcesoup.find("dt",text="IPO / Stock").text
				
			except AttributeError:
				cb_public = "."
				
			print ("cb_public:")
			print (cb_public)
			
			
			#find cb_initialinvestment_date
			try:
				cb_initialinvestment_date = str(cbsourcesoup.find('h2', {'id':'funding_rounds'}).findNext("tbody").find_all("tr"))
				cb_initialinvestment_date = str(cb_initialinvestment_date[len(cb_initialinvestment_date)-1].contents[0].text)
			except AttributeError:
				cb_initialinvestment_date = "."
				
			print ("cb_initialinvestment_date:")
			print (cb_initialinvestment_date)
			
			
			#find cb_initialinvestment_amount
			#TO DO: Split round from amount
			try:
				cb_initialinvestment_amount = str(cbsourcesoup.find('h2', {'id':'funding_rounds'}).findNext("tbody").find_all("tr"))
				cb_initialinvestment_amount = str(cb_initialinvestment_amount[len(cb_initialinvestment_amount)-1].contents[1].text)
				
				# try to split amount and round, let's just hope both is given
				try:
					cb_initialinvestment_round = str.split(cb_initialinvestment_amount," / ")[1]
					cb_initialinvestment_amount = str.split(cb_initialinvestment_amount," / ")[0]
					print("try")
					print(str.split(cb_initialinvestment_amount[len(cb_initialinvestment_amount)-1].contents[1].text," / ")[1])
				except:
					cb_initialinvestment_round = "."
					
			except AttributeError:
				cb_initialinvestment_amount = "."
				cb_initialinvestment_round = "."
				
			print ("cb_initialinvestment_amount:")
			print (cb_initialinvestment_amount)
			print ("cb_initialinvestment_round:")
			print (cb_initialinvestment_round)
			
			
			#find cb_initialinvestment_leadinvestor
			# TO DO: Beautify multiple investors
			try:
				cb_initialinvestment_leadinvestor = cbsourcesoup.find('h2', {'id':'funding_rounds'}).findNext("tbody").find_all("tr")
				cb_initialinvestment_leadinvestor = str(cb_initialinvestment_leadinvestor[len(cb_initialinvestment_leadinvestor)-1].contents[3].text)
			except AttributeError:
				cb_initialinvestment_leadinvestor = "."
				
			print ("cb_initialinvestment_leadinvestor:")
			print (cb_initialinvestment_leadinvestor)
			
			#find cb_investors
			try:
				cb_investors_raw = cbsourcesoup.find('h2', {'id':'investors'}).find_all("tbody")
				for investor in cb_investors:
					cb_investors = cb_investors.find_all("tr").find_next("td").text + ", "
			except AttributeError:
				cb_investors = "."
				
			print ("cb_investors:")
			print (cb_investors)
			
			
			#find cb_investmentstatus_round
			# TO DO: Series oder Seed
			try:
				cb_investmentstatus = str(cbsourcesoup.find('h2', {'id':'funding_rounds'}).findNext("tbody").find_all("tr"))
				
				res=""
				for inv in cb_investmentstatus:
					if ("Seed" in inv.contents[1].text) or ("Series" in inv.contents[1].text):
						res = inv.contens[1].text
						break
					
				# try to split amount and round, let's just hope both is given
				try:
					cb_investmentstatus_round = str.split(res," / ")[1]
					
				except:
					cb_investmentstatus_round = "."
					
			except AttributeError:
				cb_investmentstatus = "."
				cb_investmentstatus_round = "."
				
			print ("cb_investmentstatus_round:")
			print (cb_investmentstatus_round)
			
			#find cb_investmentstatus_amount
			try:
				cb_investmentstatus = str(cbsourcesoup.find('h2', {'id':'funding_amounts'}).findNext("tbody").find_all("tr"))
				cb_investmentstatus = str(cb_investmentstatus[0].contents[1].text)
				
				# try to split amount and round, let's just hope both is given
				try:
					cb_investmentstatus_amount = str.split(cb_investmentstatus," / ")[0]
				except:
					cb_investmentstatus_amount = "."
					
			except AttributeError:
				cb_investmentstatus_amount = "."
				
			print ("cb_investmentstatus_amount:")
			print (cb_investmentstatus_amount)
			
			#find cb_noOfInvestments
			try:
				cb_noOfInvestments = str.strip(str.strip(str.strip(str(cbsourcesoup.find('h2', {'id':'funding_amounts'}).findNext("span",{"class": "collection-count"}).text,"("),")")))
					
			except AttributeError:
				cb_noOfInvestments = "."
				
			print ("cb_noOfInvestments:")
			print (cb_noOfInvestments)
			
			
			#cb_news_visible
			try:
				cb_news_visible_source = cbsourcesoup.find('div', {'class':'base info-tab press_mentions'}).find('div',{'class':'card-content box container card-slim'}).find_next("tbody").find_all('tr')
				cb_news_visible = ""
				for news in cb_news_visible_source:
					cb_news_visible = cb_news_visible + news.contents[0].text + " " + news.contents[1].text + "\n"
				
			except AttributeError:
				cb_news_visible = "."
				
			print ("cb_news_visible:")
			print (cb_news_visible)
			
			browser.close()

	
		
			myExcelMgr.setValue(3,i,cb_timestamp)
			myExcelMgr.setValue(3,i,cb_status)
			myExcelMgr.setValue(4,i,cb_totalfunding)
			myExcelMgr.setValue(5,i,lfd)
			myExcelMgr.setValue(6,i,cb_headquarter)
			myExcelMgr.setValue(7,i,cb_founders)
			myExcelMgr.setValue(8,i,cb_description)
			myExcelMgr.setValue(9,i,cb_statusWhen)
			myExcelMgr.setValue(10,i,cb_statusBy)
			myExcelMgr.setValue(11,i,cb_noOfEmployees)
			myExcelMgr.setValue(12,i,cb_noOfAcquisitions)
			myExcelMgr.setValue(13,i,cb_noOfProducts)
			myExcelMgr.setValue(14,i,cb_dateFounded)
			myExcelMgr.setValue(15,i,cb_initialinvestment_date)
			myExcelMgr.setValue(16,i,cb_initialinvestment_amount)
			myExcelMgr.setValue(17,i,cb_initialinvestment_round)
			myExcelMgr.setValue(18,i,cb_initialinvestment_leadinvestor)
			myExcelMgr.setValue(19,i,cb_news_visible)
			myExcelMgr.setValue(20,i,cb_categories)
			myExcelMgr.setValue(21,i,cb_investors)
			myExcelMgr.setValue(22,i,cb_investmentstatus)
			myExcelMgr.setValue(23,i,cb_initialinvestment_round)
			myExcelMgr.setValue(24,i,cb_noOfInvestments)
			myExcelMgr.setValue(25,i,investmentstatus_amount)
					
			
			myExcelMgr.writeData()
		
		
	elif col == '.':
		#TO DO: if theres a dot fill all field with dots
		cb_totalfunding = "."
		b = "."
		lfd = "."
		
		myExcelMgr.setValue(3,i,b)
		myExcelMgr.setValue(4,i,cb_totalfunding)
		myExcelMgr.setValue(5,i,lfd)
		myExcelMgr.writeData()

		browser.close()

	myExcelMgr.writeData()
		
		
def exporttoexcel(excel, col):
	workbook = xlwt.Workbook()
	sheet = workbook.add_sheet("Python Output")
    
	col_index = 0    
	row_index = 0
    
	for values in excel:
		for val in values:
			sheet.write(row_index, col_index, val)
			row_index = row_index + 1
		col_index = col_index + 1
		row_index = 0

	workbook.save(xlsfile)
	print('Exported current dataset')

def collectlinks():	
	ua = "user-agent=" + UserAgent().chrome
	print("Setting UA to " + ua)
	opts.add_argument(ua)

	
	# try:
		# browser = webdriver.Chrome("C:/Users/Tobias Hlavka/AppData/Local/Programs/Python/Python36/chromedriver/chromedriver.exe", chrome_options=opts)
		# browser.get("http://crunchbase.com")
		
	# except:
		# print('Webdriver failure')
		# sys.exit()
		
	# time.sleep(60)
	# SCROLL_PAUSE_TIME = 4

	# # Get scroll height
	# last_height = browser.execute_script("return document.body.scrollHeight")

	# while True:
		# sys.stdout.write(".")
		# # Scroll down to bottom
		# browser.execute_script("window.scrollTo(0, document.body.scrollHeight+1);")

		# # Wait to load page
		# time.sleep(SCROLL_PAUSE_TIME)

		# # Calculate new scroll height and compare with last scroll height
		# new_height = browser.execute_script("return document.body.scrollHeight")
		# if new_height == last_height:
			# break
		# last_height = new_height
	# text = browser.page_source.encode("utf-8")
	text=""
	
	with open("listtestsource.txt", "r", encoding="utf-8") as text_file:
			text = text_file.read()
	print("read file from disk")
	cbsourcesoup = bs4.BeautifulSoup(text, "html.parser")
	
	#cb_names = cbsourcesoup.find_all('a', {'href': re.compile(r'^/organization/.*')})
	cb_divs = cbsourcesoup.find_all('div', {'class': 'content container'})
	
	print("parsed file. Nbr of matches found: " + str(len(cb_divs)))
		
	namelist=""
	
	print("Extracting names and hrefs...")
	
	for eachDiv in cb_divs:
		cb_name = eachDiv.find('a', {'href': re.compile(r'/organization/.*')})
		#cb_name = cb_name.content[0]
		print(cb_name)
		namelist= namelist + cb_name['title'] + ";" + cb_name['href'] + "\n"
	
	#namelist=""
	#for eachName in cb_names:
		#namelist= namelist + eachName['title'] + ";" + "https://www.crunchbase.com" + eachName['href'] + "\n"
	#	namelist= namelist + eachName['title'] + ";" + eachName['href'] + "\n"
	
	
	print("writing to file")
	with open("inputlinks.txt", "w", encoding="utf-8") as text_file:
		print(namelist, file=text_file)

class ExcelThreadManager:
	excel = []
	currn = 0
	trn = 0
	path = ""
	
	def __init__(self, path):
		self.path = path
		book = open_workbook(self.path)
		sheet = book.sheet_by_index(0)
		for col in range(sheet.ncols):
			values = []
			for row in range(sheet.nrows):
				values.append(sheet.cell(row,col).value)
			self.excel.append(values)
			
		trn = len(self.excel[0])-1
		self.trn = trn
		currn = self.excel[5].index("")
		self.currn = currn
		print(('Total number of rows in Excel is ' + str(trn)))
		print(('Current request row number is ' + str(currn)))
		
	def getAllData(self):
		return self.excel
		
	def getValue(self, col, row):
		return self.excel[col][row]
	
	def setValue(self, col,row,value):
		self.excel[col][row] = value
	
	def writeData(self):
		workbook = xlwt.Workbook()
		sheet = workbook.add_sheet("Python Output")
		
		col_index = 0    
		row_index = 0
		
		for values in self.excel:
			for val in values:
				sheet.write(row_index, col_index, val)
				row_index = row_index + 1
			col_index = col_index + 1
			row_index = 0

		workbook.save(self.path)
		print('Exported current dataset')
		
	def getStartRow(self):
		return self.currn
		
	def getLastRow(self):
		return self.trn
	
#collectlinks()
readexcel()
