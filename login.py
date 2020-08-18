	
import time, os, pytesseract
import re
import pandas as pd
from selenium import webdriver
from wand.image import Image
from selenium.webdriver.common.keys import Keys
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class Browser():

	def __init__(self):

		fp = webdriver.FirefoxProfile()

		fp.set_preference("browser.download.folderList", 2)
		fp.set_preference("browser.download.manager.showWhenStarting", False)
		fp.set_preference("browser.download.dir", DOWNLOAD_PATH)
		fp.set_preference("browser.helperApps.neverAsk.saveToDisk", PDF_TYPES)
		fp.set_preference("browser.helperApps.neverAsk.openFile", PDF_TYPES)
		fp.set_preference("plugin.disable_full_page_plugin_for_types", PDF_TYPES)
		fp.set_preference("pdfjs.disabled", True)

		self.driver = webdriver.Firefox(executable_path=GECKODRIVER, firefox_profile=fp)

	def get(self, url):
		self.driver.get(url)


	def login(self, email, password):

		self.driver.find_element_by_name("Email").send_keys(email)
		self.driver.find_element_by_name("Password").send_keys(password)
		self.driver.find_element_by_id("loginButton").click()

	def logout(self):
	    self.driver.find_element_by_link_text("Log off").click()


	def search(self, case_number):
	    
	    self.driver.find_element_by_link_text("Records Search").click()
	    self.driver.find_element_by_name("cs.CaseNumber").send_keys(case_number)
	    self.driver.find_element_by_css_selector(".btn").click()
	    self.driver.find_element_by_link_text("Petition for Administration").click()




def read_excel_file():
    
	xl_file = pd.ExcelFile(EXCEL_FILE)
	dfs = {sheet_name: xl_file.parse(sheet_name) for sheet_name in xl_file.sheet_names}

	data = dfs['LEE-FL'].dropna()
	case_number = pd.DataFrame({"Number": data['Unnamed: 1']})
	case_number = case_number.drop(case_number.index[0])

	return case_number


def img_to_text(filename):
	try:
		from PIL import Image
	except ImportError:
		import Image

	text = pytesseract.image_to_string(Image.open(filename))	
	
	get_info(text.split("\n"))



DATA = {
    "Petitioner name": [],
    "Petitioner addr": [],
    "Decedent name": [],
    "Decedent addr": []
    }

def get_info(text):

	names = format_names(text)

	DATA["Petitioner name"] += names[0]
	DATA["Decedent name"] += names[1]

	print(DATA)

def format_names(data):

	re_name_petitioner = r"Petitioner, \w+"
	re_name_decedent = r"Decedent, \w+"
	name_list = []

	for words in data:
		if re.search(re_name_petitioner, words) or re.search(re_name_decedent, words):
			name = words.split(",")[1]
			name_list.append([name[1:]])
	return name_list

def pdf_to_img(file):

	with(Image(filename=file, resolution=150)) as source:
		for i, image in enumerate(source.sequence):
			if i is 0:
				newfilename = file + '.jpeg'
				Image(image).save(filename=newfilename)
				img_to_text(newfilename)

def main():

	brs = Browser() # 1
	brs.get(WEB_SITE) # 2
	brs.login(EMAIL, PASSWORD) # 3

	case_numbers = read_excel_file()
	count = 0
	for code in case_numbers['Number']:
		count += 1
		if count > 3: break
		brs.search(code)
		time.sleep(1)

	brs.logout()

	for file in os.listdir(DOWNLOAD_PATH):
		FILE_PATH = os.path.join(DOWNLOAD_PATH, file)
		pdf_to_img(FILE_PATH)


if __name__ == "__main__":

	EMAIL = 'paul@bestloantovalue.com'
	PASSWORD = 'Rexmydog029$'
	EXCEL_FILE = "LeeCountyFL.xlsx"
	WEB_SITE = 'https://matrix.leeclerk.org/Account/Login'
	PDF_TYPES = "application/pdf,application/vnd.adobe.xfdf,application/vnd.fdf,application/vnd.adobe.xdp+xml"
	GECKODRIVER = r'C:\Users\RAGNAR\Downloads\geckodriver.exe'

	DOWNLOAD_PATH = os.path.join(os.getcwd(), 'pdf')

	if not os.path.exists(DOWNLOAD_PATH):
		os.mkdir(DOWNLOAD_PATH)

	main()