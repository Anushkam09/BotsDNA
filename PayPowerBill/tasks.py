from RPA.Browser.Selenium import Selenium
from RPA.Archive import Archive
import os
import PyPDF2 as PDF

class ZipHandling:
    def __init__(self,zip_directory,zip_file_path):
        self.zip_directory = zip_directory
        self.zip_file_path = zip_file_path
        self.zip =Archive()
    
    def extract_zip_file(self):
        self.zip.extract_archive(self.zip_directory, self.zip_file_path)
        

class PayPowerBillSite:
    def __init__(self, url, download_directory):
        self.url = url
        self.download_directory = download_directory
        self.browser = Selenium()
        self.browser.set_download_directory(self.download_directory)
        self.browser.open_chrome_browser(url = self.url)
    
    def download_file(self):
        self.browser.wait_until_element_is_visible("//a[contains(@href, 'Payments.zip')]", timeout=10)
        self.browser.click_element("//a[contains(@href, 'Payments.zip')]")
        self.waiting(f"{self.download_directory}/Payments.zip")

    def fill_form(self, details):
        self.browser.go_to(f"{self.url}/{details["url_to_go"]}")
        xpath_start = "//html/body/center/center/div/div/div/table[1]/tbody/tr"
        self.browser.input_text(f"{xpath_start}[1]/td/input", details["Installment Number"])
        self.browser.input_text(f"{xpath_start}[2]/td/input", details['Transaction_Number'])
        self.browser.input_text(f"{xpath_start}[3]/td/input", details["Amount"])
        self.browser.click_element(f"{xpath_start}[4]/td/table/tbody/tr[1]/td[2]/select[1]")
        self.browser.click_element(self.browser.find_elements(f"{xpath_start}[4]/td/table/tbody/tr[1]/td[2]/select[1]/option")[int(details["DD"])])
        self.browser.click_element(self.browser.find_elements(f"{xpath_start}[4]/td/table/tbody/tr[1]/td[2]/select[2]/option")[int(details["month"])])
        self.browser.click_element(self.browser.find_elements(f"{xpath_start}[4]/td/table/tbody/tr[1]/td[2]/select[3]/option")[int(details["Year"]) - 1990])
        self.browser.click_element(self.browser.find_elements(f"{xpath_start}[4]/td/table/tbody/tr[1]/td[2]/select[4]/option")[int(details["HH"])])
        self.browser.click_element(self.browser.find_elements(f"{xpath_start}[4]/td/table/tbody/tr[1]/td[2]/select[5]/option")[int(details["MM"])])
        self.browser.click_element(self.browser.find_elements(f"{xpath_start}[4]/td/table/tbody/tr[1]/td[2]/select[6]/option")[int(details["time"])])
        self.browser.click_element(self.browser.find_elements(f"{xpath_start}[4]/td/table/tbody/tr[2]/td[2]/select/option")[int(details["Bill collector"])-1])
        question = "".join(self.browser.get_value('id:mathExp').split()[:-1])
        valid_chars = "0123456789+-*/() "
        question_cleaned = ''.join([char for char in question if char in valid_chars])
        try:
            answer = eval(question_cleaned)
            self.browser.input_text("id:mathResult", answer)
        except SyntaxError as e:
            print(f"Error evaluating expression: {e}")
            return None
        self.browser.click_element("id:tCheck")
        self.browser.click_element("id:bill")
        self.browser.wait_until_element_is_visible("id:TransNo")
        Trans_No = self.browser.get_text("id:TransNo")
        return Trans_No

    def waiting(self, file_path):
        while not os.path.exists(file_path):
            continue
        return
    

class FileHandler:
    def __init__(self,folder_path,completed_directory):
        self.folder_path = folder_path
        self.completed_directory = completed_directory
    
    def list_file_names(self):
        return os.listdir(self.folder_path)
    
    def rename_and_move_file(self, old_file_name, new_file_name):
        current_file_path = f"{self.folder_path}/{old_file_name}"
        new_file_path = f"{self.completed_directory}/{new_file_name}"
        if not os.path.exists(self.completed_directory):
            os.makedirs(self.completed_directory)
        os.rename(current_file_path, new_file_path)

    
class PDFHandler:
    def __init__(self,folder_path):
        self.folder = folder_path
    
    def read_pdf(self, file_name):
        pdf_path = f"{self.folder}/{file_name}"
        pdf_File_Object = open(pdf_path, 'rb')
        pdf_Reader = PDF.PdfReader(pdf_File_Object)
        page_Object = pdf_Reader.pages[0]
        text = page_Object.extract_text()
        pdf_File_Object.close()
        return text

class Process:
    def __init__(self, url, download_directory,zip_directory, zip_file_path, completed_directory):
        self.url = url
        self.download_directory = download_directory
        self.zip_directory = f"{zip_directory}.zip"
        self.extracted_zip_directory = zip_directory
        self.zip_file_path = zip_file_path
        self.completed_directory = completed_directory
    
    def extract_transaction_details(self,text, file_name):
        months_dict = {"Jan": "1", "Feb": "2", "Mar": "2", "Apr": "4", "May": "5", "Jun": "6", "Jul": "7", "Aug": "8", "Sep": "9", "Oct": "10", "Nov": "11", "Dec": "12"}
        type_dict = {"DAE": "DAE.html", "DN": "DN.html", "EB": "index.html"}
        time_dict = {"AM": "0", "PM": "1"}
        bill_collector_dict = {"Charan Teja": "1", "Kiran Kumar": "2", "Ravi Rao": "3", "Rama Krishna":"4", "Suhas Babu": 5}
        lines = text.split("\n")
        details = {}
        type = ""
        number = ""
        nums = "1234567890"
        for char in file_name:
            if char not in nums:
                type = type + char
            else:
                number = number + char

        details["url_to_go"] = type_dict[type]
        details["Installment Number"] = number
        first_line = lines[-3].strip().split()
        if first_line[4] != "₹":
            details["Amount"] = first_line[4]
            details["Bill collector"] = bill_collector_dict[first_line[5] + " " + first_line[6]]
        else:
            details["Amount"] = first_line[5]
            details["Bill collector"] = bill_collector_dict[first_line[6] + " " + first_line[7]]


        transaction_no = lines[-2].strip().split("T")[-1]
        if " " in transaction_no:
            transaction_no = "".join(transaction_no.split())
        details["Transaction_Number"] = transaction_no
        date_time = lines[-1].strip()
        details["HH"] = date_time[:2]
        details["MM"] = date_time[3:5]
        details["time"] = time_dict[date_time[6:8]]
        details["DD"] = date_time[12:14]
        details["month"] =months_dict[date_time[15:18]]
        details["Year"] = date_time[-4:]
        return details
        
    def pay_power_bill(self):
        website = PayPowerBillSite(self.url, self.download_directory)
        website.download_file()

        zip_handler = ZipHandling(self.zip_directory, self.zip_file_path)
        zip_handler.extract_zip_file()

        file_handling = FileHandler(self.extracted_zip_directory, self.completed_directory)
        files_list = file_handling.list_file_names()

        pdf_handling = PDFHandler(self.extracted_zip_directory)
        for files in files_list:
            text = pdf_handling.read_pdf(files)
            details = self.extract_transaction_details(text, files[:-4])
            trans_no = website.fill_form(details)
            if trans_no:
                new_file_name = f"{files[:-4]}-{trans_no}.pdf"
                file_handling.rename_and_move_file(files, new_file_name)

        
if __name__ == "__main__":
    URL = "https://botsdna.com/PayPowerBill/"
    DOWNLOAD_DIRECTORY = "./PayPowerBill/download"
    COMPLETED_DIRECTORY = "./PayPowerBill/completed"
    ZIP_DIRECTORY = "./PayPowerBill/download/Payments"
    ZIP_FILE_PATH = "./PayPowerBill/download"

    process = Process(URL,DOWNLOAD_DIRECTORY, ZIP_DIRECTORY, ZIP_FILE_PATH, COMPLETED_DIRECTORY)
    process.pay_power_bill()

