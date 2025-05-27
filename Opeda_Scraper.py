import os
import shutil
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC


download_dir = r"C:\Users\dagarcia\Downloads"
temp_dir = r"D:\Python\PDE_Weekly\Temp"

class Opeda_Scraper:
    
    def __init__(self, url, download_dir):
        self.url = url
        self.download_dir = download_dir
        
        if not os.path.exists(download_dir):
            os.makedirs(download_dir)        
        # self.temp_dir = temp_dir
        
        self.clear_download_dir()
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option("prefs", {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        self.options.add_argument('--no-proxy-server')
        self.driver = webdriver.Chrome(options=self.options)
        self.driver.implicitly_wait(10)
        self.vars = {}

    def clear_download_dir(self):
        for filename in os.listdir(self.download_dir):
            file_path = os.path.join(self.download_dir, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f'Failed to delete {file_path}. Reason: {e}')
        os.makedirs(self.download_dir, exist_ok=True)
        
    def pull_product_data(self, prod):
        self.driver.get(self.url)
        
        element = self.driver.find_element(By.XPATH, f"//p[text()='{prod}']")

        parent_element = element.find_element(By.XPATH, "./ancestor::a")
        parent_element.click()
        time.sleep(5)
        
        download_button = self.driver.find_element(By.XPATH, "//button[@title='Download to Excel']")
        download_button.click() 
        time.sleep(5)     

        original_file = os.path.join(self.download_dir, "grid.csv")
        new_file = os.path.join(self.download_dir, f"{prod}.csv")
        os.rename(original_file, new_file)

        

        if hasattr(self, 'data'):
            csv_df = pd.read_csv(new_file)
            self.data = pd.concat([self.data, csv_df], ignore_index=True)  
        else:
            self.data = pd.read_csv(new_file)
       

    def close_driver(self):
        self.driver.quit()
        

opeda = Opeda_Scraper(r"https:\\opeda.intel.com\binsplit", temp_dir)

opeda.pull_product_data('PQFCV')
# opeda.pull_product_data('PQGCV')




