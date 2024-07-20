from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

# Initialize WebDriver
service = Service(executable_path=r"C:\Users\gald1\Desktop\chromedriver-win64\chromedriver.exe")
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("detach", True)
chrome_options.add_argument("--start-maximized")

driver = webdriver.Chrome(service=service, options=chrome_options)

driver.get('https://he.aliexpress.com/?gatewayAdapt=glo2isr')



search_bar = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'search-words')))
search_button = driver.find_element(By.XPATH, '//*[@id="_full_container_header_23_"]/div[2]/div/div[1]/div/input[2]')
search_bar.send_keys('Car Scratches remover')
search_button.click()

# Wait for results to load
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.list--galleryWrapper--29HRJT4')))

item_names = []
item_sells = []
item_links = []
columns = ['Name', 'Sells', 'Links']

# Loop through pages
# Scroll to the bottom of the page to load more items
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(10)  # Wait for items to load

# Get item elements
div_wrapper_items = driver.find_element(By.CSS_SELECTOR, '.list--galleryWrapper--29HRJT4')
item_divs = div_wrapper_items.find_elements(By.CSS_SELECTOR, '.list--gallery--C2f2tvm.search-item-card-wrapper-gallery')

# Extract item details
for div in item_divs:
    try:
        title_div = div.find_element(By.CSS_SELECTOR, '.multi--title--G7dOCj3')
        item_title = title_div.get_attribute('title')
        orders_element = div.find_element(By.CSS_SELECTOR, '.multi--trade--Ktbl2jB')
        link_div = div.find_element(By.CSS_SELECTOR, '.multi--container--1UZxxHY.cards--card--3PJxwBm.search-card-item')
        link_text = link_div.get_attribute('href')
        item_links.append(link_text)
        item_names.append(item_title)
        item_sells.append(orders_element.text)
    except Exception as e:
        print(f"Error extracting item: {e}")
    
   

# Save to Excel
excel_type = pd.DataFrame(zip(item_names, item_sells, item_links), columns=columns)
excel_type.to_excel('ItemsAndSells.xlsx', index=False)
print(excel_type)

driver.quit()