# importing library
from selenium import webdriver
import openpyxl

# managing chrome driver
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome("C:/chromedriver/chromedriver.exe",options=options)
driver.maximize_window()

# url of the product Apple iPhone 12 Pro Max 256 GB Pacific Blue from amazon, flipkart, paytmmall.
amazon_url = "https://www.amazon.in/dp/B08L5T31M6?th=1"
flipkart_url = "https://www.flipkart.com/apple-iphone-12-pro-max-pacific-blue-256-gb/p/itm3a0860c94250e?pid=MOBFWBYZ8STJXCVT&lid=LSTMOBFWBYZ8STJXCVT0OKDMO&marketplace=FLIPKART&q=iphone+12+pro+max&store=tyy%2F4io&srno=s_1_2&otracker=AS_QueryStore_OrganicAutoSuggest_1_9_na_na_na&otracker1=AS_QueryStore_OrganicAutoSuggest_1_9_na_na_na&fm=SEARCH&iid=f2c9d338-43fd-495f-92e2-b77c936cacf6.MOBFWBYZ8STJXCVT.SEARCH&ppt=sp&ppn=sp&ssid=dnarke0deo0000001627350256832&qH=5a7a12c4a730c1af"
paytmmall_url = "https://paytmmall.com/apple-iphone-12-pro-max-256-gb-pacific-blue-CMPLXMOBAPPLE-IPHONEDUMM202561B6D39AD-pdp?product_id=338975481&sid=64d9b644-02e7-431a-8169-d9f9a45124e8&src=consumer_search&svc=-1&cid=66781&tracker=autosuggest%7C%7Ciphone%2012%20pro%20max%7Cgrid%7CSearch_experimentName%3Ddemographics_location%23NA_gender%23NA%7C%7C3%7Cdemographics_location%23NA_gender%23NA&get_review_id=333122484"

# best to buy from
best_price = 0
best_website = ''
url = ''

# open workbook
work_book = openpyxl.Workbook()
worksheet = work_book.active

worksheet['B1'] = "Amazon Details"
worksheet['C1'] = "Flipkart Details"
worksheet['D1'] = "Paytmmall Details"
worksheet['A2'] = "Product Name"
worksheet['A3'] = "Product Price"
worksheet['A4'] = "Best to Buy From"
worksheet['A5'] = "specifications of the product"

# to get product details from amazon
def get_amazon_details():
    global best_price, best_website, url

    driver.get(amazon_url)
    driver.implicitly_wait(2)
    amazon_name = driver.find_element_by_id("productTitle").text
    amazon_price = driver.find_element_by_id("priceblock_ourprice").text
    amazon_specifications = driver.find_element_by_id("feature-bullets").text
    img = openpyxl.drawing.image.Image(r"C:\Users\hp\Desktop\python project 1\amazon.jpg")
    img.height = 250
    img.width = 250
    
    worksheet['B2'] = amazon_name
    worksheet['B3'] = amazon_price
    worksheet['B5'] = amazon_specifications
    img.anchor = 'B9'
    worksheet.add_image(img)
    best_website = "Amazon"
    best_price = float(amazon_price[1:].replace(',', ''))
    url = amazon_url
    print(url)

# to get product details from flipkart
def get_flipkart_details():
    global best_price, best_website, url

    driver.get(flipkart_url)
    driver.implicitly_wait(2)
    flipkart_name = driver.find_element_by_class_name("B_NuCI").text
    flipkart_price = driver.find_element_by_class_name("_25b18c").text
    flipkart_specifications = driver.find_element_by_class_name("_2418kt").text
    img = openpyxl.drawing.image.Image(r"C:\Users\hp\Desktop\python project 1\flipkart.jpeg")
    img.height = 250
    img.width = 250
    
    worksheet['C2'] = flipkart_name
    worksheet['C3'] = flipkart_price
    worksheet['C5'] = flipkart_specifications
    img.anchor = 'C9'
    worksheet.add_image(img)
    if float(flipkart_price[1:].replace(',', '')) < best_price:
        best_price = float(flipkart_price.text[1:])
        best_website = "Flipkart"
        url = flipkart_url

# to get product details from paytmmall
def get_paytmmall_details():
    global best_website, best_price, url

    driver.get(paytmmall_url)
    driver.implicitly_wait(2)
    paytmmall_name = driver.find_element_by_class_name("NZJI").text
    paytmmall_price = driver.find_element_by_class_name("_1V3w").text
    paytmmall_specifications = driver.find_element_by_class_name("_1a-K").text
    img = openpyxl.drawing.image.Image(r"C:\Users\hp\Desktop\python project 1\paymmall.jpg")
    img.height = 250
    img.width = 250

    worksheet['D2'] = paytmmall_name
    worksheet['D3'] = paytmmall_price
    worksheet['D5'] = paytmmall_specifications
    img.anchor = 'D9'
    worksheet.add_image(img)
    if float(paytmmall_price.replace(',', '')) < best_price:
        best_price = float(paytmmall_price)
        best_website = "Paytmmall"
        url = paytmmall_url

get_amazon_details()
get_flipkart_details()
get_paytmmall_details()

# update details to excel
worksheet.merge_cells('B4:D4')
worksheet['B4'] = str(best_website) + ' Rs. ' + str(best_price) + "   URL:   " + url 

# save the workbook
file_name = 'product_details.xlsx'
work_book.save(file_name)

print(f"File saved as {file_name} \nDone!")

# close work book and chrome browser
work_book.close()
driver.quit()
