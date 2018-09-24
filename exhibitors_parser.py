from time import sleep
import os

from selenium import webdriver
from bs4 import BeautifulSoup
from openpyxl import Workbook


dir = os.path.abspath(os.path.dirname(__file__))
url = "https://ces19.mapyourshow.com/7_0/alphalist.cfm?endrow=2103&alpha=*"
out_file_name = 'exhibitors_2018.xlsx'
out_file_dir = os.path.join(dir, 'data', out_file_name)
wait_time = 5
id_exhibitors = []


print("Start load page...")

driver = webdriver.Chrome()
driver.get(url)
#driver.minimize_window()

sleep(wait_time)

# loop of drop down web page
x, y = 0, 250
r = 0
for i in range(1000):
    try:
        driver.execute_script("window.scrollTo({}, {})".format(x, y))
        print("Drop down web page --> {}".format(r))
        r += 1
        x += 250
        y += 250
        sleep(wait_time)
    except:
        break

data = driver.page_source
bsObj = BeautifulSoup(data, "html.parser")

# get card list of all exhibitors
all_cards_exhibitors = bsObj.find("table", id="jq-regular-exhibitors").find("tbody")
for card in all_cards_exhibitors.find_all("tr"):
    id_exhibitors.append(card["data-exhid"])
print("Get all exhibitors card --> {}".format(len(id_exhibitors)))

driver.close()

# create out file sheet and headers
wb = Workbook()
ws = wb.create_sheet("exhibitors_2018")
ws.append([
    "Company",
    "Website",
    "Social Media",])

# scrape data from all card
print("Start get details from card ...")
n = 0
for id in id_exhibitors:
    try:
        print("Iteration --> {}".format(n))
        n += 1
        url_card = "https://ces19.mapyourshow.com/7_0/exhibitor/exhibitor-details.cfm?ExhID={}".format(id)
        driver = webdriver.Chrome()
        driver.get(url_card)
        sleep(wait_time)
        data = driver.page_source
        bsObj = BeautifulSoup(data, "html.parser")

        # get company name, website and social media from card
        company_name = bsObj.find("h1", class_="sc-Exhibitor_Name").text
        try:
            website = bsObj.find("p", class_="sc-Exhibitor_Url").find("a")["href"]
        except:
            website = ""
        try:
            social_media = bsObj.find("div", class_="sc-Exhibitor_SocialMedia").find("a")["href"]
        except:
            social_media = ""

        driver.close()

        # create line
        data_line = [company_name,
                     website,
                     social_media,
        ]
        ws.append(data_line)
        wb.save(out_file_dir)

    except:
        continue

print("Close WebDriver...")
print('Save Data in file {}'.format(out_file_dir))