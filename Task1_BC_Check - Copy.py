
#Author -- urajarat
##Task 1

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import os
import time
from time import sleep
from selenium.webdriver.support.ui import Select

driver = webdriver.Chrome(executable_path=r"C:\Users\Public\AppData\chromedriver.exe")
Asin = "B08G5Z8BTS"
driver.get("https://alaska-na.amazon.com/index.html")
time.sleep(3)
fnsku = driver.find_element_by_name('fnsku_simple')
fnsku.send_keys(Asin)                # Entering ASIN in ALASKA FNSKU Text Box
fnsku.send_keys(Keys.RETURN)
time.sleep(3)

rows = len(driver.find_elements_by_xpath("/html/body/table/tbody/tr[4]/td/div[2]/div/table[2]/tbody/tr"))
col = len(driver.find_elements_by_xpath("/html/body/table/tbody/tr[4]/td/div[2]/div/table[2]/tbody/tr[1]/th"))
# print(rows)
# print(col)
count = 0
s = 0
for r in range(4, rows - 3):
    table_value = driver.find_element_by_xpath("/html/body/table/tbody/tr[4]/td/div[2]/div/table[2]/tbody/tr["+str(r)+"]/td[1]").text
    # print(table_value)
    if(table_value == FC[1]):
        s = 0
        count += 1
        StockValue = driver.find_element_by_xpath("/html/body/table/tbody/tr[4]/td/div[2]/div/table[2]/tbody/tr["+str(r)+"]/td[13]").text
        if(int(StockValue) >= 1):
            print("There is Stock in Shippping FC ")
            print("The Stock Value in Shipping FC is " + str(StockValue))
        else:
            print("There is NO Stock in Shipping FC ")
    else:
        StockValue1 = driver.find_element_by_xpath("/html/body/table/tbody/tr[4]/td/div[2]/div/table[2]/tbody/tr["+str(r)+"]/td[13]").text
        StockValue2 = driver.find_element_by_xpath("/html/body/table/tbody/tr[4]/td/div[2]/div/table[2]/tbody/tr["+str(r+1)+"]/td[13]").text
        if(int(StockValue1) > int(s)):
            s = StockValue1
            Newr = r
        elif(int(StockValue2) > int(s)):
            s = StockValue2
            Newr = r
# print(s)
# print(Newr)
NewFC = driver.find_element_by_xpath("/html/body/table/tbody/tr[4]/td/div[2]/div/table[2]/tbody/tr["+str(Newr)+"]/td[1]").text
NewStockValue = s
NewFC1 = str(NewFC)
time.sleep(3)

if(int(s) == 0 & int(count) >= 1):
    print("There is Shipping FC in Alaska")
    FinalFC = FC[1]       # Shipping FC is sent for Bin Check Creation
elif(int(s) >= 1):
    print("New FC with highest Stock Value is :" + str(NewFC) + " and New Stock Value is : " + str(Newr))
    FinalFC = str(NewFC)          # New FC with highest Stock Value is sent for Bin Check Creation
elif(int(count) == 0 & int(s) == 0):
    print("There is Error on Alaska ")      # No FC is found on Alaska & Error in Alaska

time.sleep(5)
#
# # Bin Check For Respective FC
# driver.get("https://tt.amazon.com/")
# time.sleep(10)
# driver.refresh()
# time.sleep(3)
# driver.find_element_by_link_text("Create").click()
# time.sleep(10)
# driver.find_element_by_link_text("Manually select a CTI").click()
# time.sleep(3)
#
# # Selecting CTI
# Category = driver.find_element_by_id("category")
# cti = Select(Category)
# cti.select_by_visible_text('iss')
# time.sleep(3)
# Type = driver.find_element_by_id("type")
# cti = Select(Type)
# cti.select_by_index(153)
# time.sleep(5)
# Item = driver.find_element_by_id("item")
# cti = Select(Item)
# cti.select_by_index(7)
# time.sleep(5)
#
# # Entering the Shipping FC in Building ID Text Box
# building_id = driver.find_element_by_id("building_id")
# building_id.clear()
# time.sleep(3)
# building_id.send_keys(FinalFC)
# building_id.send_keys(Keys.RETURN)
# time.sleep(5)
#
# # Writing About Issue in the Short Description
#
# time.sleep(3)
# Description = driver.find_element_by_name("short_description")
# Description.clear()
# time.sleep(3)
# Description.send_keys("// Customer Issue is : " + CustomerIssue[1] + "\n ASIN : " + ASIN1)
# time.sleep(3)
#
# # Clicking On Submit Ticket Button
# SubmitTicket = driver.find_element_by_xpath("//*[@id='button_bar']/a[1]")
# SubmitTicket.click()
# time.sleep(5)
#
# #Bin Check Ticket Submitted
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
