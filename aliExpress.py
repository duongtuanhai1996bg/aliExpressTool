from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook
import timeit
import time
import sys
chrome_driver = ".\\chromedriver.exe"
wb = Workbook()
from tkinter import *
import tkinter as tk

filepath = ''
pageNum = 0
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:" + "9022")
driver = webdriver.Chrome(chrome_driver, chrome_options=chrome_options)
# driver = webdriver.Chrome('chromedriver')  # type: WebDriver
def scroll(driver):
    driver.execute_script("window.scrollTo(0, 487)")
    time.sleep(0.1)
    driver.execute_script("window.scrollTo(488, 1487)")
    time.sleep(0.1)
    driver.execute_script("window.scrollTo(1488, 2487)")
def jumpPage(driver, pageNum):
    # jumpPagePath = "(//a[contains(text(),'" + str(pageNum) + "')])[10]"
    linkText = str(pageNum)
    if driver.find_elements_by_link_text(linkText):
        driver.find_element_by_link_text(linkText).click()
    else:
        print("Can not find products on page {0}! Please CHECK".format(pageNum))
        sys.exit("Exit tool")
def writeToExcel(sheet, row, filepath):
    # append all rows
    sheet.append(row)
    # save file
    wb.save(filepath)
    # print('  => Export data SUCCESSED')
    return sheet

def getDataFromPage(driver, sheet, page, count):
    process = 0
    href_elements = driver.find_elements_by_xpath(
        "//*[@id='node-gallery']/div[5]/div/div/ul/li//div[1]/div[1]/a")
    for x in href_elements:
        seller_sku = "MOTO-" + str(1000 + count)
        data = []
        parentage = ''
        arrayDes = []
        img_href = x.get_attribute("href")
        driver.execute_script("window.open('');")
        time.sleep(0.5)
        driver.switch_to.window(driver.window_handles[1])
        driver.get(img_href)
        driver.execute_script("window.scrollTo(0, 487)")
        time.sleep(0.1)
        driver.execute_script("window.scrollTo(488, 1300)")
        description = driver.find_elements_by_xpath('//*[@id="product-description"]/div')
        for x in description:
            str_all = x.get_attribute('innerHTML').replace('\n', '')
            str_cut,mid, end = str_all.partition('<p><span')
            arrayDes.append(str_cut)
            # print(arrayDes)
        if driver.find_elements_by_class_name('product-price-value'):
            cost_element = driver.find_elements_by_class_name('product-price-value')
            costArr = [x.text for x in cost_element]
            # print(costArr)
            if len(costArr) == 1:
                cost_old = costArr[0]
                # print("Price same!")
            else:
                cost_old = costArr[1]
            cost = costArr[0]
        else:
            cost = "this item is no longer available!"
            cost_old = "this item is no longer available!"
        name_element = driver.find_element_by_class_name('product-title')
        name = name_element.text

        if driver.find_elements_by_xpath('//*[@id="root"]/div/div[2]/div/div[1]/div/div/div[2]/ul/li/div/img'):
            img_src_elements = driver.find_elements_by_xpath(
                '//*[@id="root"]/div/div[2]/div/div[1]/div/div/div[2]/ul/li/div/img')
            imgArray = []
            for x in img_src_elements:
                imgArray.append(x.get_attribute("src").replace('_50x50.jpg', ''))
            # driver.execute_script("window.scrollTo(0, 487)")
            # time.sleep(0.1)
            # driver.execute_script("window.scrollTo(488, 1487)")
        else:
            imgArray =[]

        if driver.find_elements_by_class_name('sku-title'):
            check_sku = driver.find_element_by_class_name('sku-title').text
            parentage = "Parent"
            # parentage_sku = seller_sku
            data.append(name)
            data.append(seller_sku)
            data.append("")
            data.append("")
            data.append(cost)
            data.append(cost_old)
            data.append(parentage)
            data.append("")
            data.append("")
            data.append("")
            data.append("")
            data.append("")
            data.append("Color")
            data.append(arrayDes[0])
            data.extend(imgArray)
            # data.append(parentage_sku)
            sheet = writeToExcel(sheet, data, filepath)
            count += 1
            data = []
            if driver.find_elements_by_class_name("sku-property-text"):
                sku_child_elements_txt = driver.find_elements_by_xpath(
                    '//*[@id="root"]/div/div[2]/div/div[2]/div[7]/div/div/ul/li/div/span')
                for x in sku_child_elements_txt:
                    x.click()
                    if driver.find_elements_by_class_name('product-price-value'):
                        cost_element = driver.find_elements_by_class_name('product-price-value')
                        costArr = [x.text for x in cost_element]
                        # print(costArr)
                        if len(costArr) == 1:
                            cost_old = costArr[0]
                            # print("Price same!")
                        else:
                            cost_old = costArr[1]
                        cost = costArr[0]
                    else:
                        cost = "this item is no longer available!"
                        cost_old = "this item is no longer available!"
                    seller_sku_tmp = "MOTO-" + str(1000 + count)
                    size = x.text
                    # print(color)
                    data.append(name)
                    data.append(seller_sku_tmp)
                    data.append("")
                    data.append(size)
                    data.append(cost)
                    data.append(cost_old)
                    data.append("Child")
                    data.append(seller_sku)
                    data.append("")
                    if check_sku.find('ize') != -1:
                        data.append(size)
                    else:
                        data.append("")
                    data.append("Variation")
                    data.append("Update")
                    data.append("Color")
                    data.append(arrayDes[0])
                    data.extend(imgArray)
                    # data.append(parentage_sku)
                    writeToExcel(sheet, data, filepath)
                    count += 1
                    # print(count)
                    data = []
            if driver.find_elements_by_class_name("sku-property-image"):
                sku_child_elements = driver.find_elements_by_xpath(
                    '//*[@id="root"]/div/div[2]/div/div[2]/div/div/div/ul/li/div/img')
                for x in sku_child_elements:
                    seller_sku_tmp = "MOTO-" + str(1000 + count)
                    color = x.get_attribute('title')
                    img_src_tmp = x.get_attribute('src').replace('_50x50.jpg', '')
                    imgArray.insert(0, img_src_tmp)
                    # print(color)
                    data.append(name)
                    data.append(seller_sku_tmp)
                    data.append("")
                    data.append(color)
                    data.append(cost)
                    data.append(cost_old)
                    data.append("Child")
                    data.append(seller_sku)
                    data.append("")
                    data.append("")
                    data.append("Variation")
                    data.append("Update")
                    data.append("Color")
                    data.append(arrayDes[0])
                    data.extend(imgArray)
                    # data.append(parentage_sku)
                    writeToExcel(sheet, data, filepath)
                    imgArray.pop(0)
                    count += 1
                    # print(count)
                    data = []


        else:
            data.append(name)
            data.append(seller_sku)
            data.append("")
            data.append("")
            data.append(cost)
            data.append(cost_old)
            data.append(parentage)
            data.append("")
            data.append("")
            data.append("")
            data.append("")
            data.append("")
            data.append("")
            data.append(arrayDes[0])
            data.extend(imgArray)
            # data.append(parentage_sku)
            writeToExcel(sheet, data, filepath)
            count += 1
            # print(data)
        process += 1
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        print(" ==> LOAD {0}/{1} items on page {2}/{4} Page. Total: {3} items Loaded".format(process, len(href_elements), page, count, pageNum))
    return sheet, count
def main():
    start = timeit.default_timer()
    url = ''
    count = 0
    #GUI
    root = tk.Tk()
    root.minsize(width=840, height=190)
    root.title("Welcome")
    lbl = Label(root, text="Input your store's link: ", font=("Arial Bold", 20))
    lbl.pack()
    txt = Entry(root, width=130)
    txt.pack(ipady=10)
    def clicked():
        # global pageNum
        global filepath
        # pageNum = int(txt_page.get())
        url = txt.get()
        print(" Connecting to {0}".format(url))
        driver.get(url)
        if driver.find_elements_by_xpath('/html/head/meta[5]'):
            filepath = driver.find_element_by_xpath('/html/head/meta[5]').get_attribute('content') + ".xlsx"
        else:
            filepath = "Store.xlsx"
        # print(filepath)
        root.destroy()

    lbl_spc = Label(root, text=" ", font=("Arial Bold", 10))
    lbl_spc.pack()
    btn = Button(root,height=3, width=15, text="OK",font=("Arial Bold", 14), command=clicked)
    btn.pack()
    root.mainloop()

    sheet = wb.active
    row1 = ['Name', 'Seller SKU','REMOVE_Main Picture', 'Color Picture', 'Sale off Price', 'Price', 'Parentage', 'Parent SKU', ' ', 'size_name',' ',' ',' ', 'Description', 'Main image']
    sheet = writeToExcel(sheet, row1, filepath)
    # print(sheet)

    window_after_title = driver.title
    print(window_after_title)
    global pageNum
    pageNum = int(driver.find_elements_by_xpath('//*[@id="pagination-bottom"]/div[1]/a')[-2].text)
    scroll(driver)
    sheet, count = getDataFromPage(driver, sheet, 1, count)
    # titles.extend(name)
    # costs.extend(cost)
    # costs_old.extend(cost_old)
    # imagelink.extend(imgLink)
    # colors.extend(color)
    # sizes.extend(size)
    # writeToExcel(titles, costs, costs_old, imagelink, colors, sizes)
    # print(" ==> Data page 1 record Done!!")
    for i in range(2, pageNum+1):
        jumpPage(driver, i)
        scroll(driver)
        sheet, count = getDataFromPage(driver, sheet, i, count)
        # print(" ==> Data page {0} record Done!!".format(i))
    stop = timeit.default_timer()
    print("RUN TIME: {0} min".format((stop-start)/60))
    sys.exit("Exit tool")


main()
