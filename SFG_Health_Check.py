from selenium import webdriver
import time
from tqdm import tqdm
import os
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from PIL import Image
from io import BytesIO
import pytesseract
from datetime import datetime, timedelta
import win32com.client
import logging

def MBIStatus():
    global MBI
    driver = webdriver.Chrome(
        executable_path="\\path\\to\\chromedriver.exe")
    try:
        driver.get("https://example.com") // organization site
        time.sleep(10)
        check1 = driver.find_elements_by_id("isc_G")
        print(len(check1))
        if len(check1) == 0:
            check1 = "Red"
            driver.quit()
        else:
            check1 = "Green"
            driver.quit()
        driver = webdriver.Chrome(
            executable_path="\\path\\to\\chromedriver.exe", options=option)
        driver.get("https://example.com") // organization site
        time.sleep(10)
        check2 = driver.find_elements_by_name("submit")
        print(len(check2))
        if len(check2) == 0:
            driver.quit()
            check2 = "Red"
        else:
            driver.quit()
            check2 = "Green"
        if check1 == "Green" and check2 == "Green":
            MBI = "Healthy - Both the links are accessible"
        else:
            MBI = "Red - Kindly validate the MBI links"
    finally:
        driver.quit()


def NodeStatus():
    global nodehealth
    driver = webdriver.Chrome(
        executable_path="\\path\\to\\chromedriver.exe", options=option)
    try:
        driver.get("https://example.com") // organization site
        driver.implicitly_wait(10)
        time.sleep(3)
        driver.find_element_by_name("autho").send_keys("userid")
        driver.find_element_by_name("password").send_keys("password")
        driver.find_element_by_name("submit").click()
        time.sleep(3)
        driver.find_element_by_link_text("Operations").click()
        time.sleep(3)
        driver.find_element_by_link_text("System").click()
        time.sleep(3)
        driver.find_element_by_link_text("Cluster").click()
        time.sleep(3)
        driver.find_element_by_link_text("Node Status").click()
        time.sleep(3)
        driver.switch_to.frame("basefrm")
        time.sleep(3)
        node1 = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td/form/table/tb"
                                             "ody/tr[2]/td[2]/table[3]/tbody/tr[6]/td[12]").text
        node1 = node1.strip()
        node2 = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td/form/ta"
                                             "ble/tbody/tr[2]/td[2]/table[3]/tbody/tr[7]/td[12]").text
        node2 = node2.strip()
        driver.quit()
        print(node1)
        print(node2)
        if node1 == "Active" and node2 == "Active":
            nodehealth = "Healthy (Node1 - %s, Node2 - %s)" %(node1, node2)
        else:
            nodehealth = "Red (Node1 - %s, Node2 - %s)" %(node1, node2)
    except Exception as e:
        driver.quit()
        logging.exception(str(e))
        nodehealth = "Not accessible"


def AdaptersCheck():
    global adapterstatus
    driver = webdriver.Chrome(
        executable_path="\\path\\to\\chromedriver.exe", options=option)
    try:
        driver.get("https://example.com") // organization site
        driver.implicitly_wait(10)
        time.sleep(3)
        driver.find_element_by_name("autho").send_keys("userid")
        driver.find_element_by_name("password").send_keys("password")
        driver.find_element_by_name("submit").click()
        time.sleep(3)
        driver.find_element_by_link_text("Deployment").click()
        time.sleep(3)
        driver.find_element_by_link_text("Services").click()
        time.sleep(3)
        driver.find_element_by_link_text("Configuration").click()
        time.sleep(3)
        driver.switch_to.frame("basefrm")
        time.sleep(3)
        driver.find_element_by_id("autoCompleteServiceNameSource").send_keys("Goo")
        time.sleep(3)
        driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td/table/tb"
                                     "ody/tr[2]/td[2]/table/tbody/tr[15]/td[3]/input").click()
        time.sleep(3)
        DropDown = driver.find_element_by_name("aserver")
        Select(DropDown).select_by_visible_text("node1")
        time.sleep(5)
        Images1 = list(driver.find_elements_by_css_selector("img[src$='selected_btn_dep.gif']"))
        Images1 = (len(Images1))  #9
        Select(DropDown).select_by_visible_text("node2")
        time.sleep(5)
        Images2 = list(driver.find_elements_by_css_selector("img[src$='selected_btn_dep.gif']"))
        Images2 = (len(Images2))  #8
        driver.quit()
        print(Images1, Images2)
        if Images1 == 10 and Images2 == 9:
            adapterstatus = "Healthy (Active Node1 adapters - %s, Active Node2 adapters - %s)" %(Images1, Images2)
        else:
            adapterstatus = "Red Active Node1 adapters - %s, Active Node2 adapters - %s)" %(Images1, Images2)
    except Exception as e:
        driver.quit()
        logging.exception(str(e))
        adapterstatus = "Not accessible"

def PerimeterServerStatus():
    global perimeterstatus
    driver = webdriver.Chrome(
        executable_path="\\path\\to\\chromedriver.exe", options=option)
    try:
        driver.get("https://example.com") // organization site
        driver.implicitly_wait(10)
        time.sleep(3)
        driver.find_element_by_name("autho").send_keys("userid")
        driver.find_element_by_name("password").send_keys("password")
        driver.find_element_by_name("submit").click()
        time.sleep(3)
        driver.find_element_by_link_text("Operations").click()
        time.sleep(3)
        driver.find_element_by_link_text("System").click()
        time.sleep(3)
        driver.find_element_by_link_text("Troubleshooter").click()
        time.sleep(3)
        driver.switch_to.frame("basefrm")
        time.sleep(3)
        DropDown = driver.find_element_by_name("aserver")
        Select(DropDown).select_by_visible_text("node1")
        window_before = driver.current_window_handle
        driver.find_element_by_link_text("[ Perimeter Server Status ]").click()
        for handle in driver.window_handles:
            if handle != window_before:
                window_after = handle
                break
        time.sleep(5)
        driver.switch_to.window(window_after)
        time.sleep(3)
        Status1 = driver.find_element_by_xpath("/html/body/table[1]/tbody/tr/td/table/tbody/t"
                                               "r[8]/td[2]/table/tbody/tr[7]/td[4]").text
        Status1 = Status1.strip() #Enabled
        Status2 = driver.find_element_by_xpath("/html/body/table[1]/tbody/tr/td/tab"
                                               "le/tbody/tr[8]/td[2]/table/tbody/tr[8]/td[4]").text
        Status2 = Status2.strip() #Enabled
        driver.close()
        driver.switch_to.window(window_before)
        driver.switch_to.frame("basefrm")
        Select(DropDown).select_by_visible_text("node2")
        time.sleep(5)
        driver.find_element_by_link_text("[ Perimeter Server Status ]").click()
        for handle in driver.window_handles:
            if handle != window_before:
                window_after = handle
                break
        time.sleep(5)
        driver.switch_to.window(window_after)
        time.sleep(1)
        Status3 = driver.find_element_by_xpath(
            "/html/body/table[1]/tbody/tr/td/table/tbody/tr[8]/td[2]/table/tbody/tr[7]/td[4]").text
        Status3 = Status3.strip()  # Enabled
        Status4 = driver.find_element_by_xpath(
            "/html/body/table[1]/tbody/tr/td/table/tbody/tr[8]/td[2]/table/tbody/tr[8]/td[4]").text
        Status4 = Status4.strip()  # Enabled
        driver.close()
        driver.quit()
        print(Status2, Status1, Status3, Status4)
        if Status1 == "Enabled" and Status2 == "Enabled":
            node1 = "Enabled"
        else:
            node1 = "Red"
        if Status3 == "Enabled" and Status4 == "Enabled":
            node2 = "Enabled"
        else:
            node2 = "Red"
        if node1 == "Enabled" and node2 == "Enabled":
            perimeterstatus = "Healthy (Node1 - %s, Node2 - %s)" %(node1, node2)
        else:
            perimeterstatus = "Red (Node1 - %s, Node2 - %s)" %(node1, node2)
    except Exception as e:
        driver.quit()
        logging.exception(str(e))
        perimeterstatus = "Not accessible"


def myFileGateWay():
    global Status
    driver = webdriver.Chrome(
        executable_path="\\path\\to\\chromedriver.exe", options=option)
    try:
        driver.get("https://example.com") // organization site
        driver.implicitly_wait(10)
        time.sleep(5)
        driver.find_element_by_name("userName").send_keys("userid")
        driver.find_element_by_name("password").send_keys("password")
        driver.find_element_by_name("password").send_keys(Keys.ENTER)
        time.sleep(3)
        driver.find_element_by_xpath("//div[text()='Tools']").click()
        time.sleep(3)
        driver.find_element_by_xpath("//div[text()='Activity Snapshot']").click()
        time.sleep(3)
        element = driver.find_element_by_xpath("//td[text()='Activity Snapshot']")
        location = element.location
        size = element.size
        png = driver.get_screenshot_as_png()
        im = Image.open(BytesIO(png))
        left = location['x'] + 140
        top = location['y'] + 50
        right = location['x'] + size['width'] + 120
        bottom = location['y'] + size['height'] + 45
        im = im.crop((left, top, right, bottom))  # defines crop points
        im.save('screenshot.png')  # saves new cropped image
        im = Image.open('screenshot.png')
        pytesseract.pytesseract.tesseract_cmd = 'C:\\path\\to\\tesseract.exe'
        text = pytesseract.image_to_string(im, config='--psm 6').split()
        file = open("value.txt", "w")
        file.write(text[0])
        file.close()
        Number = open("value.txt", "r").read().split()[0]
        print(str(Number))
        if str(Number[-1]) not in (r"1|2|3|4|5|6|7|8|9|0"):
            Number = str(Number[0:-1])
            print(Number)
        if str(Number[0]) not in (r"1|2|3|4|5|6|7|8|9|0"):
            Number = str(Number[1:])
            print(Number)
        print(Number)
        success = driver.find_elements_by_xpath("//div[text()='%s']" %Number)
        if len(success) == 0:
            driver.quit()
            time.sleep(120)
            myFileGateWay()
        else:
            driver.find_element_by_xpath("//div[text()='%s']" % Number).click()
            time.sleep(3)
            driver.find_element_by_css_selector("img[src$='/Window/close.png']").click()
            element = driver.find_element_by_xpath("//div[text()='Discovery Time']")
            location = element.location
            size = element.size
            png1 = driver.get_screenshot_as_png()
            im1 = Image.open(BytesIO(png1))
            left = location['x']
            top = location['y'] + 20
            right = location['x'] + size['width'] + 50
            bottom = location['y'] + size['height'] + 25
            im1 = im1.crop((left, top, right, bottom))  # defines crop points
            im1.save('screenshot1.png')  # saves new cropped image
            im1 = Image.open('screenshot1.png')
            text = pytesseract.image_to_string(im1, config='--psm 6')
            print(text)
            if str(text[-1]) not in (r"1|2|3|4|5|6|7|8|9|0"):
                text = str(text[0:-1])
                print(text)
            if str(text[0]) not in (r"1|2|3|4|5|6|7|8|9|0"):
                text = str(text[1:])
                print(text)
            driver.quit()
            if text >= ReqD:
                Status = "Healthy - Last successful file transfer (%s)" % text
            else:
                Status = "Red - Last successful file transfer (%s)" % text
    except Exception as e:
        driver.quit()
        logging.exception(str(e))
        Status = "Not accessible"

def MBIStatus1():
    global MBI1
    driver = webdriver.Chrome(
        executable_path="\\path\\to\\chromedriver.exe")
    try:
        driver.get("https://example.com") // organization site
        driver.implicitly_wait(15)
        time.sleep(10)
        check1 = driver.find_elements_by_id("isc_G")
        print(len(check1))
        if len(check1) == 0:
            check1 = "Red"
            driver.quit()
        else:
            check1 = "Green"
            driver.quit()
        driver = webdriver.Chrome(
            executable_path="\\path\\to\\chromedriver.exe", options=option)
        time.sleep(10)
        driver.get("https://example.com") // organization site
        driver.implicitly_wait(15)
        time.sleep(10)
        check2 = driver.find_elements_by_name("submit")

        print(len(check2))
        if len(check2) == 0:
            driver.quit()
            check2 = "Red"
        else:
            driver.quit()
            check2 = "Green"
        if check1 == "Green" and check2 == "Green":
            MBI1 = "Healthy - Both the links are accessible"
        else:
            MBI1 = "Red - Kindly validate the MBI links"
    finally:
        driver.quit()


def NodeStatus1():
    global nodehealth1
    driver = webdriver.Chrome(
        executable_path="\\path\\to\\chromedriver.exe", options=option)
    try:
        driver.get("https://example.com") // organization site
        driver.implicitly_wait(10)
        time.sleep(3)
        driver.find_element_by_name("autho").send_keys("userid")
        driver.find_element_by_name("password").send_keys("password")
        driver.find_element_by_name("submit").click()
        time.sleep(3)
        driver.find_element_by_link_text("Operations").click()
        time.sleep(3)
        driver.find_element_by_link_text("System").click()
        time.sleep(3)
        driver.find_element_by_link_text("Cluster").click()
        time.sleep(3)
        driver.find_element_by_link_text("Node Status").click()
        time.sleep(3)
        driver.switch_to.frame("basefrm")
        time.sleep(3)
        node1 = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td/form/table/tb"
                                             "ody/tr[2]/td[2]/table[3]/tbody/tr[6]/td[12]").text
        node1 = node1.strip()
        node2 = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td/form/ta"
                                             "ble/tbody/tr[2]/td[2]/table[3]/tbody/tr[7]/td[12]").text
        node2 = node2.strip()
        driver.quit()
        print(node1)
        print(node2)
        if node1 == "Active" and node2 == "Active":
            nodehealth1 = "Healthy (Node1 - %s, Node2 - %s)" %(node1, node2)
        else:
            nodehealth1 = "Red (Node1 - %s, Node2 - %s)" %(node1, node2)
    except Exception as e:
        driver.quit()
        logging.exception(str(e))
        nodehealth1 = "Not accessible"


def AdaptersCheck1():
    global adapterstatus1
    driver = webdriver.Chrome(
        executable_path="\\path\\to\\chromedriver.exe", options=option)
    try:
        driver.get("https://example.com") // organization site
        driver.implicitly_wait(10)
        time.sleep(3)
        driver.find_element_by_name("autho").send_keys("userid")
        driver.find_element_by_name("password").send_keys("password")
        driver.find_element_by_name("submit").click()
        time.sleep(3)
        driver.find_element_by_link_text("Deployment").click()
        time.sleep(3)
        driver.find_element_by_link_text("Services").click()
        time.sleep(3)
        driver.find_element_by_link_text("Configuration").click()
        time.sleep(3)
        driver.switch_to.frame("basefrm")
        time.sleep(3)
        driver.find_element_by_id("autoCompleteServiceNameSource").send_keys("LFG")
        time.sleep(3)
        driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td/table/tb"
                                     "ody/tr[2]/td[2]/table/tbody/tr[15]/td[3]/input").click()
        time.sleep(3)
        DropDown = driver.find_element_by_name("aserver")
        Select(DropDown).select_by_visible_text("node1")
        time.sleep(5)
        Images1 = list(driver.find_elements_by_css_selector("img[src$='selected_btn_dep.gif']"))
        Images1 = (len(Images1))  #9
        Select(DropDown).select_by_visible_text("node2")
        time.sleep(5)
        Images2 = list(driver.find_elements_by_css_selector("img[src$='selected_btn_dep.gif']"))
        Images2 = (len(Images2))  #8
        driver.quit()
        print(Images1, Images2)
        if Images1 == 11 and Images2 == 9:
            adapterstatus1 = "Healthy (Active Node1 adapters - %s, Active Node2 adapters - %s)" %(Images1, Images2)
        else:
            adapterstatus1 = "Red (Active Node1 adapters - %s, Active Node2 adapters - %s)" %(Images1, Images2)
    except Exception as e:
        driver.quit()
        logging.exception(str(e))
        adapterstatus1 = "Not accessible"

def PerimeterServerStatus1():
    global perimeterstatus1
    driver = webdriver.Chrome(
        executable_path="\\path\\to\\chromedriver.exe", options=option)
    try:
        driver.get("https://example.com") // organization site
        driver.implicitly_wait(10)
        time.sleep(3)
        driver.find_element_by_name("autho").send_keys("userid")
        driver.find_element_by_name("password").send_keys("password")
        driver.find_element_by_name("submit").click()
        time.sleep(3)
        driver.find_element_by_link_text("Operations").click()
        time.sleep(3)
        driver.find_element_by_link_text("System").click()
        time.sleep(3)
        driver.find_element_by_link_text("Troubleshooter").click()
        time.sleep(3)
        driver.switch_to.frame("basefrm")
        time.sleep(3)
        DropDown = driver.find_element_by_name("aserver")
        Select(DropDown).select_by_visible_text("node1")
        window_before = driver.current_window_handle
        driver.find_element_by_link_text("[ Perimeter Server Status ]").click()
        for handle in driver.window_handles:
            if handle != window_before:
                window_after = handle
                break
        time.sleep(5)
        driver.switch_to.window(window_after)
        time.sleep(3)
        Status1 = driver.find_element_by_xpath("/html/body/table[1]/tbody/tr/td/table/tbody/t"
                                               "r[8]/td[2]/table/tbody/tr[7]/td[4]").text
        Status1 = Status1.strip() #Enabled
        Status2 = driver.find_element_by_xpath("/html/body/table[1]/tbody/tr/td/tab"
                                               "le/tbody/tr[8]/td[2]/table/tbody/tr[8]/td[4]").text
        Status2 = Status2.strip() #Enabled
        driver.close()
        driver.switch_to.window(window_before)
        driver.switch_to.frame("basefrm")
        Select(DropDown).select_by_visible_text("node2")
        time.sleep(5)
        driver.find_element_by_link_text("[ Perimeter Server Status ]").click()
        for handle in driver.window_handles:
            if handle != window_before:
                window_after = handle
                break
        time.sleep(5)
        driver.switch_to.window(window_after)
        time.sleep(1)
        Status3 = driver.find_element_by_xpath(
            "/html/body/table[1]/tbody/tr/td/table/tbody/tr[8]/td[2]/table/tbody/tr[9]/td[4]").text
        Status3 = Status3.strip()  # Enabled
        Status4 = driver.find_element_by_xpath(
            "/html/body/table[1]/tbody/tr/td/table/tbody/tr[8]/td[2]/table/tbody/tr[8]/td[4]").text
        Status4 = Status4.strip()  # Enabled
        driver.close()
        driver.quit()
        print(Status2, Status1, Status3, Status4)
        if Status1 == "Enabled" and Status2 == "Enabled":
            node1 = "Enabled"
        else:
            node1 = "Red"
        if Status3 == "Enabled" and Status4 == "Enabled":
            node2 = "Enabled"
        else:
            node2 = "Red"
        if node1 and node2 == "Enabled":
            perimeterstatus1 = "Healthy (Node1 - %s, Node2 - %s)" %(node1, node2)
        else:
            perimeterstatus1 = "Red (Node1 - %s, Node2 - %s)" %(node1, node2)
    except Exception as e:
        driver.quit()
        logging.exception(str(e))
        perimeterstatus1 = "Not accessible"


def myFileGateWay1():
    global Status5
    driver = webdriver.Chrome(
        executable_path="\\path\\to\\chromedriver.exe", options=option)
    try:
        driver.get("https://example.com") // organization site
        driver.implicitly_wait(10)
        time.sleep(5)
        driver.find_element_by_name("userName").send_keys("userid")
        driver.find_element_by_name("password").send_keys("password")
        driver.find_element_by_name("password").send_keys(Keys.ENTER)
        time.sleep(3)
        driver.find_element_by_xpath("//div[text()='Tools']").click()
        time.sleep(3)
        driver.find_element_by_xpath("//div[text()='Activity Snapshot']").click()
        time.sleep(3)
        element = driver.find_element_by_xpath("//td[text()='Activity Snapshot']")
        location = element.location
        size = element.size
        png = driver.get_screenshot_as_png()
        im = Image.open(BytesIO(png))
        left = location['x'] + 140
        top = location['y'] + 50
        right = location['x'] + size['width'] + 120
        bottom = location['y'] + size['height'] + 45
        im = im.crop((left, top, right, bottom))  # defines crop points
        im.save('screenshot.png')  # saves new cropped image
        im = Image.open('screenshot.png')
        pytesseract.pytesseract.tesseract_cmd = 'C:\\path\\to\\tesseract.exe'
        text = pytesseract.image_to_string(im, config='--psm 6').split()
        file = open("value.txt", "w")
        file.write(text[0])
        file.close()
        Number = open("value.txt", "r").read().split()[0]
        print(str(Number))
        if str(Number[-1]) not in (r"1|2|3|4|5|6|7|8|9|0"):
            Number = str(Number[0:-1])
            print(Number)
        if str(Number[0]) not in (r"1|2|3|4|5|6|7|8|9|0"):
            Number = str(Number[1:])
            print(Number)
        print(Number)
        success = driver.find_elements_by_xpath("//div[text()='%s']" %Number)
        if len(success) == 0:
            driver.quit()
            time.sleep(120)
            myFileGateWay1()
        else:
            driver.find_element_by_xpath("//div[text()='%s']" % Number).click()
            time.sleep(3)
            driver.find_element_by_css_selector("img[src$='/images/Window/close.png']").click()
            element = driver.find_element_by_xpath("//div[text()='Discovery Time']")
            location = element.location
            size = element.size
            png1 = driver.get_screenshot_as_png()
            im1 = Image.open(BytesIO(png1))
            left = location['x']
            top = location['y'] + 20
            right = location['x'] + size['width'] + 50
            bottom = location['y'] + size['height'] + 25
            im1 = im1.crop((left, top, right, bottom))  # defines crop points
            im1.save('screenshot1.png')  # saves new cropped image
            im1 = Image.open('screenshot1.png')
            text = pytesseract.image_to_string(im1, config='--psm 6')
            print(text)
            if str(text[-1]) not in (r"1|2|3|4|5|6|7|8|9|0"):
                text = str(text[0:-1])
                print(text)
            if str(text[0]) not in (r"1|2|3|4|5|6|7|8|9|0"):
                text = str(text[1:])
                print(text)
            driver.quit()
            if text >= ReqD:
                Status5 = "Healthy - Last successful file transfer (%s)" % text
            else:
                Status5 = "Red - Last successful file transfer (%s)" % text
    except Exception as e:
        driver.quit()
        logging.exception(str(e))
        Status5 = "Not accessible"

def statusemail():
    Sign = open("C:\\signature\\KP Sign.html")
    HTML = Sign.read()
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail_item = outlook.CreateItem(0)
    mail_item.To = 'EmailAddress'
    mail_item.Cc = "CCEmailAddress"
    mail_item.Subject = "WebConsole/Server Health Check - %s" %ReqD1
    body1 = """<html lang="en">
<head>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous">
    <title>Table</title>
    <style>.nodehealth {color: %s;} .Status {color: %s;} .perimeterstatus {color: %s;} .adapterstatus {color: %s;} .MBI {color: %s;} .status {color: green;}</style>
</head>
<body>
    <p class="lead">Hi Team,</p>
    <p class="lead">Kindly find the health check report below,</p>
    <table class="table">
        
        <thead  class="thead-dark">
            <tr>
                <th scope="col">Sl. No</th>
                <th scope="col">Parameters Checked</th>
                <th scope="col">Status</th>
            </tr>
        </thead>
        <tr>
            <th scope="row" >1.</th>
            <td>SFG Node check</td>
            <td class = 'nodehealth'>%s</td>
        </tr>
        <tr>
            <th scope="row" >2.</th>
            <td>Tools - Activity Snapshot</td>
            <td class = 'Status'>%s</td>
        </tr>
        <tr>
            <th scope="row" >3.</th>
            <td>Perimeter Server Status</td>
            <td class = 'perimeterstatus'>%s</td>
        </tr>
        <tr>
            <th scope="row" >4.</th>
            <td>SFG- Adapters Check</td>
            <td class = 'adapterstatus'>%s</td>
        </tr>
        <tr>
            <th scope="row" >5.</th>
            <td>SFG – MBI availability</td>
            <td class = 'MBI'>%s</td>
        </tr>
        <tr>
            <th scope="row" >6.</th>
            <td>OOM/CPU Utilization</td>
            <td class = 'status'>Healthy</td>
        </tr>
        <br><caption>Prod Health Check Report</caption>
    </table>
    
    <br>
    <p></p>
    <script src="app.js"></script>
</body>
</html>""" \
           %(color1, color2, color3, color4, color5, nodehealth, Status, perimeterstatus,adapterstatus, MBI)

    body2 = """<html lang="en">
<head>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous">
    <title>Table</title>
    <style>.nodehealth {color: %s;} .Status {color: %s;} .perimeterstatus {color: %s;} .adapterstatus {color: %s;} .MBI {color: %s;} .status {color: green;}</style>
</head>
<body>
    <table class="table">
        
        <thead  class="thead-dark">
            <tr>
                <th scope="col">Sl. No</th>
                <th scope="col">Parameters Checked</th>
                <th scope="col">Status</th>
            </tr>
        </thead>
        <tr>
            <th scope="row" >1.</th>
            <td>SFG Node check</td>
            <td class = 'nodehealth'>%s</td>
        </tr>
        <tr>
            <th scope="row" >2.</th>
            <td>Tools - Activity Snapshot</td>
            <td class = 'Status'>%s</td>
        </tr>
        <tr>
            <th scope="row" >3.</th>
            <td>Perimeter Server Status</td>
            <td class = 'perimeterstatus'>%s</td>
        </tr>
        <tr>
            <th scope="row" >4.</th>
            <td>SFG- Adapters Check</td>
            <td class = 'adapterstatus'>%s</td>
        </tr>
        <tr>
            <th scope="row" >5.</th>
            <td>SFG – MBI availability</td>
            <td class = 'MBI'>%s</td>
        </tr>
        <tr>
            <th scope="row" >6.</th>
            <td>OOM/CPU Utilization</td>
            <td class = 'status'>Healthy</td>
        </tr>
        <br><caption>PreProd Health Check Report</caption>
    </table>
    
    <br>
</body>
</html>%s""" \
           %(color6, color7, color8, color9, color10, nodehealth1, Status5, perimeterstatus1, adapterstatus1, MBI1,  HTML)
    mail_item.HTMLBody = (body1 + body2)
    mail_item.Send()

option = Options()
option.add_argument("--headless")

ReqD = datetime.strftime(datetime.now() - timedelta(hours=1), '%m/%d/%Y %H:%M:%S')
ReqD1 = datetime.strftime(datetime.now(), '%m/%d/%Y')
print(ReqD)
logging.basicConfig(filename= os.path.join("\\\\path\\to\\savelogs",
                                           "SFGHealthCheck_error.log"), level=logging.ERROR)
MBIStatus()
NodeStatus()
AdaptersCheck()
PerimeterServerStatus()
myFileGateWay()
MBIStatus1()
NodeStatus1()
AdaptersCheck1()
PerimeterServerStatus1()
myFileGateWay1()

if "Healthy" not in nodehealth:
    color1 = "red"
elif "Not" in nodehealth:
    color1 = "red"
else:
    color1 = "Green"

if "Healthy" not in Status:
    color2 = "red"
elif "Not" in Status:
    color2 = "red"
else:
    color2 = "green"

if "Healthy" not in perimeterstatus:
    color3 = "red"
elif "Not" in perimeterstatus:
    color3 = "red"
else:
    color3 = "green"

if "Healthy" not in adapterstatus:
    color4 = "red"
elif "Not" in adapterstatus:
    color4 = "red"
else:
    color4 = "green"

if "Healthy" not in MBI:
    color5 = "red"
elif "Not" in MBI:
    color5 = "red"
else:
    color5 = "green"

if "Healthy" not in nodehealth1:
    color6 = "red"
elif "Not" in nodehealth1:
    color6 = "red"
else:
    color6 = "green"

if "Healthy" not in Status5:
    color7 = "red"
elif "Not" in Status5:
    color7 = "red"
else:
    color7 = "green"

if "Healthy" not in perimeterstatus1:
    color8 = "red"
elif "Not" in perimeterstatus1:
    color8 = "red"
else:
    color8 = "green"

if "Healthy" not in adapterstatus1:
    color9 = "red"
elif "Not" in adapterstatus1:
    color9 = "red"
else:
    color9 = "green"

if "Healthy" not in MBI1:
    color10 = "red"
elif "Not" in MBI1:
    color10 = "red"
else:
    color10 = "green"

statusemail()
