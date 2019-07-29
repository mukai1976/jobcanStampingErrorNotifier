#coding:UTF-8

#libraries needs to be installed
#selenium, pyyaml, slackclient, bs4, lxml
# and phantomjs

# get ChromeDriver from here
# https://sites.google.com/a/chromium.org/chromedriver/downloads

from __future__ import absolute_import, division, print_function

import sys
import json
import re

import datetime
import time

import urllib
import numpy

from selenium import webdriver
from selenium.webdriver.support.events import EventFiringWebDriver
from selenium.webdriver.support.events import AbstractEventListener
from selenium.webdriver.support.select import Select

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException

from bs4 import BeautifulSoup
import json
import yaml
import os
from slackclient import SlackClient
from time import sleep
#from datetime import datetime,date,timedelta

#FOR REAL USE set this to be True to hide Chrome screen
HEADLESSNESS = True

#defalut value
SLACK_TOKEN = ''
SLACK_USER_ID = ''
SLACK_CHANNEL: ''

#loading credentials
args = sys.argv
# credentials_mukai.yaml
with open(args[1],"r") as stream:
    try:
        credentials = yaml.load(stream, Loader=yaml.SafeLoader)
        globals().update(credentials)
    except yaml.YAMLError as exc:
        print(exc)

# SlackClient get
sc = SlackClient(SLACK_TOKEN)
_slack_users_list = sc.api_call(
"users.list"
)
filtered_members = list(filter(lambda x: (x.get('deleted') == False) and (x.get('is_bot') == False),_slack_users_list.get('members')))
slack_users_list = []
for members in filtered_members:
    _id = members[u'id']
    _real_name = members[u'real_name'].replace(" ", "") .replace("　", "")
    #print(members[u'name'],members[u'real_name'])
    if _id != "USLACKBOT":
        slack_users_list.append((_id,_real_name))

class ScreenshotListener(AbstractEventListener):
    #count for error screenshots
    exception_screenshot_count = 0

    def on_exception(self, exception, driver):
        screenshot_name = "00_exception_{:0>2}.png".format(ScreenshotListener.exception_screenshot_count)
        ScreenshotListener.exception_screenshot_count += 1
        driver.get_screenshot_as_file(screenshot_name)
        print("Screenshot saved as '%s'" % screenshot_name)

def makeDriver(*, headless=True):
    options = Options()
    if(headless):
        options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1280,800')
    _driver = webdriver.Chrome(options=options)
    return EventFiringWebDriver(_driver, ScreenshotListener())

def loginJobcan(driver):
    url = JC_URL

    driver.get(url)
    driver.implicitly_wait(5)

    userId_box = driver.find_element_by_name('client_login_id')
    managerId_box = driver.find_element_by_name('client_manager_login_id')
    pass_box = driver.find_element_by_name('client_login_password')
    userId_box.send_keys(JC_LOGINID)
    managerId_box.send_keys(JC_MANAGER_LOGINID)
    pass_box.send_keys(JC_PASSWORD)

    #driver.save_screenshot('0before login.png')
    #print( "saved before login" )

    #login
    driver.find_element_by_css_selector('body > div > div:nth-child(1) > form > div:nth-child(5) > button').click()

    #driver.save_screenshot('1after login.png')
    #print( "saved after login" )
    #print("URL:" + driver.current_url)

    #StampinError-table
    ActionChains(driver).move_to_element(driver.find_element_by_css_selector('#adit-manage-step > a')).perform()
    #driver.save_screenshot('2after StampinError-table.png')

    #待機が必要なので
    time.sleep(1)
    driver.find_element_by_css_selector('#adit-manage-menu > ul > li:nth-child(3) > dl > dd > ul > li:nth-child(1) > a').click()

    #group_id select
    driver.find_element_by_xpath("//*[@id='group_id']/option[@value=" + JC_GROUPID + "]").click()
    driver.find_element_by_css_selector('#mainForm > div > a').click()

    #driver.save_screenshot('3after StampinError-table.png')

    return driver

#打刻エラー一覧を取得して[{staff:スタッフ, errordate:日時, contents:内容}, ...] の形式で返す
def getStampinError(driver):

    stampingerror_items = []
    driver.implicitly_wait(2)
    trs = driver.find_element_by_xpath("//*[@id='wrap-basic-shift-table']/div").find_elements(By.TAG_NAME, "tr")
    for i in range(1,len(trs)):
        tds = trs[i].find_elements(By.TAG_NAME, "td")
        for j in range(0,len(tds)):
            if j < len(tds):
                #スタッフ
                if j == 0:
                    staff = tds[j].text
                #日時
                elif j == 1:
                    errordate = tds[j].text
                #内容
                elif j == 2:
                    contents = tds[j].text

        stampingerror_items.append((staff, errordate, contents))

    return stampingerror_items

################## main starts here ##################################
if __name__ == "__main__":
    print( "【start】" + SLACK_USER_ID + " " + str(datetime.datetime.now()))

    sc = SlackClient(SLACK_TOKEN)

    driver = makeDriver(headless=HEADLESSNESS)
    #print( 'driver created' )

    try:

        loginJobcan(driver)

        stampingerror_items = getStampinError(driver)

    except:
        print("Unexpected error:", sys.exc_info()[0])
        raise
    finally:
        if HEADLESSNESS:
            driver.quit()


    message = ""
    for stampingerror_item in stampingerror_items:
        staff, errordate, contents = stampingerror_item
        _staff = staff.replace(" ", "").replace("　", "")
        three_days_ago = datetime.datetime.now() - datetime.timedelta(days=3)
        dt_cnv_errordate = datetime.datetime.strptime(errordate, '%Y/%m/%d')

        #3日以前のみ出力
        if dt_cnv_errordate <= three_days_ago:
            slack_userid = ""
            slack_users = numpy.array(slack_users_list)
            r = numpy.where(slack_users[:,1]==_staff)
            slack_userid = slack_users[r,0][0][0]
            if slack_userid != "":
                sc.api_call(
                    "chat.postMessage",
                    channel=slack_userid,
                    text= "%s %s %s" % ('{:　<8}'.format(staff), '{:　<11}'.format(errordate), contents) + "\r\n",
                    username="jobcan 3日経過している打刻エラー※修正申請済み含む",
                    user=SLACK_USER_ID,
                    link_names=1
                )
            message = message + "%s %s %s" % ('{:　<8}'.format(staff), '{:　<11}'.format(errordate), contents) + "\r\n"

    sc.api_call(
       "chat.postMessage",
       channel=SLACK_CHANNEL,
       text=message,
       username="jobcan 3日経過している打刻エラー※修正申請済み含む",
       user=SLACK_USER_ID,
       link_names=1
    )
    print( "【end  】" + SLACK_USER_ID + " " + str(datetime.datetime.now()))
