# Apr. 3, 2019
# This project is aimed at help automate the information collecting of the alumni of Shanghai Jiao Tong University and University of Michigan Joint institute from the linkedin information which is open to the public
# Fangzhe Li, freshman in Joint Institute, the student assistant of JI development office

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException,NoSuchAttributeException, \
    StaleElementReferenceException,NoSuchCookieException
from selenium.webdriver.common.keys import Keys
import time
from pygame import mixer

import re
import chardet
import pypinyin
import xlrd
import string
import xlsxwriter
# from xlutils.copy import copy
# import xlwt
# from xlwt import Style
import logging
import subprocess

"""
prerequiste:
check Chrome version, make sure ChromeDriver and Chrome in your running environment are compatible
https://sites.google.com/a/chromium.org/chromedriver/downloads
"""

class Webscrapper:
    def __init__(self, url, logger):
        self.root_url = url
        self.driver = None
        self.logger = logger
        self.nearest_date = '09日09月2021年'
        self.too_late_date = '09日08月2021年'
        self.try_days = ['1','2','3','4','5']
        mixer.init(44100)
        pwd = subprocess.check_output('pwd').decode('utf-8')[:-1]
        sound_file = pwd+'/'+'bell.wav'
        self.logger.debug('sound_file:{}'.format(sound_file))
        try:
            self.sound = mixer.Sound(sound_file)
        except FileNotFoundError as e:
            self.logger.error('{}:{}'.format(e, sound_file))
            exit(-1)
            

    def run(self):
        self.driver = self.openChrome()
        self.login()
        self.direct_to_online_reservation_page()
        self.direct_to_chinese_go_abroad_study()
        self.direct_to_go_abroad_physical_exam()
        self.direct_to_reservation()
        i = 0
        self.too_late_day = self.too_late_date.split('日')[0]
        self.too_late_month = self.too_late_date.split('日')[1].split('月')[0]
        self.nearest_day = self.nearest_date.split('日')[0]
        self.nearest_month = self.nearest_date.split('日')[1].split('月')[0]

        self.logger.debug('too_late_date is {}月{}日,'
                          'nearest date is {}月{}日'.format(self.too_late_month,
                                                          self.too_late_day,
                                                          self.nearest_month,
                                                          self.nearest_day))
        while (self.nearest_month > self.too_late_month or
               (self.nearest_month == self.too_late_month and
                self.nearest_day > self.too_late_day)):
            i += 1
            _, i = divmod(i, 5)
            self.direct_to_select_date(self.try_days[i])
            self.check_nearest_date()
            self.nearest_day = self.nearest_date.split('日')[0]
            self.nearest_month = self.nearest_date.split('日')[1].split('月')[0]
            self.logger.debug('nearest date is {}月{}日'.format(self.nearest_month,
                                                              self.nearest_day))
        while True:
            self.sound.play()

        self.driver.quit()


    # open the chrome
    def openChrome(self):
        # do the start setting
        option = webdriver.ChromeOptions()
        option.add_argument('disable-infobars')
        # open the chrome
        driver = webdriver.Chrome(options=option)
        driver.set_page_load_timeout(20)
        driver.set_script_timeout(20)
        return driver


    def login(self):
        self.driver.delete_all_cookies()
        login_url = self.root_url+"/MEC/login"
        self.driver.get(login_url)
        try:
            elem = self.driver.find_element_by_id("loginName")
            elem.send_keys("shhg_li")
            elem = self.driver.find_element_by_id("password")
            elem.send_keys("Cfwlsqydckyhgl7@")
            # elem.send_keys(Keys.ENTER)
            # <button class="layui-btn layui-btn-normal" lay-submit="" lay-filter="loginByName">登录</button>
            elem = self.driver.find_element_by_class_name('layui-btn-normal')
            elem.click()
            self.cookies = self.driver.get_cookies()                #get the cookies of the website
            """
            [<class 'dict'>: {'domain': 'online.shhg12360.cn', 'expiry': 1659165409, 'httpOnly': True, 
            'name': '__jsluid_s', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': 
            '70d1565482c319870c4849fb8e82ebbc'},
            <class 'dict'>: {'domain': 'online.shhg12360.cn', 'httpOnly': True, 'name': 'JSESSIONID', 
            'path': '/MEC', 'secure': False, 'value': '71F54CA5416B4EF9719A2CF97DCDCB83'}
            ]
                {
                    'domain': cookie['domain'],
                    'name': cookie['name'],
                    'value': cookie['value'],
                    'path': cookie['path'],
                    'expires': None
                }
            """
            for cookie in self.cookies:
                self.driver.add_cookie(cookie)
        except Exception as e:
            self.logger.error("exception happened in login():{}".format(e))
        else:
            self.logger.info('login succeed')
        return



    # direct to online reservation
    def direct_to_online_reservation_page(self):
        # self.driver.page_source contains the raw html
        try:
            post_login_page = self.driver.current_url
            self.logger.debug('post_login_page is {}'.format(post_login_page))
            time.sleep(5)
            post_login_page = self.driver.current_url
            self.logger.debug('after sleep 5 seconds, post_login_page is {}'.format(post_login_page))
            if '/MEC/user/mec/recordList' in post_login_page:
                # <li><a href="/MEC/user/mec/choose">在线预约</a></li>
                # <a href="/MEC/user/mec/choose">在线预约</a>
                # elements = self.driver.find_elements_by_tag_name('a')
                # for element in elements:
                #     value = element.get_attribute('href')
                #     if '在线预约' in value:
                #         element.click()
                #         break

                element = self.driver.find_element_by_partial_link_text('在线预约')
                element.click()
                time.sleep(2)
                # <div class="layui-unselect layui-form-checkbox layui-form-checked"
                # lay-skin="primary"><span>已阅读</span><i class="layui-icon layui-icon-ok"></i></div>
                element = self.driver.find_element_by_class_name("layui-form-checkbox")
                element.click()

                # <button id="continue" class="layui-btn layui-btn-normal" lay-submit="">继续</button>
                element = self.driver.find_element_by_id('continue')
                element.click()
            else:
                self.logger.error('not in correct page')
                exit(-1)
        except (NoSuchElementException,NoSuchAttributeException) as e:
            self.logger.error(e)


    def direct_to_chinese_go_abroad_study(self):
        try:
            post_page = self.driver.current_url
            self.logger.debug('post_page is {}'.format(post_page))
            time.sleep(5)
            post_page = self.driver.current_url
            self.logger.debug('after sleep 5 seconds, post_page is {}'.format(post_page))

            if '/MEC/user/mec/choose' in post_page:
                # select '中国公民出境留学'
                # <a class="layui-btn layui-btn-sm  layui-btn-normal "
                # href="javascript:void(0)" style="white-space: normal !important;"
                # onclick="Mec_choose.refreshWeb('16','1')">中国公民出境留学</a>
                elements = self.driver.find_elements_by_class_name('layui-btn-normal')
                for element in elements:
                    # onclick_value = element.get_attribute('onclick')
                    # if "Mec_choose.refreshWeb('16','1')" in onclick_value:
                    if element.text == '中国公民出境留学':
                        element.click()
                        break
            else:
                self.logger.error('wrong page')
                exit(-2)
        except (NoSuchElementException,NoSuchAttributeException) as e:
            self.logger.error(e)

    def direct_to_go_abroad_physical_exam(self):
        try:
            post_page = self.driver.current_url
            self.logger.debug('post_page is {}'.format(post_page))
            time.sleep(5)
            post_page = self.driver.current_url
            self.logger.debug('after sleep 5 seconds, post_page is {}'.format(post_page))

            if '/MEC/user/mec/choose?id=16' in post_page:
            # select 'jinbanglu'
            # <a class="layui-btn layui-btn-sm layui-btn-normal"
            # href="javascript:void(0)" style="white-space: normal !important;"
            # onclick="Mec_choose.chooseRes(4)">出境留学体检（地址：长宁区金浜路15号）</a>
                elements = self.driver.find_elements_by_class_name('layui-btn-normal')
                for element in elements:
                    if '长宁区金浜路15号' in element.text:
                        element.click()
                        break


                post_page = self.driver.current_url
                self.logger.debug('post_page is {}'.format(post_page))
                time.sleep(3)
                post_page = self.driver.current_url
                self.logger.debug('after sleep 3 seconds, post_page is {}'.format(post_page))
                if 'MEC/user/mec' in post_page:
                    # click 已阅读 checkbox
                    #<div class="layui-unselect layui-form-checkbox" lay-skin="primary"><span>已阅读</span><i class="layui-icon layui-icon-ok"></i></div>
                    element = self.driver.find_element_by_class_name("layui-form-checkbox")
                    element.click()

                    # <button id="continue" class="layui-btn layui-btn-normal" lay-submit="">继续</button>
                    element = self.driver.find_element_by_id('continue')
                    element.click()
                else:
                    self.logger.error('no prompt page, maybe right,wrong page -4')
                    # exit(-4)
            else:
                self.logger.error('wrong page -3')
                exit(-3)
        except (NoSuchElementException,NoSuchAttributeException) as e:
            self.logger.error(e)


    def direct_to_reservation(self):
        try:
            post_page = self.driver.current_url
            self.logger.debug('post_page is {}'.format(post_page))
            time.sleep(2)
            post_page = self.driver.current_url
            self.logger.debug('after sleep 2 seconds, post_page is {}'.format(post_page))

            if '/MEC/user/mec' in post_page:        # 展示个人信息，待确认
                #<button class="layui-btn layui-btn-normal"
                # lay-submit="" lay-filter="mec">继续</button>
                element = self.driver.find_element_by_class_name('layui-btn-normal')
                if element.text =='继续':
                    element.click()
                else:
                    logger.error('unsync')
            else:
                self.logger.error('wrong page -5')
                exit(-5)
        except (NoSuchElementException,NoSuchAttributeException) as e:
            self.logger.error(e)

    def direct_to_select_date(self,target_date):
        try:
            post_page = self.driver.current_url
            self.logger.debug('post_page is {}'.format(post_page))
            time.sleep(5)
            post_page = self.driver.current_url
            self.logger.debug('after sleep 5 seconds, post_page is {}'.format(post_page))

            if '/MEC/user/mec' in post_page:
                #<input type="text" id="numDate" name="numDate"
                # placeholder="请选择预约日期" autocomplete="off" class="layui-input"
                # lay-filter="numDate" lay-vertype="tips" lay-key="1">
                element = self.driver.find_element_by_id('numDate')
                element.click()
                time.sleep(5)

                # element = self.driver.find_element_by_class_name('laydate-btns-now')
                # if element.text == '现在':        #select 现在 button
                #     element.click()
                # time.sleep(2)

                #<td lay-ymd="2021-8-6" class="laydate-day-next">6</td>
                elements = self.driver.find_elements_by_class_name('laydate-day-next')
                for element in elements:        #select target_date
                    self.logger.debug('element.text={}'.format(element.text))
                    if element.text == target_date:
                        element.click()
                        break
                time.sleep(2)

                # #<button class="layui-btn layui-btn-normal"
                # # lay-submit="" lay-filter="res">提交</button>
                # element = self.driver.find_element_by_class_name('layui-btn-normal')
                # if element.text == '提交':        #select button
                #     element.click()
                # time.sleep(2)

                element = self.driver.find_element_by_class_name('layui-layer-btn0')
                if element.text == '确定':        #select ‘确定’ trying to go to nearest date webpage
                    element.click()

            else:
                self.logger.error('wrong page -6')
                exit(-6)
        except (NoSuchElementException,NoSuchAttributeException,StaleElementReferenceException) as e:
            self.logger.error(e)


    def check_nearest_date(self):
        try:
            post_page = self.driver.current_url
            self.logger.debug('post_page is {}'.format(post_page))
            time.sleep(3)
            post_page = self.driver.current_url
            self.logger.debug('after sleep 3 seconds, post_page is {}'.format(post_page))

            if '/MEC/user/mec' in post_page:
                element = self.driver.find_element_by_id('numDate')
                # get actual value of the text input,which is filled by js.
                self.nearest_date = element.get_attribute('value')
                self.logger.info('nearest date is {}'.format(self.nearest_date))
            else:
                self.logger.error('wrong page -7')
                exit(-7)
        except (NoSuchElementException,NoSuchAttributeException,
                StaleElementReferenceException) as e:
            self.logger.error(e)







if __name__ == '__main__':
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)

    log_handler = logging.StreamHandler()
    formatter = logging.Formatter('%(asctime)s: %(levelname)s %(message)s')
    log_handler.setFormatter(formatter)
    logger.addHandler(log_handler)

    root_url = 'https://online.shhg12360.cn'
    instance = Webscrapper(root_url, logger)
    instance.run()
