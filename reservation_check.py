# Apr. 3, 2019
# This project is aimed at help automate the information collecting of the alumni of Shanghai Jiao Tong University and University of Michigan Joint institute from the linkedin information which is open to the public
# Fangzhe Li, freshman in Joint Institute, the student assistant of JI development office

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException,NoSuchAttributeException,NoSuchCookieException
from selenium.webdriver.common.keys import Keys
import time
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

    def run(self):
        self.driver = self.openChrome()
        self.login()
        self.direct_to_online_reservation_page()
        self.direct_to_chinese_go_abroad_study()
        self.direct_to_go_abroad_physical_exam()

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
        # resp = driver.get(login_url)
        # print ('resp={}'.format(resp))
        # elem = driver.find_element_by_id("login-email")
        # print("exception happened:{}".format(e))
        # elem.send_keys("ji-alumni@sjtu.edu.cn")
        # elem = driver.find_element_by_id("login-password")
        # elem.send_keys("UMSJTUJI2006")
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
            time.sleep(10)
            post_login_page = self.driver.current_url
            self.logger.debug('after sleep 10 seconds, post_login_page is {}'.format(post_login_page))
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
            time.sleep(10)
            post_page = self.driver.current_url
            self.logger.debug('after sleep 10 seconds, post_page is {}'.format(post_page))

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
            time.sleep(10)
            post_page = self.driver.current_url
            self.logger.debug('after sleep 10 seconds, post_page is {}'.format(post_page))

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
                time.sleep(10)
                post_page = self.driver.current_url
                self.logger.debug('after sleep 10 seconds, post_page is {}'.format(post_page))
                if 'MEC/user/mec' in post_page:
                    # click 已阅读 checkbox
                    #<div class="layui-unselect layui-form-checkbox" lay-skin="primary"><span>已阅读</span><i class="layui-icon layui-icon-ok"></i></div>
                    element = self.driver.find_element_by_class_name("layui-form-checkbox")
                    element.click()

                    # <button id="continue" class="layui-btn layui-btn-normal" lay-submit="">继续</button>
                    element = self.driver.find_element_by_id('continue')
                    element.click()
                else:
                    self.logger.error('wrong page -4')
                    exit(-4)
            else:
                self.logger.error('wrong page -3')
                exit(-3)
        except (NoSuchElementException,NoSuchAttributeException) as e:
            self.logger.error(e)










# direct to the page of reservation
def direct_to_reserve_page(driver,English_name_of_alumni):
    driver.get('https://www.linkedin.com')
    time.sleep(4)
    mainpage = driver.current_url
    if (mainpage[0:29]=='https://www.linkedin.com/feed'):
        try:
            elem = driver.find_element_by_xpath("//*[@id='nav-search-artdeco-typeahead']/artdeco-typeahead-deprecated-input/input")  # get to the search button
            elem.send_keys(English_name_of_alumni)
            elem.send_keys(Keys.ENTER)
        except Exception as e:
            print("exception happened in direct_to_reserve_page():{}".format(e))
        driver.implicitly_wait(20)
        time.sleep(3)
        # for link in driver.find_elements_by_xpath("(//*[@href])"):
        #     print (link.get_attribute('href'))
        element = driver.find_element_by_xpath("(//*[@href])[31]")
        # element = driver.find_element_by_class_name("search-results__list").find_elements_by_xpath("(//*[@href])[1]")
        # print("in direct_to_reserve_page: start to get url")
        url = element.get_attribute('href')
        starttime = time.time()
        driver.implicitly_wait(5)
        try:
            driver.get(url)
        except:
            print('！！！！！！time out after 10 seconds when loading page！！！！！！')
            driver.execute_script("window.stop()")
        endtime = time.time()
        print ('time elapsed:{} seconds'.format(endtime-starttime))
        print(driver.current_url)
        print('in direct_to_reserve_page:already direct to the personal page')
    else:
        url = 0
    return url




#
# # get the page source of the alumni page
# def get_page_source():
#     content = driver.page_source
#     return content
#
# #
# # def writeExcel(row, col, str, styl=Style.default_style):
# #     file = r'C:\Users\tople\.PyCharmCE2018.3\config\scratches\temporary.xls'
# #     rb = xlrd.open_workbook(file, formatting_info=True,ragged_rows=True)
# #     wb = copy(rb)
# #     ws = wb.get_sheet(format_dict['chart_sheet'])
# #     ws.write(row, col, str, styl)
# #     wb.save(file)
# #     rb.release_resources()
# #
# #
# # style = xlwt.easyxf('font: name Times New Roman')
# #
#
# def parse(content,url,index):
#     profile_txt = ''.join(re.findall('"data":{"\*profile":.*', content))
#     profile_txt = bytes(profile_txt, 'utf-8')
#     # print(profile_txt)
#     encode_type = chardet.detect(profile_txt)
#     profile_txt = profile_txt.decode(encode_type['encoding'])  # decode the text
#     firstname = re.findall('"firstName":"(.*?)"', profile_txt)
#     lastname = re.findall('"lastName":"(.*?)"', profile_txt)
#     if firstname and lastname:
#         print('Name: %s%s    Linkedin: %s' % (lastname[0], firstname[0], url))
#
#         occupation = re.findall('"headline":"(.*?)"', profile_txt)
#         if occupation:
#             print('Occupation: %s' % occupation[0])
#
#         locationName = re.findall('"locationName":"(.*?)"', profile_txt)
#         if locationName:
#             print('Location: %s' % locationName[0])
#
#         networkInfo_txt = ' '.join(
#             re.findall('({[^{]*?profile\.ProfileNetworkInfo"[^\}]*?\})', content))
#         connectionsCount = re.findall('"connectionsCount":(\d+)', networkInfo_txt)
#         if connectionsCount:
#             print('Connection number: %s' % connectionsCount[0])
#     # educations = re.findall('({[^{]*?profile\.Education"[^\}]*?\})', content)
#     educations = re.findall('"description.*?degreeUrn', content)
#     if educations:
#         print('Education:')
#     schoolNamelist=[]
#     fieldOfStudylist = []
#     degreeNamelist = []
#     company_name_list=[]
#     bachelor_degree_list=[]
#     bachelor_school_list = []
#     bachelor_field_list=[]
#     master_degree_list=[]
#     master_school_list = []
#     master_field_list = []
#     PhD_degree_list=[]
#     PhD_school_list = []
#     PhD_field_list = []
#     for one in educations:
#         schoolName = re.findall('"schoolName":"(.*?)"', one)
#         schoolName = ''.join(schoolName)
#         schoolNamelist.append(schoolName)
#         fieldOfStudy = re.findall('"fieldOfStudy":"(.*?)"', one)
#         fieldOfStudy = ''.join(fieldOfStudy)
#         fieldOfStudylist.append(fieldOfStudy)
#         degreeName = re.findall('"degreeName":"(.*?)"', one)
#         degreeName = ''.join(degreeName)
#         degreeNamelist.append(degreeName)
#         if schoolName:
#             fieldOfStudy = '   %s' % fieldOfStudy if fieldOfStudy else ''
#             degreeName = '   %s' % degreeName if degreeName else ''
#             print('     %s %s %s' % (schoolName,  fieldOfStudy, degreeName))               #screen show part
#     for i in range(len(educations)):
#             if degreeNamelist[i].find('B.S.')>=0 or degreeNamelist[i].find('本')>=0 or degreeNamelist[i].find('学')>=0 or degreeNamelist[i].find('Bachelor')>=0 or degreeNamelist[i].find('BS')>=0:
#                 bachelor_degree_list.append(degreeNamelist[i])
#                 bachelor_school_list.append(schoolNamelist[i])
#                 bachelor_field_list.append(fieldOfStudylist[i])
#             elif degreeNamelist[i].find('M')>=0 or degreeNamelist[i].find('硕')>=0:
#                 master_degree_list.append(degreeNamelist[i])
#                 master_school_list.append(schoolNamelist[i])
#                 master_field_list.append(fieldOfStudylist[i])
#             elif degreeNamelist[i].find('D')>=0 or degreeNamelist[i].find('博')>=0:
#                 PhD_degree_list.append(degreeNamelist[i])
#                 PhD_school_list.append(schoolNamelist[i])
#                 PhD_field_list.append(fieldOfStudylist[i])
#             else:
#                 pass
#     try:
#         if bachelor_school_list[0].find('Jiao')>=0 or bachelor_school_list[0].find('Michigan')>=0 or bachelor_school_list[0].find('交')>=0 or bachelor_school_list[0].find('SJTU')>=0:
#             print("this is our alumni")
#             writeExcel(index + 2, 24, 'finished', style)  # flag to identify if the information has been found
#             try:
#                 writeExcel(index + 2, format_dict['column_of_url_linkedin'], url, style)
#                 writeExcel(index + 2, 25, lastname[0] + firstname[0],style)                                                                           # to check if it is the alumni we are going to find
#                 writeExcel(index + 2, format_dict['column_of_position'], occupation[0], style)
#                 writeExcel(index + 2, format_dict['column_of_place'], locationName[0], style)
#                 writeExcel(index + 2, 26, bachelor_school_list[0], style)                                                                             # to check if it is our alumni
#                 writeExcel(index + 2, format_dict['column_of_master_degree_university'], master_school_list[0], style)
#                 writeExcel(index + 2, 6, "master", style)
#                 writeExcel(index + 2, format_dict['column_of_master_field'], master_field_list[0], style)
#                 try:
#                     writeExcel(index + 2, format_dict['column_of_PhD_degree_university'], PhD_school_list[0], style)
#                     writeExcel(index + 2, format_dict['column_of_PhD_field'], PhD_field_list[0], style)
#                     writeExcel(index + 2, 9, "PhD", style)
#                 except:
#                     position = re.findall('({[^{]*?"companyName"[^\}]*?\})', content)
#                     if position:
#                         print('Work Experience:')
#                     for one in position:
#                         companyName = re.findall('"companyName":"(.*?)"', one)
#                         company_name_list.append(companyName)
#                         try:
#                             writeExcel(index + 2, format_dict['column_of_company'], company_name_list[0],style)
#                             title = re.findall('"title":"(.*?)"', one)
#                             locationName = re.findall('"locationName":"(.*?)"', one)
#                             if companyName:
#                                 title = '   %s' % title[0] if title else ''
#                                 locationName = '   %s' % locationName[0] if locationName else ''
#                                 print('    %s %s %s' % (companyName[0], title, locationName))
#                         except:
#                             print('No company information is been found')
#             except:
#                 print('some information is missing here')
#         else:
#             print('this is not our alumni')
#             writeExcel(index + 2, 24, 'finished', style)  # flag to identify if the information has been found
#     except:
#         print('We cannot get his bachelor\'s degree information so it is need to be checked by human')
#         writeExcel(index + 2, 24, 'to be checked', style)
#         writeExcel(index + 2, format_dict['column_of_url_linkedin'], url, style)
#

if __name__ == '__main__':
    # alumnilist = alumni_list_input()
    # data = xlrd.open_workbook(r'C:\Users\tople\.PyCharmCE2018.3\config\scratches\temporary.xls',ragged_rows=True)
    # table = data.sheets()[format_dict['chart_sheet']]

    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)

    log_handler = logging.StreamHandler()
    formatter = logging.Formatter('%(asctime)s: %(levelname)s %(message)s')
    log_handler.setFormatter(formatter)
    logger.addHandler(log_handler)

    root_url = 'https://online.shhg12360.cn'
    instance = Webscrapper(root_url, logger)
    instance.run()
    #
    # url = direct_to_reserve_page(driver,alumnilist[j])
    #
    # for j in range(len(alumnilist)):
    #     # print(table.cell(j+2,24).value)
    #     if ((table.cell(j+2,24)).value!='finished') and (table.cell(j+2,24).value!='error')and (table.cell(j+2,24).value!='to be checked'):
    #         # print('before entering direct_to_alumni')
    #         url = direct_to_reserve_page(driver,alumnilist[j])
    #         # print('exit  direct_to_alumni')
    #         try:
    #             if (url!=0):
    #                 content = get_page_source()
    #                 # print(content)                  #for the test use
    #                 parse(content, url,j)
    #             # driver.quit()
    #         except TypeError:
    #             print('We cannot get the information of this alumni')
    #             writeExcel(j+2, 24, 'error', style)
    #             # driver.quit()
    #         except Exception as e:
    #             print("exception happened in main:{}".format(e))
    #         else:
    #             print("write the content into the excel successfully!\n\n")
    # data.release_resources()
