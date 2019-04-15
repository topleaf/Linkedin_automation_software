# Apr. 3, 2019
# This project is aimed at help automate the information collecting of the alumni of Shanghai Jiao Tong University and University of Michigan Joint institute from the linkedin information which is open to the public
# Fangzhe Li, freshman in Joint Institute, the student assistant of JI development office

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import re
import chardet
import pypinyin
import xlrd
import string
from xlutils.copy import copy
import xlwt
from xlwt import Style


format_dict = {'column_of_master_degree_university':4 ,'column_of_master_field':5 , 'column_of_PhD_degree_university':7 ,'column_of_PhD_field':8 ,'column_of_alumni_names':0, 'column_of_company':15,'column_of_position':16 ,'column_of_place':18,'column_of_url_linkedin':23,'chart_sheet':1}
# to input the excel of alumnilist to the list 'alumnilist' and to interpret the chinese names to their pinyin names
def alumni_list_input():
    data = xlrd.open_workbook(r'C:\Users\tople\.PyCharmCE2018.3\config\scratches\temporary.xls',ragged_rows=True)
    table = data.sheets()[format_dict['chart_sheet']]
    nrows = table.nrows
    alumnilist = []
    for i in range(nrows-2):
        temp = table.cell(i+2,format_dict['column_of_alumni_names']).value
        temp = "".join(temp)                                        # temp: list >>>str
        length_of_alumni_name = len(temp)
        if length_of_alumni_name == 3:
            last_name = temp[0]
            first_name = temp[1:]
            last_name = string.capwords(pypinyin.slug(last_name, separator=''))
            #this pinyin switch can' t deal with the situations that the same characters can have different pronunciations
            first_name = string.capwords(pypinyin.slug(first_name, separator=''))
            temp = first_name + ' ' + last_name
        elif length_of_alumni_name == 2:
            last_name = temp[0:1]
            first_name = temp[1:]
            last_name = string.capwords(pypinyin.slug(last_name, separator=''))
            first_name = string.capwords(pypinyin.slug(first_name, separator=''))
            temp = first_name + ' ' + last_name
        elif length_of_alumni_name == 4:
            last_name = temp[0:2]
            first_name = temp[2:]
            # here the situation that the person has a character for his last name while has three characters for his first name is not taken into account
            last_name = string.capwords(pypinyin.slug(last_name, separator=''))
            first_name = string.capwords(pypinyin.slug(first_name, separator=''))
            temp = first_name + ' ' + last_name
        else:
            print("nothing I can do in interpretation")
        alumnilist.append(temp)
    data.release_resources()
    return alumnilist


# open the chrome
def openChrome():
    # do the start setting
    option = webdriver.ChromeOptions()
    option.add_argument('disable-infobars')
    # open the chrome
    driver = webdriver.Chrome(chrome_options=option)
    return driver


# direct to the page of the alumni
def direct_to_alumni_page(driver,English_name_of_alumni):
    elem = driver.find_element_by_xpath("//*[@id='nav-search-artdeco-typeahead']/artdeco-typeahead-deprecated-input/input")  # get to the search button
    elem.send_keys(English_name_of_alumni)
    elem.send_keys(Keys.ENTER)
    time.sleep(3)
    # for link in driver.find_elements_by_xpath("(//*[@href])"):
    #     print (link.get_attribute('href'))
    element = driver.find_element_by_xpath("(//*[@href])[31]")     #hard code here
    url = element.get_attribute('href')
    driver.get(url)
    time.sleep(3)
    print('already direct to the personal page')
    return url


# login in
def login():
    linkedin_url = "https://www.linkedin.com"
    driver.get(linkedin_url)
    elem = driver.find_element_by_id("login-email")
    elem.send_keys("ji-alumni@sjtu.edu.cn")
    elem = driver.find_element_by_id("login-password")
    elem.send_keys("UMSJTUJI2006")
    driver.find_element_by_id("login-submit").click()
    driver.get_cookies()                #get the cookie of the website
    return


# get the page source of the alumni page
def get_page_source():
    content = driver.page_source
    return content


def writeExcel(row, col, str, styl=Style.default_style):
    file = r'C:\Users\tople\.PyCharmCE2018.3\config\scratches\temporary.xls'
    rb = xlrd.open_workbook(file, formatting_info=True,ragged_rows=True)
    wb = copy(rb)
    ws = wb.get_sheet(format_dict['chart_sheet'])
    ws.write(row, col, str, styl)
    wb.save(file)
    rb.release_resources()


style = xlwt.easyxf('font: name Times New Roman')


def parse(content,url,index):
    profile_txt = ''.join(re.findall('"data":{"\*profile":.*', content))
    profile_txt = bytes(profile_txt, 'utf-8')
    # print(profile_txt)
    encode_type = chardet.detect(profile_txt)
    profile_txt = profile_txt.decode(encode_type['encoding'])  # decode the text
    firstname = re.findall('"firstName":"(.*?)"', profile_txt)
    lastname = re.findall('"lastName":"(.*?)"', profile_txt)
    if firstname and lastname:
        print('Name: %s%s    Linkedin: %s' % (lastname[0], firstname[0], url))

        occupation = re.findall('"headline":"(.*?)"', profile_txt)
        if occupation:
            print('Occupation: %s' % occupation[0])

        locationName = re.findall('"locationName":"(.*?)"', profile_txt)
        if locationName:
            print('Location: %s' % locationName[0])

        networkInfo_txt = ' '.join(
            re.findall('({[^{]*?profile\.ProfileNetworkInfo"[^\}]*?\})', content))
        connectionsCount = re.findall('"connectionsCount":(\d+)', networkInfo_txt)
        if connectionsCount:
            print('Connection number: %s' % connectionsCount[0])
    educations = re.findall('({[^{]*?profile\.Education"[^\}]*?\})', content)
    if educations:
        print('Education:')
    schoolNamelist=[]
    fieldOfStudylist = []
    degreeNamelist = []
    company_name_list=[]
    bachelor_degree_list=[]
    bachelor_school_list = []
    bachelor_field_list=[]
    master_degree_list=[]
    master_school_list = []
    master_field_list = []
    PhD_degree_list=[]
    PhD_school_list = []
    PhD_field_list = []
    for one in educations:
        schoolName = re.findall('"schoolName":"(.*?)"', one)
        schoolName = ''.join(schoolName)
        schoolNamelist.append(schoolName)
        fieldOfStudy = re.findall('"fieldOfStudy":"(.*?)"', one)
        fieldOfStudy = ''.join(fieldOfStudy)
        fieldOfStudylist.append(fieldOfStudy)
        degreeName = re.findall('"degreeName":"(.*?)"', one)
        degreeName = ''.join(degreeName)
        degreeNamelist.append(degreeName)
        if schoolName:
            fieldOfStudy = '   %s' % fieldOfStudy if fieldOfStudy else ''
            degreeName = '   %s' % degreeName if degreeName else ''
            print('     %s %s %s' % (schoolName,  fieldOfStudy, degreeName))               #screen show part
    for i in range(len(educations)):
            if degreeNamelist[i].find('B.S.')>=0 or degreeNamelist[i].find('本')>=0 or degreeNamelist[i].find('学')>=0 or degreeNamelist[i].find('Bachelor')>=0 or degreeNamelist[i].find('BS')>=0:
                bachelor_degree_list.append(degreeNamelist[i])
                bachelor_school_list.append(schoolNamelist[i])
                bachelor_field_list.append(fieldOfStudylist[i])
            elif degreeNamelist[i].find('M')>=0 or degreeNamelist[i].find('硕')>=0:
                master_degree_list.append(degreeNamelist[i])
                master_school_list.append(schoolNamelist[i])
                master_field_list.append(fieldOfStudylist[i])
            elif degreeNamelist[i].find('D')>=0 or degreeNamelist[i].find('博')>=0:
                PhD_degree_list.append(degreeNamelist[i])
                PhD_school_list.append(schoolNamelist[i])
                PhD_field_list.append(fieldOfStudylist[i])
            else:
                pass
    try:
        if bachelor_school_list[0].find('Jiao')>=0 or bachelor_school_list[0].find('Michigan')>=0 or bachelor_school_list[0].find('交')>=0 or bachelor_school_list[0].find('SJTU')>=0:
            print("this is our alumni")
            writeExcel(index + 2, 24, 'finished', style)  # flag to identify if the information has been found
            try:
                writeExcel(index + 2, format_dict['column_of_url_linkedin'], url, style)
                writeExcel(index + 2, 25, lastname[0] + firstname[0],style)                                                                           # to check if it is the alumni we are going to find
                writeExcel(index + 2, format_dict['column_of_position'], occupation[0], style)
                writeExcel(index + 2, format_dict['column_of_place'], locationName[0], style)
                writeExcel(index + 2, 26, bachelor_school_list[0], style)                                                                             # to check if it is our alumni
                writeExcel(index + 2, format_dict['column_of_master_degree_university'], master_school_list[0], style)                                                                                #regular expression may help here
                writeExcel(index + 2, format_dict['column_of_master_field'], master_field_list[0], style)
                try:
                    writeExcel(index + 2, format_dict['column_of_PhD_degree_university'], PhD_school_list[0], style)
                    writeExcel(index + 2, format_dict['column_of_PhD_field'], PhD_field_list[0], style)
                except:
                    position = re.findall('({[^{]*?"companyName"[^\}]*?\})', content)
                    if position:
                        print('Work Experience:')
                    for one in position:
                        companyName = re.findall('"companyName":"(.*?)"', one)
                        company_name_list.append(companyName)
                        try:
                            writeExcel(index + 2, format_dict['column_of_company'], company_name_list[0],style)
                            title = re.findall('"title":"(.*?)"', one)
                            locationName = re.findall('"locationName":"(.*?)"', one)
                            if companyName:
                                title = '   %s' % title[0] if title else ''
                                locationName = '   %s' % locationName[0] if locationName else ''
                                print('    %s %s %s' % (companyName[0], title, locationName))
                        except:
                            print('No company information is been found')
            except:
                print('some information is missing here')
        else:
            print('this is not our alumni')
            writeExcel(index + 2, 24, 'finished', style)  # flag to identify if the information has been found
    except:
        print('We cannot get his bachelor\'s degree information so it is need to be checked by human')
        writeExcel(index + 2, 24, 'to be checked', style)


if __name__ == '__main__':
    alumnilist = alumni_list_input()
    data = xlrd.open_workbook(r'C:\Users\tople\.PyCharmCE2018.3\config\scratches\temporary.xls',ragged_rows=True)
    table = data.sheets()[format_dict['chart_sheet']]
    for j in range(len(alumnilist)):
        # print(table.cell(j+2,24).value)
        if ((table.cell(j+2,24)).value!='finished') and (table.cell(j+2,24).value!='error')and (table.cell(j+2,24).value!='to be checked'):
            driver = openChrome()
            login()
            try:
                url = direct_to_alumni_page(driver,alumnilist[j])
                content = get_page_source()
                # print(content)                  #for the test use
                parse(content, url,j)
                driver.quit()
            except TypeError:
                print('We cannot get the information of this alumni')
                writeExcel(j+2, 24, 'error', style)
                driver.quit()
            else:
                print("write the content into the excel successfully!\n\n")
    data.release_resources()
