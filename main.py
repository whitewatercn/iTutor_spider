import time

import xlrd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains

# 待查询信息

# begin_year = input("请输入最早年份:")
current_timestamp = time.time()
time_tuple = time.localtime(current_timestamp)
end_year = time_tuple.tm_year

book = xlrd.open_workbook("setting/search.xls")
sheet = book.sheet_by_index(0)
author = sheet.cell_value(rowx=1,colx=0)
work_unit = sheet.cell_value(rowx=1,colx=1)

wait_time = 0.5 #等待时间
action_pixel = 100 #鼠标滚动像素

# get网站
wd = webdriver.Chrome(service=Service(r"tool/chromedriver"))
wd.get('https://kns.cnki.net/kns8/AdvSearch?dbcode=CFLS')
wd.implicitly_wait(2)

# 切换至专业检索
switch_majorsearch=wd.find_element(By.CSS_SELECTOR,'li[name="majorSearch"]')
switch_majorsearch.click()

# 输入检索式
time.sleep(wait_time)
switch_input_majorsearch = wd.find_element(By.CSS_SELECTOR,'.textarea-major')
search_text =  "AU = "+ "'"+ author +"' "+ "AND AF % "+ "'"+ work_unit +"'" 
switch_input_majorsearch.send_keys(search_text)
wd.find_element(By.CSS_SELECTOR,'.btn-search').click()


file = open("output.csv",'w',encoding='utf-8')
file.write('编号 论文标题 发表期刊 摘要 \n')
file.close()
article_number=0

while True:
 
    # 本页的所有结果
    time.sleep(2)
    main_window = wd.current_window_handle
    
    element = wd.find_element(By.CLASS_NAME,'result-table-list')
    buttons = element.find_elements(By.CLASS_NAME,'fz14')
    for button in buttons:
        actions = ActionChains(wd)
        actions.move_to_element(button).perform()
        button.click()
        # time.sleep(wait_time)
        # 切换到新窗口并获取信息，然后关闭窗口
        wd.switch_to.window(wd.window_handles[1])
        article_number+=1
        title= wd.find_element(By.CLASS_NAME,'wx-tit').find_element(By.CSS_SELECTOR,'h1').text
        # author_elements = wd.find_element(By.CSS_SELECTOR, '#authorpart').find_elements(By.CSS_SELECTOR, 'span')
        # author = ' '.join([element.text for element in author_elements])
        journal= wd.find_element(By.CLASS_NAME,'top-tip').find_element(By.CSS_SELECTOR,'span').find_element(By.CSS_SELECTOR,'a').text
        abstract= wd.find_element(By.CLASS_NAME,'abstract-text').text
        # keywords= wd.find_element(By.CLASS_NAME,'keywords').find_elements(By.CSS_SELECTOR,'a')
        # keyword = ' '.join([element.text for element in keywords])
        line = (str(article_number) + ',' + "'"+title +"'"+',' +"'"+ journal+"'" +',' +"'"+ abstract+"'" + '\n')

        # line = (str(article_number) + ' ' + title + ' ' + author + ' ' + journal + ' ' + abstract + ' ' + keywords + '\n')
        file=open('output.csv','a',encoding='utf-8')
        file.write(line)
        wd.close()
        wd.switch_to.window(main_window)
        # time.sleep(wait_time)
    try:
        # time.sleep(wait_time)
        switch_next = wd.find_element(By.ID,'PageNext').click()
    except NoSuchElementException:
        break


wd.quit()

