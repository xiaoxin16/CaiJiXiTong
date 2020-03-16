#! /usr/bin/python
# -*- coding: utf8 -*-

import re, time, ipdb, math, socket
from multiprocessing import Pool, Lock, Manager
from urllib.parse import urlparse
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Alignment


# 1.
def isIP(str):
    p = re.compile('^((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$')
    if p.match(str):
        return True
    else:
        return False


# 2. url formalization 标准化
def get_url_normalize_single(url):
    default_scheme = "http"
    if urlparse(url).scheme == '':
        url_new = default_scheme + "://" + url
    else:
        url_new = url
    if urlparse(url_new).hostname is None:
        url_new = "异常"
    elif "." not in urlparse(url_new).hostname:
        url_new = "异常"
    elif urlparse(url_new).hostname[0] is ".":
        url_new = "异常"
    elif len(urlparse(url_new).hostname) < 4:
        url_new = "异常"
    else:
        {}
    url_site = url_new
    return url_site


# 2. 非空IP 列表
def update_task_excel(da, conf_data, add_head):
    wb = load_workbook(conf_data["dst"] + "/" + conf_data["fn"])
    ws = wb.active
    alig_s = Alignment(horizontal='left', vertical='center')
    god_sheet_name = '核查结果_正常'
    bad_sheet_name = '核查结果_异常'
    if god_sheet_name not in wb.sheetnames:
        ws_1 = wb.create_sheet(god_sheet_name, 0)
    else:
        ws_1 = wb[god_sheet_name]
    if bad_sheet_name not in wb.sheetnames:
        ws_2 = wb.create_sheet(bad_sheet_name, 1)
    else:
        ws_2 = wb[bad_sheet_name]
    index_1 = 1
    index_2 = 1
    if add_head == 1:
        for i in range(len(conf_data["title"])):
            ws_1.cell(row=1, column=i + 1).value = conf_data["title"][i]
            ws_1.cell(row=1, column=i + 1).alignment = alig_s
        index_1 = 2
    for i in sorted(da):
        # print(da[i])
        if da[i][4].strip() == "":
            for ii in range(1, len(da[i]) + 1):
                ws_2.cell(row=index_2, column=ii).value = da[i][ii - 1]
                ws_2.cell(row=index_2, column=ii).alignment = alig_s
            index_2 = index_2 + 1
        else:
            for ii in range(1, len(da[i]) + 1):
                ws_1.cell(row=index_1, column=ii).value = da[i][ii - 1]
                ws_1.cell(row=index_1, column=ii).alignment = alig_s
            index_1 = index_1 + 1
    wb.save(conf_data["dst"] + "/" + conf_data["fn"])


# 3. multiprocess_fun
def multiprocess_fun(d_a, task_kind, conf_data):
    poll_num = conf_data["poll"]
    pool = Pool(poll_num)
    key_dic = {}
    value_dic = {}
    # 分组
    for i in range(poll_num):
        key_dic[i] = []
        value_dic[i] = []
    for key, value in d_a.items():
        for i in range(poll_num):
            if int(key % poll_num) == i:
                key_dic[i].append(key)
                value_dic[i].append(value)
                continue
    new_dict = {}
    res_dict = []
    maneger = Manager()
    lock = maneger.Lock()
    for i in range(poll_num):
        new_dict[i] = dict(zip(key_dic[i], value_dic[i]))
        print("分组:", new_dict[i])
        if task_kind == 1:
            res_dict.append(pool.apply_async(dns_process,  (new_dict[i], conf_data)))
        elif task_kind == 2:
            res_dict.append(pool.apply_async(get_title_by_selenium, (new_dict[i], conf_data)))
        elif task_kind == 3:
            res_dict.append(pool.apply_async(get_alexa_rank_by_link114, (new_dict[i],)))
        elif task_kind == 4:
            res_dict.append(pool.apply_async(get_alexa_rank_by_link114_multi, (new_dict[i], conf_data)))
        elif task_kind == 5:
            res_dict.append(pool.apply_async(get_alexa_rank_by_alexa, (new_dict[i],)))
    pool.close()
    pool.join()

    res_dict_N = {}
    for res in res_dict:
        res_dict_N.update(res.get())
    if task_kind == 1:
        update_task_excel(res_dict_N, conf_data, 0)
    elif task_kind == 2:
        update_task_excel(res_dict_N, conf_data, 0)
    elif task_kind == 3:
        update_task_excel(res_dict_N, conf_data, 1)
    elif task_kind == 4:
        update_task_excel(res_dict_N, conf_data, 1)
    elif task_kind == 5:
        update_task_excel(res_dict_N, conf_data, 1)
    return res_dict_N


# 4. [3, 'http://01014688.comxxx', 'http://01014688.comxxx', '01014688.comxxx', '', '', '', '']
def dns_process(d_a, conf_data):
    for key, value in d_a.items():
        norm_url = get_url_normalize_single(value[1])
        value.append(norm_url)
        domain = urlparse(norm_url).hostname
        myaddr = []
        if isIP(domain):
            value.append("")
            value.append(domain)
            myaddr.append(domain)
        else:
            value.append(domain)
            try:
                # print(domain)
                A = socket.gethostbyname(domain)
                myaddr.append(A)
            except socket.error:
                myaddr.append("")
            value.append(myaddr[0])
        time.sleep(0.1)
        if myaddr[0].strip() == "":
            value.append("")
            value.append("")
            value.append("")
        else:
            dbpath = conf_data["conf"] + "/ipipfree.ipdb"
            db = ipdb.City(dbpath)
            city_str = db.find(myaddr[0], "CN")
            jing_nei = "中国,香港,澳门,台湾"
            if city_str[0] in jing_nei:
                if city_str[1] in jing_nei:
                    value.append("境外")
                else:
                    value.append("境内")
            else:
                value.append("境外")
            value.append(city_str[0] + "·" + city_str[1])
    return d_a


# 5.
def get_title_by_selenium(d_a, conf_data):
    chrome_options = Options()
    # 设置浏览器窗口大小
    chrome_options.add_argument('--window-size=1366,768')
    # chrome_options.add_argument('--disable-gpu')
    # 限制图片+JavaScript
    # prefs ={
    #     'profile.managed_default_content_setting_values': {
    #         'images': 2,
    #         'javascript': 2
    #     },
    #     'permissions.default.stylesheet': 2
    # }
    # chrome_options.add_experimental_option('prefs', prefs)
    # 浏览器不提供可视化页面. linux下如果系统不支持可视化不加这条会启动失败
    # chrome_options.add_argument('--headless')
    # 禁用浏览器弹窗
    prefs = {
        'profile.default_content_setting_values': {
            'notifications': 2
        }
    }
    chrome_options.add_experimental_option('prefs', prefs)

    # 启动浏览器
    browser = webdriver.Chrome(chrome_options=chrome_options)
    timeout_s = 15
    browser.implicitly_wait(timeout_s)
    # browser.set_page_load_timeout(timeout_s - 1)
    # browser.set_script_timeout(timeout_s - 1)
    all_win = browser.window_handles
    screen_shot_dir = conf_data["dst"] + "/截图/"

    col_l = conf_data["col"]
    yum = len(d_a) % col_l
    counts = math.ceil(len(d_a)/col_l)
    key_list = list(d_a.keys())
    value_list = list(d_a.values())
    col_add = col_l
    # print("***进入进程:", key_list)
    # print("***进入进程:", value_list)
    for i in range(counts):
        if (i == (counts - 1)) and (yum > 0):
            col_add = yum
        print("单个进程轮询:", len(d_a), yum, counts)
        for j in range(col_add):
            index = value_list[i*col_l + j][0]
            url = value_list[i*col_l + j][2]
            # print(index, url)
            # 开始请求
            try:
                js = "window.open(\"" + url + "\");"
                browser.execute_script(js)
            except {socket.timeout, TimeoutException}:
                print("1. socket, 超时:", url)
                browser.execute_script("window.stop();")
                url_refer = "超时"
        time.sleep(timeout_s*0.8)
        for j in range(col_add):
            index = value_list[i*col_l + j][0]
            url = value_list[i*col_l + j][2]
            all_win = browser.window_handles
            browser.switch_to.window(all_win[-1])
            domain = urlparse(url).hostname
            # print("*****", url)
            try:
                b_title = browser.title
                if b_title == "":
                    b_title = "补" + url
                print(index, "\t", url, "\t", b_title)
                url_con = browser.current_url.rstrip('/')
                if domain == urlparse(url_con).hostname:
                    url_refer = ""
                else:
                    url_refer = browser.current_url
                browser.get_screenshot_as_file(screen_shot_dir + str(index) + "_" + str(domain) + ".png")
            except TimeoutException:
                url_refer = "超时"
                b_title = "超时或者无法访问"
                url_con = ""
                try:
                    browser.get_screenshot_as_file(screen_shot_dir + str(index) + "_" + str(domain) + ".png")
                    print(index, " browser 超时:", url, "截图成功")
                    # print("截图成功")
                except BaseException as msg:
                    print(index, ":", url, msg)
                    screen_shot_file_name = screen_shot_dir + str(index) + "_" + str(domain) + ".txt"
                    fpt = open(screen_shot_file_name, 'w')
                    fpt.write("超时了哇")
                    print(index, " browser 超时:", url, "截图失败")
            browser.close()
            browser.switch_to.window(all_win[0])

            value_list[i * col_l + j].append(b_title)
            value_list[i * col_l + j].append(url_refer)
            value_list[i * col_l + j].append(url_con)
    browser.close()
    browser.quit()
    return d_a


# 6. link114 get Alexa one by one
def get_alexa_rank_by_link114(da):
    url_114 = "http://www.link114.cn/alexa/"
    chrome_options = Options()
    # chrome_options.add_argument('--headless')
    browser = webdriver.Chrome(options=chrome_options)
    timeout_s = 20
    browser.implicitly_wait(timeout_s)
    # 开始请求
    # js = "window.open(\"" + url_114 + "\");"
    # browser.execute_script(js)
    browser.get(url_114)
    for key, value in da.items():
        url = value[3]
        if url is not None:
            browser.find_element_by_id("ip_websites").clear()
            browser.find_element_by_id("ip_websites").send_keys(url)
            browser.find_element_by_id("tj").click()
            trlist = browser.find_elements_by_tag_name("tr")
            for tr in trlist:
                tdlist = tr.find_elements_by_tag_name("td")
                if len(tdlist) > 0:
                    trid_1 = tr.find_elements_by_tag_name("td")[1].get_attribute("value")
                    trid_alexa = tr.find_elements_by_tag_name("td")[2].text
                    trid_alexa = trid_alexa.replace("Alexa:", "")
                    value.append(trid_1)
                    value.append(trid_alexa)
        time.sleep(0.5)
        print(value)
    browser.close()
    browser.quit()
    return da


# 7. link114 get Alexa column
def get_alexa_rank_by_link114_multi(d_a, conf_data):
    url_114 = "http://www.link114.cn/"
    browser = webdriver.Chrome()
    timeout_s = 5
    browser.implicitly_wait(timeout_s)
    # 开始请求
    browser.get(url_114)
    # unselect 1
    str1 = "//*[@id='chk_baidu_qz_zz']"
    check_conditions = browser.find_element_by_xpath(str1)
    browser.execute_script("$(arguments[0]).click();", check_conditions)
    browser.execute_script("$(arguments[0]).attr('checked',false);", check_conditions)
    # unselect 2
    str2 = "//*[@id='chk_baidu_qz_ai']"
    check_conditions = browser.find_element_by_xpath(str2)
    browser.execute_script("$(arguments[0]).click();", check_conditions)
    browser.execute_script("$(arguments[0]).attr('checked',false);", check_conditions)
    # unselect 3
    str3 = "//*[@id='chk_baidu_sl']"
    check_conditions = browser.find_element_by_xpath(str3)
    browser.execute_script("$(arguments[0]).click();", check_conditions)
    browser.execute_script("$(arguments[0]).attr('checked',false);", check_conditions)
    # unselect 4
    str4 = "//*[@id='chk_so360_qz_zz']"
    check_conditions = browser.find_element_by_xpath(str4)
    browser.execute_script("$(arguments[0]).click();", check_conditions)
    browser.execute_script("$(arguments[0]).attr('checked',false);", check_conditions)

    # select 5
    # str5 = "//*[@id='chk_title']"
    # check_conditions = browser.find_element_by_xpath(str5)
    # browser.execute_script("$(arguments[0]).click();", check_conditions)
    # browser.execute_script("$(arguments[0]).attr('checked',true);", check_conditions)
    # select 6
    str6 = "//*[@id='chk_alexa']"
    check_conditions = browser.find_element_by_xpath(str6)
    browser.execute_script("$(arguments[0]).click();", check_conditions)
    browser.execute_script("$(arguments[0]).attr('checked',true);", check_conditions)
    # select 7
    # str7 = "//*[@id='chk_ip']"
    # check_conditions = browser.find_element_by_xpath(str7)
    # browser.execute_script("$(arguments[0]).click();", check_conditions)
    # browser.execute_script("$(arguments[0]).attr('checked',true);", check_conditions)

    col_l = conf_data["col"]
    yum = len(d_a) % col_l
    counts = math.ceil(len(d_a) / col_l)
    key_list = list(d_a.keys())
    value_list = list(d_a.values())
    col_add = col_l
    # print("***进入进程:", key_list)
    # print("***进入进程:", value_list)
    for i in range(counts):
        if (i == (counts - 1)) and (yum > 0):
            col_add = yum
        print("单个进程轮询:", len(d_a), yum, counts)
        domain_set = []
        for j in range(col_add):
            index = value_list[i * col_l + j][0]
            domain = value_list[i * col_l + j][3]
            domain_set.append(domain)
            print(index, "\t", domain)
        domain_str = ',n'.join(domain_set)
        print(domain_str)
        browser.find_element_by_id("ip_websites").clear()
        browser.find_element_by_id("ip_websites").send_keys(domain_str)
        browser.find_element_by_id("tj").click()
        time.sleep(10)
        trlist = browser.find_elements_by_tag_name("tr")
        data_dict = {}
        index = 0
        for tr in trlist:
            # 获取tr中的所有td
            tdlist = tr.find_elements_by_tag_name("td")
            data_dict[index] = []
            if len(tdlist) > 0:
                # 获取td[0]的文本
                text_1 = tr.find_elements_by_tag_name("td")[0].text
                text_1 = text_1.replace(".", "")
                # trid_1 = tr.get_attribute("id")
                trid_1 = tr.find_elements_by_tag_name("td")[1].get_attribute("value")
                trid_alexa = tr.find_elements_by_tag_name("td")[2].text
                trid_alexa = trid_alexa.replace("Alexa:", "")
                data_dict[index].append(text_1)
                data_dict[index].append(trid_1)
                data_dict[index].append(trid_alexa)
                index = index + 1
        alexa_set = data_dict.values()
        for d in alexa_set:
            for key, value in d_a.items():
                if d[1] == value[3]:
                    value.append(d[1])
                    value.append(d[2])
                    break
        time.sleep(timeout_s)
    browser.close()
    browser.quit()
    return d_a


# 根据html获取 alexa_rank
def get_alexa_rank_from_html(html):
    reg = r' <div class="rankmini-rank">.*?<span>#</span>([\d,]{0,20}).*?</div>'
    urlre = re.compile(reg, re.I | re.M | re.S)
    urllist = urlre.findall(html)
    if len(urllist):
        rank_n = urllist[0].replace(',', '')
    else:
        rank_n = "无排名"
    return rank_n

# 8. Alexa get Alexa one by one
def get_alexa_rank_by_alexa(da):
    url_alexa_head = "https://www.alexa.com/"
    chrome_options = Options()
    # chrome_options.add_argument('--headless')
    browser = webdriver.Chrome(options=chrome_options)
    timeout_s = 10
    browser.implicitly_wait(timeout_s)

    # browser.set_page_load_timeout(timeout_s - 1)
    # browser.set_script_timeout(timeout_s - 1)

    # js = "window.open(\"" + "https://www.baidu.com" + "\");"
    # browser.execute_script(js)
    # time.sleep(1)
    all_win = browser.window_handles
    print("total len = ", len(da))
    for key, value in da.items():
        print(key, "\t", value)
        domain = value[3]
        print(domain)
        url_alexa = url_alexa_head + "siteinfo/" + domain

        try:
            js = "window.open(\"" + url_alexa + "\");"
            browser.execute_script(js)
            time.sleep(10)
            html = browser.page_source
            rank_n = get_alexa_rank_by_alexa(html)
        except TimeoutException:
            rank_n = "Timeout"
            browser.execute_script("window.stop();")
            # print("Alexa Timeout")
        print(domain, "\t", rank_n)
        # value.append(domain)
        # value.append(rank_n)
        time.sleep(5)
        browser.close()
        browser.switch_to.window(all_win[0])
    browser.quit()
    return da