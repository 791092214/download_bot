from selenium import webdriver
import time
import datetime
import os


def getDatesByTimes(sDateStr, eDateStr):
    '''
    得到一个区间的所有日期

    getDatesByTimes('20190226','20190401') 用例展示

    '''
    list = []
    datestart = datetime.datetime.strptime(sDateStr, '%Y%m%d')
    dateend = datetime.datetime.strptime(eDateStr, '%Y%m%d')
    list.append(datestart.strftime('%Y%m%d'))
    while datestart < dateend:
        datestart += datetime.timedelta(days=1)
        list.append(datestart.strftime('%Y%m%d'))
    return list


################### 在这里输入变量 ################################

download_file_name = '精英贷业务客户授信信息明细表'

date_period = getDatesByTimes('20200521', '20200528')  # 第一个输入起始日期，第二个输入截止日期

date_special = ['20170131', '20170228', '20170331', '20170430', '20170531', '20170630', '20170731', '20170831',
                '20170930', '20171031', '20171130', '20171231',
                '20180131', '20180228', '20180331',
                '20180430']  # 如果你需要下载特定日期的报表。在这里输入，特定的日期，比如： date_special = ['20200215','20200626','20200830'],并将下面的 for循环 改成 for i in date_special:

'''
'20170131','20170228','20170331','20170430','20170531','20170630','20170731','20170831','20170930','20171031','20171130','20171231',
                '20180131','20180228','20180331','20180430','20180531','20180630','20180731','20180831','20180930','20181031','20181130','20181231',
                '20190131','20190228','20190331','20190430','20190531','20190630','20190731','20190831','20190930','20191031','20191130','20191231',
                '20200131','20200229','20200331','20200430','20200531','20200630','20200731','20200831','20200930','20201031','20201130','20201231',

'''

################## 在上面的区域输入变量 ############################

# 用json格式的文件来储存表格路径，信息
table_infos_json = {'普惠金融事业部授信风险指标信息表(客户经理组表)':
                        ['//*[@id="button-1029-btnInnerEl"]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td/div/img[1]',
                         '//*[@id="treeview-1064"]/table/tbody/tr[6]/td/div',
                         '8b80a09a6010354f016010bf0230033b',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[5]/td/div/img[2]',
                         'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
                         ],
                    '普惠金融事业部三大压降任务进度表(事业部)':
                        ['//*[@id="button-1029-btnInnerEl"]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td/div/img[1]',
                         '//*[@id="treeview-1064"]/table/tbody/tr[7]/td/div',
                         '9a888c956c1e8a98016c6b5d23be3b47',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[5]/td/div/img[2]',
                         'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
                         ],
                    '普惠金融事业部表内不良贷款压降任务处置进度表':
                        ['//*[@id="button-1029-btnInnerEl"]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td/div/img[1]',
                         '//*[@id="treeview-1064"]/table/tbody/tr[8]/td/div',
                         '9a888c956892a7f601696c23bd0203e6',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[5]/td/div/img[2]',
                         'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
                         ],
                    '普惠金融事业部不良及重大风险隐患贷款任务处置进度表':
                        ['//*[@id="button-1029-btnInnerEl"]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td/div/img[1]',
                         '//*[@id="treeview-1064"]/table/tbody/tr[9]/td/div',
                         '9a888c9569a3879d016a4483abf16355',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[5]/td/div/img[2]',
                         'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
                         ],
                    '普惠金融事业部出表及专项任务处置压降进度表':
                        ['//*[@id="button-1029-btnInnerEl"]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td/div/img[1]',
                         '//*[@id="treeview-1064"]/table/tbody/tr[10]/td/div',
                         '9a888c9569a3879d016a4485a23c6363',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[5]/td/div/img[2]',
                         'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
                         ],
                    '普惠金融事业部重要风险指标情况表':
                        ['//*[@id="button-1029-btnInnerEl"]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td/div/img[1]',
                         '//*[@id="treeview-1064"]/table/tbody/tr[11]/td/div',
                         '9a888c956e4934aa016e6e43d59417e8',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[5]/td/div/img[2]',
                         'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
                         ],
                    '普惠金融事业部重要风险指标情况_客户经理组版':
                        ['//*[@id="button-1029-btnInnerEl"]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td/div/img[1]',
                         '//*[@id="treeview-1064"]/table/tbody/tr[12]/td/div',
                         '9a888c9570bddd5f0170be8469e5092c',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[5]/td/div/img[2]',
                         'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
                         ],
                    '普惠金融事业部重要风险指标情况_产品版':
                        ['//*[@id="button-1029-btnInnerEl"]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td/div/img[1]',
                         '//*[@id="treeview-1064"]/table/tbody/tr[13]/td/div',
                         '9a888c9570bddd5f0170be8266ce0919',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[5]/td/div/img[2]',
                         'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
                         ],
                    '风险客户日报表':
                        ['//*[@id="button-1029-btnInnerEl"]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td/div/img[1]',
                         '//*[@id="treeview-1064"]/table/tbody/tr[12]/td/div',
                         '9a888c956b0d65e6016b40fbcc7f28dd',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[7]/td/div/img[2]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[11]/td/div/img[3]'
                         ],
                    '风险客户日报表（剔除出表借据）':
                        ['//*[@id="button-1029-btnInnerEl"]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[3]/td/div/img[1]',
                         '//*[@id="treeview-1064"]/table/tbody/tr[13]/td/div',
                         '9a888c956b0d65e6016b69f92a194194',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[7]/td/div/img[2]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/table/tbody/tr[11]/td/div/img[3]'
                         ],
                    '精英贷业务客户授信信息明细表':
                        ['//*[@id="button-1048-btnInnerEl"]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[2]/div[2]/div/table/tbody/tr[3]/td/div/img[1]',
                         '//*[@id="treeview-1061"]/table/tbody/tr[22]/td/div',
                         '9a888c9567e6cbb40167e87354fb09b3',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[2]/div[2]/div/table/tbody/tr[17]/td/div/img[2]',
                         'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
                         ],
                    '精英贷业务借款信息明细表':
                        ['//*[@id="button-1048-btnInnerEl"]',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[2]/div[2]/div/table/tbody/tr[3]/td/div/img[1]',
                         '//*[@id="treeview-1061"]/table/tbody/tr[23]/td/div',
                         '9a888c9567e6cbb40167e8603c6108f7',
                         '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[1]/div[2]/div[2]/div[2]/div/table/tbody/tr[17]/td/div/img[2]',
                         'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
                         ]
                    }  # list第一个值：部门相对路径，第二个值：第一个“+”号的绝对路径， 第三个值：待下载表的相对路径， 第四个值：待下载表的特殊id，第五个值：第二个“+”号的绝对路径（如果没有二级目录，就是乱码）, 第六个值：第三个“+”号的绝对路径，如果不存在第三个加号，则是一串乱码

driver = webdriver.Chrome()

root_path = 'http://login.tf.cn/cas/login'

driver.get(root_path)

account_login = '/html/body/div[1]/div/div[3]/div[2]/div/div/div[1]/b[3]'

account_path = '/html/body/div[1]/div/div[3]/div[2]/div/div/form/div/div[4]/section/div/input[2]'

password_path = '/html/body/div[1]/div/div[3]/div[2]/div/div/form/div/div[4]/div[1]/section[1]/div/input[2]'

login_button_path = '/html/body/div[1]/div/div[3]/div[2]/div/div/form/div/div[4]/div[1]/section[3]/button/p'

data_integration_platform_path = '/html/body/div[1]/div/div[2]/ul/a[11]/li'

table_area_path = '/html/body/div[2]/div/div[1]/div/div[2]/ul/li[7]/a'  # 进入报表平台的path /html/body/div[2]/div/div[1]/div/div[2]/ul/li[7]/a

branch_path = table_infos_json[download_file_name][0]  # 进入相关部门的路径

driver.find_element_by_xpath(account_login).click()

driver.find_element_by_xpath(account_path).send_keys('001477')

driver.find_element_by_xpath(password_path).send_keys('ZXCvbn456258')

driver.find_element_by_xpath(login_button_path).click()

time.sleep(2)

driver.find_element_by_xpath(data_integration_platform_path).click()

n = driver.window_handles

driver.switch_to.window(n[1])

time.sleep(1)

driver.find_element_by_xpath(table_area_path).click()

# driver.find_element_by_xpath(branch_path).click()   # 进入相关部门

# 将我们想要的报表逐步打开

# 如果有三级目录
if table_infos_json[download_file_name][1][::-1][3:6] == table_infos_json[download_file_name][4][::-1][3:6] and \
        table_infos_json[download_file_name][1][::-1][3:6] == table_infos_json[download_file_name][5][::-1][3:6]:

    tmp_plus_button_path_1 = table_infos_json[download_file_name][1]

    driver.find_element_by_xpath(tmp_plus_button_path_1).click()

    tmp_plus_button_path_2 = table_infos_json[download_file_name][4]

    driver.find_element_by_xpath(tmp_plus_button_path_2).click()

    tmp_plus_button_path_3 = table_infos_json[download_file_name][5]

    driver.find_element_by_xpath(tmp_plus_button_path_3).click()

# 如果是二级目录

elif table_infos_json[download_file_name][1][::-1][3:6] == table_infos_json[download_file_name][4][::-1][3:6] and \
        table_infos_json[download_file_name][1][::-1][3:6] != table_infos_json[download_file_name][5][::-1][
                                                              3:6]:  # 如果存在二级目录

    tmp_plus_button_path_1 = table_infos_json[download_file_name][1]

    driver.find_element_by_xpath(tmp_plus_button_path_1).click()

    tmp_plus_button_path_2 = table_infos_json[download_file_name][4]

    driver.find_element_by_xpath(tmp_plus_button_path_2).click()


#   如果是一级目录
else:

    tmp_plus_button_path_1 = table_infos_json[download_file_name][1]

    driver.find_element_by_xpath(tmp_plus_button_path_1).click()

tmp_table_path = table_infos_json[download_file_name][2]

driver.find_element_by_xpath(tmp_table_path).click()  # 点击待下载的报表，等待加载

time.sleep(1)

file_nums = 0  # 路径里的文件数量

tmp_button_nums = 1064  # 下载文件的地址，会从1064开始递增，就是，文件下载的关闭按钮
for i in date_special:

    file_nums += 1  # 每一次循环，路径文件的数量，应该是要加1

    tmp_frame_id = table_infos_json[download_file_name][3]

    tmp_frame_id = 'frame_tab_' + tmp_frame_id  # 'frame_tab_report-region_report_' 原来报表平台的frame id

    driver.switch_to.frame(driver.find_element_by_id(tmp_frame_id))  # 切换frame

    driver.find_element_by_id('datefield-1066-inputEl').click()  # 这里的Input inputid不一样！！！ 原来是datefield-1011-inputEl

    driver.find_element_by_id('datefield-1066-inputEl').clear()

    driver.find_element_by_id('datefield-1066-inputEl').send_keys(i)  # 填入第一个日期

    '''
    driver.find_element_by_id('datefield-1070-inputEl').click() # 这里的Input inputid不一样！！！ 原来是datefield-1011-inputEl

    driver.find_element_by_id('datefield-1070-inputEl').clear()

    # driver.find_element_by_id('datefield-1070-inputEl').send_keys(i) # 填入第三个日期 第三个日期不填

    driver.find_element_by_id('datefield-1068-inputEl').click() # 这里的Input inputid不一样！！！ 原来是datefield-1011-inputEl

    driver.find_element_by_id('datefield-1068-inputEl').clear()

    driver.find_element_by_id('datefield-1068-inputEl').send_keys('20170101') # 填入第二个日期
    '''

    time.sleep(1)

    driver.find_element_by_id('button-1075-btnInnerEl').click()  # 要等一下,查询Button，有变化,中间的数字不一样

    time.sleep(2)

    # driver.switch_to.frame('report_frame') 即席查询里面，没有report_frame
    while True:
        try:
            driver.find_element_by_xpath('//*[@id="button-1057-btnInnerEl"]').click()  # 点击查询
            break
        except:
            print("下载中！！！别着急。。。")
            continue

    while True:  # 这里容易出错，所以加一个try,处理异常

        try:
            time.sleep(3)  # 这里不用睡，可以调快一点
            driver.find_element_by_xpath('//*[@id="menuitem-1060-textEl"]').click()  # 点击下载文本的按钮
            break

        except:

            print("正在加载，请耐心等候，φ(≧ω≦*)♪。。。。。。")

            continue

    driver.find_element_by_xpath('//*[@id="button-1005-btnInnerEl"]').click()  # 也要等一下 //*[@id="button-1005-btnInnerEl"]

    while True:
        try:
            driver.find_element_by_xpath('//*[@id="button-1005-btnInnerEl"]').click()
            break
        except:
            print("正在努力下载中，n(*≧▽≦*)n")
            continue

    # n = driver.window_handles

    # driver.switch_to.window(n[2])

    # 下面需要，点出文件下载，再下载，再退出当前frame
    driver.switch_to.default_content()
    driver.find_element_by_xpath('//*[@id="treeview-1061"]/table/tbody/tr[2]/td/div').click()  # 进入文件下载
    driver.switch_to.frame('frame_tab_free-query-platform-query-download')  # 切换至文件下载的frame
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="gridview-1031"]/table/tbody/tr[2]/td[7]/div/img').click()
    driver.find_element_by_xpath('//*[@id="button-1005-btnInnerEl"]').click()
    driver.switch_to.parent_frame()
    # ActionChains(driver).double_click(driver.find_element_by_xpath('//*[@id="tab-1065"]'))
    time.sleep(1)
    tmp_button_address = '//*[@id="tab-' + str(tmp_button_nums) + '-closeEl"]'  # //*[@id="tab-1064-closeEl"]
    driver.find_element_by_xpath(tmp_button_address).click()  # 关闭文件下载
    tmp_button_nums = tmp_button_nums + 1
    # 下面是给文件改名字
    # while True:     # 只有，当目标文件夹里，文件的数量增加的时候，才可以关掉下载窗口

    time.sleep(1)
    download_path = r'C:\Users\Administrator.USER-20190125QU\Downloads'
    file_nums_list = os.listdir(download_path)
    '''
    file_nums_now = len(file_nums_list)
    if file_nums == file_nums_now-1:
        break
    else:
        continue

    if file_nums == 1:  # 记录下，被下载的文件的名字。
        final_file_name = os.listdir(download_path)[-1][0:-13]
    '''
    # driver.close()

    # n = driver.window_handles

    # driver.switch_to.window(n[1])

    src_file_path_1 = r'C:/Users/Administrator.USER-20190125QU/Downloads'
    src_file_path_2 = src_file_path_1 + "/"
    src_file_path_3 = src_file_path_2 + '精英贷业务客户授信信息明细表_20210129.txt'
    dst_file_path = src_file_path_2 + download_file_name + i + '.txt'

    # 给文件改名字,要直到文件下载到位了，名字改成功了，再走下一步
    while True:
        try:
            os.rename(src_file_path_3, dst_file_path)
            break
        except:
            print("等待文件下载到位。。。。。")
            continue