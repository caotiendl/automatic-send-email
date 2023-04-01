# automatic-send-email
only apply for yahoo mail, If use google , please use a suitable API
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium_stealth import stealth
from selenium.webdriver.common.by import By
import time
from openpyxl import load_workbook
import smtplib
import re

def send_email(list_email,text):
    yahoo_user = 'caoquangtiendl@yahoo.com'  # nhập tài khoản email dùng để gửi
    yahoo_app_password = '*******'        # nhập mật khẩu email dùng để gửi
    sent_from = yahoo_user
    sent_to = list_email
    sent_subject = 'update ty gia ngay 8-6'
    sent_body = text

    email_text = "From: %s\r\nTo: %s\r\nSubject: %s\r\n\r\n%s\r\n" % (
    sent_from, ', '.join(sent_to), sent_subject, sent_body)

    try:
        server = smtplib.SMTP_SSL('smtp.mail.yahoo.com', 465)
        server.ehlo()
        server.login(yahoo_user, yahoo_app_password)
        server.sendmail(sent_from, sent_to, email_text.encode('utf-8'))
        server.quit()
        print('Email sent!')
    except Exception as exception:
        print("Error: %s!\n\n" % exception)

def read_file_xlsx():

    wb = load_workbook("C:\\Users\caoqu\Desktop\email.xlsx")
    ws = wb.active
    first_column = ws['A']
    list_email = []
    for x in range(1, len(first_column)):
        list_email.append(first_column[x].value)
    return list_email


def get_data():
    data = []
    data_vc = []
    data_ag = []
    data_os = []
    data_hs = []
    options = webdriver.ChromeOptions()
    options.add_argument("start-maximized")

    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    driver = webdriver.Chrome(executable_path='C:\\Users\caoqu\Desktop\chromedriver.exe')

    # VietComBank
    driver.get("https://portal.vietcombank.com.vn/Personal/TG/Pages/ty-gia.aspx?devicechannel=default")
    driver.implicitly_wait(4)
    tr_usd = driver.find_elements(By.XPATH, '//tr[@class="odd"]')[19]
    time_vc = tr_usd.get_attribute("data-time")
    data_vc.append(time_vc)
    ref = tr_usd.find_elements(By.XPATH, './td')
    for i in ref:
        data_vc.append(i.text)
    data.append(data_vc)

    # Agribank
    driver.get("https://www.agribank.com.vn/vn/ty-gia")
    time.sleep(3)
    _string_time_ag = driver.find_element(By.XPATH, '//div[@class="luu_ycc"]').text
    abc = _string_time_ag.split(" ")
    string_time_ag = ""
    for i in abc:
        if (re.search(r'\d', i)):
            if string_time_ag == "":
                string_time_ag += i
            else:
                string_time_ag = string_time_ag + ", " + i
    data_ag.append(string_time_ag)
    tr_usd_ag = driver.find_elements(By.XPATH, '//tr')[1]
    ref_ag = tr_usd_ag.find_elements(By.XPATH, './td')
    for i in ref_ag:
        data_ag.append(i.text)
    data.append(data_ag)

    # Overseas Bank
    driver.get("https://www.uob.com.vn/general/online-rates/foreign-exchange-rates.page")
    time.sleep(3)

    _string_time_os = driver.find_element(By.XPATH, '//span[@class="rates"]').text
    abc = _string_time_os.split(" ")
    string_time_os = ""
    for i in abc:
        if (re.search(r'\d', i)):
            if string_time_os == "":
                string_time_os += i
            else:
                string_time_os = string_time_os + ", " + i

    data_os.append(string_time_os)
    tr_usd_os = driver.find_elements(By.XPATH, '//tr[@id]')[9]
    ref_os = tr_usd_os.find_elements(By.XPATH, './td/p[@class="small"]')
    for i in ref_os:
        data_os.append(i.text)
    data.append(data_os)

    # HSBC bank
    driver.get("https://www.hsbc.com.vn/foreign-exchange/rate/")
    time.sleep(3)
    string_time_hs = driver.find_element(By.XPATH,
                                         '//div[@id="content_main_basicTable_1"]/table[@class="desktop"]/tbody/tr/td').text
    data_hs.append(string_time_hs)
    tr_usd_hs = \
    driver.find_elements(By.XPATH, '//div[@id="content_main_basicTable_2"]/table[@class="desktop"]/tbody/tr')[0]
    ref_hs = tr_usd_hs.find_elements(By.XPATH, './td')
    for i in ref_hs:
        data_hs.append(i.text)
    data.append(data_hs)

    driver.quit()
    return data


# main()
from datetime import datetime

while True:
    now = datetime.now()
    print(now)
    if now.hour < 17:
        list_email = read_file_xlsx()
        data = get_data()
        text = """ Thời gian lấy dữ liệu: {thoigian}
        TỶ GIÁ cụ thể:
        Ngân hàng VietComBank: lúc {time_vc} ngoại tệ USD mua bằng tiền mặt {tien_mat_vc}, mua bằng chuyển khoản {chuyen_khoan_vc}, giá bán {gia_ban_vc}
        Ngân hàng Agribank: lúc {time_ag} ngoại tệ USD mua bằng tiền mặt {tien_mat_ag}, mua bằng chuyển khoản {chuyen_khoan_ag}, giá bán {gia_ban_ag}
        Ngân hàng Overseas Bank: lúc {time_os} ngoại tệ USD mua bằng tiền mặt {tien_mat_os}, mua bằng chuyển khoản {chuyen_khoan_os}, giá bán {gia_ban_os}
        Ngân hàng HSBC bank: lúc {time_hs} ngoại tệ USD mua bằng tiền mặt {tien_mat_hs}, mua bằng chuyển khoản {chuyen_khoan_hs}, giá bán {gia_ban_hs}

        """.format(thoigian=str(now), time_vc=data[0][0], tien_mat_vc=data[0][3], chuyen_khoan_vc=data[0][4],
                   gia_ban_vc=data[0][5], \
                   time_ag=data[1][0], tien_mat_ag=data[1][2], chuyen_khoan_ag=data[1][3], gia_ban_ag=data[1][4], \
                   time_os=data[2][0], tien_mat_os=data[2][1], chuyen_khoan_os=data[2][2], gia_ban_os=data[2][3], \
                   time_hs=data[3][0], tien_mat_hs=data[3][1], chuyen_khoan_hs=data[3][2], gia_ban_hs=data[3][3])

        send_email(list_email, text)
        time.sleep(3600)
    else:
        miu = (60 - now.minute) * 60
        time.sleep(miu)
