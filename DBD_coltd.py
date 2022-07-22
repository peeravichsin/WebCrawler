from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
import time
from datetime import date
today = date.today()

import pandas as pd

def coltd(file_name):
    # For using Chrome
    browser = webdriver.Chrome(r'.Cdriver\chromedriver.exe')
    action = webdriver.ActionChains(browser)


    #----------------------------------------------------------------------------------------

    ##Function

    def ThaiMonth2Num(month):
        months = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']
        Month_Num = months.index(month) + 1
        mn = f"{Month_Num:02}"
        return str(mn)

    def BE2AD(year):
        AD = int(year) - 543
        return str(AD)

    #-----------------------------------------------------------------------------------------

    try:
        # Webpage
        browser.get('https://datawarehouse.dbd.go.th/login')

        ####Captcha Page(Close Pop-up)
        close_button = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.CLASS_NAME, 'btn.btn-secondary')))
        browser.execute_script("arguments[0].click();", close_button)

        ####For type captcha by-hand
        time.sleep(20)

        ####Summit Button
        summit_button = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.CLASS_NAME, 'btn.btn-purple.btn-submit')))
        browser.execute_script("arguments[0].click();", summit_button)

        ####First Page(Searchbox)
        search_box = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.CLASS_NAME, 'form-control.tempTxt')))
        search_box.send_keys('จำกัด')

        search = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.CLASS_NAME, 'btn.btn-link')))
        time.sleep(1)
        browser.execute_script("arguments[0].click();", search)

        sort = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.CLASS_NAME, 'btn.btn-outline.btn-filter')))
        time.sleep(1)
        browser.execute_script("arguments[0].click();", sort)

        #----------------------------------------------------------------------------------------
        #XPATH ประเภทของบริษัท
        ##หากต้องการเข้าถึงข้อมูลบริษัทประเภทใดก็ให้ใช้เครื่องหมายสี่เหลี่ยม (#) ไว้ด้านหน้าประเภทบริษัทอื่นๆที่ไม่ต้องการจะดึงข้อมูล

        ####เลือกประเภทนิติบุคคล
        Choose_type = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="filterBoxContent"]/div[2]/h4')))
        time.sleep(1)
        browser.execute_script("arguments[0].click();", Choose_type)

        #----------------------------------------------------------------------------------------

        ##ประเภทของนิติบุคคล
        ####บริษัทจำกัด Company Limited
        CoLtd = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="collapse2"]/div[2]/label')))
        time.sleep(1)
        browser.execute_script("arguments[0].click();", CoLtd)

        #----------------------------------------------------------------------------------------


        ####เลือกสถานะ
        choose_status = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[4]/div[2]/div/div[1]/div[1]/ul/li[3]/div/form/div[1]/div[3]/h4')))
        browser.execute_script("arguments[0].click();", choose_status)
        time.sleep(1)

        ####สถานะยังดำเนินกิจการอยู่
        status = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[4]/div[2]/div/div[1]/div[1]/ul/li[3]/div/form/div[1]/div[3]/div/ul/li[1]/div[1]/label')))
        browser.execute_script("arguments[0].click();", status)
        time.sleep(1)

        confirm_button = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.CLASS_NAME, 'btn.btn-link.btn-confirm')))
        browser.execute_script("arguments[0].click();", confirm_button)
        browser.implicitly_wait(60)

        
        #----------------------------------------------------------------------------------------


        DBD = []
        check = 0
        page_num = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[4]/div[2]/div/div[4]/div[1]/nav/ul/li[3]/span')))
        time.sleep(3)
        page = int((page_num.text).translate(str.maketrans({'/': '', ',': ''})))
        print(page)

        while check != page:

            link = WebDriverWait(browser, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="fixTable"]/tbody/tr[/]')))

            #province = WebDriverWait(browser, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="fixTable"]/tbody/tr[/]/td[8]')))


            ####Go&Get Data in linkpage
            for i in range(len(link)):

                browser.set_page_load_timeout(60)

                if check >=1:
                    time.sleep(3)
                    num_box = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[4]/div[2]/div/div[4]/div[1]/nav/ul/li[2]/input')))
                    browser.execute_script("arguments[0].click();", num_box)
                    num_box.clear()
                    time.sleep(3)
                    #num_box.send_keys(Keys.DELETE)
                    ans_page = page - (page - (check + 1))
                    num_box.send_keys(str(ans_page))
                    num_box.send_keys(Keys.ENTER)
                    browser.set_page_load_timeout(30)

                #ตรวจสอบสถานะ
                status = WebDriverWait(browser, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="fixTable"]/tbody/tr[/]/td[5]')))
                if status[i].text == 'ยังดำเนินกิจการอยู่':
                    pass
                else:
                    break

                #ชื่อนิติบุคคล
                j_name = WebDriverWait(browser, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="fixTable"]/tbody/tr[/]/td[3]')))
                browser.set_page_load_timeout(60)
                j_name = j_name[i].text

                #เลขนิติบุคคล
                j_ID = WebDriverWait(browser, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="fixTable"]/tbody/tr[/]/td[2]')))
                browser.set_page_load_timeout(30)
                j_ID = j_ID[i].text

                #ประเภทนิติบุคคล
                j_type = WebDriverWait(browser, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="fixTable"]/tbody/tr[/]/td[4]')))
                browser.set_page_load_timeout(30)
                j_type = j_type[i].text
                
                #ทุนจดทะเบียน
                capital = WebDriverWait(browser, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="fixTable"]/tbody/tr[/]/td[9]')))
                browser.set_page_load_timeout(30)
                capital = capital[i].text

                if capital == '-' or capital == '':
                    capital = 0.00
                else:
                    capital = capital.replace(',', '')
                    capital = float(capital)

                #วันที่กรรมการมีผลฯ
                directors_date = today.strftime("%Y-%m-%d")

                #วันดึงข้อมูล
                scraped_date = today.strftime("%Y-%m-%d")

                #รหัสประเภทธุรกิจ
                b_code = WebDriverWait(browser, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="fixTable"]/tbody/tr[/]/td[6]')))
                browser.set_page_load_timeout(30)
                b_code = b_code[i].text

                #ประเภทธุรกิจ
                b_type = WebDriverWait(browser, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="fixTable"]/tbody/tr[/]/td[7]')))
                browser.set_page_load_timeout(30)
                b_type = b_type[i].text

                #------------------------------------------------------------------------------------------------------------------------------------------------
                #####ข้อมูลภายในหน้าเว็บของแต่ละบริษัท

                link = WebDriverWait(browser, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="fixTable"]/tbody/tr[/]')))
                #link[i].click()
                browser.execute_script("arguments[0].click();", link[i])


                #วันที่จดทะเบียนจัดตั้ง
                regis_date = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/table/tbody/tr[4]/td[2]')))
                regis_date = (regis_date.text).split(' ')
                regis_date = BE2AD(regis_date[2]) + '-' + ThaiMonth2Num(regis_date[1]) + '-' + regis_date[0]

                #ที่ตั้ง
                locate = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[2]/td')))
                locate = locate.text

                #หมายเลขโทรศัพท์
                tel = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[3]/td[2]')))
                tel = tel.text

                #อีเมล
                email = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[5]/td[2]')))
                email = email.text

                #รายชื่อคณะกรรมการ
                director_list = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[2]/div/div')))
                browser.set_page_load_timeout(30)
                director_list = director_list.text
                director_list = director_list.replace('/', '')
                director_list = director_list.split('\n')

                #กรรมการลงชื่อผูกพัน
                authorized_director = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[3]/div[1]/div/p')))
                browser.set_page_load_timeout(120)
                authorized_director = (authorized_director.text).translate(str.maketrans({'/': '', '\n': '', '-': ''}))

                #------------------------------------------------------------------------------------------------------------------------------------------------
                ####หน้าข้อมูลงบการเงิน

                financial_statement = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[4]/div[2]/nav/div/div[2]/div/ul/li[2]/a')))
                browser.execute_script("arguments[0].click();", financial_statement)

                #หน้างบกำไรขาดทุน
                income_statement = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[4]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div/ul/li[2]/a')))
                browser.execute_script("arguments[0].click();", income_statement)

                #รายได้หลัก
                main_income = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[4]/div[2]/div[1]/div[2]/div[1]/div[2]/div[2]/div/div[2]/table/tbody/tr[1]/td[6]')))
                browser.set_page_load_timeout(30)
                main_income = main_income.text
                
                if main_income == '-' or main_income == '':
                    main_income = 0.00
                else:
                    main_income = main_income.replace(',', '')
                    main_income = float(main_income)

                #ต้นทุนขาย (Cost of Sales)
                CoS = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[4]/div[2]/div[1]/div[2]/div[1]/div[2]/div[2]/div/div[2]/table/tbody/tr[3]/td[6]')))
                browser.set_page_load_timeout(30)
                CoS = CoS.text

                if CoS == '-' or CoS == '':
                    CoS = 0.00
                else:
                    CoS = CoS.replace(',', '')
                    CoS = float(CoS)

                #ต้นทุนขายคิดเป็น % จากรายได้ คิดได้จาก (ต้นทุนขาย / รายได้หลัก) *100
                if CoS !=0 and main_income != 0:
                    percentage1 = "{:.2f}".format((CoS / main_income) * 100)
                else:
                    percentage1 = 0.00

                #กำไร(ขาดทุน)สุทธิ (Net Profit/Loss)
                NetPL = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[4]/div[2]/div[1]/div[2]/div[1]/div[2]/div[2]/div/div[2]/table/tbody/tr[10]/td[6]')))
                browser.set_page_load_timeout(30)
                NetPL = NetPL.text

                if NetPL == '-' or NetPL == '':
                    NetPL = 0.00
                else:
                    NetPL = NetPL.replace(',', '')
                    NetPL = float(NetPL)

                #กำไรคิดเป็น % จากรายได้ คิดได้จาก (กำไร(ขาดทุน)สุทธิ / รายได้หลัก) *100
                if NetPL !=0 and main_income != 0:
                    percentage2 = "{:.2f}".format((NetPL / main_income) * 100)
                else:
                    percentage2 = 0.00

                #------------------------------------------------------------------------------------------------------------------
                ##หน้าอัตราส่วนทางการเงิน

                Financial_Ratio = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[4]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div/ul/li[3]/a')))
                browser.execute_script("arguments[0].click();", Financial_Ratio)

                #อัตราส่วนทุนหมุนเวียน(หน่วย: เท่า) [Current Ratio]
                CR = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[4]/div[2]/div[1]/div[2]/div[1]/div[2]/div[2]/div/div[2]/table/tbody[2]/tr[1]/td[5]')))
                browser.set_page_load_timeout(30)
                CR = CR.text

                if CR == '-' or CR == '':
                    CR = 0.00
                else:
                    CR = CR.replace(',', '')
                    CR = float(CR)

                #อัตราส่วนหนี้สินรวมต่อส่วนของผู้ถือหุ้น (หน่วย: เท่า) [Debt to Equity]
                DE = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[4]/div[2]/div[1]/div[2]/div[1]/div[2]/div[2]/div/div[2]/table/tbody[4]/tr[3]/td[5]')))
                browser.set_page_load_timeout(30)
                DE = DE.text
                
                if DE == '-' or DE == '':
                    DE = 0.00
                else:
                    DE = DE.replace(',', '')
                    DE = float(DE)

                #อัตราผลตอบแทนจากส่วนของผู้ถือหุ้น (หน่วย: %) [Return of Equity]
                ROE = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[4]/div[2]/div[1]/div[2]/div[1]/div[2]/div[2]/div/div[2]/table/tbody[1]/tr[2]/td[5]')))
                browser.set_page_load_timeout(30)
                ROE = ROE.text

                if ROE == '-' or ROE == '':
                    ROE = 0.00
                else:
                    ROE = ROE.replace(',', '')
                    ROE = float(ROE)

                
                detail = {
                    'juristic_ID' : j_ID,  #เลขนิติบุคคล
                    'juristic_name' : j_name,   #ชื่อนิติบุคคล
                    'juristic_type' : j_type, #ประเภทนิติบุคคล
                    'registered_capital' : capital, #ทุนจดทะเบียน
                    'main_income' : main_income, #รายได้หลัก
                    'cost_of_sales': CoS,  #ต้นทุนขาย (Cost of Sales)
                    'percentage1' : percentage1,  #ต้นทุนขายคิดเป็น % จากรายได้ คิดได้จาก (ต้นทุนขาย / รายได้หลัก) *100
                    'Net_PL' : NetPL,  #กำไร(ขาดทุน)สุทธิ (Net Profit/Loss)
                    'percentage2' : percentage2,  #กำไรคิดเป็น % จากรายได้ คิดได้จาก (กำไร(ขาดทุน)สุทธิ / รายได้หลัก) *100
                    'current_ratio' : CR,  #อัตราส่วนทุนหมุนเวียน(หน่วย: เท่า) [Current Ratio]
                    'debt_to_equity' : DE,  #อัตราส่วนหนี้สินรวมต่อส่วนของผู้ถือหุ้น (หน่วย: เท่า) [Debt to Equity]
                    'return_of_equity' : ROE,  #อัตราผลตอบแทนจากส่วนของผู้ถือหุ้น (หน่วย: %) [Return of Equity]
                    'registered_date' : regis_date,  #วันที่จดทะเบียนจัดตั้ง
                    'directors_list' : director_list,  #Directors,  #รายชื่อคณะกรรมการ
                    'directors_date' : directors_date,  #วันที่กรรมการมีผลฯ
                    'authorized_director' : authorized_director,   #กรรมการลงชื่อผูกพัน
                    'location' : locate,   #ที่ตั้ง
                    'email' : email,  #อีเมล
                    'phone_number' : tel,   #หมายเลขโทรศัพท์
                    'business_code' : b_code,   #รหัสประเภทธุรกิจ
                    'business_type' : b_type,   #ประเภทธุรกิจ
                    'scraped_date' : scraped_date, #วันที่ดึงข้อมูล
                }

                print("***************** ข้อมูลทั้งหมด *****************")
                print(detail)
                print("***************** ข้อมูลทั้งหมด *****************")
                DBD.append(detail)
                print(len(DBD))

                browser.back()
                browser.back()
                browser.back()
                browser.back()

                #-------------------------------------------------------------------------------------------------
                ##แก้ปัญหาการกลับมาหน้าแรกแล้วฟังก์ชันการกรองของเว็บหายด้วยการกดใหม่อีกรอบ

                sort = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[4]/div[2]/div/div[1]/div[1]/ul/li[3]/button')))
                browser.execute_script("arguments[0].click();", sort)

                Choose_type = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="filterBoxContent"]/div[2]/h4')))
                browser.execute_script("arguments[0].click();", Choose_type)

                ####บริษัทจำกัด Company Limited
                CoLtd = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="collapse2"]/div[2]/label')))
                time.sleep(2)
                action.double_click(CoLtd).perform()


                ####เลือกสถานะ
                choose_status = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[4]/div[2]/div/div[1]/div[1]/ul/li[3]/div/form/div[1]/div[3]/h4')))
                browser.execute_script("arguments[0].click();", choose_status)
                

                ####สถานะยังดำเนินกิจการอยู่
                status = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[4]/div[2]/div/div[1]/div[1]/ul/li[3]/div/form/div[1]/div[3]/div/ul/li[1]/div[1]/label')))
                time.sleep(2)
                action.double_click(status).perform()

                confirm_button = browser.find_element(By.CLASS_NAME, 'btn.btn-link.btn-confirm')
                browser.execute_script("arguments[0].click();", confirm_button)
                browser.set_page_load_timeout(30)
                time.sleep(3)

                #------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
            
            if check == 0:
                next = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[4]/div[2]/div/div[4]/div[1]/nav/ul/li[6]/button')))
                browser.execute_script("arguments[0].click();", next)
                browser.set_page_load_timeout(30)

            elif check >=1:
                num_box = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[4]/div[2]/div/div[4]/div[1]/nav/ul/li[2]/input')))
                browser.execute_script("arguments[0].click();", num_box)
                num_box.clear()
                time.sleep(3)
                #num_box.send_keys(Keys.DELETE)
                ans_page = 63748 - (63748 - (check + 1))
                num_box.send_keys(str(ans_page))
                num_box.send_keys(Keys.ENTER)
                browser.set_page_load_timeout(30)

            check +=1

        df = pd.DataFrame(DBD)
        df.to_excel(f'{file_name}.xlsx', index = False, header=True)

        #-----------------------------------------------------------------------------------------------------------------
        ##Change file format

        df2 = pd.read_excel(f'{file_name}.xlsx')

        DBD2 = []
        for i in range(len(df2)):
            director_list2 = df2.iloc[i][13].strip("'][").split(', ')
            
            for j in range(len(director_list2)):
                detail2 = {
                'juristic_ID' : df2.iloc[i][0],  #เลขนิติบุคคล
                'juristic_name' : df2.iloc[i][1],   #ชื่อนิติบุคคล
                'juristic_type' : df2.iloc[i][2], #ประเภทนิติบุคคล
                'registered_capital' : df2.iloc[i][3], #ทุนจดทะเบียน
                'main_income' : df2.iloc[i][4], #รายได้หลัก
                'cost_of_sales': df2.iloc[i][5],  #ต้นทุนขาย (Cost of Sales)
                'percentage1' : df2.iloc[i][6],  #ต้นทุนขายคิดเป็น % จากรายได้ คิดได้จาก (ต้นทุนขาย / รายได้หลัก) *100
                'Net_PL' : df2.iloc[i][7],  #กำไร(ขาดทุน)สุทธิ (Net Profit/Loss)
                'percentage2' : df2.iloc[i][8],  #กำไรคิดเป็น % จากรายได้ คิดได้จาก (กำไร(ขาดทุน)สุทธิ / รายได้หลัก) *100
                'current_ratio' : df2.iloc[i][9],  #อัตราส่วนทุนหมุนเวียน(หน่วย: เท่า) [Current Ratio]
                'debt_to_equity' : df2.iloc[i][10],  #อัตราส่วนหนี้สินรวมต่อส่วนของผู้ถือหุ้น (หน่วย: เท่า) [Debt to Equity]
                'return_of_equity' : df2.iloc[i][11],  #อัตราผลตอบแทนจากส่วนของผู้ถือหุ้น (หน่วย: %) [Return of Equity]
                'registered_date' : df2.iloc[i][12],  #วันที่จดทะเบียนจัดตั้ง
                'directors_list' : director_list2[j].replace("'","").strip(),  #Directors,  #รายชื่อคณะกรรมการ
                'directors_date' : df2.iloc[i][14],  #วันที่กรรมการมีผลฯ
                'authorized_director' : df2.iloc[i][15],   #กรรมการลงชื่อผูกพัน
                'location' : df2.iloc[i][16],   #ที่ตั้ง
                'email' : df2.iloc[i][17],  #อีเมล
                'phone_number' : df2.iloc[i][18],   #หมายเลขโทรศัพท์
                'business_code' : df2.iloc[i][19],   #รหัสประเภทธุรกิจ
                'business_type' : df2.iloc[i][20],   #ประเภทธุรกิจ
                'scraped_date' : df2.iloc[i][21], #วันที่ดึงข้อมูล
                }
                
                DBD2.append(detail2)

        df3 = pd.DataFrame(DBD2)
        df3.to_excel(f'{file_name}.xlsx', index = False, header=True)

    finally:
        print('Finished Scraping')
        browser.quit()