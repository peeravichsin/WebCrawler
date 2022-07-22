import pandas as pd
import re
from datetime import date as dadadate
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from time import sleep
import datetime
def govspend(st_date, en_date, file_name):
    PATH = r'.\Cdriver\chromedriver.exe'
    driver = webdriver.Chrome(PATH)
    action = webdriver.ActionChains(driver)

    driver.get("http://www.gprocurement.go.th/new_index.html")
    
    # CTRL+SHIFT+I

    try:

        #temporary use until 21 june 2022 
        Close_popup = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.ID, "home-popup-close-button")))

        if Close_popup.is_displayed():
            Close_popup.click()

        else: 
            pass
        ###  Get into website ###

        search_link = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.CLASS_NAME, "btn-org.ff"))
        )
        search_link.click() 
        
                            ###  User input here ###

        select_type = Select(WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.ID, "announceType"))
        ))

        select_type.select_by_index(1)

        select_method = Select(WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.ID, "methodId"))
        ))

        select_method.select_by_index(2)

        stl_date = str(st_date).split('-')
        enl_date = str(en_date).split('-')
        print(st_date)
        print(en_date)
        stc_date = dadadate(int(stl_date[0]), int(stl_date[1]) , int(stl_date[2]))
        enc_date = dadadate(int(enl_date[0]), int(enl_date[1]) , int(enl_date[2]))
        delta = enc_date - stc_date

        if delta.days <= 90:
            st_date = stl_date[2] + stl_date[1] + str(int(stl_date[0]) + 543)
            en_date = enl_date[2] + enl_date[1] + str(int(enl_date[0]) + 543)

        else:
            print('day is out of range for searching rule')


        S_date = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "announceSDate"))
        )
        S_date.send_keys(st_date)
        E_date = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "announceEDate"))
        )
        E_date.send_keys(en_date)

                            ###  User input here ###
        # project_status = Select(WebDriverWait(driver, 2).until(
        # EC.presence_of_element_located((By.ID, "projectStatus"))
        # ))

        # project_status.select_by_index(2)

        search = WebDriverWait(driver, 1).until(
            EC.presence_of_element_located((By.CLASS_NAME, "btnCommon"))
        )

        search.click()
        sleep(60)

        ###  Get into website ###

        ### Assign parameter ###

        Govspend = []       # the main list that collect all data
        btn=[2,3,4,5,6,7]   # list of index for change page
        itn=['ข้อมูลรายชื่อผู้ยื่นเอกสาร','ข้อมูลรายชื่อผู้ผ่านการพิจารณาคุณสมบัติและเทคนิค','ข้อมูลสาระสำคัญในสัญญา']
        check = 0
        no_of_proj = 1
    
        
        while check != no_of_proj :
            per_page = 0
            no_link = WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/form/table[1]/tbody/tr/td/table[3]/tbody/tr')))
            no_link = len(no_link)

            if no_link >= 50:
              per_page = 50
            elif no_link < 50 :
              per_page = no_link - 1

            for i in range(per_page):
                driver.implicitly_wait(60)
                links = WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="form1"]/table[1]/tbody/tr/td/table[3]/tbody/tr/td[7]/a')))
                seen = links[i].is_displayed()
                print("Is webdriver see this : "+str(seen))

                if seen is False :
                    page = btn[0]
                    change_page = driver.find_elements(By.XPATH, '/html/body/form/table[1]/tbody/tr/td/table[4]/tbody/tr/td[2]/div/table/tbody/tr/td')
                    change_page[page].click()
                    print('Click again') 

                links[i].click()
                proj_no = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="trProjectId"]/td[2]/input'))).get_attribute("value")
                corp_name = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/table[1]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr[2]/td[2]/input'))).get_attribute("value")
                provide_method = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="trMethod"]/td[2]/input'))).get_attribute("value")
                provide_type = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="trType"]/td[2]/input'))).get_attribute("value")
                proj_type = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="trGovStatus"]/td[2]/input'))).get_attribute("value")
                province = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="moiName"]'))).get_attribute("value")
                
                proj_cost = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="trPriceBuild"]/td[2]/input'))).get_attribute("value")
                proj_cost = proj_cost.replace(',','')
                proj_cost = float(proj_cost)

                proj_status = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="trProjectStatus"]/td[2]/input'))).get_attribute("value")
                
                applicants = []
                buy_date = []
                send_date = []
                corp_pass = []

                DS_date = {}

                proj_name = ''
                win_tin = ''
                winner = ''
                eGP_no = ''
                contact_no = ''
                contact_date = ''
                contact_price = ''
                contact_status = ''
                description = ''

                proj_name = ''
                bidder_tin = ''
                bidder = ''
                bid_price = ''


                for j in itn:

                    if j == 'ข้อมูลรายชื่อผู้ยื่นเอกสาร':
                        try:
                            applicants_link = driver.find_element_by_partial_link_text(j)
                            if applicants_link.is_displayed():
                                applicants_link.click()
                                proj_applicants = WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH,'/html/body/form/table[1]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr[14]/td/table/tbody/tr/td[1]')))
                                for proj_applicant in proj_applicants:
                                    applicants.append(proj_applicant.text)
                                applicants.pop(0)
                                
                                
                                receives_date = WebDriverWait(driver, 5).until(EC.presence_of_all_elements_located((By.XPATH,'/html/body/form/table[1]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr[14]/td/table/tbody/tr/td[2]')))
                                for receive_date in receives_date:
                                    buy_date.append(receive_date.text)
                                buy_date.pop(0)

                                # print('<------------ วันที่รับซื้อเอกสาร ------------>')
                                # print(buy_date)


                                acceptances_date = WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH,'/html/body/form/table[1]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr[14]/td/table/tbody/tr/td[3]')))
                                for acceptance_date in acceptances_date:
                                    send_date.append(acceptance_date.text)
                                send_date.pop(0)

                                # print('<------------ วันที่ยื่นเอกสาร  ------------>')
                                # print(send_date) 
                            # DS_date = {applicants[i].strip() : [buy_date[i].strip(),send_date[i].strip()] for i in range(len(applicants))}

                            driver.back()
                        except NoSuchElementException:
                            applicants = 'ไม่มีข้อมูล'
                            buy_date = 'ไม่มีข้อมูล'
                            send_date = 'ไม่มีข้อมูล'

                #         # print('<------------ ข้อมูลรายชื่อผู้ยื่นเอกสาร ------------>')
                #         # print(applicants)
                            
                    elif j == 'ข้อมูลรายชื่อผู้ผ่านการพิจารณาคุณสมบัติและเทคนิค':
                        try:
                            bidder_pass_link = driver.find_element_by_partial_link_text(j)
                            if bidder_pass_link.is_displayed():
                                bidder_pass_link.click()
                                spec_passes = WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH,'/html/body/form/table[1]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr[14]/td/table/tbody/tr/td[1]')))
                                for spec_pass in spec_passes:
                                    corp_pass.append(spec_pass.text)
                                corp_pass.pop(0)

                                # print('<------------ ข้อมูลรายชื่อผู้ผ่านการพิจารณาคุณสมบัติและเทคนิค ------------>')
                                # print(corp_pass)

                            
                            driver.back()
                        except NoSuchElementException:
                            corp_pass = 'ไม่มีข้อมูล'
                            
                        

                    elif j == 'ข้อมูลสาระสำคัญในสัญญา':
                        try:
                            summary_link = driver.find_element_by_partial_link_text(j)
                            if summary_link.is_displayed():
                                summary_link.click()
                            
                                winner_detail = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH,'/html/body/form/table[1]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr[14]/td/table/tbody/tr[2]/td')))
                                win_tin = winner_detail[1].text
                                winner = winner_detail[2].text
                                eGP_no = winner_detail[3].text
                                contact_no = winner_detail[4].text
                                contact_date = winner_detail[5].text

                                contact_price = winner_detail[6].text

                                if contact_price == '' or re.match(r'\s',contact_price):
                                    contact_price = 0

                                else:    
                                    contact_price = contact_price.replace(',','')
                                    contact_price = float(contact_price)

                                contact_status = winner_detail[7].text
                                description = winner_detail[8].text

                                # print('<------------ รายละเอียดผู้ชนะ ------------>')
                                # print(f'เลขภาษี : {win_tin}, ชื่อผู้ชนะ : {winner}, เลขคุมสัญญาในระบบ e-GP : {eGP_no}, เลขที่สัญญา : {contact_no}, วันที่ทำสัญญา : {contact_date}, จำนวนเงิน : {contact_price}, สถานะสัญญา : {contact_status}, เหตุผล : {description}')
                                
                                bidding_detail = WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH,'/html/body/form/table[1]/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td')))

                                proj_name = bidding_detail[1].text

                                bidder_tin = bidding_detail[2].text
                                bidder_tin = bidder_tin.split('\n')

                                bidder = bidding_detail[3].text
                                bidder = bidder.split('\n')

                                bid_price = bidding_detail[4].text
                                bid_price = bid_price.split('\n')
                            driver.back()

                                # print(f'ชื่อโครงการ : {proj_name}, เลขที่เสียภาษีชื่อผู้เสนอราคา : {bidder_tin}, ชื่อผู้เสนอ : {bidder}, ราคาที่เสนอ : {bid_price}')
                        except NoSuchElementException:
                            winner_detail = 'ไม่มีข้อมูล'

                            win_tin = winner_detail
                            winner = winner_detail
                            eGP_no = winner_detail
                            contact_no = winner_detail
                            contact_date = winner_detail
                            contact_price = 0.0
                            contact_status = winner_detail
                            description = winner_detail

                            bidding_detail = 'ไม่มีข้อมูล'

                            proj_name = bidding_detail
                            bidder_tin = bidding_detail
                            bidder = bidding_detail
                            bid_price = bidding_detail
                driver.back()
                scrap_date = datetime.datetime.now()
                scrap_date = scrap_date.strftime("%d/%m/%Y")
                
                
                detail = {
                            "proj_no" : proj_no,                 #เลขที่โครงการ
                            "proj_name" : proj_name,             #ชื่อโครงการ
                            "subdep_name" : corp_name,           #หน่วยงาน
                            "mthd_name" : provide_method,        #วิธีการจัดหา
                            "typ_name" : provide_type,           #ประเภทการจัดหา
                            "proj_type" : proj_type,             #ประเภทโครงการ
                            "province" : province,               #จังหวัด
                            "proj_cost" : proj_cost,             #ราคากลาง
                            "proj_status" : proj_status,         #สถานะโครงการ
                            "applicants" : applicants ,          #รายชื่อผู้ยื่น
                            "corp_pass"  : corp_pass,            #รายชื่อผู้ผ่านการพจารณา list
                            "bidder_tin" : bidder_tin,           #เลขที่เสียภาษีชื่อผู้เสนอราคา list
                            "bidder" : bidder,                   #ชื่อผู้เสนอ list
                            "buy_date" : buy_date,               #วันที่รับซื้อเอกสาร  list
                            "send_date" : send_date,             #วันที่ยื่นเอกสาร list
                            "bid_price" : bid_price,             #ราคาที่เสนอ list 
                            "win_tin" : win_tin,                 #เลขที่เสียภาษีของผู้ชนะ
                            "winner" : winner,                   #ชื่อผู้ชนะ
                            "eGP_no" : eGP_no,                   #เลขคุมสัญญาในระบบ e-GP
                            "contact_no" : contact_no,           #เลขที่สัญญา
                            "contact_date" : contact_date,       #วันที่ทำสัญญา
                            "contact_price" : contact_price,     #จำนวนเงิน
                            "contact_status" : contact_status,   #สถานะสัญญา
                            "descr" : description,               #เหตุผล
                            "scrap_date" : scrap_date            #วันที่ดึงข้อมูล
                            }

                print("***************** สรุปข้อมูลทั้งหมดของโครงการ *****************")
                print(detail)
                print("***************** สรุปข้อมูลทั้งหมดของโครงการ *****************")
                Govspend.append(detail)
                print(len(Govspend))

            
                driver.back()

                
                if i % 10 == 9:
                    btn.pop(0)
                    if len(btn) > 0:
                        page = btn[0]
                        if page == 7:
                            try:
                                tb = driver.find_element(By.XPATH, '/html/body/form/table[1]/tbody/tr/td/table[4]/tbody/tr/td[2]/div/table/tbody/tr/td[8]/b')
                                if tb.is_displayed():
                                    tb.click()
                                    print("-----------------> Next")
                                    btn=[2,3,4,5,6,7]
                                    sleep(30)
                            except:
                                fb = driver.find_element(By.XPATH, '/html/body/form/table[1]/tbody/tr/td/table[4]/tbody/tr/td[2]/div/table/tbody/tr/td[8]/font')
                                if fb.is_displayed():
                                    print("XXXXXXXX    Stop   XXXXXXXX")
                                    check += 1
                                    break
                            
                        elif page != 7:
                            change_page = driver.find_elements(By.XPATH, '/html/body/form/table[1]/tbody/tr/td/table[4]/tbody/tr/td[2]/div/table/tbody/tr/td')
                            change_page[page].click()
                            

            
        rdf = pd.DataFrame(Govspend)
        rdf.to_excel(rf'.\sele\Data\{file_name}.xlsx', index = False, header=True)
        sleep(5)
        df = pd.read_excel(rf'.\sele\Data\{file_name}.xlsx')

        gov = []
        for i in range(len(df)):
            ds_date = {}
            if df.iloc[i][1] != 'ไม่มีข้อมูล':
                    applicants = df.iloc[i][9].strip("'][").split(', ')
                    buy_date = df.iloc[i][13].strip("'][").split(', ')
                    send_date = df.iloc[i][14].strip("'][").split(', ')
                    bidder_tin = df.iloc[i][11].strip("'][").split(', ')
                    bidder = df.iloc[i][12].strip("'][").split(', ')
                    bid_price =	df.iloc[i][15].strip("'][").split(', ')
                            
                    for j in range(len(bidder)):
                        detail = {
                                    "proj_no" : df.iloc[i][0],                                      #เลขที่โครงการ
                                    "proj_name" :  df.iloc[i][1],                                   #ชื่อโครงการ
                                    "subdep_name" :  df.iloc[i][2],                                 #หน่วยงาน
                                    "mthd_name" : df.iloc[i][3],                                    #วิธีการจัดหา
                                    "typ_name" : df.iloc[i][4],                                     #ประเภทการจัดหา
                                    "proj_type" : df.iloc[i][5],                                    #ประเภทโครงการ
                                    "province" : df.iloc[i][6],                                     #จังหวัด
                                    "proj_cost" : df.iloc[i][7],                                    #ราคากลาง
                                    "proj_status" : df.iloc[i][8],                                  #สถานะโครงการ
                                    "applicants" : df.iloc[i][9],                                   #ผู้ยื่นราคา
                                    "buy_date" : df.iloc[i][13],                                    #วันที่รับซื้อเอกสาร
                                    "send_date" : df.iloc[i][14],                                   #วันที่ยื่นเอกสาร
                                    "bidder_tin" : bidder_tin[j].replace("'","").strip(),           #เลขที่เสียภาษีชื่อผู้เสนอราคา
                                    "bidder" : bidder[j].replace("'","").strip(),                   #ชื่อผู้เสนอ
                                    "bid_price" : bid_price[j].replace("'","").strip(),             #ราคาที่เสนอ
                                    "corp_pass" : df.iloc[i][10],                                   #ผู้ผ่านการคัดเลือก
                                    "win_tin" : df.iloc[i][16],                                     #เลขที่เสียภาษีของผู้ชนะ
                                    "winner" : df.iloc[i][17],                                      #ชื่อผู้ชนะ
                                    "eGP_no" : df.iloc[i][18],                                      #เลขคุมสัญญาในระบบ e-GP
                                    "contact_no" : df.iloc[i][19],                                  #เลขที่สัญญา
                                    "contact_date" : df.iloc[i][20],                                #วันที่ทำสัญญา
                                    "contact_price" : df.iloc[i][21],                               #จำนวนเงิน
                                    "contact_status" : df.iloc[i][22],                              #สถานะสัญญา
                                    "descr" : df.iloc[i][23],                                       #เหตุผล
                                    "scrap_date" : df.iloc[i][24]                                   #วันที่ดึงข้อมูล
                                    }
                        # print(detail)
                        gov.append(detail)
            elif df.iloc[i][1] == 'ไม่มีข้อมูล':
                detail = {
                                    "proj_no" : df.iloc[i][0],                                      #เลขที่โครงการ
                                    "proj_name" :  df.iloc[i][1],                                   #ชื่อโครงการ
                                    "subdep_name" :  df.iloc[i][2],                                 #หน่วยงาน
                                    "mthd_name" : df.iloc[i][3],                                    #วิธีการจัดหา
                                    "typ_name" : df.iloc[i][4],                                     #ประเภทการจัดหา
                                    "proj_type" : df.iloc[i][5],             #ประเภทโครงการ
                                    "province" : df.iloc[i][6],               #จังหวัด
                                    "proj_cost" : df.iloc[i][7],             #ราคากลาง
                                    "proj_status" : df.iloc[i][8],         #สถานะโครงการ
                                    "applicants" : df.iloc[i][9],           #ผู้ยื่นราคา
                                    "corp_pass"  :  df.iloc[i][10],        #ผู้ผ่านการคัดเลือก
                                    "buy_date" : df.iloc[i][13],               #วันที่รับซื้อเอกสาร
                                    "send_date" : df.iloc[i][14],             #วันที่ยื่นเอกสาร
                                    "bidder_tin" : df.iloc[i][11],           #เลขที่เสียภาษีชื่อผู้เสนอราคา
                                    "bidder" : df.iloc[i][12],                   #ชื่อผู้เสนอ
                                    "bid_price" : df.iloc[i][15],             #ราคาที่เสนอ
                                    "win_tin" : df.iloc[i][16],                 #เลขที่เสียภาษีของผู้ชนะ
                                    "winner" : df.iloc[i][17],                   #ชื่อผู้ชนะ
                                    "eGP_no" : df.iloc[i][18],                   #เลขคุมสัญญาในระบบ e-GP
                                    "contact_no" : df.iloc[i][19],           #เลขที่สัญญา
                                    "contact_date" : df.iloc[i][20],       #วันที่ทำสัญญา
                                    "contact_price" : df.iloc[i][21],     #จำนวนเงิน
                                    "contact_status" : df.iloc[i][22],   #สถานะสัญญา
                                    "descr" : df.iloc[i][23],               #เหตุผล
                                    "scrap_date" : df.iloc[i][24]            #วันที่ดึงข้อมูล
                                    }
                # print(detail)
                gov.append(detail)
        gdf = pd.DataFrame(gov)
        gov2 = []
        ad = 0
        for i in range(len(gdf)):
            ds_date = {}
            if gdf.iloc[i][1] != 'ไม่มีข้อมูล':
                No_pro = gdf.iloc[i][0]
                applicants = gdf.iloc[i][9].strip("'][").split(', ')
                buy_date = gdf.iloc[i][10].strip("'][").split(', ')
                send_date = gdf.iloc[i][11].strip("'][").split(', ')
                corp_pass = gdf.iloc[i][15].strip("'][").split(', ')
                corp_pass =[corp_pass[i].replace("'","").strip() for i in range((len(corp_pass)))]
            
                ds_date = { applicants[i].replace("'","").strip() :  [buy_date[i].replace("'","").strip(),  send_date[i].replace("'","").strip()]  for i in range(len(applicants)) }
                bidder = gdf.iloc[i][13]
                if bidder in corp_pass:
                    pass_stat = True
                else:
                    pass_stat = False

                for key,value in ds_date.items():
                    if key.startswith(f'{bidder}'):
                        date = ds_date[key]
                
                detail = {
                                    "proj_no" : gdf.iloc[i][0],                 #เลขที่โครงการ
                                    "proj_name" :  gdf.iloc[i][1],             #ชื่อโครงการ
                                    "subdep_name" :  gdf.iloc[i][2],           #หน่วยงาน
                                    "mthd_name" : gdf.iloc[i][3],        #วิธีการจัดหา
                                    "typ_name" : gdf.iloc[i][4],           #ประเภทการจัดหา
                                    "proj_type" : gdf.iloc[i][5],             #ประเภทโครงการ
                                    "province" : gdf.iloc[i][6],               #จังหวัด
                                    "proj_cost" : gdf.iloc[i][7],             #ราคากลาง
                                    "proj_status" : gdf.iloc[i][8],         #สถานะโครงการ
                                    # "applicants" : df.iloc[i][9],           #ผู้ยื่นราคา
                                    "corp_pass"  :  pass_stat,        #ผู้ผ่านการคัดเลือก
                                    "buy_date" : date[0],               #วันที่รับซื้อเอกสาร
                                    "send_date" : date[1],             #วันที่ยื่นเอกสาร
                                    "bidder_tin" : gdf.iloc[i][12],           #เลขที่เสียภาษีชื่อผู้เสนอราคา
                                    "bidder" : gdf.iloc[i][13],                   #ชื่อผู้เสนอ
                                    "bid_price" : gdf.iloc[i][14],             #ราคาที่เสนอ
                                    "win_tin" : gdf.iloc[i][16],                 #เลขที่เสียภาษีของผู้ชนะ
                                    "winner" : gdf.iloc[i][17],                   #ชื่อผู้ชนะ
                                    "eGP_no" : gdf.iloc[i][18],                   #เลขคุมสัญญาในระบบ e-GP
                                    "contact_no" : gdf.iloc[i][19],           #เลขที่สัญญา
                                    "contact_date" : gdf.iloc[i][20],       #วันที่ทำสัญญา
                                    "contact_price" : gdf.iloc[i][21],     #จำนวนเงิน
                                    "contact_status" : gdf.iloc[i][22],   #สถานะสัญญา
                                    "descr" : gdf.iloc[i][23],               #เหตุผล
                                    "scrap_date" : gdf.iloc[i][24]            #วันที่ดึงข้อมูล
                                    }
                # print(detail)
                gov2.append(detail)
            elif gdf.iloc[i][1] == 'ไม่มีข้อมูล':
                detail = {
                                    "proj_no" : gdf.iloc[i][0],                 #เลขที่โครงการ
                                    "proj_name" :  gdf.iloc[i][1],             #ชื่อโครงการ
                                    "subdep_name" :  gdf.iloc[i][2],           #หน่วยงาน
                                    "mthd_name" : gdf.iloc[i][3],        #วิธีการจัดหา
                                    "typ_name" : gdf.iloc[i][4],           #ประเภทการจัดหา
                                    "proj_type" : gdf.iloc[i][5],             #ประเภทโครงการ
                                    "province" : gdf.iloc[i][6],               #จังหวัด
                                    "proj_cost" : gdf.iloc[i][7],             #ราคากลาง
                                    "proj_status" : gdf.iloc[i][8],         #สถานะโครงการ
                                    # "applicants" : gdf.iloc[i][9],           #ผู้ยื่นราคา
                                    "corp_pass"  :  gdf.iloc[i][10],        #ผู้ผ่านการคัดเลือก
                                    "buy_date" : gdf.iloc[i][13],               #วันที่รับซื้อเอกสาร
                                    "send_date" : gdf.iloc[i][14],             #วันที่ยื่นเอกสาร
                                    "bidder_tin" : gdf.iloc[i][11],           #เลขที่เสียภาษีชื่อผู้เสนอราคา
                                    "bidder" : gdf.iloc[i][12],                   #ชื่อผู้เสนอ
                                    "bid_price" : gdf.iloc[i][15],             #ราคาที่เสนอ
                                    "win_tin" : gdf.iloc[i][16],                 #เลขที่เสียภาษีของผู้ชนะ
                                    "winner" : gdf.iloc[i][17],                   #ชื่อผู้ชนะ
                                    "eGP_no" : gdf.iloc[i][18],                   #เลขคุมสัญญาในระบบ e-GP
                                    "contact_no" : gdf.iloc[i][19],           #เลขที่สัญญา
                                    "contact_date" : gdf.iloc[i][20],       #วันที่ทำสัญญา
                                    "contact_price" : gdf.iloc[i][21],     #จำนวนเงิน
                                    "contact_status" : gdf.iloc[i][22],   #สถานะสัญญา
                                    "descr" : gdf.iloc[i][23],               #เหตุผล
                                    "scrap_date" : gdf.iloc[i][24]            #วันที่ดึงข้อมูล
                                    }
                # print(detail)
                gov2.append(detail)
        
        g2df = pd.DataFrame(gov2)
        g2df.to_excel(rf'.\sele\Data\{file_name}.xlsx', index = False, header=True)

    finally:
        print('Finished Scraping')
        driver.quit()


