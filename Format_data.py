import pandas as pd

def format_data(file_name):
    df = pd.read_excel(f'{file_name}.xlsx')
    if df.isnull().values.any():
        df.drop(df.index[len(df)-1], axis=0, inplace=True)
    gov = []
    for i in range(len(df)):
        ds_date = {}
        if df.iloc[i][1] != 'ไม่มีข้อมูล':
                applicants = df.iloc[i][10].strip("'][").split(', ')
                buy_date = df.iloc[i][14].strip("'][").split(', ')
                send_date = df.iloc[i][15].strip("'][").split(', ')
                bidder_tin = df.iloc[i][12].strip("'][").split(', ')
                bidder = df.iloc[i][13].strip("'][").split(', ')
                bid_price =	df.iloc[i][16].strip("'][").split(', ')
                        
                for j in range(len(bidder)):
                    detail = {
                                "proj_no" : df.iloc[i][0],                                      #เลขที่โครงการ
                                "proj_name" :  df.iloc[i][1],                                   #ชื่อโครงการ
                                "subdep_name" :  df.iloc[i][2],                                 #หน่วยงาน
                                "mthd_name" : df.iloc[i][3],                                    #วิธีการจัดหา
                                "typ_name" : df.iloc[i][4],                                     #ประเภทการจัดหา
                                "proj_type" : df.iloc[i][5],                                    #ประเภทโครงการ
                                "province" : df.iloc[i][6],                                     #จังหวัด
                                "proj_money" : df.iloc[i][7],                                   #งบประมาณ   
                                "proj_cost" : df.iloc[i][8],                                    #ราคากลาง
                                "proj_status" : df.iloc[i][9],                                  #สถานะโครงการ
                                "applicants" : df.iloc[i][10],                                   #ผู้ยื่นราคา
                                "buy_date" : df.iloc[i][14],                                    #วันที่รับซื้อเอกสาร
                                "send_date" : df.iloc[i][15],                                   #วันที่ยื่นเอกสาร
                                "bidder_tin" : bidder_tin[j].replace("'","").strip(),           #เลขที่เสียภาษีชื่อผู้เสนอราคา
                                "bidder" : bidder[j].replace("'","").strip(),                   #ชื่อผู้เสนอ
                                "bid_price" : bid_price[j].replace("'","").strip(),             #ราคาที่เสนอ
                                "corp_pass" : df.iloc[i][11],                                   #ผู้ผ่านการคัดเลือก
                                "win_tin" : df.iloc[i][17],                                     #เลขที่เสียภาษีของผู้ชนะ
                                "winner" : df.iloc[i][18],                                      #ชื่อผู้ชนะ
                                "eGP_no" : df.iloc[i][19],                                      #เลขคุมสัญญาในระบบ e-GP
                                "contact_no" : df.iloc[i][20],                                  #เลขที่สัญญา
                                "contact_date" : df.iloc[i][21],                                #วันที่ทำสัญญา
                                "contact_price" : df.iloc[i][22],                               #จำนวนเงิน
                                "contact_status" : df.iloc[i][23],                              #สถานะสัญญา
                                "descr" : df.iloc[i][24],                                       #เหตุผล
                                "scrap_date" : df.iloc[i][25]                                   #วันที่ดึงข้อมูล
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
                                "proj_money" : df.iloc[i][7],                                   #งบประมาณ
                                "proj_cost" : df.iloc[i][8],             #ราคากลาง
                                "proj_status" : df.iloc[i][9],         #สถานะโครงการ
                                "applicants" : df.iloc[i][10],           #ผู้ยื่นราคา
                                "corp_pass"  :  df.iloc[i][11],        #ผู้ผ่านการคัดเลือก
                                "buy_date" : df.iloc[i][14],               #วันที่รับซื้อเอกสาร
                                "send_date" : df.iloc[i][15],             #วันที่ยื่นเอกสาร
                                "bidder_tin" : df.iloc[i][12],           #เลขที่เสียภาษีชื่อผู้เสนอราคา
                                "bidder" : df.iloc[i][13],                   #ชื่อผู้เสนอ
                                "bid_price" : df.iloc[i][14],             #ราคาที่เสนอ
                                "win_tin" : df.iloc[i][17],                 #เลขที่เสียภาษีของผู้ชนะ
                                "winner" : df.iloc[i][18],                   #ชื่อผู้ชนะ
                                "eGP_no" : df.iloc[i][19],                   #เลขคุมสัญญาในระบบ e-GP
                                "contact_no" : df.iloc[i][20],           #เลขที่สัญญา
                                "contact_date" : df.iloc[i][21],       #วันที่ทำสัญญา
                                "contact_price" : df.iloc[i][22],     #จำนวนเงิน
                                "contact_status" : df.iloc[i][23],   #สถานะสัญญา
                                "descr" : df.iloc[i][24],               #เหตุผล
                                "scrap_date" : df.iloc[i][25]            #วันที่ดึงข้อมูล
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
            applicants = gdf.iloc[i][10].strip("'][").split(', ')
            buy_date = gdf.iloc[i][11].strip("'][").split(', ')
            send_date = gdf.iloc[i][12].strip("'][").split(', ')
            corp_pass = gdf.iloc[i][16].strip("'][").split(', ')
            corp_pass =[corp_pass[i].replace("'","").strip() for i in range((len(corp_pass)))]
        
            ds_date = { applicants[i].replace("'","").strip() :  [buy_date[i].replace("'","").strip(),  send_date[i].replace("'","").strip()]  for i in range(len(applicants)) }
            bidder = gdf.iloc[i][14]
            # print(f'B : {bidder}')
            # print(f'c : {corp_pass}')
            if bidder in corp_pass:
                pass_stat = True
            else:
                pass_stat = False

            for key,value in ds_date.items():
                if key.startswith(f'{bidder}'):
                    ds_date[key]
            
            detail = {
                                "proj_no" : gdf.iloc[i][0],                 #เลขที่โครงการ
                                "proj_name" :  gdf.iloc[i][1],             #ชื่อโครงการ
                                "subdep_name" :  gdf.iloc[i][2],           #หน่วยงาน
                                "mthd_name" : gdf.iloc[i][3],        #วิธีการจัดหา
                                "typ_name" : gdf.iloc[i][4],           #ประเภทการจัดหา
                                "proj_type" : gdf.iloc[i][5],             #ประเภทโครงการ
                                "province" : gdf.iloc[i][6],               #จังหวัด
                                "proj_money" : gdf.iloc[i][7],              #งบประมาณ
                                "proj_cost" : gdf.iloc[i][8],             #ราคากลาง
                                "proj_status" : gdf.iloc[i][9],         #สถานะโครงการ
                                "corp_pass"  :  pass_stat,        #ผู้ผ่านการคัดเลือก
                                "buy_date" : ds_date[key][0],               #วันที่รับซื้อเอกสาร
                                "send_date" : ds_date[key][1],             #วันที่ยื่นเอกสาร
                                "bidder_tin" : gdf.iloc[i][13],           #เลขที่เสียภาษีชื่อผู้เสนอราคา
                                "bidder" : gdf.iloc[i][14],                   #ชื่อผู้เสนอ
                                "bid_price" : gdf.iloc[i][15],             #ราคาที่เสนอ
                                "win_tin" : gdf.iloc[i][17],                 #เลขที่เสียภาษีของผู้ชนะ
                                "winner" : gdf.iloc[i][18],                   #ชื่อผู้ชนะ
                                "eGP_no" : gdf.iloc[i][19],                   #เลขคุมสัญญาในระบบ e-GP
                                "contact_no" : gdf.iloc[i][20],           #เลขที่สัญญา
                                "contact_date" : gdf.iloc[i][21],       #วันที่ทำสัญญา
                                "contact_price" : gdf.iloc[i][22],     #จำนวนเงิน
                                "contact_status" : gdf.iloc[i][23],   #สถานะสัญญา
                                "descr" : gdf.iloc[i][24],               #เหตุผล
                                "scrap_date" : gdf.iloc[i][25]            #วันที่ดึงข้อมูล
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
                                "proj_money" : gdf.iloc[i][7],                                   #งบประมาณ
                                "proj_cost" : gdf.iloc[i][8],             #ราคากลาง
                                "proj_status" : gdf.iloc[i][9],         #สถานะโครงการ
                                "corp_pass"  :  gdf.iloc[i][11],        #ผู้ผ่านการคัดเลือก
                                "buy_date" : gdf.iloc[i][14],               #วันที่รับซื้อเอกสาร
                                "send_date" : gdf.iloc[i][15],             #วันที่ยื่นเอกสาร
                                "bidder_tin" : gdf.iloc[i][12],           #เลขที่เสียภาษีชื่อผู้เสนอราคา
                                "bidder" : gdf.iloc[i][13],                   #ชื่อผู้เสนอ
                                "bid_price" : gdf.iloc[i][14],             #ราคาที่เสนอ
                                "win_tin" : gdf.iloc[i][17],                 #เลขที่เสียภาษีของผู้ชนะ
                                "winner" : gdf.iloc[i][18],                   #ชื่อผู้ชนะ
                                "eGP_no" : gdf.iloc[i][19],                   #เลขคุมสัญญาในระบบ e-GP
                                "contact_no" : gdf.iloc[i][20],           #เลขที่สัญญา
                                "contact_date" : gdf.iloc[i][21],       #วันที่ทำสัญญา
                                "contact_price" : gdf.iloc[i][22],     #จำนวนเงิน
                                "contact_status" : gdf.iloc[i][23],   #สถานะสัญญา
                                "descr" : gdf.iloc[i][24],               #เหตุผล
                                "scrap_date" : gdf.iloc[i][25]            #วันที่ดึงข้อมูล
                                }
            # print(detail)
            gov2.append(detail)
    
    g2df = pd.DataFrame(gov2)
    g2df.to_excel(f'{file_name}.xlsx', index = False, header=True)