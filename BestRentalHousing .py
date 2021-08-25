import xlrd
import requests
import json
import xlsxwriter
import time


API_key = "<自己的高德题图API_key>"

tables = []


def get_location(address):
    """获取经纬度"""
    url = "https://restapi.amap.com/v3/geocode/geo?address={}&output=json&key={}".format(address,API_key)
    headers = {
        'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36'
    }
    print(url)
    response = requests.get(url, headers=headers).text

    res = json.loads(response)
    geocodes = (res['geocodes'])
    location = ""
    #取第一个，也是最后一个
    for detail in geocodes[:] :
        location = detail['location']
        # print (detail['location'])

    return location

def get_way(outDoorDate,outDoorTime,from_address,to_address,city_code):
    """获取路径 可以考虑加上时间因素"""
    url = "https://restapi.amap.com/v3/direction/transit/integrated?&outDoorDate={}&outDoorTime={}&origin={}&destination={}&city={}&output=json&key={}".format(outDoorDate,outDoorTime,from_address,to_address,city_code,API_key)
    headers = {
        'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36'
    }
    print(url)
    response = requests.get(url, headers=headers).text
    res = json.loads(response)
    transits = (res['route']['transits'])
    costs = []
    #遍历出行方案
    for detail in transits[:] :
        #没有金额的情况时，默认5元
        if detail['cost'] == None or detail['cost'] == []:
            #默认5块钱通勤
            cost = "5"
        else:
            cost = detail['cost']
        cost = detail['cost']
        duration = detail['duration']
        walking_distance = detail['walking_distance']
        #花费，时间，步行距离
        costs.append((cost,duration,walking_distance))
        # print("cost：" + cost + "   duration：" + duration +"   walking_distance：" +walking_distance+ "\n")
        # print(costs)
        # print(""+cost+"元")
        # m, s = divmod(int(duration), 60)
        # h, m = divmod(m, 60)

        # print(""+str(h)+"小时"+str(m)+"分钟")
        # print("步行"+walking_distance+"米")
        # print("\n")
    return costs


def read_excel():
    # 打开文件
    path = r'./2beike_20210824012029833.xls'
    workbook = xlrd.open_workbook(path)
    # workbook = read_excel(path,engine='openpyxl')
    # 获取所有sheet
    sheet_name = workbook.sheet_names()[0]

    # 根据sheet索引或者名称获取sheet内容
    sheet = workbook.sheet_by_index(0) # sheet索引从0开始
    # sheet = workbook.sheet_by_name('Sheet1')

    #print (workboot.sheets()[0])
    # sheet的名称，行数，列数
    print (sheet.name,sheet.nrows,sheet.ncols)


    # 获取整行和整列的值（数组）
    # rows = sheet.row_values(1) # 获取第2行内容
    # cols = sheet.col_values(2) # 获取第3列内容
    # print (rows)
    # print (cols)

    for rown in range(sheet.nrows):
        array = {'城市':'','房源':'','房源网址':'','区县':'','商圈':'','小区':'',
                 '大小':'','朝向':'','户型':'','租金':'','来源':'','租房具体地址':''
                ,'租房经纬度':''
                ,'公司位置':'','公司经纬度':''
                ,'另一半公司位置':'','另一半公司经纬度':''
                ,'最短时间':'','最短时间花费':''
                ,'另一半最短时间':'','另一半最短时间花费':''
                ,'综合最短时间':'','综合最短时间花费':''
                ,'最少花费时间':'','最少花费':''
                ,'另一半最少花费时间':'','另一半最少花费':''
                ,'综合最少花费时间':'','综合最少花费':''}
        array['城市'] = sheet.cell_value(rown,0)
        array['房源'] = sheet.cell_value(rown,1)
        array['房源网址'] = sheet.cell_value(rown,2)
        array['区县'] = sheet.cell_value(rown,3)
        array['商圈'] = sheet.cell_value(rown,4)
        array['小区'] = sheet.cell_value(rown,5)
        array['大小'] = sheet.cell_value(rown,8) \
            if "㎡" in sheet.cell_value(rown,8) \
            else \
                sheet.cell_value(rown,7)  \
                    if "㎡" in sheet.cell_value(rown,7) \
                    else \
                        sheet.cell_value(rown,6) \
                            if "㎡" in sheet.cell_value(rown,6) \
                            else ""
        array['朝向'] = sheet.cell_value(rown,7) if sheet.cell_value(rown,6) else sheet.cell_value(rown,8)
        room = sheet.cell_value(rown,6) if "室" in sheet.cell_value(rown,6) else sheet.cell_value(rown,8)
        array['户型'] = room if "室" in room else ""
        #租金是范围的，算平均值
        amt = sheet.cell_value(rown,9).split(' ')[0]
        if "-" in amt:
            amt = (float(amt.split('-')[0])+float(amt.split('-')[1]))/2
        array['租金'] = amt
        array['来源'] = sheet.cell_value(rown,10)
        address = ""
        if array['区县'] :
            address = "广东省深圳市" + array['区县']+array['商圈']+array['小区']
        else:
            address = "广东省深圳市" + array['来源']
        array['租房具体地址'] = address

        tables.append(array)

    # print (len(tables))
    # print (tables)
    return tables

def excel_storage(tbname,response):
    """将字典数据写入excel"""
    workbook = xlsxwriter.Workbook("./out_{}Detail.xls".format(tbname))
    worksheet = workbook.add_worksheet()
    """设置标题加粗"""
    bold_format = workbook.add_format({'bold': True})
    worksheet.write('A1', '城市', bold_format)
    worksheet.write('B1', '房源', bold_format)
    worksheet.write('C1', '房源网址', bold_format)
    worksheet.write('D1', '区县', bold_format)
    worksheet.write('E1', '商圈', bold_format)
    worksheet.write('F1', '小区', bold_format)
    worksheet.write('G1', '房屋面积', bold_format)
    worksheet.write('H1', '房屋朝向', bold_format)
    worksheet.write('I1', '房屋户型', bold_format)
    worksheet.write('J1', '租金', bold_format)
    worksheet.write('K1', '来源', bold_format)
    worksheet.write('L1', '租房具体地址', bold_format)
    worksheet.write('M1', '租房经纬度', bold_format)
    worksheet.write('N1', '公司位置', bold_format)
    worksheet.write('O1', '公司经纬度', bold_format)
    worksheet.write('P1', '另一半公司位置', bold_format)
    worksheet.write('Q1', '另一半公司经纬度', bold_format)
    worksheet.write('R1', '最短时间', bold_format)
    worksheet.write('S1', '最短时间花费', bold_format)
    worksheet.write('T1', '另一半最短时间', bold_format)
    worksheet.write('U1', '另一半最短时间花费', bold_format)
    worksheet.write('V1', '综合最短时间', bold_format)
    worksheet.write('W1', '综合最短时间花费', bold_format)
    worksheet.write('X1', '最少花费时间', bold_format)
    worksheet.write('Y1', '最少花费', bold_format)
    worksheet.write('Z1', '另一半最少花费时间', bold_format)
    worksheet.write('AA1', '另一半最少花费', bold_format)
    worksheet.write('AB1', '综合最少花费时间', bold_format)
    worksheet.write('AC1', '综合最少花费', bold_format)
    row = 1
    col = 0
    for item in response:
        worksheet.write_string(row, col, str(item['城市']))
        worksheet.write_string(row, col + 1, str(item['房源']))
        worksheet.write_string(row, col + 2, str(item['房源网址']))
        worksheet.write_string(row, col + 3, str(item['区县']))
        worksheet.write_string(row, col + 4, str(item['商圈']))
        worksheet.write_string(row, col + 5, str(item['小区']))
        worksheet.write_string(row, col + 6, str(item['大小']))
        worksheet.write_string(row, col + 7, str(item['朝向']))
        worksheet.write_string(row, col + 8, str(item['户型']))
        worksheet.write_string(row, col + 9, str(item['租金']))
        worksheet.write_string(row, col + 10, str(item['来源']))
        worksheet.write_string(row, col + 11, str(item['租房具体地址']))
        worksheet.write_string(row, col + 12, str(item['租房经纬度']))
        worksheet.write_string(row, col + 13, str(item['公司位置']))
        worksheet.write_string(row, col + 14, str(item['公司经纬度']))
        worksheet.write_string(row, col + 15, str(item['另一半公司位置']))
        worksheet.write_string(row, col + 16, str(item['另一半公司经纬度']))
        worksheet.write_string(row, col + 17, str(item['最短时间']))
        worksheet.write_string(row, col + 18, str(item['最短时间花费']))
        worksheet.write_string(row, col + 19, str(item['另一半最短时间']))
        worksheet.write_string(row, col + 20, str(item['另一半最短时间花费']))
        worksheet.write_string(row, col + 21, str(item['综合最短时间']))
        worksheet.write_string(row, col + 22, str(item['综合最短时间花费']))
        worksheet.write_string(row, col + 23, str(item['最少花费时间']))
        worksheet.write_string(row, col + 24, str(item['最少花费']))
        worksheet.write_string(row, col + 25, str(item['另一半最少花费时间']))
        worksheet.write_string(row, col + 26, str(item['另一半最少花费']))
        worksheet.write_string(row, col + 27, str(item['综合最少花费时间']))
        worksheet.write_string(row, col + 28, str(item['综合最少花费']))
        row += 1
    workbook.close()

if __name__ == '__main__':
    # 读取Excel

    out_table = []
    tables = read_excel();
    del tables[0]
    i =0
    address1 = "自己的公司地址 越精确越好"
    address2 = "另一半的公司地址 越精确越好"
    #所在城市区号
    areaCode = "0755"
    ll1 = get_location(address1)
    ll2 = get_location(address2)
    #出发时间
    outDoorDate = "2021-08-25"
    outDoorTime = "08:00"
    for row in tables[:]:

        # 休眠0.06秒
        time.sleep(0.08)

        fromStr = get_location(row['租房具体地址'])
        row['租房经纬度'] = fromStr
        row['公司位置'] = address1
        row['公司经纬度'] = ll1
        row['另一半公司位置'] = address2
        row['另一半公司经纬度'] = ll2

        #自己的交通成本
        costs1 = get_way(outDoorDate,outDoorTime,fromStr,ll1,areaCode)

        #按照时间排序
        costs1.sort(key = lambda x:x[1])
        # print("按照时间排序"+str(costs))
        row["最短时间"]= float(costs1[0][1])
        # row["最短时间花费"]=str(cost)+"---"+str(costs[0][0])
        row["最短时间花费"]=float(costs1[0][0])

        #按照花费排序
        costs1.sort(key = lambda x:x[0])
        # print("按照花费排序"+str(costs))
        row["最少花费时间"]= float(costs1[0][1])
        # row["最少花费"]=str(cost)+"---"+str(costs[0][0])
        row["最少花费"]=float(costs1[0][0])


        #另一半的
        costs2 = get_way(outDoorDate,outDoorTime,fromStr,ll2,areaCode)

        #不需要计算另一半的，将另一半相关的都直接改为0

        #按照时间排序
        costs2.sort(key = lambda x:x[1])
        # print("按照时间排序"+str(costs))
        row["另一半最短时间"]= float(costs2[0][1])
        # row["最短时间花费"]=str(cost)+"---"+str(costs[0][0])
        row["另一半最短时间花费"]=float(costs2[0][0])

        #按照花费排序
        costs2.sort(key = lambda x:x[0])
        # print("按照花费排序"+str(costs))
        row["另一半最少花费时间"]= float(costs2[0][1])
        # row["最少花费"]=str(cost)+"---"+str(costs[0][0])
        row["另一半最少花费"]=float(costs2[0][0])

        if row["最短时间"] == None or row["最短时间"] == '':
            row["最短时间"] = 9999999.0
        if row["最少花费"] == None or row["最少花费"] == '':
            row["最少花费"] = 9999999.0

        if row["另一半最短时间"] == None or row["另一半最短时间"] == '':
            row["另一半最短时间"] = 9999999.0
        if row["另一半最少花费"] == None or row["另一半最少花费"] == '':
            row["另一半最少花费"] = 9999999.0

        row["综合最短时间"] = float(row["最短时间"])+float(row["另一半最短时间"])
        row["综合最短时间花费"] = float(row["租金"])+(float(row["最短时间花费"])+float(row["另一半最短时间花费"]))*2*22

        row["综合最少花费时间"] = float(row["最少花费时间"])+float(row["另一半最少花费时间"])
        row["综合最少花费"] = float(row["租金"])+(float(row["最少花费"])+float(row["另一半最少花费"]))*2*22

        out_table.append(row)

        i = i+1
        if(i>10):
            break
        print("当前进度  : "+ str(i) +"  /  "+ str(len(tables)))


    out_table.sort(key = lambda x:float(x["综合最短时间"]))
    excel_storage("0825最短时间",out_table)
    out_table.sort(key = lambda x:float(x["综合最少花费"]))
    excel_storage("0825最少花费",out_table)


    # fromStr = get_location("广东省深圳市南山区蛇口宏宝花园")
    # print(fromStr)
    # toStr = get_location("深圳湾科技生态园12栋B座裙楼8层")
    # print(toStr)
    #
    # res = get_way(fromStr,toStr,"0755")
    #获取具体的路费与交通时间
    # get_way("113.927941,22.492820","113.952696,22.530438","0755")
    # print(res)

    print ('读取成功')
