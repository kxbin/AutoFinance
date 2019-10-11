import sys,re,time,datetime,shutil,requests,json
import xlwings as xw
import tushare as ts

class Bill: # 账单类
    def __init__(self, timestamp, amount, blance, description):
        self.timestamp = timestamp
        self.amount = float(amount)
        self.blance = float(blance)
        self.description = description
    
    def toString(self):
        print('交易时间：%s\t\t交易金额：%6.2f\t账户余额：%6.2f\t交易描述：%s' %(self.timestamp, self.amount, self.blance, self.description))

class Bills: # 账单列表类
    def __init__(self, bills):
        self.bills = bills
    
    def toString(self):
        for bill in self.bills:
            bill.toString()

    def sliceByTimestamp(self, startTimestamp, endTimestamp): # 根据时间戳切片
        bills = list()
        for bill in self.bills:
            if bill.timestamp >= startTimestamp and bill.timestamp < endTimestamp:
                bills.append(bill)
        return Bills(bills)
    
    def search(self, fields): # 根据条件搜索账单
        bills = list()
        for bill in self.bills:
            if len(fields) == 1:
                if re.search(fields[0], bill.description):
                    bills.append(bill)
            elif len(fields) == 2:
                if bill.description == fields[0] and abs(bill.amount) > float(fields[1]):
                    bills.append(bill)
            elif len(fields) == 3:
                if bill.description == fields[0] and abs(bill.amount) > float(fields[1]) and abs(bill.amount) < float(fields[2]):
                    bills.append(bill)
            else:
                print("元数据解析规则不能@分割后不能大于4个！")
                sys.exit(1)
        return Bills(bills)
    
    def sum(self): # 账单金额求和
        result = 0
        for bill in self.bills:
            result = result + bill.amount
        return result

class MetaData: # 元数据类
    def __init__(self, metaDataType, ID, name, value1, value2, value3):
        self.metaDataType = metaDataType
        self.ID = int(ID)
        self.name = name
        self.value1 = value1
        self.value2 = float(value2)
        self.value3 = value3
    
    def toString(self):
        print('元数据类型：%s\tID：%d\t名称：%s\tValue1：%s\tValue2：%f\tValue3：%s' %(self.metaDataType, self.ID, self.name, self.value1, self.value2, self.value3))

    def isMDID(self, MDID): # 判断元数据是否为某个MDID
        temp = re.findall(r'\D+|\d+',MDID)
        if self.metaDataType == temp[0] and self.ID == int(temp[1]):
            return True
        else:
            return False

    def write(self, excelName, sheetName, value=None, flag=None): # 元数据目标值写入excel
        if value is None:
            value = self.value2
        if self.value1 is not None:
            sheetRange = self.value1
            if self.value1[0] == '!':
                if flag == '!':
                    sheetRange = self.value1[1:].split(',')[0]
                else:
                    sheetRange = self.value1[1:].split(',')[1]
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(excelName)
            sht = wb.sheets[sheetName]
            sht.range(sheetRange).value = value
            wb.save()
            wb.close()
            app.kill()

class MetaDatas: # 元数据列表类
    def __init__(self, metaDatas):
        self.metaDatas = metaDatas

    def toString(self):
        for metaData in self.metaDatas:
            metaData.toString()

    def selectByMDID(self, MDID): # 根据MDID选择元数据
        for metaData in self.metaDatas:
            if metaData.isMDID(MDID):
                return metaData
        return None

    def sliceByType(self, metaDataType): # 根据元数据类型切片
        metaDatas = list()
        for metaData in self.metaDatas:
            if metaData.metaDataType == metaDataType:
                metaDatas.append(metaData)
        return MetaDatas(metaDatas)

    def write(self, excelName, sheetName):
        for metaData in self.metaDatas:
            if metaData.metaDataType == 'Data' and metaData.value1 is not None:
                metaData.write(excelName, sheetName)

    def parseFromBills(self, bills): # 根据账单解析元数据目标值
        for metaData in self.metaDatas:
            if metaData.metaDataType == 'Data':
                if metaData.value3 is not None:
                    value = 0
                    rules = metaData.value3.split('&AF&')
                    for rule in rules:
                        fields = rule.split('&')
                        if len(fields) >= 2:
                            linkMetaData = fields.pop(1)
                            bsum = bills.search(fields).sum()
                            if linkMetaData[0:1] == '+':
                                value = value + bsum + self.selectByMDID(linkMetaData[1:]).value2
                            elif linkMetaData[0:1] == '-':
                                value = value + bsum - self.selectByMDID(linkMetaData[1:]).value2
                            else:
                                print("连接元数据时只能用+或-！")
                                sys.exit(1)
                        else:
                            value = value + bills.search(fields).sum()
                    metaData.value2 = abs(value)
            else:
                continue

class RealPrice: # 实时资产价格获取类
    api = {
        'Cash': '',
        'Fund': 'https://api.doctorxiong.club/v1/fund/detail?code=',
        'Stock': 'https://api.doctorxiong.club/v1/stock/detail?code=',
        'Coin': 'https://www.okex.me/api/index/v3/',
        'FE': 'https://api.shenjian.io/exchange/currency/?appid=48c15a32562e29faab4dfdf6affa26a2&form='
    }

    def __init__(self, params):
        self.assetType = params[0]
        self.code = params[1]
        self.ercode = None
        if len(params) == 3:
            self.ercode = params[2]
        if self.assetType in self.api:
            if self.assetType == 'Coin':
                self.url = self.api.get(self.assetType) + self.code + '/constituents'
            elif self.assetType == 'FE':
                self.url = self.api.get(self.assetType) + self.code + '&to=CNY'
            else:
                self.url = self.api.get(self.assetType) + self.code
        else:
            print('没有%s资产类型,无法实例化RealPrice' %(self.assetType))
            sys.exit(1)

    def getPrice(self, date=None): # 获取资产价格
        price = 0
        result = self.getUrlResult()
        if date is not None:
            if self.assetType == 'Fund':
                r = json.loads(result)
                if r.get('code') == 200:
                    netWorthData = r.get('data').get('netWorthData')
                    for nwd in netWorthData:
                        if nwd[0] == date:
                            price = float(nwd[1])
                else:
                    print('%s资产类型接口json返回异常%s' %(self.assetType, result))
                    sys.exit(1)
            elif self.assetType == 'Cash':
                price = 1.0
            else:
                print('%s资产类型不支持按日期获取价格' %(self.assetType))
                sys.exit(1)
        else:
            if self.assetType == 'Cash':
                price = 1.0
            elif self.assetType == 'Fund':
                r = json.loads(result)
                if r.get('code') == 200:
                    price = float(r.get('data').get('netWorth'))
                else:
                    print('%s资产类型接口json返回异常%s' %(self.assetType, result))
                    sys.exit(1)
            elif self.assetType == 'Stock':
                r = json.loads(result)
                if r.get('code') == 200:
                    price = float(r.get('data').get('price'))
                else:
                    print('%s资产类型接口json返回异常%s' %(self.assetType, result))
                    sys.exit(1)
            elif self.assetType == 'Coin':
                r = json.loads(result)
                if r.get('code') == 0:
                    price = float(r.get('data').get('last'))
                else:
                    print('%s资产类型接口json返回异常%s' %(self.assetType, result))
                    sys.exit(1)
            elif self.assetType == 'FE':
                r = json.loads(result)
                if r.get('error_code') == 0:
                    price = float(r.get('data').get('rate'))
                else:
                    print('%s资产类型接口json返回异常%s' %(self.assetType, result))
                    sys.exit(1)
        if self.ercode is not None:
            return price * RealPrice(['FE', self.ercode]).getPrice()
        else:
            return price

    def getUrlResult(self):
        if self.assetType == 'Cash':
            return None
        r = requests.get(self.url)
        if r.status_code == 200:
            return r.text
        else:
            print('调用Api返回状态码不为200,请检查接口%s！' %(self.url))
            sys.exit(1)

class AutoFinance: # 自动财务类-主类
    def __init__(self, historyBillFile, metaDataFile, target):
        self.historyBillFile = historyBillFile
        self.metaDataFile = metaDataFile
        self.target = target
        self.bills = None
        self.latelyBills = None
        self.metaDatas = None
        self.startTimestamp = None
        self.endTimestamp = None
        self.initiationDate = None

    def do(self):
        self.bills = self.loadBill(self.historyBillFile.split('|')[0], self.historyBillFile.split('|')[1])
        self.metaDatas = self.loadMetaData(self.metaDataFile.split('|')[0], self.metaDataFile.split('|')[1])
        self.doMeta1()
        self.doMeta2()
        self.doDatas()
        self.doMeta3()
        self.doMeta4()
        self.doAssets()

    def doMeta1(self): # 更新账期并筛选账单
        meta1 = self.metaDatas.selectByMDID('Meta1')
        today = datetime.date.today()
        if today.day >= meta1.value2:
            startTime = str(today.year) + '/' + str(today.month-1) + '/' + str(int(meta1.value2)) + '  00:00:00'
            endTime = str(today.year) + '/' + str(today.month) + '/' + str(int(meta1.value2)) + '  00:00:00'
        else:
            startTime = str(today.year) + '/' + str(today.month-2) + '/' + str(int(meta1.value2)) + '  00:00:00'
            endTime = str(today.year) + '/' + str(today.month-1) + '/' + str(int(meta1.value2)) + '  00:00:00'
        self.startTimestamp = int(time.mktime(time.strptime(startTime, '%Y/%m/%d  %H:%M:%S')))
        self.endTimestamp = int(time.mktime(time.strptime(endTime, '%Y/%m/%d  %H:%M:%S')))
        self.latelyBills = self.bills.sliceByTimestamp(self.startTimestamp, self.endTimestamp)
        self.latelyBills.toString()
        newFileName = '个人现金流报表(' + time.strftime('%Y.%m.%d', time.localtime(self.startTimestamp)) + '-' + time.strftime('%Y.%m.%d', time.localtime(self.endTimestamp-1)) + ').xlsx'
        shutil.copy2(self.target.split('|')[0], newFileName)
        self.target = newFileName + '|' + self.target.split('|')[1]
        meta1.write(self.target.split('|')[0], self.target.split('|')[1], time.strftime('%Y/%m/%d', time.localtime(self.startTimestamp)) + '-' + time.strftime('%Y/%m/%d', time.localtime(self.endTimestamp-1)))

    def doMeta2(self): # 计算定投日和定投日最近的交易日
        meta2 = self.metaDatas.selectByMDID('Meta2')
        meta2.value2 = time.strftime('%Y%m', time.localtime(self.startTimestamp)) + str(int(meta2.value2))
        pro = ts.pro_api('e14bc7d11830addb61f9284ccc2cc8e739c4bdc08a83e7fe34465160')
        df = pro.trade_cal(start_date=time.strftime('%Y%m%d', time.localtime(self.startTimestamp)), end_date=time.strftime('%Y%m%d', time.localtime(self.endTimestamp)))
        for index, row in df.iterrows():
            if int(row['cal_date']) < int(meta2.value2) or row['is_open'] != 1:
                continue
            else:
                self.initiationDate = df['cal_date'][index-1]
                break

    def doDatas(self): # 根据解析规则计算Data类型元数据的目标值
        self.metaDatas.parseFromBills(self.latelyBills)
        self.metaDatas.write(self.target.split('|')[0], self.target.split('|')[1])

    def doMeta3(self): # 计算日常消费
        meta3 = self.metaDatas.selectByMDID('Meta3')
        meta3.value2 = meta3.value2 + self.latelyBills.bills[-1].blance
        meta3.write(self.target.split('|')[0], self.target.split('|')[1])  

    def doMeta4(self): # 计算期末余额
        meta4 = self.metaDatas.selectByMDID('Meta4')
        meta4.value2 = abs(self.latelyBills.sum() - self.metaDatas.selectByMDID('Data1').value2 - self.metaDatas.selectByMDID('Data2').value2 + self.metaDatas.selectByMDID('Data3').value2 + self.metaDatas.selectByMDID('Data4').value2 + self.metaDatas.selectByMDID('Data5').value2 + self.metaDatas.selectByMDID('Data6').value2 + self.metaDatas.selectByMDID('Data7').value2 + self.metaDatas.selectByMDID('Data8').value2)
        meta4.write(self.target.split('|')[0], self.target.split('|')[1])

    def doAssets(self): # 计算资产
        assets = self.metaDatas.sliceByType('Asset')
        for asset in assets.metaDatas:
            if asset.value3[0] == '!':
                temp = asset.value3[1:].split(':')
                money = float(temp.pop(0))
                rp = RealPrice(temp)
                print('元数据%s%d%sValue2原来的值为%f,现在将要反写该元数据,请注意保存该值！！！' %(asset.metaDataType, asset.ID, asset.name, asset.value2))
                asset.value2 = round(asset.value2 + money / rp.getPrice(self.initiationDate[2:]), 2)
                asset.write(self.metaDataFile.split('|')[0], self.metaDataFile.split('|')[1], flag='!')
                asset.value2 = round(asset.value2 * rp.getPrice(), 2)
                asset.write(self.target.split('|')[0], self.target.split('|')[1])
            else:
                rp = RealPrice(asset.value3.split(':'))
                asset.value2 = round(asset.value2 * rp.getPrice(), 2)
                asset.write(self.target.split('|')[0], self.target.split('|')[1])

    def loadBill(self, excelName, sheetName): # 加载账单
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(excelName)
        sht = wb.sheets[sheetName]
        rowCount = sht.used_range.last_cell.row
        colCount = sht.used_range.last_cell.column
        bills = list()
        if colCount != 6:
            print("账单源数据结构错误，必须是6列！")
            sys.exit(1)
        else:
            for i in range(3,rowCount + 1):
                timestamp = int(time.mktime(time.strptime(sht.range((i,2)).value, '%Y/%m/%d  %H:%M:%S')))
                if sht.range((i,4)).value != '' and sht.range((i,5)).value == '':
                    amount = sht.range((i,4)).value
                elif sht.range((i,4)).value == '' and sht.range((i,5)).value != '':
                    amount = -float(sht.range((i,5)).value)
                else:
                    print("账单源数据结构错误，交易金额为空！")
                    sys.exit(1)
                blance = sht.range((i,6)).value
                description = sht.range((i,3)).value
                bill = Bill(timestamp, amount, blance, description)
                print("Loading Bill ......")
                bills.append(bill)
        wb.close()
        app.kill()
        print("Load Bills Complete!")
        return Bills(bills)

    def loadMetaData(self, excelName, sheetName): # 加载元数据
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(excelName)
        sht = wb.sheets[sheetName]
        rowCount = sht.used_range.last_cell.row
        colCount = sht.used_range.last_cell.column
        metaDatas = list()
        if colCount != 6:
            print("元数据数据结构错误，必须是6列！")
            sys.exit(1)
        else:
            for i in range(1,rowCount + 1):
                if sht.range((i,1)).value != 'Meta' and sht.range((i,1)).value != 'Data' and sht.range((i,1)).value != 'Asset':
                    continue
                else:
                    metaDataType = sht.range((i,1)).value
                    ID = sht.range((i,2)).value
                    name = sht.range((i,3)).value
                    value1 = sht.range((i,4)).value
                    value2 = sht.range((i,5)).value
                    value3 = sht.range((i,6)).value
                    metaData = MetaData(metaDataType, ID, name, value1, value2, value3)
                    print("Loading MetaData ......")
                    metaDatas.append(metaData)
        wb.close()
        app.kill()
        print("Load MetaDatas Complete!")
        return MetaDatas(metaDatas)

# 实例化自动财务并生成报表
ins = AutoFinance('HistoryBill.xls|sheet', 'MetaData.xlsx|Sheet', 'Template.xlsx|Sheet')
ins.do()