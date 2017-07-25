# encoding=utf-8

import os
import os.path
import hashlib
import ssl
import urllib
import urllib2
import json
import base64

import requests
from openpyxl import load_workbook


def md5str(str):  # md5加密字符串
    m = hashlib.md5(str.encode(encoding="utf-8"))
    return m.hexdigest()


def md5(byte):  # md5加密byte
    return hashlib.md5(byte).hexdigest()


class DamatuApi():
    ID = '40838'
    KEY = 'ca9507e17e8d5ddf7c57cd18d8d33010'
    HOST = 'http://api.dama2.com:7766/app/'

    def __init__(self, username, password):
        self.username = username
        self.password = password

    def getSign(self, param=b''):
        return (md5(bytes(self.KEY).encode('utf-8') + bytes(self.username).encode('utf-8') + param))[:8]

    def getPwd(self):
        return md5str(self.KEY + md5str(md5str(self.username) + md5str(self.password)))

    def post(self, path, params={}):
        data = urllib.urlencode(params).encode('utf-8')
        url = self.HOST + path
        response = urllib2.Request(url, data)
        return urllib2.urlopen(response).read()

    # 查询余额 return 是正数为余额 如果为负数 则为错误码
    def getBalance(self):
        data = {'appID': self.ID,
                'user': self.username,
                'pwd': self.getPwd(),
                'sign': self.getSign()
                }
        res = self.post('d2Balance', data)
        res = str(res).encode('utf-8')
        jres = json.loads(res)
        if jres['ret'] == 0:
            return jres['balance']
        else:
            return jres['ret']

    # 上传验证码 参数filePath 验证码图片路径 如d:/1.jpg type是类型，查看http://wiki.dama2.com/index.php?n=ApiDoc.Pricedesc  return 是答案为成功 如果为负数 则为错误码
    def decode(self, filePath, type):
        f = open(filePath, 'rb')
        fdata = f.read()
        filedata = base64.b64encode(fdata)
        f.close()
        data = {'appID': self.ID,
                'user': self.username,
                'pwd': self.getPwd(),
                'type': type,
                'fileDataBase64': filedata,
                'sign': self.getSign(filePath)
                }# self.getSign(base64.b64decode(filedata))
        res = self.post('d2File', data)
        res = str(res).encode('utf-8')
        jres = json.loads(res)
        print('jres', jres)
        if jres['ret'] == 0:
            # 注意这个json里面有ret，id，result，cookie，根据自己的需要获取
            return (jres['result'])
        else:
            return jres['ret']

    # url地址打码 参数 url地址  type是类型(类型查看http://wiki.dama2.com/index.php?n=ApiDoc.Pricedesc) return 是答案为成功 如果为负数 则为错误码
    def decodeUrl(self, url, type):
        data = {'appID': self.ID,
                'user': self.username,
                'pwd': self.getPwd(),
                'type': type,
                'url': urllib.quote(url),
                'sign': self.getSign(url.encode(encoding="utf-8"))
                }
        res = self.post('d2Url', data)
        res = str(res).encode('utf-8')
        jres = json.loads(res)
        if jres['ret'] == 0:
            # 注意这个json里面有ret，id，result，cookie，根据自己的需要获取
            return (jres['result'])
        else:
            return jres['ret']

    # 报错 参数id(string类型)由上传打码函数的结果获得 return 0为成功 其他见错误码
    def reportError(self, id):
        # f=open('0349.bmp','rb')
        # fdata=f.read()
        # print(md5(fdata))
        data = {'appID': self.ID,
                'user': self.username,
                'pwd': self.getPwd(),
                'id': id,
                'sign': self.getSign(id.encode(encoding="utf-8"))
                }
        res = self.post('d2ReportError', data)
        res = str(res).encode('utf-8')
        jres = json.loads(res)
        return jres['ret']


class Bill(object):
    def __init__(self, bills):
        self.bill_id = ''
        self.bill_num = ''
        self.bill_date = ''
        self.bill_money = ''

        if len(bills) == 4:
            self.bill_id = bills[0]
            self.bill_num = bills[1]
            self.bill_date = bills[2]
            self.bill_money = bills[3]

    def _get_two_dimension_picture(self):
        try:
            callback = "jQuery1102037866965332068503_1500810751838"
            fpdm = self.bill_id
            ssl._create_default_https_context = ssl._create_unverified_context
            resp = requests.get('https://fpcyweb.tax.sh.gov.cn:1001/WebQuery/yzmQuery?callback=%s'
                                '&fpdm=%s&r=0.8492015565279871'
                                % (callback, fpdm), verify=False)
        except Exception as e:
            print(u'获取验证码图片错误 - %s' % str(e))
        else:
            return json.loads(resp.content[len(callback) + 1:-1])
        return None

    def __str__(self):
        return "bill_id:%s bill_num:%s bill_date:%s bill_money:%s " % (
            self.bill_id, self.bill_num, self.bill_date, self.bill_money
        )


class Excel(object):
    def __init__(self):
        self.path = '/Users/wangyangyang/Desktop/test.bmp'

    def save(self, text):
        with open(self.path, 'w') as f:
            f.write(base64.b64decode(text))
            # f.write(text)

    def setup(self, path):
        if not os.path.isfile(path):
            raise Exception(u'文件%s不存在' % path)

        wb = load_workbook(path)
        bills = []

        for sheetname in wb.sheetnames:
            sheet = wb.get_sheet_by_name(sheetname)
            # if len(list(sheet["A"])) != 4:
            #     raise Exception(u'Excel表格不是4行，请检查')
            #
            # if sheet["A1"] != '发票代码' or sheet["A2"] != '发票号码' or \
            #         sheet["A3"] != '开票日期' or sheet["A4"] != '开具金额（不含税）':
            #     raise Exception(u'Excel格式错误，请检查')

            for i in range(1, sheet.max_column):
                bills.append(Bill([row.value for row in sheet[chr(65 + i)]]))

        self.bills = bills

    def check(self):
        for bill in self.bills:
            keys = bill._get_two_dimension_picture()
            # print(keys)
            # print(type(keys))

            dmt = DamatuApi("xiaolongchang", "858993460")
            print('*********')
            if keys:
                self.save(keys['key1'])
                print(dmt.decode(self.path, 60))  # 上传打码
                break




# 调用类型实例：
# 1.实例化类型 参数是打码兔用户账号和密码
if __name__ == '__main__':
    # dmt = DamatuApi("xiaolongchang", "858993460")
    # # 2.调用方法：
    # print(dmt.getBalance())  # 查询余额
    # print(dmt.decode('0349.bmp', 4))  # 上传打码
    # print(
    # dmt.decodeUrl('http://captcha.qq.com/getimage?aid=549000912&r=0.7257105156128585&uin=3056517021', 200))  # 上传打码
    # print(dmt.reportError('894657096')) #上报错误

    excel = Excel()
    # 文件路径
    excel.setup('./发票测试.xlsx')
    excel.check()
