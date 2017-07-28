# encoding=utf-8

import os
import os.path
import hashlib
import random
import string
import urllib
import urllib2
import json
import base64
from binascii import b2a_hex

import execjs
import requests
from openpyxl import load_workbook
from Crypto.Cipher import AES
from Crypto import Random
from requests.packages.urllib3.exceptions import InsecureRequestWarning

# 禁用安全请求警告
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

# DaMaTu相关配置
DamatuUserName = 'xiaolongchang'
DamatuUserPassWD = '858993460'

# Excel表格所在位置
Excel_filepath = './发票测试.xlsx'

# 验证码保存路径
PictureFileSavePath = '/Users/wangyangyang/Desktop/test.bmp'


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
    def decode(self, filedata, type):
        # f = open(filePath, 'rb')
        # fdata = f.read()
        # filedata = base64.b64encode(fdata)
        # f.close()
        data = {'appID': self.ID,
                'user': self.username,
                'pwd': self.getPwd(),
                'type': type,
                'fileDataBase64': filedata,
                'sign': self.getSign(base64.b64decode(filedata))
                }  # self.getSign(base64.b64decode(filedata))
        res = self.post('d2File', data)
        # res = str(res).encode('utf-8')
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
    def __init__(self, bills, row=None):
        self.bill_id = ''
        self.bill_num = ''
        self.bill_date = ''
        self.bill_money = ''
        self.verification_code = None
        self.row = row

        if len(bills) == 4:
            self.bill_id = bills[0]
            self.bill_num = bills[1]
            self.bill_date = bills[2]
            self.bill_money = bills[3]

    def save(self, text):
        with open(PictureFileSavePath, 'w') as f:
            f.write(base64.b64decode(text))
            print(u'存储验证码')

    def _get_verification_picture(self):
        try:
            print('查询验证码')
            callback = "jQuery%s_%s" % (
                ''.join([random.choice(string.digits) for _ in range(22)]),
                ''.join([random.choice(string.digits) for _ in range(13)]))
            r = '0.%s' % ''.join([random.choice(string.digits) for _ in range(16)])
            fpdm = self.bill_id
            resp = requests.get('https://fpcyweb.tax.sh.gov.cn:1001/WebQuery/yzmQuery?callback=%s'
                                '&fpdm=%s&r=%s' % (callback, fpdm, r), verify=False)
        except Exception as e:
            print(u'获取验证码图片错误 - %s' % str(e))
        else:
            return json.loads(resp.content[len(callback) + 1:-1]), callback
        return None, ''

    def check_task(self):
        keys, callback = self._get_verification_picture()
        if keys:
            print(keys, callback)
            print('key4:%s' % keys.get('key4'))

        if keys and keys.get('key4') == '00':
            # keys['key4'] == '00' 为最正常验证码，其他奇葩的他们解析不了
            self.save(keys['key1'])

            verification_code = dmt.decode(keys['key1'], 200)
            print("verification_code", verification_code)
            try:
                verification_code = int(verification_code)
                return False
            except ValueError:
                if verification_code == 'ERROR':
                    return False
                self.verification_code = verification_code
                self._post_request(callback, keys.get('key2'), keys.get('key3'), keys.get('key4'))
                return True

        return False

    def _post_request(self, callback, key2, key3, key4):
        try:
            url = 'https://fpcyweb.tax.sh.gov.cn:1001/WebQuery/query'
            data = {
                'callback': callback,
                'fpdm': str(self.bill_id),
                'fphm': str(self.bill_num),
                'kprq': str(self.bill_date),
                'fpje': str(self.bill_money),
                'fplx': key4,
                'yzm': self.verification_code,
                'yzmSj': key2, #str(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
                'index': key3,
                'iv': self.get_crypto(), #b2a_hex(Random.get_random_bytes(128 / 8)),
                'salt': self.get_crypto(), #b2a_hex(Random.get_random_bytes(128 / 8)),
                '_': ''.join([random.choice(string.digits) for _ in range(13)])
            }
            print(url)
            print(json.dumps(data))
            resp = requests.post(url, data=json.dumps(data), verify=False)
            status_code = resp.status_code
            print('*************')
            print(resp.content, resp.text)
            print(status_code)
        except Exception as e:
            print(u'请求验证失败，正在重试请稍等！- %s' % str(e))
            return False
        else:
            pass
        return False

    def get_crypto(self):
        htmlstr = ''
        with open("./crypto-js-develop/src/core.js", 'r') as f:
            line = f.readline()
            while line:
                htmlstr = htmlstr + line
                line = f.readline()

        ctx = execjs.compile(htmlstr)
        return ctx.call('test', 16)

    def __str__(self):
        return "bill_id:%s bill_num:%s bill_date:%s bill_money:%s " % (
            self.bill_id, self.bill_num, self.bill_date, self.bill_money
        )


class Excel(object):
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

            # 按列读取
            for i in range(1, sheet.max_column):
                bills.append(Bill([row.value for row in sheet[chr(65 + i)]]))

            # 按行读取
            # for i in range(1, sheet.max_row):
            #     a = [col.value for col in sheet[str(i)]]
            #     if len(a) == 4:
            #         bills.append(Bill(a, i))

        self.bills = bills

    def check(self):
        if not hasattr(self, 'bills'):
            print(u'请先执行setup...')
            return

        for bill in self.bills:
            is_ok = False

            while not is_ok:
                is_ok = bill.check_task()
                break
            break


dmt = DamatuApi(DamatuUserName, DamatuUserPassWD)

if __name__ == '__main__':
    excel = Excel()
    excel.setup(Excel_filepath)
    excel.check()
