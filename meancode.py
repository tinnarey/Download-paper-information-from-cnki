import requests  # 导入requests 模块
import re
import xlwt
import time
import xlrd
from xlrd import open_workbook
from xlutils.copy import copy


class BeautifulPicture():
    def get_pic(self):
        data = xlrd.open_workbook(tablename+'.xls','r')  # 打开xls文件
        table = data.sheets()[0]  # 打开第一张表
        i = table.nrows #上一次爬到的表1的行数
        i1 = 0
        used = 0
        rb = open_workbook(tablename+'.xls','utf-8')
        wb = copy(rb) # 将上一次爬到的复制到新表里，并在新表里继续添加纪录
        # 通过get_sheet()获取的sheet有write()方法
        ws = wb.get_sheet(0)
        p = 1
        for num in range(p, p + page):
            # 这里的num是页码 ,url需要在所需爬取的页面二次加载复制
            web_url='http://kns.cnki.net/kns/brief/brief.aspx?curpage=%s&RecordsPerPa' \
                    'ge=20&QueryID=3&ID=&turnpage=1&tpagemode=L&dbPrefix=SCDB' \
                    '&Fields=&DisplayMode=listmode&PageName=ASP.brief_default_result_aspx&isinEn=1#J_ORDER&'% num

            # 这里开始是时间控制
            t = int(time.clock())
            useTime = t - used
            # 如果一个周期的时间使用太短，则等待一段时间
            # 主要用于防止被禁
            if (useTime < 120 and useTime > 10):
                print("useTime=%s" % useTime)
                whiteTime = 120 - useTime
                print("等待%s秒" % whiteTime)
                time.sleep(whiteTime)
            used = int(time.clock())
            print(useTime)
            print('开始网页get请求')
            r = self.request(web_url)
            # 这里是报错的解释，能知道到底是因为什么不能继续爬了
            yan = re.search(r'参数错误', r.text)
            if yan != None:
                print("参数")
                break
            yan = re.search(r'验证码', r.text)
            if yan != None:
                print("验证")
                break

            #这里开始使用正则抓列表里每一个文献的url
            soup = re.findall(r'<TR([.$\s\S]*?)</TR>', r.text)
            for a in soup:
                i1 += 1
                name = re.search(r'_blank.*<', a)
                name = name.group()[8:-1]
                name = re.sub(r'<font class=Mark>', '', name)
                name = re.sub(r'</font>', '', name)

                url = re.search(r'href=.*? ', a)#将’‘看做一个子表达式，惰性匹配一次就可以了
                url = url.group()

                # 将爬来的相对地址，补充为绝对地址
                url = "http://kns.cnki.net/KCMS/" + url[11:-2] #数字是自己数的。。。

                #下面是参考文献详情的URL
                FN = re.search(r'FileName.*?&', url)
                if FN !=None:
                    FN = re.search(r'FileName.*?&', url).group()
                DN = re.search(r'DbName.*?&', url)
                if DN !=None:
                    DN=re.search(r'DbName.*?&', url).group()
                DC = re.search(r'DbCode.*?&', url).group()
                DUrl = "http://kns.cnki.net/KCMS/detail/frame/list.aspx?%s%s%sRefType=1" % (FN, DN, DC)
                R = self.request(DUrl)
                #如果没有参考文献，则认为是劣质文献，不爬，转爬下一篇
                isR = re.search(r'参考文献', R.text)
                if i1 == 1:
                    print("name:%s" % name)
                if isR == None:
                    continue
                d = self.request(url).text
                type = re.search(r'"\).html\(".*?"', d)
                type = type.group()[9:-1]
                ins = re.search(r'TurnPageToKnet\(\'in\',\'.*?\'', d)
                if ins == None:
                    continue
                ins = ins.group()[21:-1]
                wt = re.findall(r'TurnPageToKnet\(\'au\',\'.*?\'', d)
                writer = ""
                for w in wt:
                    writer = writer + "," + w[21:-1]
                writer = writer[1:]

                #文献摘要
                summary = re.search(r'(?<=name="ChDivSummary">).+?(?=</span>)', d)
                summary = summary.group()

                ws.write(i, 0, name)    #文献名
                ws.write(i, 1, writer)  #作者名
                ws.write(i, 2, type)    #文献类别
                ws.write(i, 16, summary)  #摘要

                # 期刊 以及是否为核心期刊
                sourinfo = re.search(r'sourinfo([.$\s\S]*?)</div', d)
                if sourinfo != None:
                    sourinfo = sourinfo.group()
                    # print(sourinfo)
                    from_ = re.search(r'title.*</a', sourinfo).group()
                    from_ = re.sub(r'title">.*?>', '', from_)
                    from_ = re.sub(r'</a', '', from_)
                    ws.write(i, 3, from_)
                    core = re.search(r'中文核心期刊', sourinfo)
                    if core != None:
                        # print(core.group())
                        ws.write(i, 4, "中文核心期刊")

                 # 这里是文献的来源基金
                fund = re.search(r'TurnPageToKnet\(\'fu\',\'.*?\'', d)
                if fund != None:
                    fund = fund.group()[21:-1]
                    ws.write(i, 5, fund)

                # 这里是文献的关键词，最多可以记录8个关键词
                kw = re.findall(r'TurnPageToKnet\(\'kw\',\'.*?\'', d)
                tnum = 0
                for tkw in kw:
                    tnum += 1
                    tkw = tkw[21:-1]
                    if tnum > 8:
                        break
                    ws.write(i, 5 + tnum, tkw)

                i += 1 # 增加页码的计数
        wb.save(tablename+'.xls') #

    def request(self, url):  # 返回网页的response


        # 这里是伪造浏览器信息，和伪造来源
        user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Ch' \
                     'rome/58.0.3029.81 Safari/537.36'
        referer = "http://kns.cnki.net/kns/brief/result.aspx"

        # 这里是伪造cookie，你要从浏览器里复制出来，粘贴在这里
        # 可以把里面的时间不断更新，这样能爬久一点 &ot后的时间为未来的时间

        cookie = 'Ecp_notFirstLogin=jUvOHH; Ecp_ClientId=4180829215704786267; UM_distinctid=16586024542a' \
                 'b-08263edfa73478-5b193413-15f900-165860245444b2; cnkiUserKey=c8a62aaf-e45e-4265-8546-5' \
                 'b0e4417f76f; Ecp_session=1; ASP.NET_SessionId=dal4gjh2d0g2ymyvlgghhcru; SID_klogin=125143; ' \
                 'SID_kns=123118; SID_crrs=125134; SID_krsnew=125133; KNS_SortType=; RsPerPage=20; Ecp_Login' \
                 'Stuts=%7B%22IsAutoLogin%22%3Afalse%2C%22UserName%22%3A%22KT1008%22%2C%22ShowName%22%3A%22%' \
                 '25E8%25A5%25BF%25E5%258D%2597%25E7%259F%25B3%25E6%25B2%25B9%25E5%25A4%25A7%25E5%25AD%25A6%' \
                 '22%2C%22UserType%22%3A%22bk%22%2C%22r%22%3A%22jUvOHH%22%7D; _pk_ref=%5B%22%22%2C%22%22%2C1' \
                 '546830331%2C%22http%3A%2F%2Fwww.cnki.net%2F%22%5D; _pk_ses=*; LID=WEEvREcwSlJHSldRa1FhcEE0' \
                 'QVRCZ1g2NkluUHltb2J3ek9aYWpKS0gzQT0=$9A4hF_YAuvQ5obgVAqNKPCYcEjKensW4IQMovwHtwkF4VYPoHbKxJ' \
                 'w!!; c_m_LinID=LinID=WEEvREcwSlJHSldRa1FhcEE0QVRCZ1g2NkluUHltb2J3ek9aYWpKS0gzQT0=$9A4hF_YA' \
                 'uvQ5obgVAqNKPCYcEjKensW4IQMovwHtwkF4VYPoHbKxJw!!&ot=01/08/2019 12:27:50; c_m_expire=2019-01-07 14:10:50'
        headers = {'User-Agent': user_agent,
                   "Referer": referer,
                   "cookie": cookie}
        r = requests.get(url, headers=headers, timeout=80)
        return r

tablename = 'nlp'
work_book=xlwt.Workbook(encoding='utf-8')
sheet=work_book.add_sheet('文献信息')
sheet.write(0,0,'文章名字')
sheet.write(0,1,'作者')
sheet.write(0,2,'文献类别')
sheet.write(0,3,'期刊')
sheet.write(0,4,'是否中文核心')
sheet.write(0,5,'基金来源')
sheet.write(0,6,'关键字')
sheet.write(0,16,'摘要')
work_book.save(tablename+'.xls')

print('开始时间',time.clock())
page = 10 #一共需要爬取的页数
beauty = BeautifulPicture()  # 创建类的实例
beauty.get_pic()  # 执行类中的方法
print('结束时间',time.clock())
