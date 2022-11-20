from datetime import date, timedelta
import datetime
import ddddocr
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import cv2
import easyocr
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from app01.migrations.spider.MysqlOperation import *
from app01.migrations.spider.ExcelOperation import *
import random

Mysql = sql()
Excel = write_excel("8")
reader = easyocr.Reader(['ch_sim', 'en'])
accounts = Mysql.query_all_accounts()
sleep_base = 0
ocr = ddddocr.DdddOcr(old=True)
det = ddddocr.DdddOcr(det=True)
current_directory = os.path.dirname(os.path.abspath(__file__))


# os.environ["CUDA_VISIBLE_DEVICES"] = "True"
# os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'

def GetTime(Time):
    global sleep_base
    if Time > 3:
        a = random.randint(-1, 5)
        Time += a
    Time += sleep_base
    time.sleep(Time)
    print('sleep:' + str(Time))


def sleepToTomorrow():
    now = time.localtime()
    now_time = time.strftime("%Y-%m-%d %H:%M:%S", now)
    # print(now_time)
    tomorrow = (date.today() + timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")
    # print(tomorrow)
    startTime2 = datetime.datetime.strptime(now_time, "%Y-%m-%d %H:%M:%S")
    endTime2 = datetime.datetime.strptime(tomorrow, "%Y-%m-%d %H:%M:%S")
    endTime2 = endTime2 + timedelta(hours=4)
    total_seconds = (endTime2 - startTime2).total_seconds()
    # print(total_seconds)
    return total_seconds


def JudgeMid(metrix):
    num = 0
    for i in metrix:
        for j in i:
            if j < 200:
                num = num + 1
    if num < 30:
        return "-"
    else:
        return "+"


def loginCodeVerify(path, text):
    poses = coordinate(path)
    im = cv2.imread(path)
    map = {}
    for box in poses:
        x1, y1, x2, y2 = box
        t = im[y1:y2, x1:x2]
        cv2.imwrite('word.jpg', t)
        word = distinguishWord('word.jpg')
        x = (x1 + x2) / 2
        y = (y1 + y2) / 2
        map[word] = [x, y]
    print(map)
    location = []
    for i in range(3):
        w = map.get(text[i])
        if w is None:
            raise Exception('识别错误')
        location.append(w)
    return location


def coordinate(path):
    with open(path, 'rb') as f:
        image = f.read()
    poses = det.detection(image)
    return poses


def distinguishWord(path):
    with open(path, 'rb') as f:
        image = f.read()

    res = ocr.classification(image)
    f.close()
    return res


def RecognizePicture(path):
    str = distinguishWord(path)
    if len(str) < 3:
        raise Exception('识别错误')
    one = int(str[0])
    two = int(str[2])
    if str[1] == '+':
        return one + two
    else:
        return one - two


class GetInvoice:
    def __init__(self):
        self.pic_id = None
        self.account = None
        browser = webdriver.Firefox(executable_path=current_directory + "\geckodriver.exe")
        GetTime(3)
        self.browser = browser

    def VerificationCode(self, ele_selectyzm, selectyzm_text):
        jcaptchaimage = self.browser.find_element_by_id("jcaptchaimage")
        move_to_jcaptchaimage = ActionChains(self.browser).move_to_element(jcaptchaimage)  # 移动到该元素
        move_to_jcaptchaimage.perform()  # 执行
        time.sleep(2)
        jcaptchaimage.screenshot('jcaptchaimage.png')
        GetTime(1)
        text = selectyzm_text.replace(' ', '').replace('"', '').replace('请依次点击', '')
        location = loginCodeVerify('jcaptchaimage.png', text)
        GetTime(2)
        ActionChains(self.browser).move_to_element_with_offset(ele_selectyzm, 5, 5).perform()
        GetTime(1)
        print('location:')
        print(location)

        for i in range(3):
            ActionChains(self.browser).move_to_element_with_offset(jcaptchaimage, location[i][0] * 0.8,
                                                                   location[i][1] * 0.8).click().perform()
            time.sleep(1)
        GetTime(5)

    def GetAccount(self):
        Mysql.recon()
        global accounts
        accounts = Mysql.query_all_accounts()
        for account in accounts:
            if account['flag'] == '0':
                return account
        # 所有账号已爬完
        print('所有账号已爬完')
        t = sleepToTomorrow()
        print('sleep:' + str(t))
        time.sleep(t)
        Mysql.recon()
        time.sleep(2)
        accounts = Mysql.query_all_accounts()
        return accounts[0]

    def Login(self):
        global sleep_base
        self.account = self.GetAccount()
        res = self.LoginPerfom()
        if res == True:
            return
        else:
            # 登录失败
            print('停止5分钟')
            time.sleep(60 * 5)
            sleep_base = 0
            Mysql.recon()
            self.Login()

            # 发出失败通知

    def LoginPerfom(self):
        global sleep_base
        for TryTime in range(6):
            try:
                print('进入网站')
                self.browser.get("http://cpquery.cnipa.gov.cn/")
                self.browser.maximize_window()
                GetTime(6)
                self.browser.find_element_by_id("username1").send_keys(self.account['username'])
                self.browser.find_element_by_id("password1").send_keys(self.account['password'])

                for i in range(10):
                    try:
                        selectyzm_text = self.browser.find_element_by_id("selectyzm_text")
                        move_to_selectyzm_text = ActionChains(self.browser).move_to_element(selectyzm_text)
                        move_to_selectyzm_text.perform()
                        GetTime(1)
                        # 识别验证码
                        self.VerificationCode(selectyzm_text, selectyzm_text.text)
                        # 登录
                        if (selectyzm_text.text == '验证成功'):
                            login = self.browser.find_element_by_id("publiclogin")
                            move_to_login = ActionChains(self.browser).move_to_element(login).click(login)
                            move_to_login.perform()
                            break
                        else:
                            raise Exception('验证失败')
                    except:
                        # sleep_base += 1
                        print('重试')
                        js = 'document.getElementById("reload").click()'
                        self.browser.execute_script(js)
                        GetTime(4)

                WebDriverWait(self.browser, 30, 0.5).until(
                    EC.visibility_of_element_located((By.CLASS_NAME, 'intruduction_box')))
                time.sleep(2)
                # 声明确认
                self.browser.execute_script("window.scrollBy(0,1000)")
                time.sleep(2)
                agree = self.browser.find_element_by_id("agreeid")
                goBtn = self.browser.find_element_by_id("goBtn")
                ActionChains(self.browser).move_to_element(agree).click(agree).pause(2).move_to_element(goBtn).click(
                    goBtn).perform()

                WebDriverWait(self.browser, 30, 0.5).until(
                    EC.visibility_of_element_located((By.ID, 'select-key:shenqingh')))
                time.sleep(2)
                sleep_base = 0
                return True
            except Exception:
                if sleep_base < 3:
                    sleep_base += 1
                print('登录失败')
                self.browser.get("https://www.baidu.com")
                time.sleep(5)
                traceback.print_exc()
        return False

    def SendID(self, ID):
        WebDriverWait(self.browser, 30, 0.5).until(
            EC.visibility_of_element_located((By.ID, 'select-key:shenqingh')))
        time.sleep(2)

        InputForm = self.browser.find_element_by_id("select-key:shenqingh")
        InputForm.send_keys(ID)
        self.Reflesh = self.browser.current_url

    def GetVeryPic(self):
        pic = self.browser.find_element_by_id("authImg")
        pic.screenshot('freeze.png')
        time.sleep(2)
        print('getverpic')

    def Recognize(self):
        for i in range(10):
            try:
                abspath = os.path.abspath('.') + r'\freeze.png'
                VeryCode = RecognizePicture(abspath)
                print('VeryCode' + str(VeryCode))
                break
            except:
                print('验证码识别错误')
                self.browser.find_element_by_id("authImg").click()
                traceback.print_exc()
                time.sleep(1)
                self.GetVeryPic()
        time.sleep(2)
        self.browser.find_element_by_id("very-code").send_keys(VeryCode)
        time.sleep(1)
        print('recognize')
        # 避免右键菜单未取消
        zhuanlimc = self.browser.find_element_by_id("tips")
        ActionChains(self.browser).move_to_element(zhuanlimc).click().perform()
        time.sleep(1)

    def Query(self):
        self.browser.find_element_by_id("query").click()
        # GetTime(6)
        print('query')

    def NeedPayItem(self):
        # 费用信息
        self.browser.find_element_by_link_text("费用信息").click()
        WebDriverWait(self.browser, 30, 0.5).until(
            EC.visibility_of_element_located((By.ID, 'djftitle')))
        time.sleep(2)
        # GetTime(5)

        # 选已缴费信息
        NeedPayItem = self.browser.find_elements_by_class_name("imfor_part1")[0]
        Title = NeedPayItem.find_element_by_class_name("imfor_table_grid").find_element_by_tag_name(
            "tr").find_elements_by_tag_name("th")
        Title = [i.text for i in Title]
        Content = NeedPayItem.find_element_by_class_name("imfor_table_grid").find_elements_by_tag_name("tr")[1:]
        Results = []
        for line in Content:
            # ColsText = [cols.text for cols in line.find_elements_by_tag_name("td")]
            # ColsText = [cols.find_elements_by_tag_name("span").get_attribute("title") for cols in line]
            ColsText = str(line.text).split(' ')
            Result = dict(zip(Title, ColsText))
            Results.append(Result)
        if len(Results) == 0:
            return None
        else:
            print('results:')
            print(Results)
            return Results

    def AlreadyPayItem(self):
        AlreadyPayItem = self.browser.find_elements_by_class_name("imfor_part1")[1]
        Title = AlreadyPayItem.find_element_by_class_name("imfor_table_grid").find_element_by_tag_name(
            "tr").find_elements_by_tag_name("th")
        Title = [i.text for i in Title]
        Content = AlreadyPayItem.find_element_by_class_name("imfor_table_grid").find_elements_by_tag_name("tr")[1:]
        Results = []
        for line in Content:
            # ColsText = [cols.text for cols in line.find_elements_by_tag_name("td")]
            ColsText = str(line.text).split(' ')
            Result = dict(zip(Title, ColsText))
            Results.append(Result)
        if len(Results) == 0:
            return None
        else:
            print('results:')
            print(Results)
            return Results

    def Base2Dic(self, BaseInfo):
        BaseInfos = BaseInfo.find_elements_by_tag_name("tr")
        BaseInfosDic = {}
        for item in BaseInfos:
            key = item.find_elements_by_tag_name("td")[0].text
            key = str(key).replace(':', '').replace('：', '').replace(' ', '')
            values = item.find_elements_by_tag_name("td")[1].text
            BaseInfosDic[key] = values
        if len(list(filter(None, BaseInfosDic.values()))) == 0:
            return {}
        return BaseInfosDic

    def Table2Dic(self, Table):
        TableDic = []
        Title = [i.text for i in Table.find_elements_by_tag_name("th")]
        if (len(Table.find_elements_by_tag_name("tr")) > 1):
            Content = Table.find_elements_by_tag_name("tr")[1:]
            for item in Content:
                tds = item.find_elements_by_tag_name("td")
                values = [i.text for i in tds]
                if len(list(filter(None, values))) == 0:
                    continue

                TableDic.append(dict(zip(Title, values)))
        return TableDic

    def InsertSQL(self, items, TableName, id):
        if type(items) == type({}):
            if len(items) == 0: return
            try:
                items["申请号"] = id
                Mysql.update(TableName, {"Flag": "1"}, {"申请号": id})
            except:
                pass
            items["Flag"] = 0
            Mysql.insert_add(TableName, items)
        elif type(items) == type([]):
            if len(items) == 0: return
            try:
                Mysql.update(TableName, {"Flag": "1"}, {"申请号": id})
            except:
                pass

            for item in items:
                item["申请号"] = id
                item["Flag"] = 0
                Mysql.insert_add(TableName, item)
        else:
            print("格式不对")

    def Date(self, ID):
        self.browser.find_element_by_id("gbgg").click()
        # GetTime(2)
        WebDriverWait(self.browser, 30, 0.5).until(
            EC.visibility_of_element_located((By.ID, 'gkggtitle')))
        time.sleep(2)

        DateInfo = self.browser.find_elements_by_class_name("imfor_part1")[0]
        DateInfos = self.Table2Dic(DateInfo)
        for item in DateInfos:
            try:
                del item[""]
            except:
                pass
        self.InsertSQL(DateInfos, "发布公告授权公告", ID)
        GetTime(1)

    def TansBaseInfo(self, ID):
        self.browser.find_element_by_class_name("tab_first").click()
        # GetTime(2)
        WebDriverWait(self.browser, 30, 0.5).until(
            EC.visibility_of_element_located((By.ID, 'zlxtitle')))
        time.sleep(2)

        BaseInfo = self.browser.find_elements_by_class_name("imfor_part1")[0]
        Person = self.browser.find_elements_by_class_name("imfor_part1")[1]
        DesignPersion = self.browser.find_elements_by_class_name("imfor_part1")[2]

        BaseInfos = self.Base2Dic(BaseInfo)
        print('BaseInfos:')
        print(BaseInfos)
        Persons = self.Table2Dic(Person)
        print('Persons:')
        print(Persons)
        DesignPersions = self.Base2Dic(DesignPersion)
        print('Design:')
        print(DesignPersions)

        GetTime(2)
        self.InsertSQL(BaseInfos, "著录项目信息", ID)
        self.InsertSQL(Persons, "申请人", ID)
        self.InsertSQL(DesignPersions, "发明人设计人", ID)
        # 更新维持状态
        Mysql.state(BaseInfos, ID)

    def closeBrow(self):
        self.browser.close()


def start(startTime):
    Mysql.recon()
    global sleep_base
    global accounts
    AlreadyHave = open("成功.txt", "r", encoding='utf-8').readlines()
    AlreadyHave = [i.replace("\n", "") for i in AlreadyHave]
    GetInvoiceModel = GetInvoice()
    GetInvoiceModel.Login()
    Flag = False
    # List = Excel.table_item("申请号")
    List = Mysql.query_all_IDs()

    # WaitAdd = open("成功.txt", "a", encoding="utf-8")
    # WaitAdd.truncate(0)
    # WaitAdd.close()

    for IDs in List:
        if sleep_base > 0:
            sleep_base -= 1
        Log = open("log.txt", "a", encoding="utf-8")
        WaitAdd = open("成功.txt", "a", encoding="utf-8")
        if IDs in AlreadyHave: continue
        ID = IDs.replace(".", "")
        TryTime = 0
        while TryTime < 5:
            try:
                if Flag == True:
                    GetInvoiceModel.browser.get(GetInvoiceModel.Reflesh)
                    GetTime(4)
                GetInvoiceModel.SendID(ID)
                Flag = True

                # 通过验证码
                for pic in range(10):
                    GetTime(2)
                    GetInvoiceModel.GetVeryPic()
                    GetInvoiceModel.Recognize()
                    GetInvoiceModel.Query()
                    noty = None
                    try:
                        noty = WebDriverWait(GetInvoiceModel.browser, 2, 0.5).until(
                            EC.visibility_of_element_located((By.CLASS_NAME, 'noty_message')))
                        GetTime(1)
                    except:
                        pass

                    # 验证码错误
                    if noty:
                        print('验证码不通过')
                        GetInvoiceModel.browser.find_element_by_id("authImg").click()
                        continue

                    try:
                        query_wait = 12 + sleep_base
                        WebDriverWait(GetInvoiceModel.browser, query_wait, 0.5).until(
                            EC.visibility_of_element_located((By.LINK_TEXT, '费用信息')))
                        print('验证码通过')
                        GetTime(2)
                        break
                    except TimeoutException:
                        img = GetInvoiceModel.browser.find_element_by_tag_name("img")
                        img_path = str(img.get_attribute("src"))
                        print(img_path)
                        if img_path == "http://cpquery.cnipa.gov.cn/images/reach_top.png":
                            print("达上限")
                            Mysql.update_accounts('flag', '1', GetInvoiceModel.account['username'])
                            TryTime = 4
                            accounts = Mysql.query_all_accounts()
                            GetTime(1)
                            raise Exception('达到上限，抛异常')
                        else:
                            Flag = True

                        empty_date = GetInvoiceModel.browser.find_elements_by_class_name("empty_date")
                        if empty_date is not None:
                            if pic < 9:
                                print('查询未果，重试')
                                GetInvoiceModel.browser.find_element_by_id("authImg").click()
                                continue
                            else:
                                Mysql.fail(IDs,startTime)
                                TryTime = 5
                                raise Exception('未找到结果')

                needResult = GetInvoiceModel.NeedPayItem()
                if needResult is not None:
                    print('result')
                    # 删除之前的数据
                    Mysql.update("应缴费信息", {"Flag": "1"}, {"申请号": IDs})
                    for item in needResult:
                        item["申请号"] = IDs
                        item["Flag"] = "0"
                        Mysql.insert_add("应缴费信息", item)

                alreadyResult = GetInvoiceModel.AlreadyPayItem()
                if alreadyResult is not None:
                    # 删除之前的数据
                    Mysql.update("已缴费信息", {"Flag": "1"}, {"申请号": IDs})
                    for item in alreadyResult:
                        item["申请号"] = IDs
                        item["Flag"] = "0"
                        Mysql.insert_add("已缴费信息", item)

                print(needResult)
                print(alreadyResult)
                # 更新专利库缴费信息
                Mysql.updatePayInfo(IDs, needResult, alreadyResult)

                GetInvoiceModel.TansBaseInfo(IDs)
                GetInvoiceModel.Date(IDs)
                GetInvoiceModel.browser.find_element_by_xpath('//*[@id="header_query"]').click()

                Mysql.db.commit()
                print("处理成功" + IDs)
                Log.write("成功：" + IDs + "\n")
                WaitAdd.write(IDs + "\n")
                Log.close()
                WaitAdd.close()
                AlreadyHave.append(IDs)
                Flag = False
                break
            except:
                time.sleep(1)
                Mysql.db.rollback()
                sleep_base += 1
                traceback.print_exc()
                if TryTime == 4:
                    Log.write("失败！ID：" + IDs + "\n")
                    Log.close()
                    print('重新登录')

                    # 数据库写法
                    List.append(IDs)

                    # 重新登录
                    GetInvoiceModel.Login()
                TryTime += 1

    GetInvoiceModel.browser.close()
    WaitAdd = open("成功.txt", "a", encoding="utf-8")
    WaitAdd.truncate(0)
    WaitAdd.close()
