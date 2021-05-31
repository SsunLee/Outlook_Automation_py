from datetime import date
import datetime
import time
import win32com.client as win32


class Outlook(object):

    def __init__(self,title, mailto, mailcc, textbody):
        self.title = title
        self.mailto = mailto
        self.mailcc = mailcc
        self.textbody = textbody
    def getTitle(self):
        return self.title
    def getMailTo(self):
        return self.mailto
    def getMailCC(self):
        return self.mailcc
    def getBody(self):
        return self.textbody


    def mailContents(self):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        textTitle = self.title
        textMailTo = self.mailto
        textMailCC = self.mailcc
        textBody = self.textbody

        mail.Subject = textTitle
        mail.To = textMailTo
        mail.CC = textMailCC
        mail.GetInspector
        mail.HTMLBody = textBody + mail.HTMLBody
        mail.Display()

# LG전자 외주지원반 공식 Mail
def num_1():
    today = str(date.today().strftime("%Y.%m.%d"))
    texttitle = "[공유] TASK 작업 내역 전달(" + today + ")"
    textMailTo = "'홍길동' <abcd.zz@aaa.com>; '홍길순' <abc.aab@aaa.com>"
    textMailCC = "''박보검' <bogum.park@abcd.com>; "
    textBody = """
      <body>
        <p>안녕하세요. 이순배 입니다.<br>
           금일 작업 내역전달 드립니다.<br>
           <br>
           <br>
           감사합니다. <br>
           이순배 드림
        </p>
      </body>
    """
    _mailinfo = {'title': texttitle, 'mailTo': textMailTo, 'mailCC': textMailCC, 'htmlBody': textBody}
    return _mailinfo
# 내부 TASK 메일
def num_2():
    today = str(date.today().strftime("%Y.%m.%d"))
    texttitle = "[TASK] 일일 업무 진행내역_이순배_(" + today + ")"
    textMailTo = "'홍길동' <abcd.zz@aaa.com>; '홍길순' <abc.aab@aaa.com>"
    textMailCC = "''박보검' <bogum.park@abcd.com>; "
    textBody = """
      <body>
        <p>안녕하세요. 이순배 입니다.<br>
           금일 작업 내역전달 드립니다.<br>
           <br>
           <br>
           감사합니다. <br>
           이순배 드림
        </p>
      </body>
    """
    _mailinfo = {'title': texttitle, 'mailTo': textMailTo, 'mailCC': textMailCC, 'htmlBody': textBody}
    return _mailinfo
# 보안당직메일
def num_3():
    today = str(date.today().strftime("%Y.%m.%d"))
    texttitle = "[공유][보안점검결과]_TASK_" + str(date.today().strftime("%m")) + "_" + str(date.today().strftime("%d")) + "_(" + getDay() + ")"
    textMailTo = "'홍길동' <abcd.zz@aaa.com>; '홍길순' <abc.aab@aaa.com>"
    textMailCC = "''박보검' <bogum.park@abcd.com>; "

    now = time.localtime()
    text1 = "<body> <p>안녕하세요. 이순배 입니다.<br>"
    text2 = str(date.today().strftime("%m")) + "/" + str(date.today().strftime("%d")) + "_(" + getDay() + ")" + "보안점검 결과 전달 드립니다. <br><br>"
    text3 = "점검 일시 : " + str(date.today().strftime("%y")) + "년 " + str(date.today().strftime("%m")) + "월 " + str(date.today().strftime("%d")) + "일  "
    a_Time = str(now.tm_hour) + ":" + str(now.tm_min) + " ~ " + str(now.tm_hour) + ":" + str(now.tm_min) + "<br>"
    text4 = "<br> 점검 예외인원 <br> \n 다음 근무자 : <br><br>\n 감사합니다. <br> \n 이순배 드림\n </p>\n</body>"
    textBody =text1+text2+text3+a_Time+text4

    _mailinfo = {'title': texttitle, 'mailTo': textMailTo, 'mailCC': textMailCC, 'htmlBody': textBody}
    return _mailinfo


def getDay_c(a, b, c):
    daylist = ['월', '화', '수', '목', '금', '토', '일']
    return daylist[datetime.date(a,b,c).weekday()]

# 부재중 메일
def num_4(li):
    sYear = int(date.today().strftime("%Y"))
    sMonth = int(li[0]);sDay = int(li[1])
    swday = getDay_c(sYear, sMonth, sDay)
    textMailTo =""
    textMailCC =""
    texttitle = f"[부재중 공지]  {sMonth}/{sDay}({swday}) 부재중 임을 알려드립니다."
    line1 = f"""
            <body> 
                <p> 안녕하세요. 이순배 입니다. 
                    <br> 개인 사정으로 인하여 {sMonth}/{sDay}({swday}) 부재 중임을 알려드립니다. 
                    <br>
                    <br> 급한 용무가 있으신분은 아래 번호로 연락 주시기 바랍니다.
                    <br>
                    <br> ● 날짜 : {sYear}. {sMonth}/{sDay}({swday}) 
                    <br> ● 전화번호 : 010-7377-7753 
                    <br> ● 업무 대행자 : 없음. 
                    <br>
                    <br> 감사합니다. 
                    <br> 이순배 드림 
                </p> 
            </body>"""
    textBody =line1
    _mailinfo = {'title': texttitle, 'mailTo': textMailTo, 'mailCC': textMailCC, 'htmlBody': textBody}
    return _mailinfo
def num_5(li):
    sYear = int(date.today().strftime("%Y"))
    sMonth = int(li[0]);sDay = int(li[1])
    swday = getDay_c(sYear, sMonth, sDay)

    sMonth2 = int(li[2]);sDay2 = int(li[3])
    swday2 = getDay_c(sYear, sMonth2, sDay2)
    textMailTo =""
    textMailCC =""
    texttitle = f"[부재중 공지]  {sMonth}/{sDay}({swday}) ~ {sMonth2}/{sDay2}({swday}) 부재중 임을 알려드립니다."
    line1 = f"""
            <body> 
                <p> 안녕하세요. 이순배 입니다. 
                    <br> 개인 사정으로 인하여 {sMonth}/{sDay}({swday}) ~ {sMonth2}/{sDay2}({swday}) 부재 중임을 알려드립니다. 
                    <br>
                    <br> 급한 용무가 있으신분은 아래 번호로 연락 주시기 바랍니다.
                    <br>
                    <br> ● 날짜 : {sYear}. {sMonth}/{sDay}({swday}) ~ {sMonth2}/{sDay2}({swday}) 
                    <br> ● 전화번호 : 010-7377-7753 
                    <br> ● 업무 대행자 : 없음. 
                    <br>
                    <br> 감사합니다. 
                    <br> 이순배 드림 
                </p> 
            </body>"""
    textBody =line1
    _mailinfo = {'title': texttitle, 'mailTo': textMailTo, 'mailCC': textMailCC, 'htmlBody': textBody}
    return _mailinfo

def getDay():
    now = time.localtime()
    week = ('월', '화', '수', '목', '금', '토', '일')
    a = week[now.tm_wday]
    return a

menu = """
★ developer by Sunbae.lee
1. TASK 업무 메일 보고 
2. 내부 TASK 
3. 보안당직 메일
4. 부재중 메일
5. 나가기 (메뉴 외 번호 입력시 나가기)
"""
li = []
while True:
    print(menu)
    num = int(input("실행 할 메뉴를 선택하세요."))
    if num == 1:
        print("1번을 선택했습니다.")
        # TASK 업무 메일 보고 
        print("실행 중... \n")
        a = Outlook(num_1().get('title'), num_1().get('mailTo'), num_1().get('mailCC'), num_1().get('htmlBody'))
        a.mailContents()
        time.sleep(0.3)
        print("실행 완료... \n")
    elif num == 2:
        print("2번을 선택했습니다.")
        print("실행 중... \n")
        # 내부 task 메일
        a = Outlook(num_2().get('title'), num_2().get('mailTo'), num_2().get('mailCC'), num_2().get('htmlBody'))
        a.mailContents()
        time.sleep(0.3)
        print("실행 완료... \n")
    elif num == 3:
        print("3번을 선택했습니다.")
        print("실행 중... \n")
        # 보안당직 메일
        a = Outlook(num_3().get('title'), num_3().get('mailTo'), num_3().get('mailCC'), num_3().get('htmlBody'))
        a.mailContents()
        time.sleep(0.3)
        print("실행 완료... \n")

    elif num ==4:
        # 부재중 메일
        print("4번을 선택했습니다.")
        print("실행 중... \n")
        print("  (1) 연차를 하루 사용합니다. (반차포함) ex) 12/24(월)")
        print("  (2) 연차를 하루 이상 사용합니다. (반차포함) ex) 12/24(월) ~ 12/26(수)")
        print("연차를 보낼 날짜를 위의 번호로 선택하세요.")
        dnum = int(input(""))
        if dnum == 1:
            print("ex) 12/24(월) 이라면 '12' 입력")
            stmonth = int(input(">> "))
            print("ex) 12/24(월) 이라면 '24' 입력")
            stday = int(input(">> "))
            li = [stmonth, stday,stmonth, stday]
            a = Outlook(num_4(li).get('title'), num_4(li).get('mailTo'), num_4(li).get('mailCC'), num_4(li).get('htmlBody'))
            a.mailContents()
        else:
            print("ex) 12/24(월) 이라면 '12' 입력")
            stmonth = int(input(">> "))
            print("ex) 12/24(월) 이라면 '24' 입력")
            stday = int(input(">> "))
            print("ex) 12/25(화) 이라면 '12' 입력")
            stmonth2 = int(input(">> "))
            print("ex) 12/25(화) 이라면 '25' 입력")
            stday2 = int(input(">> "))
            li = [stmonth, stday,stmonth2, stday2]
            a = Outlook(num_5(li).get('title'), num_5(li).get('mailTo'), num_5(li).get('mailCC'), num_5(li).get('htmlBody'))
            a.mailContents()

    else:
        break
