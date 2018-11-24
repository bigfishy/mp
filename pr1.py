# -*-coding:utf-8-*-
import requests
import sys
import telegram
import telepot
from lxml import html, cssselect
from datetime import datetime, timedelta, time
import time as tm
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters
import xlwt
import xlrd
import pandas as pd
import xml.etree.ElementTree as ET



def output(filename, members):
    book = xlwt.Workbook()
    sh = book.add_sheet('member')
    col_name = 'member'

    i = 0
    for member in members:
        sh.write(i, 0, member)
        i += 1
    book.save(filename)


def input(filename):
    book = xlrd.open_workbook(filename)
    book = pd.read_excel(filename, engine='xlrd')

    sh = book.sheet_by_index(0)
    nrows = sh.nrows

    row_val = []
    for row_num in range(nrows):
        row_val.append(sh.row_values(row_num)[0])

    return row_val


def send_to_members(mytoken, stock_code, members):
    # 토큰 받기
    bot = telepot.Bot(mytoken)
    # bot정보 인식
    bot.getMe()
    response = bot.getUpdates()
    print(response)


def id_list_update(mytoken):
    bot = telepot.Bot(mytoken)
    bot.getMe()
    response = bot.getUpdates()

    id_list = [i['message']['chat']['id'] for i in response]
    id_list = list(set(id_list))

    return id_list


def send(mytoken, members, message):
    bot = telepot.Bot(mytoken)
    invalid_chats = list()

    for chat_id in members:
        try:
            bot.sendMessage(chat_id=chat_id, parse_mode='MarkDown', text=message)
        except:
            print('에러', chat_id)
            invalid_chats.append(chat_id)
    return


def get_enabled_chats():
    query = EnableStatus.query(EnableStats.enabled == True)
    return query.fetch()


# 웹 페이지 크롤링
def url_elem(u):
    url = u
    resp = requests.get(url)
    return html.fromstring(resp.text)

# 국내 지수 불러오기
def ko_index_price(ko_index_code):
    url = 'http://m.finance.daum.net/m/quote/' + ko_index_code
    elem = url_elem(url)

    k_price = elem.cssselect('em.price')[0].text_content().strip()  # 주가
    k_point = elem.cssselect('em.price_fluc')[0].text_content().strip()  # 등락
    k_rate = elem.cssselect('em.rate_flucs')[0].text_content().strip()  # 등락률
    return [k_price[3:], k_point, k_rate]


# 개별 종목 불러오기
def stock_price(stock_code):
    url = 'http://m.finance.daum.net/m/item/main.daum?code=' + stock_code
    elem = url_elem(url)

    price = elem.cssselect('span.price')[0].text_content().strip()  # 주가
    price_gap = elem.cssselect('span.price_fluc')[0].text_content().strip()  # 등락
    price_gap_rate = elem.cssselect('span.rate_fluc')[0].text_content().strip()  # 등락률
    return [price, price_gap, price_gap_rate]

# 국가별 환율 불러오기
def exchange_rate(country_code):
    url = 'http://m.exchange.daum.net/mobile/exchange/exchangeDetail.daum?code=' + country_code
    elem = url_elem(url)

    e_price = elem.cssselect('em.price')[0].text_content().strip()  # 환율
    e_point = elem.cssselect('em.price_fluc')[0].text_content().strip()  # 등락
    e_rate = elem.cssselect('em.rate_flucs')[0].text_content().strip()  # 등락률
    return [e_price, e_point, e_rate]

# 국가별 주가지수 불러오기
def index_price(index_code):
    url = 'http://m.finance.daum.net/m/world/indexDetail.daum?ric=/.' + index_code
    elem = url_elem(url)

    i_price = elem.cssselect('em.price')[0].text_content().strip()  # 주가
    i_point = elem.cssselect('em.price_fluc')[0].text_content().split()  # 주가
    i_rate = elem.cssselect('em.rate_flucs')[0].text_content().strip()  # 주가
    return [i_price, "".join(i_point), i_rate]

# 1. 당일 마감지수 : 코스피/코스닥/미래에셋대우/주요국환율 (주중 오후 3시 40분)
def close_msg(mytoken, members):
    print('------------------------- close_msg start ---------------------------')
    bot = telepot.Bot(mytoken)
    # 국내지수
    ta = ' '.join(ko_index_price('kospi.daum'))
    tb = ' '.join(ko_index_price('kosdaq.daum'))

    # 미래에셋대우
    tc = ' '.join(stock_price('006800'))
    
    # 국가별환율
    td = ' '.join(exchange_rate('USD'))
    te = ' '.join(exchange_rate('JPY'))
    tf = ' '.join(exchange_rate('CNY'))


    
    for member in members:
        try:
            bot.sendMessage(member, '\n'.join(['[ 당일마감시황 ]', '(1) 코스피', ta, '(2) 코스닥', tb, '(3) 미래에셋대우', tc ,'(4) 환율 UDS', td, '(5) 환율 JPY', te, '(6) 환율 CNY', tf]))
        except:
            print(member, '\n'.join(['당일마감시황 오류건']))

    print('------------------------- close_msg end ---------------------------')


# 2. 전일 해외지수 : 다우, 나스닥, H, 상해, 니케이 (주중 오전 8시 30분)
def yesterday_msg(mytoken, members):
    print('------------------------- yesterday_msg start ---------------------------')
    bot = telepot.Bot(mytoken)
    # 다음 증권 세계지수

    iaa = ' '.join(index_price('DJI'))
    iab = ' '.join(index_price('IXIC'))
    iac = ' '.join(index_price('GSPC'))
    iad = ' '.join(index_price('N225'))
    iae = ' '.join(index_price('SSEC'))
    iaf = ' '.join(index_price('HSCE'))
    iag = ' '.join(index_price('BSESN'))
    iah = ' '.join(index_price('FTSE'))
    iai = ' '.join(index_price('GDAXI'))
    iaj = ' '.join(index_price('BVSP'))


    for member in members:
        try:
            bot.sendMessage(member,
                            '\n'.join(['[ 전일 해외지수 ]', '(1) 다우지수', iaa, '(2) 나스닥', iab, '(3) S&P500', iac, '(4) 니케이', iad, '(5) 상해종합', iae, '(6) 홍콩H', iaf, '(7) 인도BSE', iag, '(8) 영국FTSE', iah, '(9) 독일DAX', iai, '(10) 브라질BVSP', iaj]))
        except:
            print(member, '\n'.join(['전일 해외지수 오류건']))

    print('------------------------- yesterday_msg End ---------------------------')




# 3. 네이버 금융섹션 헤드라인
def bestread_news_msg(mytoken, members):
    print('------------------------- bestread_news_msg start ---------------------------')
    bot = telepot.Bot(mytoken)
    today_txt = datetime.now().strftime('%Y%m%d')
    url = 'http://finance.naver.com/news/news_list.nhn?mode=RANK&date=' + today_txt

    elem = url_elem(url)
    st_text = ''
    iSeq = 0

    for i in elem.cssselect('div.hotNewsList a'):
        print(i)
        iSeq += 1
        if i.text_content() != '':
            href = 'http://finance.naver.com' + i.get('href')
            st_text = st_text + '(' + str(iSeq) + ') <a href="' + href + '">' + i.text_content().strip() + '</a>\n'
        if iSeq >= 10:
            break

    st_text = '[ 네이버 금융섹션 헤드라인 ]\n' + st_text
    for member in members:
        try:
            bot.sendMessage(chat_id=member, text=st_text, parse_mode=telegram.ParseMode.HTML)
        except:
            print('Error:', member)

    print('------------------------- bestread_news_msg End ---------------------------')


    # 4. 금주의 미래에셋대우 ELS
def els_thisweek_msg(mytoken, members):
    print('------------------------- els_thisweek_msg start ---------------------------')
    bot = telepot.Bot(mytoken)

    xml = """<?xml version="1.0" encoding="utf-8"?>
    <message>
      <proframeHeader>
        <pfmAppName>FS-DIS2</pfmAppName>
        <pfmSvcName>DISDlsOfferSO</pfmSvcName>
        <pfmFnName>selectSubscribing</pfmFnName>
      </proframeHeader>
      <systemHeader></systemHeader>
        <DISDlsDTO>
        <val1></val1>
        <val2></val2>
        <val3></val3>
        <val4></val4>
        <val5></val5>
        <val6>0</val6>
    </DISDlsDTO>
    </message>"""
    url = "http://dis.kofia.or.kr/proframeWeb/XMLSERVICES"
    response = requests.post(url, data=xml).text

    els_list = []
    els_list.append("상품명 ㅣ 기초자산 ㅣ 제시수익률")
    for entity in ET.fromstring(response).iter('DISDlsDTO'):
        if entity[3].text == "미래에셋대우":
            els_list.append(
                entity[5].text.replace('미래에셋대우 ', '').replace(' 기타파생결합증권', 'DLS').replace(' 파생결합증권', 'ELS').replace(
                    ' 기타파생결합사채', 'DLB').replace(' 파생결합사채', 'ELB') + 'ㅣ' + entity[7].text.replace('<br/>', '/').replace(
                    ' 선물 최근월물', '').replace(' Index', '').replace(' ', '') + 'ㅣ' + entity[14].text + '%')
            els2str = ''
            for els in els_list:
                els2str += str(els) + '\n'

            st_text = str(els2str)




    st_text = '[ 이번주 미래에셋대우 ELS ]\n' + st_text
    for member in members:
        try:
            bot.sendMessage(chat_id=member, text=st_text, parse_mode=telegram.ParseMode.HTML)
        except:
            print('Error:', member)

    print('------------------------- els_thisweek_msg end ---------------------------')

    # 5. 금주의 타사 ELS
def els_etc_msg(mytoken, members):
    print('------------------------- els_etc_msg start ---------------------------')
    bot = telepot.Bot(mytoken)

    xml = """<?xml version="1.0" encoding="utf-8"?>
    <message>
      <proframeHeader>
        <pfmAppName>FS-DIS2</pfmAppName>
        <pfmSvcName>DISDlsOfferSO</pfmSvcName>
        <pfmFnName>selectSubscribing</pfmFnName>
      </proframeHeader>
      <systemHeader></systemHeader>
        <DISDlsDTO>
        <val1></val1>
        <val2></val2>
        <val3></val3>
        <val4></val4>
        <val5></val5>
        <val6>0</val6>
    </DISDlsDTO>
    </message>"""
    url = "http://dis.kofia.or.kr/proframeWeb/XMLSERVICES"
    response = requests.post(url, data=xml).text

    els_list2 = []
    els_list2.append("상품명 ㅣ 기초자산 ㅣ 제시수익률")
    for entity in ET.fromstring(response).iter('DISDlsDTO'):
        if entity[3].text != "미래에셋대우":
            els_list2.append(entity[5].text.replace('투자증권','').replace('해피플러스','').replace('(고위험, 원금비보장형)','').replace('세이프','').replace('(저위험, 원금지급형)','').replace('스마트','').replace(' MY ELS','').replace(' MY ELB','').replace('[Balance]','').replace(' able','').replace('트루 ','한투').replace('종금증권','').replace(' 홈런S ','').replace(' 홈런D ','').replace('주가연계증권','ELS').replace('기타파생결합증권','DLS').replace('파생결합증권','ELS').replace('기타파생결합사채','DLB').replace('파생결합사채','ELB').replace('증권','').replace('(','').replace(')','').replace('ELS','').replace('DLS','').replace('ELB','').replace('DLB','').replace('MY','').replace('주가연계','') +'ㅣ' + entity[7].text.replace('<br/>', '/').replace(' 선물 최근월물','').replace(' Index','').replace(' ','')+'ㅣ' + entity[14].text + '%')
            els2str2 = ''
            for els in els_list2:
                els2str2 += str(els) + '\n'
            " ".join(els2str2)
            sorted(els2str2, key=lambda els2str2: els2str2[0])

            st_text2 = str(els2str2)




    st_text2 = '[ 이번주 타사 ELS ]\n' + st_text2
    for member in members:
        try:
            bot.sendMessage(chat_id=member, text=st_text2, parse_mode=telegram.ParseMode.HTML)
        except:
            print('Error:', member)

    print('------------------------- els_etc_thisweek_msg end ---------------------------')


def voice_handler(mytoken, update):
    bot = telepot.Bot(mytoken)
    file = bot.getFile(update.message.voice.file_id)
    print("file_id: " + str(update.message.voice.file_id))
    file.download('voice.ogg')
    return


def main(argv):
    # URL_KEYWORD1 = argv[1]

    ##### NEW 로직 : 메시지 발송 시간 초기값 ######
    push_time_msg1 = time(8, 40, 1)  # 1. 당일 마감 시황 (주중 오후 3시 40분)
    push_time_msg2 = time(8, 10, 1)  # 2. 전일 해외지수 : 다우, 나스닥, H, 상해, 니케이 (주중 오전 8시 10분)
    push_time_msg3 = time(1, 00, 1)  # 3. 네이버 금융섹션 헤드라인 (주중 오전 11시)
    push_time_msg4 = time(9, 00, 1)  # 4. 금주의 당사 ELS (주중 오전 9시)
    push_time_msg5 = time(10, 00, 1)  # 5. 금주의 타사 ELS (주중 오전 10시)

    # while loop 발송 완료 상태로 초기화
    # 어느 시간에라도 실행시켜도 "다음 날" 아침부터 순차적으로 돌아가도록 ...
    push_res_msg1 = 0
    push_res_msg2 = 0
    push_res_msg3 = 0
    push_res_msg4 = 0
    push_res_msg5 = 0

    # 봇 접속 토큰
    mytoken = "427780087:AAEpCew3pQ4_1sDWwfQVaVoJwBjmljgBX3o"
    ####### 발송대상자 : 김,빈
    members_id = [ ]

    bot = telepot.Bot(mytoken)
    users = bot.getUpdates()
    updater = Updater(token=mytoken)
    dispatcher = updater.dispatcher
    dispatcher.add_handler(MessageHandler(Filters.voice, voice_handler))

    filename = 'C:/Anaconda3/Project/IDs.csv'
    tcMatFile = open(filename, "r")

    tcMat = tcMatFile.read()
    print(tcMat)
    tcMatFile.close()

    members_id2 = tcMat.split("\n")
    members_id.extend(members_id2)

    while True:
        print('멤버리스트 이전', members_id)

        members_id.extend(id_list_update(mytoken))
        members_id = list(set(members_id))
        # members_id = [422784265,442481183]

        tcMatFile = open(filename, "w")

        for member_id in members_id:
            tcMatFile.writelines(str(member_id) + '\n')

        tcMatFile.close()

        print('멤버리스트 이후', members_id)

        # 발송이력 초기화 : 하루 한번만 발송되도록 새벽 0시 기준으로 0값 세팅 -> 발송 후 1로 변경#
        if datetime.now().time().hour == 0:
            push_res_msg1 = 0
            push_res_msg2 = 0
            push_res_msg3 = 0
            push_res_msg4 = 0
            push_res_msg5 = 0

        now_weekday = datetime.now().weekday()  # 현재 요일 : 월요일 0, 일요일 7
        now_time = datetime.now().time()  # 현재 시

        ##################################################################################
        # 1. 당일 마감 시황 (주중 오후 3시 40분)
        try:
            if now_time >= push_time_msg1 and now_weekday <= 4:  # 알람시점 도달 & 주중
                if push_res_msg1 == 0:
                    close_msg(mytoken, members_id)
                    push_res_msg1 += 1  # 당일은 더 이상 발송 안되도록
        except:
            print("Error : # 1. 당일 마감 시황 (주중 오후 3시 40분)")
            print("오류 시간 : ", datetime.now())
            push_res_msg1 += 1  # 당일은 더 이상 발송 안되도록

        ##################################################################################
        # 2. 전일 해외지수 : 다우, 나스닥, H, 상해, 니케이 (주중 오전 8시 00분)
        try:
            if now_time >= push_time_msg2 and now_weekday <= 4:  # 알람시점 도달 & 주중
                if push_res_msg2 == 0:
                    yesterday_msg(mytoken, members_id)
                    push_res_msg2 += 1  # 당일은 더 이상 발송 안되도록
        except:
            print("Error : # 2. 전일 해외지수 : 다우, 나스닥, H, 상해, 니케이 (주중 오전 8시 30분)")
            print("오류 시간 : ", datetime.now())
            push_res_msg2 += 1  # 당일은 더 이상 발송 안되도록


        ##################################################################################
        # 3. 네이버 금융섹션 헤드라인 (주중 오전 11시)
        try:
            if now_time >= push_time_msg3 and now_weekday <= 4:  # 알람시점 도달 & 주중
                if push_res_msg3 == 0:
                    bestread_news_msg(mytoken, members_id)
                    push_res_msg3 += 1  # 당일은 더 이상 발송 안되도록
        except:
            print("Error : # 3. 네이버 금융섹션 헤드라인 (주중 오전 11시)")
            print("오류 시간 : ", datetime.now())
            push_res_msg3 += 1


        ##################################################################################
        # 4. 금주 미래에셋대우 ELS (주중 오전 9시)
        try:
            if now_time >= push_time_msg4 and now_weekday <= 4:  # 알람시점 도달 & 주중
                if push_res_msg4 == 0:
                    els_thisweek_msg(mytoken, members_id)
                    push_res_msg4 += 1  # 당일은 더 이상 발송 안되도록
        except:
            print("Error : # 4. 금주 미래에셋대우 ELS (주중 오전 9시)")
            print("오류 시간 : ", datetime.now())
            push_res_msg4 += 1


        ##################################################################################

        # 5. 금주 타사 ELS (주중 오전 10시)
        try:
            if now_time >= push_time_msg5 and now_weekday <= 4:  # 알람시점 도달 & 주중
                if push_res_msg5 == 0:
                    els_etc_msg(mytoken, members_id)
                    push_res_msg5 += 1  # 당일은 더 이상 발송 안되도록
        except:
            print("Error : # 5. 금주 타사 ELS (주중 오전 10시)")
            print("오류 시간 : ", datetime.now())
            push_res_msg5 += 1


        ##################################################################################

        tm.sleep(60)  # 1분간 쉰다


if __name__ == '__main__':
    main(sys.argv)
