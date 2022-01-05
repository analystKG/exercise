from PyQt5.QAxContainer import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import time
import pandas as pd
from util.const import *

class Kiwoom(QAxWidget): #QAxWidget이라는 클래스 상속 QAxWidget은 Open API를 사용할 수 있도록 연결하는 기능 제공
    def __init__(self): # 프로그램 실행 시 함수들이 자동으로 실행되도록 호출
        super().__init__() # Super: Kiwoom 클래스가 상속받는 QAxWidget 클래스의미 # __init__(): 클래스 초기화
        self._make_kiwoom_instance() # 여러 번 수행하지 않고 Kiwoom 클래스가 생성될 때 자동으로 한 번만 호출하도록 초기화 함수에 넣어줌
        self._set_signal_slots() # self._login_slot을 사용할 수 있게 하기 위해 self._set_signal_slots 함수 먼저 호출
        self._comm_connect()
        self.account_number = self.get_account_number() # 계좌번호 저장
        self.tr_event_loop = QEventLoop() # tr 요청에 대한 응답 대기를 위한 변수

        self.order = {} # 종목 코드를 키 값으로 해당 종목의 주문 정보를 담은 딕셔너리
        self.balace = {} # 종목 코드를 키 값으로 해당 종목의 매수 정보를 담은 딕셔너리

     # PC에서 Kiwoom API를 사용할 수 있도록 설정하는 함수
    def _make_kiwoom_instance(self): # _: 클래스 외부에서 명시적으로 호출해서 사용하지 않는 함수를 클래스 내부에서만 사용한다는 의미
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1") # API 식별자를 전달하여 호출 - Kiwoom 클래스가 Open API가 제공하는 API 제공 함수를을 사용할 수 있게 함

    def _set_signal_slots(self): # API로 보내는 요청들을 받아 올 슬롯을 등록하는 함수
        self.OnEventConnect.connect(self._login_slot) # 로그인 시도에 대한 응답을 _login_slot으로 받도록 설정
        # self._login_slot: 로그인이 성공했는지 실패했는지에 대한 응답을 확인할 수 있음 - 로그인 응답 처리를 받을 때 사용하는 slot 함수
        # self.OnEventConnect.connect: 매개변수로 전달하는 이름(_login_slot)을 가진 함수를 로그인 처리에 대한 응답 slot 함수로 지정
        self.OnReceiveTrData.connect(self._on_receive_tr_data) # 요청했던 TR조회가 성공했을 때 _on_receive_tr_data 함수를 호출하겠다
        self.OnReceiveMsg.connect(self._on_receive_msg) # TR/주문 메시지를 _on_receive_msg로 받도록 설정
        self.OnReceiveChejanData.connect(self._on_chejan_slot) # 주문 접수/체결 결과를 _on_chejan_slot으로 받도록 설정

    def _login_slot(self, err_code): # 로그인 시도 결과에 대한 응답을 얻는 함수
        if err_code == 0: # 로그인 성공
            print("connected")
        else: # 로그인 실패
            print("not connected")

        self.login_event_loop.exit()

    def _comm_connect(self): # 로그인 함수: 로그인 요청 신호를 보낸 이후 응답 대기를 설정하는 함수
        self.dynamicCall("CommConnect()") # API 서버로 로그인 요청을 보냄 # CommConnect: API에서 제공하는 함수, 키움증권 로그인 화면 팝업 기능

        self.login_event_loop = QEventLoop() # 로그인 시도 결과에 대해 응답 대기 상태로 만듦.
        self.login_event_loop.exec_()

    # 계좌 정보 얻어오는 함수
    def get_account_number(self, tag="ACCNO"): # ACCNO: 계좌번호
        # dynamicCall을 사용하여 로그인에 성공한 사용자 정보를 얻어 오는 API 함수 GetLoginInfo 호출
        # tag(여기선 계좌번호)값을 전달함
        account_list = self.dynamicCall("GetLoginInfo(QString)", tag) # tag로 전달한 요청에 대한 응답을 받아옴
        account_number = account_list.split(';')[0] # 국내주식 계좌번호를 account_number 변수에 저장
        print(account_number)
        return account_number

    # 종목 코드 받아 오는 함수
    def get_code_list_by_market(self, market_type): # market_type: 어떤 시장에 해당하는 종목을 얻어 올지 의미히는 구분 값
        code_list = self.dynamicCall("GetCodeListByMarket(Qstring)", market_type) # 구분 값(market_type)에 해당하는 시장에 속한 종목 코드 저장
        code_list = code_list.split(';')[:-1] # 마지막에 빈 값이 저장되므로 제거 후 반환
        return code_list

    # 종목명 받아 오는 함수
    def get_master_code_name(self, code): # 종목 코드를 받아 종목명을 반환한다
        code_name = self.dynamicCall("GetMasterCodeName(Qstring)", code)
        return code_name

    # 종목의 상장일부터 가장 최근 일자까지 일봉 정보를 가져오는 함수
    def get_price_data(self, code):
        self.dynamicCall("SetInputValue(Qstring, Qstring)", "종목코드", code)
        self.dynamicCall("SetInputValue(Qstring, Qstring)", "수정주가구분", "1")
        self.dynamicCall("CommRqData(Qstring, Qstring, int, Qstring)", "opt10081_req", "opt10081", 0, "0001")

        self.tr_event_loop.exec_()

        ohlcv = self.tr_data

        while self.has_next_tr_data: # 추가로 받아올 데이터가 남았는지 확인
            self.dynamicCall("SelfInputValue(Qstring, Qstring)", "종목코드", code)
            self.dynamicCall("SetInputValue(Qstring, Qstring", "수정주가구분", "1")
            self.dynamicCall("CommRqData(Qstring, Qstring, int, Qstring)", "opt10081_req", "opt10081", 2, "0001")
            self.tr_event_loop.exec_() # 응답 대기 상태 진입

            # self.tr_data(현재 호출로 받아 온 데이터)만큼 반복하여 ohlcv(받아 온 데이터를 모아 둔 딕셔너리)의 마지막 부분 [-1:]에 넣는다
            # 한 번에 모든 일봉 데이터를 받아 올 수 없어 호출할 때마다 받아 온 데이터를 합치는 과정
            for key, val in self.tr_data.items():
                ohlcv[key][-1:] = val

        df = pd.DataFrame(ohlcv, columns=['open','high','low','close','volume'], index=ohlcv['date'])
        # DataFrame은 가로축과 세로축을 가진 엑셀 테이블 형태로 데이터 저장 가능
        # DataFrame를 활용하여 가격 정보 저장에 사용할 데이터베이스를 쉽게 사용 가능

        return df[::-1] # 일봉 데이터의 날짜 순서를 뒤집음 # 데이터베이스에 날짜를 오름차순으로 정리하기 위함


    # TR별로 데이터를 가져오는 함수(slot 함수)
    # TR 응답은 모두 _on_receive_tr_data 하나의 함수에서 처리한다
    def _on_receive_tr_data(self, screen_no, rqname, trcode, record_name, next, unused1, unused2, unused3, unused4):
        print("[Kiwoom] _on_receive_tr_data is called {} / {} / {}".format(screen_no, rqname, trcode))
        tr_data_cnt = self.dynamicCall("GetRepeatCnt(Qstring, Qstring)", trcode, rqname)

        if next == '2':
            self.has_next_tr_data = True
        else:
            self.has_next_tr_data = False

        if rqname == "opt10081_req":
            ohlcv = {'date': [], 'open': [], 'high': [], 'low': [], 'close': [], 'volume': []}

            for i in range(tr_data_cnt):
                date = self.dynamicCall("GetCommData(Qstring, Qstring, int, Qstring", trcode, rqname, i, "일자")
                open = self.dynamicCall("GetCommData(Qstring, Qstring, int, Qstring", trcode, rqname, i, "시가")
                high = self.dynamicCall("GetCommData(Qstring, Qstring, int, Qstring", trcode, rqname, i, "고가")
                low = self.dynamicCall("GetCommData(Qstring, Qstring, int, Qstring", trcode, rqname, i, "저가")
                close = self.dynamicCall("GetCommData(Qstring, Qstring, int, Qstring", trcode, rqname, i, "현재가")
                volume = self.dynamicCall("GetCommData(Qstring, Qstring, int, Qstring", trcode, rqname, i, "거래량")

                # 받아 온 일봉 데이터를 딕셔너리 형태로 저장
                ohlcv['date'].append(date.strip())
                ohlcv['open'].append(int(open))
                ohlcv['high'].append(int(high))
                ohlcv['low'].append(int(low))
                ohlcv['close'].append(int(close))
                ohlcv['volume'].append(int(volume))

            self.tr_data = ohlcv
            # Kiwoom 객체의 속성으로 저장하여 객체를 만든 영역에서 접근해서 사용할 수 있도록 하기 위해 ohlcv를 self.tr_data를 저장
            # 받아온 일봉 데이터를 외부에서 사용할 수 있도록 만들고자 값을 옮김
        # rqname이 예수금 요청(opw00001_req)일 때 처리할 코드
        # TR 요청을 만들 때마다 TR이름에 해당하는 Elif구문을 추가하여 해당 TR에 대한 응답 로직을 만들 수 있음.
        elif rqname == "opw00001_req":
            deposit = self.dynamicCall("GetCommData(Qstring,Qstring,int,Qstring", trcode, rqname, 0, "주문가능금액")
            # 예수금 자료형을 int로 바꾸어 self.tr_data로 옮겨담음
            # 옮겨 담는 이유: TR을 요청한 함수(get_deposit)에서 이 값에 저근하여 사용할 수 있도록 하기 위
            self.tr_data = int(deposit)
            print(self.tr_data)

        self.tr_event_loop.exit() # TR 요청을 보내고 응답을 대기시키는 데 사용하는 self.tr.event_loop를 종료하는 역할
        time.sleep(0.5) # 프로그램을 0.5초만큼 쉬도록 하겠다. 이코드가 실행되는 순간부터 바로 0.5초를 대기

    # 주문 접수 및 체결 확인하기
    def send_order(self, rqname, screen_no, order_type, code, order_quantity, order_price, order_classification, origin_order_number=""):
        order_result = self.dynamicCall("SendOrder(Qstring, Qstring, Qstring, int , Qstring, int, int, Qstring, Qstring)",
                                        [rqname, screen_no, self.account_number, order_type, code, order_quantity, order_price, order_classification, origin_order_number])
        return order_result

    # 주문 메시지 수신
    def _on_receive_msg(self, screen_no, rqname, trcode, msg):
        print("[Kiwoom] _on_receive_msg is called {} / {} / {} / {}".format(screen_no, rqname, trcode, msg))

    # 주문 접수/체결
    def _on_chejan_slot(self, s_gubun, n_item_cnt, s_fid_list):
        print("[Kiwoom _on_chejan_slot is called {} / {} / {}".format(s_gubun, n_item_cnt, s_fid_list))

        for fid in s_fid_list.split(";"): # fid 리스트를 ; 기준으로 구
            if fid in FID_CODES:
                code = self.dynamicCall("GetChejanData(int)", '9001')[1:]분 # 종목코드를 얻어 와 A007700처럼 앞자리에 오는 문자를 제거
                data = self.dynamicCall("GetChejanData(int)", fid) # fid를 사용하여 데이터 얻어 오기 (ex) fid:9203을 전달하면 주문 번호를 수신하여 data에 저장

                data = data.strip().lstrip('+').lstrip('-') 데이터에 +,- 가 붙어 있으면 제거

                if data.isdigit(): # 수신한 문자형 데이터 중 숫자인 항목(ex:매수가)를 숫자로 바꿈
                    data = int(data)

                item_name = FID_CODES[fid] # fid코드에 해당하는 항목(item_name)을 찾음(ex) fid=9201 -> item_name = 계좌번호
                print("{}: {}".format(item_name, data)) # 얻어온 데이터를 출력함

                if int(s_gubun) == 0: # 접수/체결(s_gubun)이면 self_order, 잔고 이동이면 self_balance에 값 저장
                    if code not in self.order.keys(): # 아직 order에 종목 코드가 없다면 신규 생성하는 과정
                        self.order[code] = {}

                    self.order[code].update({item_name: data}) # order 딕셔너리에 데이터 저장
                elif int(s_gubun) == 1:
                    if code not in self.balance.keys(): # 아직 balace에 종목 코드가 없다면 신규 생성하는 과정
                        self.balance[code] = {}
                    self.balace[code].update({item_name: data}) # order 딕셔너리에 데이터 저장

        if int(s_gubun) == 0:
            print("* 주문 출력(self.order)")
            print(self.order)
        elif int(s_gubun) == 1:
            print("* 잔고 출력(self.balance)")
            print(self.balance)









    # 예수금 얻어오기
    def get_deposit(self):
        # 입력값 설정
        self.dynamicCall("SetInputValue(Qstring, Qstring", "계좌번호", self.account_number)
        self.dynamicCall("SetInputValue(Qstring, Qstring)","비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(Qstring, Qstring)", "조회구분", "2")
        # 호출
        self.dynamicCall("CommRqData(Qstring,Qstring,int,Qstring)","opw00001_req",'opw00001', 0, "0002")
        # 응답 대기 상태로 만들기
        self.tr_event_loop.exec_()
        return self.tr_data



















