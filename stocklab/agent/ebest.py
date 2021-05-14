import configparser
import win32com.client
import pythoncom
from datetime import datetime 
import time 

# XASession 객체의 이벤트 구현
# XASession : 사용자의 연결 상태 정보를 관리하기 위한 객체.
class XASession:
  # 로그인 상태를 확인하기 위한 클래스 변수
  login_state = 0

  def OnLogin(self,code,msg):
    # 로그인 시도 후 호출되는 이벤트. code가 0000이면 로그인 성공
    if code == "0000":
      print(code,msg)
      XASession.login_state = 1 # 로그인 성공
    else:
      print(code,msg)

  def OnDisconnect(self):
    # 서버와의 연결이 끊어지면 발생하는 이벤트
    print("Session disconnected")
    XASession.login_state = 0

class EBest:
  QUERY_LIMIT_10MIN = 200
  LIMIT_SECONDS = 600 # 10min 


  def __init__(self,mode=None):
    # config.ini 파일을 로드해 사용자, 서버 정보 저장
    # query_cnt는 10분당 200개의 TR수행을 관리하기위한 리스트
    # xa_session_client는 XASession 객체 : param mode:str - 모의서버는 DEMO 실서버는 PROD로 구분

    if mode not in ["PROD","DEMO"]:
      raise Exception("Need to run_mode(PROD or DEMO)")

    run_mode = "EBEST_"+mode 
    config = configparser.ConfigParser()
    config.read('conf/config.ini')
    self.user = config[run_mode]['user']
    self.passwd = config[run_mode]['password']
    self.cert_passwd = config[run_mode]['cert_passwd']
    self.host = config[run_mode]['host']
    self.port = config[run_mode]['port']
    self.account = config[run_mode]['account']

    # 앞에서 만든 XASession의 인스턴스 생성 (login()과 logout()에서 사용)
    self.xa_session_client = win32com.client.DispatchWithEvents("XA_Session.XASession",XASession)

    self.query_cnt = [] 

  # 현재 xingAPI를 사용한 TR의 조회는 10분(600초)에 200회로 제한. 제한 위반시 해당 프로그램 블록처리
  # 블록된 프로그램은 종료한 후에 다시 실행해야 TR 실행가능
  # 블록 방지용 메서드
  def _execute_query(self,res,in_block_name,out_block_name,*out_fields,**set_fields):
    # out_fields: 출력할 필드  / * => 위치 인자
    # set_fields: TR 호출에 필요한 필드값 / ** => 키워드 인자
    """ 
    TR코드를 실행하기 위한 메서드입니다.
    :param res:str 리소스 이름(TR)
    :param in_block_name:str 인 블록 이름
    :param_out_block_name:str 아웃 블록 이름
    :param out_params:list 출력 필드 리스트
    :param in_params:dict 인 블록에 설정할 필드 딕셔너리
    :return result:list 결과를 list에 담아 반환
    """

    time.sleep(1)
    print("current query cnt:",len(self.query_cnt))
    print(res,in_block_name,out_block_name)

    while len(self.query_cnt) >= EBest.QUERY_LIMIT_10MIN:
      time.sleep(1)
      print("waiting for execute query... current query cnt:",len(self.query_cnt))
      # 1초가 지나면 queru_cnt 리스트의 각 값과 현재 시각의 차이를 계산한 다음 filter를 사용하여 
      # 600초(LIMIT_SECONDS)가 넘지 않은 요소들만 리스트(query_cnt)에 담는다.
      # query_cnt에는 TR을 호출한지 10분이 지나지 않은 시각 정보만 담기며, 리스트의 갯수가 200개를 넘지 않으면 while문을 통과해 다음 TR 수행 가능.
      # 만약 리스트의 갯수가 200개가 넘ㅇ면 호출된지 10분이 지난 시각 정보를 제거하는 과정을 반복 수행한다.

      self.query_cnt = list(filter(lambda x: (datetime.today() - x).total_seconds() < EBest.LIMIT_SECONDS,self.query_cnt))
      xa_query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery",XAQuery) # XAQuery 객체 생성
      xa_query.LoadFromResFile(XAQuery.RES_PATH + res + ".res") # 리소스 파일 불러오기
     
      #in_block_name 셋팅
      for key,value in set_fields.items():
        xa_query.SetFieldData(in_block_name,key,0,value) # DevCenter의 TR 모의 실행에서 TR을 실행하기 전에 필요한 값을 채우는 부분과 같은 역할(필드 데이터 채우기)
      errorCode = xa_query.Request(0) # TR 요청하기

      # 요청 후 대기
      waiting_cnt = 0
      while xa_query.tr_run_state == 0: # 결과가 수신되지 않은 상태인 tr_run_state의 값이 0일 동안은 while문을 반복하며 기다린다.
        waiting_cnt = 1
        if waiting_cnt%100000 == 0: # while문을 수행하는동안 100,000번에 한번씩만 waiting 메시지 출력
          print("Waiting...",self.xa_session_client.GetLastError())
        pythoncom.PumpWaitingMessages()

      # 결과 블록
      result = []
      count = xa_query.GetBlockCount(out_block_name) # 결과가 몇개인지 확인

      # 결과의 수만큼 for문 반복
      for i in range(count):
        item = {}
        # 메서드에서 위치인자로 전달받은 out_fields에 정의한 필드의 값만 가져온 후 result에 담기
        for field in out_fields:
          value = xa_query.GetFieldData(out_block_name,field,i)
          item[field] = value 
        result.append(item)

      # 제약시간 체크
      XAQuery.tr_run_state = 0
      self.query_cnt.append(datetime.today()) # 현재 시각 추가

      # 영문 필드명을 한글 필드명으로 변환
      for item in result:
        for field in list(item.keys()): # 각 항목의 필드명을 리스트 형태로 가져온다.
          if getattr(Field,res,None): # 수행하려는 res(ex> t1305)가 Field 클래스의 속성에 있는지 확인 
            res_field = getattr(Field,res,None)
            if out_block_name in res_field:
              field_hname = res_field[out_block_name]
              if field in field_hname:
                item[field_hname[field]] = item[field]
                item.pop(field)
      

  def login(self):
    self.xa_session_client.ConnectServer(self.host,self.port)
    self.xa_session_client.Login(self.user,self.passwd,self.cert_passwd,0,0)
    while XASession.login_state == 0: # 로그인 x 
      pythoncom.PumpWaitingMessages()

  def logout(self):
    #result = self.xa_session_client.Logout()
    #if result:
    XASession.login_state = 0
    self.xa_session_client.DisconnectServer() # 서버와의 연결 종료

class XAQuery:
  # TR별로 Res 파일 존재
  RES_PATH = "C:\\eBEST\\xingAPI\\RES\\"
  # 현재 tr이 실행 중인지 확인 용도
  tr_run_state = 0 

 # 요청한 API에 대해 데이터를 수신했을 때 발생하는 이벤트
  def onReceiveData(self,code):
    print("OnReceiveData",code)
    XAQuery.tr_run_state = 1 

 # 요청한 데이터가 정상/오류인지 구분
  def onReceiveMessage(self,error,code,message):
    print("OnreceiveMessage",error,code,message)


class Field:
  t1101 = {
    "t1101OutBlock":{
      "hname":"한글명",
      "price":"현재가"
    }
  }
  t1305 = {
    "t1305OutBlock1":{
      "date":"날짜",
      "open":"시가",
      "close":"종가",
      "marketcap":"시가총액"
    }
  }
  t8436 = {
    "t8436OutBlock":{
      "hname":"종목명",
      "shcode":"단축코드",
      "expcode":"확장코드",
      "spac_gubun":"기업인수목적회사여부",
      "filler":"filler(미사용)"
    }
  }

def get_code_list(self,market=None):
  """
  TR: t8436 코스피, 코스닥의 종목 리스트를 가져온다.
  : param market:str 전체(0), 코스피(1), 코스닥(2)
  : return result:list 시장별 종목 리스트
  """

  if market!="ALL" and market !="KOSPI" and market!="KOSDAQ":
    raise Exception("Need to market param(ALL,KOSPI,KOSDAQ")

  market_code = {"ALL":"0","KOSPI":"1","KOSDAQ":"2"}
  in_params = {"gubun":market_code[market]}
  out_params = {'hname','shcode','expcode','spac_gubun'}
  result = self._execute_query("t8436","t8436InBlock","t8436OutBlock",*out_params,**in_params)
  return result