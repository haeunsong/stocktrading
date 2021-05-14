import unittest
from stocklab.agent.ebest import EBest 
import inspect 
import time 

class TestEBest(unittest.TestCase):
  # 하나의 테스트 케이스가 실행되기 전에는 setUp 메서드가 호출되고,
  # 테스트 케이스가 수행된 이후에는 tearDown 메서드가 호출.
  def setUp(self):
    # EBest 클래스의 인스턴스(self.ebest) 생성
    self.ebest = EBest("DEMO")
    self.ebest.login() 

  def tearDown(self):
    self.ebest.logout()

def test_get_code_list(self):
  print(inspect.stack()[0][3])
  all_result = self.ebest.get_code_list("ALL")
  assert all_result is not None 
  kosdaq_result = self.ebest.get_code_list("KOSDAQ")
  assert kosdaq_result is not None 
  kospi_result = self.ebest.get_code_list("KOSPI")
  assert kospi_result is not None 

  try:
    error_result = self.ebest.get_code_list("KOS")
  except:
    error_result = None 
  assert error_result is None 
  print("result:",len(all_result),len(kosdaq_result),len(kospi_result))


