import win32com.client
from pywinauto import application
import locale
import os
import time
locale.setlocale(locale.LC_ALL, 'ko_KR')

class Cybos:
  g_objCpStatus = None

  def __init__(self):
    self.g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')

  def kill_client(self):
    print("########## 기존 CYBOS 프로세스 강제 종료")
    os.system('taskkill /IM ncStarter* /F /T')
    os.system('taskkill /IM CpStart* /F /T')
    os.system('taskkill /IM DibServer* /F /T')
    os.system('wmic process where "name like \'%ncStarter%\'" call terminate')
    os.system('wmic process where "name like \'%CpStart%\'" call terminate')
    os.system('wmic process where "name like \'%DibServer%\'" call terminate')

  def connect(self, id_, pwd):
    if not self.connected():
      self.disconnect()
      self.kill_client()
      print("########## CYBOS 프로세스 자동 접속")
      app = application.Application()
      # cybos plus를 정보 조회로만 사용했기 때문에 인증서 비밀번호는 입력하지 않았다.
      app.start(
        'C:\Daishin\Starter\\ncStarter.exe /prj:cp /id:{id} /pwd:{pwd} /autostart'.format(id=id_, pwd=pwd)
      )
  

  def connected(self):
    b_connected = self.g_objCpStatus.IsConnect
    if b_connected == 0:
      return False
    return True

  def disconnect(self):
    if self.connected():
      self.g_objCpStatus.PlusDisconnect()

  def waitForRequest(self):
    remainCount = self.g_objCpStatus.GetLimitRemainCount(1)
    if remainCount <= 0:
      time.sleep(self.g_objCpStatus.LimitRequestRemainTime / 1000)

if __name__ == '__main__':
  cybos = Cybos()
  id = 'your ID'
  password = 'your PW'
  cybos.connect(id, password)
 
