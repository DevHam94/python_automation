import pyautogui as pg
import pyperclip
from time import time, sleep

# 카카오톡 클릭
pg.moveTo(1710, 392)
pg.click()

# 검색 버튼 클릭
pg.hotkey('ctrl', 'f')

# 이름 붙여넣기
pyperclip.copy('w사 이아름 차장')
pg.hotkey('ctrl', 'v')

# 사용자 더블클릭 하기
sleep(0.5)
pg.moveTo(1722, 549)
pg.doubleClick()

# 메세지 보내기
pg.moveTo(1689, 949)
pg.click()
pyperclip.copy('안녕하세요. KY마케팅 담당자 김경록입니다.')
pg.hotkey('ctrl', 'v')
pg.hotkey('enter')

