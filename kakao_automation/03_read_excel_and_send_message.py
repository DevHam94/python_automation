from time import sleep
import pyautogui as pg
import pandas as pd
import pyperclip

df = pd.read_excel('./메세지 자동화.xlsx')

for i, row in df.head(1).df.iterrows():
    print(i, row['이름'])
    print(i, row['메세지'])

    # 카카오톡 클릭
    pg.moveTo(1710, 392) # x 1728 y 457
    pg.click()

    # 검색 버튼 클릭
    pg.hotkey('ctrl', 'f')

    # 이름 붙여넣기
    pyperclip.copy(row['이름'])
    pg.hotkey('ctrl', 'v')

    # 사용자 더블클릭 하기
    sleep(0.5)
    pg.moveTo(1722, 549)
    pg.doubleClick()

    # 메세지 보내기
    pg.moveTo(1689, 949)
    pg.click()
    pyperclip.copy(row['메세지'])
    pg.hotkey('ctrl', 'v')
    pg.hotkey('enter')


