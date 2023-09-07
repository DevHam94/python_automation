import pyautogui as pg

x, y = pg.position()
print(x, y)
# login - id입력 1645 664

# 마우스 특정 위치로 이동
pg.moveTo(1645, 664)

# 마우스 클릭
pg.click()

