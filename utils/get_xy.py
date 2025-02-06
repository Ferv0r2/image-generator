import pyautogui
import time

print("5초 후에 마우스를 원하는 위치로 이동하세요...")
time.sleep(5)
x, y = pyautogui.position()
print(f"현재 마우스 좌표: ({x}, {y})")
