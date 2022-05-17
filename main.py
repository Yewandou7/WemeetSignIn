# author: skywoodsz

import os
import pandas as pd
import pyautogui
import time
from datetime import datetime
import cv2
'''
	利用opencv模板匹配进行入会像素坐标获取，执行鼠标相应动作与密码验证存在判断
	@param tempFile: 模板匹配图像
		   whatDo：鼠标执行动作
		   debug： debug
'''
def ImgAutoClick(tempFile, whatDo, debug=False):
    pyautogui.screenshot('screen.png')
    gray = cv2.imread('screen.png', 0)
    img_templete = cv2.imread(tempFile, 0)
    w, h = img_templete.shape[::-1]
    res = cv2.matchTemplate(gray, img_templete, cv2.TM_SQDIFF)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    top = min_loc[0]
    left = min_loc[1]
    x = [top, left, w, h]
    top_left = min_loc
    bottom_right = (top_left[0] + w, top_left[1] + h)
    pyautogui.moveTo(top+h/2, left+w/2)

    if debug:
        print("?")
        img = cv2.imread("screen.png",1)
        cv2.rectangle(img,top_left, bottom_right, (0,0,255), 2)
        img = cv2.resize(img, (0, 0), fx=0.5, fy=0.5, interpolation=cv2.INTER_NEAREST)
        cv2.imshow("processed",img)
        cv2.waitKey(0)
        cv2.destroyAllWindows()
        os.remove("screen.png")

    whatDo(x)

    if debug:
        img = cv2.imread("screen.png",1)
        cv2.rectangle(img,top_left, bottom_right, (0,0,255), 2)
        img = cv2.resize(img, (0, 0), fx=0.5, fy=0.5, interpolation=cv2.INTER_NEAREST)
        cv2.imshow("processed",img)
        cv2.waitKey(0)
        cv2.destroyAllWindows()
    os.remove("screen.png")
    
    return True

'''
	会议自动登录
	@param meeting_id: 会议号
		   password：密码，若则保持默认NULL
'''
def SignIn(meeting_id, password=None):
    os.startfile("E:\腾讯会议\wemeetapp.exe") # Your own wemeetapp.exe path
    print("Start Sign, please waiting.")
    time.sleep(7)
    ImgAutoClick("JoinMeeting.png", pyautogui.click, False)
    time.sleep(1)
    ImgAutoClick("meeting_id.png", pyautogui.click, False)
    time.sleep(2)
    pyautogui.write(meeting_id)
    time.sleep(2)
    ImgAutoClick("final.png", pyautogui.click, False)
    time.sleep(1)
    # res = ImgAutoClick("password.png", pyautogui.moveTo, False)
    if password != "xxxxxx":
        res = ImgAutoClick("password.png", pyautogui.moveTo, False)
        pyautogui.write(password)
        time.sleep(1)
        ImgAutoClick("passwordJoin.png", pyautogui.click, False)
        time.sleep(1)
    return True

def load_schdule(class_address):
    data = pd.read_excel(class_address)
    data.sort_values(by=['day', 'start_time'], inplace=True)
    data = data.reset_index(drop=True)
    # pd.set_option('max_rows', None)
    # pd.set_option('max_columns', None)
    pd.set_option('expand_frame_repr', False)
    pd.set_option('display.unicode.east_asian_width', True)
    print(data,"\n")
    return data

def get_class(schdule):
    NowDay = datetime.now().weekday()
    today_lessons = schdule.loc[schdule['day'] == NowDay+1]
    today_lessons = today_lessons.reset_index(drop=True)
    # print(today_lessons)
    now = datetime.now()
    for i in range(len(today_lessons)):
        # print(today_lessons.values[i][1].hour, today_lessons.values[i][1].minute, now.hour, now.minute)
        Hour = today_lessons.values[i][1].hour
        Minute = today_lessons.values[i][1].minute
        Second = today_lessons.values[i][1].second
        if now.hour < Hour or (now.hour == Hour and now.minute <= Minute):
            res_time = (Hour - now.hour) * 3600 + (Minute - now.minute) * 60 + Second - now.second
            # res_time = (today_lessons[i][1].hour - now.hour)*3600 + (today_lessons[i][1].minute - now.minute)*60 - now.second
            # print(type(today_lessons.values[i][1].hour), today_lessons.values[i][1].minute, now.hour, now.minute)
            class_name, start_time, meet_id, password = today_lessons.values[i][0], today_lessons.values[i][1], today_lessons.values[i][2], today_lessons.values[i][3]
            print("Next Class: " + class_name + "\n")
            print("Start time: " + str(start_time) + "\n")
            print("Meeting id: " + meet_id + "\n")
            print("Password: " + str(password) + "\n")
            while res_time > 60:
                print("Sleeping, waiting for next lesson. Res time: {} hour, {} min, {} seconds".format(res_time // 3600, (res_time % 3600) // 60, res_time % 60))
                now = datetime.now()
                print("Now time: {} hour, {} min, {} seconds".format(now.hour, now.minute, now.second))
                if res_time > 3600:
                    time.sleep(3600)
                    res_time -= 3600
                else:
                    time.sleep(res_time-30)
                    print("Start trying to join the class")
                    break
            # return lessons_name, lessons_start_time, lessons_meet_id, lessons_password
            return today_lessons.values[i][0], today_lessons.values[i][1], today_lessons.values[i][2], today_lessons.values[i][3]
    return None, None, None, None


if __name__ == "__main__":
    class_address = 'test.xlsx'
    schdule = load_schdule(class_address)
    while True:
        class_name, start_time, meet_id, password = get_class(schdule)
        if class_name == None:
            print("No class today!")
            now = datetime.now()
            time.sleep(86400 - now.hour*3600 - now.minute*60 - now.second)
            print("New day!")
            continue
        now = datetime.now()
        SignIn(meet_id, str(password))
        print("Signed in successfully!")
        time.sleep(200)