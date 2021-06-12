from os import supports_bytes_environ
from threading import main_thread
import pyautogui as pag
import random , time , subprocess
import win32api
import mss ,cv2
import numpy as np
from pytesseract import pytesseract
import smtplib
from email.message import EmailMessage
import imghdr


pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
#please change the path if you are not using window or you changed the path 

x1 =0
y1 =0
x2 = 0
y2 = 0

def get_location_image():
    print("Please found the left up location of your image")
    a = 100
    global x1 ,x2, y1, y2

    while a >= 0 :
        time.sleep(0.1)
        a = win32api.GetKeyState(0x01)
        x1, y1 = pag.position()
        #print(type(x))
        location = 'X : ' + str(x1) + ' Y : ' + str(y1)
        if a< 0:
            print(location)
            break    
    
    
   

    print("Please found the right bottom location of your image")
    a1=100
    while a1 >= 0 :
        time.sleep(0.1)
        a1 = win32api.GetKeyState(0x01)
        x2, y2 = pag.position()
        location1 = 'X : ' + str(x2) + ' Y : ' + str(y2)
        if a1< 0:
            print(location1)
            break


#################################################################################
#main part
#################################################################################

print("Welcome to picture to the text!")
print("please enter 1 if you just want to use picture to text")
print("Or if you want to use auto mail alarm for diffrence, enter 2!")
user_input = input()
if int(user_input) == 1 :
    get_location_image()
    image_pos = { 'left' : x1 , 'top' : y1 , 'width' : (x2 - x1) , 'height' : (y2 - y1)}
    words = ""

    with mss.mss() as sct:
    
        follow_img = np.array(sct.grab(image_pos))[:,:,:3]
    
        cv2.imshow('follow_img', follow_img)
        cv2.waitKey()
        words = pytesseract.image_to_string(follow_img)
        #if you want to use korean here, change the upper line to 
        #words = pytesseract.image_to_string(follow_img, lang = "kor")

    print(words)

elif int(user_input) == 2:
    lx =0
    ly = 0
    email_from = 'Notification board updated@noreply.com'
    email_to = "your email here" # your email here!
    email_subject = "Notification board updated!"
    email_content = "Notification board changed!"

    msg = EmailMessage()
    msg.set_content(email_content)
    
    msg['From'] = email_from
    msg['To'] = email_to
    msg['Subject'] = email_subject

    print("please click the refresh button")
    mouse_click = 100
    prev_words =""

    while mouse_click >= 0 :
        time.sleep(0.1)
        mouse_click = win32api.GetKeyState(0x01)
        lx, ly = pag.position()
        location = 'X : ' + str(lx) + ' Y : ' + str(ly)
        if mouse_click< 0:
            print(location)
            break
    
    get_location_image()
    image_pos = { 'left' : x1 , 'top' : y1 , 'width' : (x2 - x1) , 'height' : (y2 - y1)}
    words = ""
    

    with mss.mss() as sct:
    
        follow_img = np.array(sct.grab(image_pos))[:,:,:3]
    
        cv2.imwrite('image.png', follow_img)
        cv2.imshow('follow_img', follow_img)
        
        cv2.waitKey()

         
        words = pytesseract.image_to_string(follow_img, lang = 'kor+eng')
        msg.set_content(words)
        

    print(words)

    while True:
        if words != prev_words:
            print("send mail!")
            prev_words = words
            msg.set_content(words)
            with open ('image.png', 'rb') as fp:
                img_data = fp.read()
                msg.add_attachment(img_data, maintype = 'image', subtype =imghdr.what(None, img_data), filename = 'image.png')
                
            
            smtp = smtplib.SMTP('smtp.gmail.com', 587)
            smtp.starttls()
            smtp.login('your email here' , 'your app password here') #your email , your app password
            smtp.send_message(msg)
            msg.clear_content()

            smtp.quit()
            print("success to quit")
        else :
            prev_words = words

        time.sleep(5)

        pag.moveTo(lx,ly,duration=0)
        pag.mouseDown(lx,ly,button= 'left')
        pag.mouseUp(lx,ly,button= 'left')
        pag.moveTo(lx,ly+100,duration = 0)
        time.sleep(1)
        
        pag.scroll(-5000)
        
        with mss.mss() as sct:
            follow_img = np.array(sct.grab(image_pos))[:,:,:3]
            words = pytesseract.image_to_string(follow_img, lang = 'kor+eng')
        cv2.imwrite('image.png', follow_img)
        #msg.set_content(words)
            



else :
    print("Wrong input!")