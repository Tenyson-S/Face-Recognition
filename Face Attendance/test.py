from sklearn.neighbors import KNeighborsClassifier
import cv2
import numpy as np
import os
import pickle
import csv 
import time
from datetime import datetime

from win32com.client import Dispatch

def speak(str1):
    speak=Dispatch(("SAPI.spVoice"))
    speak.speak(str1)

Data=cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')
video=cv2.VideoCapture(0)



with open('data/names.pkl','rb') as w:
    LABELS=pickle.load(w)
with open('data/face_data.pkl','rb') as f:
    FACES=pickle.load(f) 

print("Shape of face matrix :",FACES.shape)

COL_NAMES=['NAMES','TIME']

model=KNeighborsClassifier(n_neighbors=5)
model.fit(FACES,LABELS)

while True:
    ret,frame=video.read()
    if ret==True:
        gray_image=cv2.cvtColor(frame,cv2.COLOR_BGR2GRAY)
        faces=Data.detectMultiScale(gray_image,1.3,5)
        for x,y,w,h in faces:
            crop_img=frame[y:y+h,x:x+w,:]
            resized_img=cv2.resize(crop_img,(50,50)).flatten().reshape(1,-1)
            prediction=model.predict(resized_img)
            ts=time.time()
            date=datetime.fromtimestamp(ts).strftime("%d-%m-%y")
            timestamp=datetime.fromtimestamp(ts).strftime("%H:&M:%S")
            exist=os.path.isfile("Attendance/Attendance_"+date+".csv")
            cv2.putText(frame,(str(prediction[0])),(x,y-15),cv2.FONT_HERSHEY_COMPLEX,1,(255,255,255),2)
            cv2.rectangle(frame,(x,y),(x+w,y+h),(50,50,255),1)
            attendance=[str(prediction[0]),str(timestamp)]
        cv2.imshow('Frame',frame)
        key=cv2.waitKey(1)
        if key==ord('o'):
            speak("Yes Present")
            time.sleep(2)
            if exist:
                with open("Attendance/Attendance_"+date+".csv","+a") as csvFile:
                    writer=csv.writer(csvFile)
                    writer.writerow(attendance)
                csvFile.close()
            else:
                with open("Attendance/Attendance_"+date+".csv","+a") as csvFile:
                    writer=csv.writer(csvFile)
                    writer.writerow(COL_NAMES)
                    writer.writerow(attendance)
                csvFile.close()         
        if (key==65 or key==97):
            break
video.release()
cv2.destroyAllWindows()