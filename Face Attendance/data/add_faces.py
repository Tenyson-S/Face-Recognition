import cv2
import numpy as np
import os
import pickle
Data=cv2.CascadeClassifier('haarcascade_frontalface_default.xml')
video=cv2.VideoCapture(0)

faces_data=[]
i=0
name=input("Enter your name :")

while True:
    ret,frame=video.read()
    if ret==True:
        gray_image=cv2.cvtColor(frame,cv2.COLOR_BGR2GRAY)
        faces=Data.detectMultiScale(gray_image,1.3,5)
        for x,y,w,h in faces:
            crop_img=frame[y:y+h,x:x+w,:]
            resized_img=cv2.resize(crop_img,(50,50))
            if len(faces_data)<=100 and i%10==0:
                faces_data.append(resized_img)
            cv2.putText(frame,str(len(faces_data)),(50,50), cv2.FONT_HERSHEY_COMPLEX, 2, (50,50,255))
            cv2.rectangle(frame,(x,y),(x+w,y+h),(50,50,255),2)
        cv2.imshow('video',frame)
        key=cv2.waitKey(1)
        if (key==65 or key==97) or len(faces_data)==100:
            break
video.release()
cv2.destroyAllWindows()

faces_data=np.asarray(faces_data).reshape(100,-1)

if 'names.pkl' not in os.listdir('data/'):
    names=[name]*100
    with open('data/names.pkl','wb') as f:
        pickle.dump(names,f)
else:
    with open('data/names.pkl','rb') as f:
        names=pickle.load(f)
    names=names+[name]*100
    with open('data/names.pkl','wb') as f:
        pickle.dump(names,f)
        

if 'face_data.pkl' not in os.listdir('data/'):
    with open('data/face_data.pkl','wb') as f:
        pickle.dump(faces_data,f)
else:
    with open('data/face_data.pkl','rb') as f:
        faces=pickle.load(f)
    faces=np.append(faces,faces_data,axis=0)
    with open('data/face_data.pkl','wb') as f:
        pickle.dump(faces,f)