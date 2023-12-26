import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2
import time

app = win32com.client.Dispatch("PowerPoint.Application")
presentation = app.Presentations.Open(FileName=u'C:\Users\Chen\Documents\D\code\python\AI_mediapipe_use\ppp.pptx', ReadOnly=1)
presentation.SlideShowSettings.Run()

delay_find = 0
flag = 0

detector = HandDetector(detectionCon=0.5, maxHands=1)
cap = cv2.VideoCapture(0)

while cap.isOpened():
    success, img = cap.read()
    hands, img = detector.findHands(img)

    if flag == 0:
        if hands:
            hand = hands
            bbox = hand
            fingers = detector.fingersUp(hand)
            totalFingers = fingers.count(1)
            msg = "None"
            if totalFingers == 1:
                presentation.SlideShowWindow.View.Next()
            cv2.putText(img, msg, (bbox + 200, bbox - 30), cv2.FONT_HERSHEY_PLAIN, 2, (0, 255, 0), 2)
            flag = 1
            delay_find = 0

    cv2.imshow("Image", img)

    delay_find = delay_find + 1
    if delay_find > 30:
        flag = 0

    if cv2.waitKey(1) & 0xFF == ord("q"):
        presentation.SlideShowWindow.View.Exit()
        app.Quit()
        break

cap.release()
cv2.destroyAllWindows()