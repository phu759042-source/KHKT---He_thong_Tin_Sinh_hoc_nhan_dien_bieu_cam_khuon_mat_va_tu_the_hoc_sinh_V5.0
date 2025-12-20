pyinstaller ^
  --windowed ^
  --name="he_thong_tin_sinh_hoc" ^
  --distpath ".\dist\He thong Tin - Sinh hoc nhan dien bieu cam va tu the hoc sinh" ^
  --icon="Emotion + Posture Detector v5.0.ico" ^
  --add-data "haarcascade_frontalface_default.xml;." ^
  --add-data "emotion_detection.h5;." ^
  --add-data "Emotion + Posture Detector v5.0.ico;." ^
  --add-data "static;static" ^
  --add-data "Emotion + Posture Detector v3.0 Camera.ico;." ^
  --add-data "Emotion + Posture Detector v3.0 Fullscreen Capture.ico;." ^
  --add-data "ARIALBD 1.ttf;." ^
  --add-data "C:\Users\ADMIN\AppData\Local\Programs\Python\Python310\Lib\site-packages\mediapipe\modules;mediapipe/modules" ^
  emotion_posture_detector_v5.0.py
