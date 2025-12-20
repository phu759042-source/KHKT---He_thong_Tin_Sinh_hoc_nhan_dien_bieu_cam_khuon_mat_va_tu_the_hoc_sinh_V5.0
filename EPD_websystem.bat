pyinstaller ^
  --windowed ^
  --icon="AI Smart Monitor.ico" ^
  --name "AISmartMonitor" ^
  --add-data "AI Smart Monitor.ico;." ^
  --hidden-import=PyQt6.QtWidgets ^
  --hidden-import=PyQt6.QtGui ^
  --hidden-import=PyQt6.QtCore ^
  --hidden-import=PyQt6.QtMultimedia ^
  --collect-all PyQt6 ^
  --collect-all mediapipe ^
  --collect-all cv2 ^
  client.py
