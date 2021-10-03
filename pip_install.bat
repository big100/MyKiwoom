chcp 65001
pip install psutil pyqt5 pandas pyttsx3 pyqtgraph matplotlib BeautifulSoup4 lxml python-telegram-bot
@echo off
@echo ====================================================================================
@echo 다음 두개의 파일내에서 디렉토리 경로를 수정해야 프로그램이 올바르게 작동합니다.
@echo ====================================================================================
@echo utility/setting.py
@echo openapi_path = 'D:/OpenAPI'
@echo system_path = 'D:/PythonProjects/MyKiwoom'
@echo ====================================================================================
@echo mykiwoom.bat
@echo cd /D D:/PythonProjects/MyKiwoom
@echo ====================================================================================
pause