@echo off
pyinstaller -i icon.ico -F -w --hidden-import PyQt5.sip pdf2img.py
pause