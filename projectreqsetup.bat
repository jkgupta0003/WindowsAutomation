REM Installing pywinauto dependencies
pip install -U pywinauto

REM Installing pytest dependencies
pip install pytest

REM Installing Pillow dependencies for image capture and screenshots in python
pip install --no-index -f https://dist.plone.org/thirdparty/ -U PIL
pip install Pillow

REM MSelenium webdriver installation
pip install -U selenium

REM Install Pynput
pip install pynput

REM Installing webdriver manager dependencies
REM pip install webdriver_manager


REM Installing pytest-allure dependencies
pip install allure-pytest

REM Install pyperclip
pip install pyperclip

REM Install win32 api library
pip install pywin32

REM Install python docx
pip install python-docx

REM install pip install pytest-html for HTML report
pip install pytest-html

REM intstall pip install html-testRunner
pip install html-testRunner

echo "Required Setup dependencies have been done"
