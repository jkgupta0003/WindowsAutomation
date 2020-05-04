"""
Created on Mar 20, 2020

@author: Gupta Jitendra
"""

from pywinauto.application import Application
import time
import pyperclip
from pynput.keyboard import Controller, Key
from pywinauto.timings import always_wait_until_passes
import subprocess
from util import utility_method_class


# from app import msofficeUutility

class MSWord:
    subprocess.call("taskkill /IM WINWORD.EXE /F", shell=True)

    def __init__(self):
        self.app = Application(backend='uia').start(r'C:\\Program Files\\Microsoft Office\\root\\Office16\\winword.exe')
        self.main = self.app.window(title='Document1 - Word')
        # self.main = self.app.active()
        utility_method_class.customLog().info("Word application is launched and first document page is opened")
        time.sleep(10)

    def maximize_app_window(self):
        # Access app's window object
        self.main.maximize()
        utility_method_class.customLog().info("Word application is launched and Maximized")

    def typeHello(self):
        self.main.child_window(title="Page 1 content", auto_id="Body", control_type="Edit").type_keys("Hello Pywin")
        utility_method_class.customLog().info("Hello is written")

    def saveWord(self):
        self.main.child_window(title="Save", control_type="Button").click_input()
        time.sleep(4)
        # self.main.FileNameEdit.type_keys("Jitendra")
        pyperclip.copy("Jitendra")
        keyboard = Controller()
        keyboard.press(Key.ctrl)
        keyboard.press('v')
        keyboard.release('v')
        keyboard.release(Key.ctrl)
        self.main.OpenButton.click_input()
        self.main.PersonalListItem.click_input()
        self.main.SaveButton.click_input()
        time.sleep(2)
        #self.app.window(title_re='.*Jitendra.docx.*').CloseButton.click_input()

    def typeBold(self):
        keyboard = Controller()
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
        self.main.BoldButton.click_input()
        self.main.child_window(title="Page 1 content", auto_id="Body", control_type="Edit").type_keys(
            " Hello Bold Pywin")

    def closeFile(self):
        self.app.window(title_re='.*Jitendra.docx.*').CloseButton.click_input()