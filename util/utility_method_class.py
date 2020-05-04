'''
Created on Mar 12, 2020

@author: Gupta Jitendra
'''
from pathlib import Path
import os
import logging
import datetime
import time

projectpath = Path(__file__).resolve().parent.parent


def projectDir():
    try:
        # print(projectpath)
        return projectpath
    except FileNotFoundError as e:
        print(e)


def logFilePath():
    logfilepath = os.path.join(projectDir(), 'logs\\application.log')
    try:
        return logfilepath
    except FileNotFoundError as e:
        print(e)


def testDataPath():
    test_data_file_path = os.path.join(projectDir(), 'testdata\\Book1.xlsx')
    try:
        return test_data_file_path
    except FileNotFoundError as e:
        print(e)


def customLog():
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s:%(name)s:%(message)s')
    # print('Path of the log file ------>', relative_path_of_project.logFilePath())
    file_handler = logging.FileHandler(logFilePath())
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    return logger


def screenshot_file_Name_and_path(name):
    currentdateAndTime = datetime.datetime.now().strftime('%d-%m-%y %H-%M-%S')
    print('Current date and time---->', currentdateAndTime)
    screenshotDir = os.path.join(projectDir(), 'screenshots\\')
    screenShotFileName = str(screenshotDir) + str(name) + '_' + str(currentdateAndTime) + '.png'
    return screenShotFileName


def customExplicitWaitForDesktop(max_wait_time, actual_msg, expected_msg):
    for i in range(0, max_wait_time):
        print(i)
        if actual_msg in expected_msg:
            print("Expected message connection message" + expected_msg)
            break
        else:
            time.sleep(5)
            print("VPN connection is failed    " + expected_msg)
