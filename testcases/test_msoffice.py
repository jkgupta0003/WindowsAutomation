"""
Created on Mar 20, 2020

@author: Gupta Jitendra
"""

from util import utility_method_class
# from util import config_file_reader
# import unittest
import pytest
import time
import allure
import os
# from allure_commons.types import AttachmentType
# from util import constantvariables
# import subprocess
from pywinauto.timings import always_wait_until_passes
from util import constantvariables
# from pages.egnyte import Egnyte
from app.msword import MSWord

word = MSWord()
word.maximize_app_window()


# @pytest.mark.skip(reason='No need to run this test case')
# @allure.step('1.1_1. Verify laptop is not connected to BCG network')
@pytest.mark.run(order=1)
@always_wait_until_passes(constantvariables.maximum_wait_time_for_windows_element,
                          constantvariables.retrival_wait_time_for_windows_element)
def test_wordHello():
    # subprocess.call("taskkill /IM WINWORD.EXE /F", shell=True)
    with allure.step("TC_ID_001: Launch Word to verify if the user is able to write"):
        try:
            word.typeHello()
            utility_method_class.customLog().info(
                "Verify that Hello is written in word ")
        except AssertionError as e:
            utility_method_class.customLog().error('unable to write hello' + e)
            raise e


@pytest.mark.run(order=2)
@always_wait_until_passes(constantvariables.maximum_wait_time_for_windows_element,
                          constantvariables.retrival_wait_time_for_windows_element)
def test_typeBold():
    with allure.step("TC_ID_002: Launch Word to verify if the bold and Save functionality is working "):
        try:
            word.typeBold()
            word.saveWord()
            word.closeFile()
            utility_method_class.customLog().info(
                "Verify that Hello is written in bold in word ")
        except AssertionError as e:
            utility_method_class.customLog().error('unable to write bold hello' + e)
            raise e
