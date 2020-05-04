SET projectDir=%CD%



cd %projectDir%\report
DEL /F/S/Q *.* > NUL
RMDIR /S/Q allure-report\
cd ..
cd %projectDir%\testcases
echo "**********Test case execution start using pytest*********"
pytest -v test_msoffice.py --alluredir "%projectDir%\report"


echo "*********Allure Report Generation***************"
cd %projectDir%\report
allure generate --clean "%projectDir%\report"	

