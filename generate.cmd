@echo off
rem
rem prerequisite:
rem   1. download Python 3 for Windows
rem   https://www.python.org/downloads/windows/
rem
rem ver 0.5
rem 
rem Author:      <plix1014@gmail.com>
rem
rem Created:     17.04.2021
rem Copyright:   (c) 2021
rem Licence:     CC BY-NC-SA http://creativecommons.org/licenses/by-nc-sa/4.0/
rem -------------------------------------------------------------------------------

rem virtual env dir
set penv=py_icu_env

rem run prepare if it does not exists
if not exist %systemdrive%%homepath%\%penv% (
  call prepare_icu_venv.cmd %~n0
)


echo activating virtual env
call %systemdrive%%homepath%\%penv%\Scripts\activate.bat


python at_hosp_csv2excel.py

echo.
echo 'AT_Hospitalisierung.xlsx' fuer Excel/Libre Office fertig
echo.
timeout /t 15

call deactivate


