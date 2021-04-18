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
  echo creating virtual env
  python -m venv %systemdrive%%homepath%\%penv%
)

echo activating virtual env
call %systemdrive%%homepath%\%penv%\Scripts\activate.bat

echo install required packages
pip install wheel
pip install -r requirements.txt 

call deactivate

rem if called from main script, do not output info
if %1x==x (
  echo.
  echo virtual Environment fertig. Nun bitte 'generate.cmd' aufrufen
  echo.
  timeout /t 15
)


