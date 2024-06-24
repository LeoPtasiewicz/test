@echo off

:: Create virtual environment
python -m venv venv

:: Activate virtual environment
call venv\Scripts\activate

:: Install required packages using python -m pip to avoid PATH issues
python -m pip install pandas undetected-chromedriver selenium python-dotenv openpyxl

:: Deactivate virtual environment
deactivate
