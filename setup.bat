@echo off

:: Create virtual environment
python -m venv venv

:: Activate virtual environment
call venv\Scripts\activate

:: Install required packages
pip install pandas undetected-chromedriver selenium python-dotenv

:: Deactivate virtual environment
deactivate
