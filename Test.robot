*** Settings ***
Library        Selenium2Library
Library        ExcelLibrary
Library        stringUtils.py
Library        BuiltIn
Library        String

*** Variables ***
${URL}                         https://www.google.com/
${BROWSER}                     gc
${excel}                       C:/Users/Piyanut/Development/Robot-Script/2023/excel/20230725_006-Vitanature
${ANDROID_HOME}
${JAVA_HOME}
${language}


*** Keywords ***
Open Website
    Open Browser    ${URL}    ${BROWSER}    options=add_experimental_option("detach", True)
    Maximize Browser Window
    Set selenium speed    1 seconds

    
    


*** Test Cases ***
TC#1
    Open Website

