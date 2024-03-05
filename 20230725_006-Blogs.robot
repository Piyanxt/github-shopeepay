*** Settings ***
Library        Selenium2Library
Library        ExcelRobot
Library        stringUtils.py
Library        BuiltIn
Library        String

*** Variables ***
${URL}                         http://192.168.100.251:4006/
${BROWSER}                     gc
${excel}                       C:/Users/Piyanut/Development/Robot-Script/2023/excel/20230725_006-Vitanature
${language}          
${expected_result}       
${result}
${th}            
${en}
${row}
${column}
${title}
${desciption}
${header}
${detail}
${index}    
${excel_type}







*** Keywords ***
Open Website
        Open Browser          ${url}    ${browser}   options=add_experimental_option("detach", True)
        Maximize Browser Window
        Set selenium speed    0.5 seconds
    
Read Excel
       [Arguments]     ${excel_type}              
        Open Excel     ${excel}.${excel_type} 
        ${CountRow}=    Get Row Count    TC5
        
        FOR    ${index}    IN RANGE    6    ${CountRow} -2
        Open Website 
            ${A1}    Read Cell Data By Name    TC5    A${index}         
            Set Global Variable    ${row}              ${A1}   
            ${B1}    Read Cell Data By Name    TC5    B${index}
            Set Global Variable    ${column}            ${B1}
            ${C1}    Read Cell Data By Name    TC5    C${index}
            Set Global Variable    ${title}             ${C1}
            ${D1}    Read Cell Data By Name    TC5    D${index}
            Set Global Variable    ${desciption}        ${D1}
            ${E1}    Read Cell Data By Name    TC5    E${index}
            Set Global Variable    ${header}            ${E1}
            ${F1}    Read Cell Data By Name    TC5    F${index}
            Set Global Variable    ${detail}            ${F1}
            ${G1}    Read Cell Data By Name    TC5    G${index}
            Set Global Variable    ${language}   ${G1}
            ${H1}    Read Cell Data By Name    TC5    H${index}
            Set Global Variable    ${expected_result}   ${H1}

            Click Element    xpath=/html/body/div[1]/div/div/div[1]/div[2]/div[1]/div/div[2]/div/div/span[3]
            Blogs 
        END

           
            

Language

        Sleep    3s

        Click Element     xpath=/html/body/div[1]/div/div/div[1]/div[2]/div[1]/div/div[1]/div/div[2]/div[1]/div/div/div/div/div
        Click Element     xpath=/html/body/div[1]/div/div/div[1]/div[2]/div[1]/div/div[1]/div/div[2]/div[1]/div/div/div/div/div/div


Blogs
        Run Keyword If     '${language}' == 'EN'
        ...    Language

        execute javascript    window.scrollTo(0, 570)

        ${title}=   Replace String     ${title}    ${SPACE}        ${EMPTY}
        ${title}=   Get Text     xpath=/html/body/div/div/div/div[1]/div[6]/div[2]/div/div/div[2]/div/div/h3









        Click Element    xpath=/html/body/div[1]/div/div/div[1]/div[3]/div/div[3]/div[${column}]/div[2]/div[2]/div[2]





Blogs Detial

        Click Element    xpath=/html/body/div[1]/div/div/div[1]/div[3]/div/div/div[2]/img[1]

        ${link_fb}    Run Keyword And Return Status    Switch Window         Vitanatureplus | Facebook
        Run Keyword If    '${link_fb}' == 'True'
    ...    Change Text     ${link_fb}    J${index}    ${excel}    ${excel_type}  

        Switch Window    MAIN    
        
        Click Element    xpath=/html/body/div[1]/div/div/div[1]/div[3]/div/div/div[2]/img[2]


        ${link_ig}    Run Keyword And Return Status    Switch Window         Page couldn't load • Instagram  # Vitanature+ (@vitanatureplus) • Instagram photos and videos
        Run Keyword If    '${link_ig}' == 'True'
        ...    Change Text     ${link_ig}    K${index}    ${excel}    ${excel_type}

        Switch Window    MAIN

        Click Element    xpath=/html/body/div[1]/div/div/div[1]/div[3]/div/div/div[2]/img[3]
        
        ${link_ln}    Run Keyword And Return Status    Switch Window         LINE Add Friend
        Run Keyword If    '${link_ln}' == 'True'
        ...    Change Text     ${link_ln}    L${index}    ${excel}    ${excel_type}
    
        Switch Window    MAIN





Change Text  
    [Arguments]    ${check}    ${index}    ${excel}    ${excel_type}    
    Run Keyword If    '${check}' == 'True'
    ...            editFile    ${index}    TC4    ${excel}.${excel_type}    PASSED    alignment_value=center    font_color=1E8449        
    ...    ELSE    editFile    ${index}    TC4    ${excel}.${excel_type}    FAILED    alignment_value=center    font_color=FF0000    








*** Test Cases ***
TC#1
    Read Excel    excel_type=xlsx
