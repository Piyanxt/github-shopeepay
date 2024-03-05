*** Settings ***
Library        Selenium2Library
Library        ExcelRobot
Library        stringUtils.py
Library        BuiltIn
Library        String




*** Variables ***
${url}           http://192.168.100.251:4006/
${browser}       gc
${excel}         C:/Users/Piyanut/Development/Robot-Script/2023/excel/20230725_006-Vitanature
${case}          
${th}            
${en}
${language}           
${expected_result}       
${result}

${th_in_space}
${th_output_space}

${th_string}
${th_output_stting}
${th_output_sytem_space}

${en_in_space}
${en_output_space}

${en_string}
${en_output_stting}
${en_output_sytem_space}  
      


    



*** Keywords ***
Open Website
    Open Browser          ${url}    ${browser}   options=add_experimental_option("detach", True)
    Maximize Browser Window
    Set selenium speed    0.5 seconds
    
Read Excel
    [Arguments]    ${excel_type}               
        Open Excel     ${excel}.${excel_type} 
        ${CountRow}=    Get Row Count    TC3
        FOR    ${index}    IN RANGE    6    ${CountRow} -2
    Open Website
            ${A1}    Read Cell Data By Name    TC3    A${index}         
            Set Global Variable    ${th}              ${A1}   
            ${B1}    Read Cell Data By Name    TC3    B${index}
            Set Global Variable    ${en}              ${B1}
            ${C1}    Read Cell Data By Name    TC3    C${index}
            Set Global Variable    ${th_in_space}     ${C1}
            ${D1}    Read Cell Data By Name    TC3    D${index}
            Set Global Variable    ${th_output_space}  ${D1}
            ${E1}    Read Cell Data By Name    TC3    E${index}
            Set Global Variable    ${en_in_space}      ${E1}
            ${F1}    Read Cell Data By Name    TC3    F${index}
            Set Global Variable    ${en_output_space}   ${F1}
            ${G1}    Read Cell Data By Name    TC3    G${index}
            Set Global Variable    ${expected_result}   ${G1}
           
            Sleep    2s
            Click Element         xpath=/html/body/div[1]/div/div/div[1]/div[2]/div[1]/div/div[2]/div/div/span[4]
            Contact Us TH         ${th}    ${th_in_space}    ${th_output_space}    ${index}    ${excel_type}
            Contact Us EN         ${en}    ${en_in_space}    ${en_output_space}    ${index}    ${excel_type}
            
        END

Contact Us TH

    [Arguments]     ${th}    ${th_in_space}    ${th_output_space}    ${index}    ${excel_type}
    
    execute javascript    window.scrollTo(0, -document.body.scrollHeight)

    #ส่งค่าเข้าไปเพื่อทำการตัดบรรทัด
    ${th_string_1}=   Replace String     ${th}    \n    ${EMPTY} 

    #ส่งค่าเข้าไปเพื่อเป็นการตัดช่องว่างของข้อความ    
    ${th_string_2}=   Replace String     ${th_string_1}    ${SPACE}    ${EMPTY}   
        Set Test Variable    ${th_string_2}
        Return th string     ${index}    ${excel_type}    ${th_string_2}
     
    Sleep    3s

    #ดึงข้อมูลข้อความออกมาเพื่อนดูว่าตรงกับที่ส่งเข้าไปหรือไม่
    ${th_output_sytem_space}=    Get Text    xpath=/html/body/div[1]/div/div/div[1]/div[3]/div/div/div[1]

     #ส่งค่าด้านหลังไปแทนที่ค่าข้างหน้า
    ${th_output_string1}=   Replace String     ${th_output_sytem_space}    \n    ${EMPTY} 
    

    #ส่งค่าเข้าไปเพื่อเป็นการตัดช่องว่างของข้อความ
    ${th_output_string2}=   Replace String     ${th_output_string1}    ${SPACE}    ${EMPTY}  
        Return th output system      ${index}    ${excel_type}    ${th_output_string2}
    
    #เช็คความถูกต้องของ Callum C และ D ว่าตรงกันหรือไม่
    ${compate_th_c1}    Run Keyword And Return Status    Should Be Equal    ${th_string_2}    ${th_output_string2}

    #แปลง boolean เป็น string  
    ${compate_th_c1}     Convert To String    ${compate_th_c1}
        Set Test Variable    ${compate_th_c1}
    
    #เช็คความถูกต้องของการเปรียบระหว่าง C และ D แล้วนำผลมาเช็คกับ Callum G  เพื่อที่จะเอาผลที่ถูกต้องไปเก็บไว้ที่ Callun H
    ${check_compate_c1}    Run Keyword And Return Status    Should Be Equal    ${compate_th_c1}    ${expected_result}
    Log To Console         ${check_compate_c1}       
        Run Keyword If    '${check_compate_c1}' == 'True' or '${check_compate_c1}' == 'False'
    ...    Change Text     ${check_compate_c1}    H${index}    ${excel}    ${excel_type}
    # Check Social    ${index}     ${excel_type}  

Contact Us EN

    [Arguments]     ${en}    ${en_in_space}    ${en_output_space}   ${index}    ${excel_type}

    Sleep    3s

    Click Element    xpath=/html/body/div[1]/div/div/div[1]/div[2]/div[1]/div/div[1]/div/div[2]/div[1]/div/div/div/div/div/button/div/span
    Click Element    xpath=/html/body/div[1]/div/div/div[1]/div[2]/div[1]/div/div[1]/div/div[2]/div[1]/div/div/div/div/div/div/button


    execute javascript    window.scrollTo(0, -document.body.scrollHeight)

    #ส่งค่าด้านหลังไปแทนที่ค่าข้างหน้า
    ${en_string_1}=   Replace String     ${en}    \n          ${EMPTY} 

    #ส่งค่าเข้าไปเพื่อเป็นการตัดช่องว่างของข้อความ    
    ${en_string_2}=   Replace String     ${en_string_1}    ${SPACE}    ${EMPTY}   
        Set Test Variable    ${en_string_2}
        Return en string     ${index}    ${excel_type}    ${en_string_2}
     
    Sleep    2s
    #ดึงข้อมูลข้อความออกมาเพื่อนดูว่าตรงกับที่ส่งเข้าไปหรือไม่
    ${en_output_sytem_space}=    Get Text    xpath=/html/body/div[1]/div/div/div[1]/div[3]/div/div/div[1]

     #ส่งค่าด้านหลังไปแทนที่ค่าข้างหน้า
    ${en_output_string1}=   Replace String     ${en_output_sytem_space}    \n        ${EMPTY} 

    #ส่งค่าเข้าไปเพื่อเป็นการตัดช่องว่างของข้อความ
    ${en_output_string2}=   Replace String     ${en_output_string1}    ${SPACE}    ${EMPTY}  
        Return en output system      ${index}    ${excel_type}    ${en_output_string2}

    #เช็คความถูกต้องของ Callum C และ D ว่าตรงกันหรือไม่
    ${compate_en_c1}    Run Keyword And Return Status    Should Be Equal    ${en_string_2}    ${en_output_string2}
    ${compate_en_c1}     Convert To String    ${compate_en_c1}
        Set Test Variable    ${compate_en_c1}  

    #เช็คความถูกต้องของการเปรียบระหว่าง C และ D แล้วนำผลมาเช็คกับ Callum G  เพื่อที่จะเอาผลที่ถูกต้องไปเก็บไว้ที่ Callun H
    ${check_compate_c1}    Run Keyword And Return Status    Should Be Equal    ${compate_en_c1}    ${expected_result}
    Log To Console     ${check_compate_c1}    
        Run Keyword If    '${check_compate_c1}' == 'True' or '${check_compate_c1}' == 'False'
    ...    Change Text     ${check_compate_c1}    I${index}    ${excel}    ${excel_type}
    # Check Social    ${index}     ${excel_type}   


# Check Social

#     [Arguments]        ${index}     ${excel_type}

#     execute javascript    window.scrollTo(0, 500) 

#     Click Element    xpath=/html/body/div[1]/div/div/div[2]/a[1]/div

#     ${link_fb}    Run Keyword And Return Status    Switch Window         Vitanatureplus | Facebook
#     Run Keyword If    '${link_fb}' == 'True'
#    ...    Change Text     ${link_fb}    J${index}    ${excel}    ${excel_type}  

#     Switch Window    MAIN    
    
#     Click Element    xpath=/html/body/div[1]/div/div/div[2]/a[2]/div
    

#     ${link_ig}    Run Keyword And Return Status    Switch Window         Page couldn't load • Instagram  # Vitanature+ (@vitanatureplus) • Instagram photos and videos
#     Run Keyword If    '${link_ig}' == 'True'
#     ...    Change Text     ${link_ig}    K${index}    ${excel}    ${excel_type}

#     Switch Window    MAIN

#     Click Element    xpath=/html/body/div[1]/div/div/div[2]/a[3]/div
    
#     ${link_ln}    Run Keyword And Return Status    Switch Window         LINE Add Friend
#     Run Keyword If    '${link_ln}' == 'True'
#     ...    Change Text     ${link_ln}    L${index}    ${excel}    ${excel_type}
   
#    Switch Window    MAIN



Return th string 
    [Arguments]    ${index}    ${excel_type}    ${th_string_2}
    editFile           C${index}     TC3     ${excel}.${excel_type}     ${th_string_2}


Return th output system 
    [Arguments]    ${index}    ${excel_type}    ${th_output_string2}
    editFile           D${index}     TC3     ${excel}.${excel_type}     ${th_output_string2}


Return en string 
    [Arguments]    ${index}    ${excel_type}    ${en_string_2}
    editFile           E${index}     TC3     ${excel}.${excel_type}     ${en_string_2}


Return en output system 
    [Arguments]    ${index}    ${excel_type}    ${en_output_string2}
    editFile           F${index}     TC3     ${excel}.${excel_type}     ${en_output_string2}


Change Text  
    [Arguments]    ${check}    ${index}    ${excel}    ${excel_type}    
    Run Keyword If    '${check}' == 'True'
    ...            editFile    ${index}    TC3    ${excel}.${excel_type}    PASSED    alignment_value=center     font_color=1E8449      
    ...    ELSE    editFile    ${index}    TC3    ${excel}.${excel_type}    FAILED    alignment_value=center     font_color=FF0000    



*** Test Cases ***
TC #1
    Read Excel    excel_type=xlsx
                                                        
       



   








