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
${check_text}
${check_position}
${product_title}
${no_product}
${position_product_title_row}
${position_product_title_col}
${desciption}
${price}
${title}
${administration}
${additional_instructions}
${certification standards}
${expected_result_title_home}
${expected_result_desciption_home}
${expected_result_price_home}
${expected_result_title_category}
${expected_result_desciption_category}
${expected_result_price_category}
${expected_result_title_detail}
${expected_result_desciption_detail}
${expected_result_price_detail}
${expected_result_administration}
${expected_result_additional_instructions}
${expected_result_certification_standards}
${result}
${language}

*** Keywords ***
Open Website
    Open Browser    ${URL}    ${BROWSER}    options=add_experimental_option("detach", True)
    Maximize Browser Window
    Set selenium speed    1 seconds
    
Language

    Sleep    3s
    Click Element     xpath=/html/body/div[1]/div/div/div[1]/div[2]/div[1]/div/div[1]/div/div[2]/div[1]/div/div/div/div/div
    Click Element    xpath=/html/body/div[1]/div/div/div[1]/div[2]/div[1]/div/div[1]/div/div[2]/div[1]/div/div/div/div/div/div

Read Excel
    [Arguments]    ${excel_type}    
    Open Excel    ${excel}.${excel_type} 
    ${CountRow}=    Get Row Count   TC2.3
    FOR    ${index}    IN RANGE    6    ${CountRow} - 2
        Open Website

        #ชื่อsheet,column,row

        ${A1}    Read Cell Data By Name   TC2.3    A${index}           
        Set Global Variable    ${check_text}                                      ${A1}
        ${B1}    Read Cell Data By Name   TC2.3    B${index}
        Set Global Variable    ${check_position}                                  ${B1}
        ${C1}    Read Cell Data By Name   TC2.3    C${index}
        Set Global Variable    ${product_title}                                   ${C1}
        ${D1}    Read Cell Data By Name   TC2.3    D${index}
        Set Global Variable    ${no_product}                                      ${D1}
        ${E1}    Read Cell Data By Name   TC2.3    E${index}
        Set Global Variable    ${position_product_title_row}                      ${E1}
        ${F1}    Read Cell Data By Name   TC2.3    F${index}
        Set Global Variable    ${position_product_title_col}                      ${F1}
        ${G1}    Read Cell Data By Name   TC2.3    G${index}
        Set Global Variable    ${desciption}                                      ${G1}
        ${H1}    Read Cell Data By Name   TC2.3    H${index}
        Set Global Variable    ${price}                                           ${H1}
        ${I1}    Read Cell Data By Name   TC2.3    I${index}
        Set Global Variable    ${title}                                           ${I1}
        ${J1}    Read Cell Data By Name   TC2.3    J${index}
        Set Global Variable    ${administration}                                  ${J1}
        ${K1}    Read Cell Data By Name   TC2.3    K${index}
        Set Global Variable    ${additional_instructions}                         ${K1}
        ${L1}    Read Cell Data By Name   TC2.3    L${index}
        Set Global Variable    ${certification_standards}                         ${L1}
        ${M1}    Read Cell Data By Name   TC2.3    M${index}
        Set Global Variable    ${language}                                        ${M1}
        ${N1}    Read Cell Data By Name   TC2.3    N${index}
        Set Global Variable    ${expected_result_title_home}                      ${N1}
        ${P1}    Read Cell Data By Name   TC2.3    P${index}
        Set Global Variable    ${expected_result_desciption_home}                 ${P1}
        ${R1}    Read Cell Data By Name   TC2.3    R${index}
        Set Global Variable    ${expected_result_price_home}                      ${R1}
        ${T1}    Read Cell Data By Name   TC2.3    T${index}
        Set Global Variable    ${expected_result_title_category}                  ${T1}
        ${V1}    Read Cell Data By Name   TC2.3    V${index}
        Set Global Variable    ${expected_result_desciption_category}             ${V1}
        ${X1}    Read Cell Data By Name   TC2.3    X${index}
        Set Global Variable    ${expected_result_price_category}                  ${X1}
        ${Z1}    Read Cell Data By Name   TC2.3    Z${index}
        Set Global Variable    ${expected_result_title_detail}                    ${Z1}
        ${AB1}    Read Cell Data By Name   TC2.3    AB${index}
        Set Global Variable    ${expected_result_desciption_detail}               ${AB1}
        ${AD1}    Read Cell Data By Name   TC2.3    AD${index}
        Set Global Variable    ${expected_result_price_detail}                    ${AD1}
        ${AF1}    Read Cell Data By Name   TC2.3    AF${index}
        Set Global Variable    ${expected_result_administration}                  ${AF1}
        ${AH1}    Read Cell Data By Name   TC2.3    AH${index}
        Set Global Variable    ${expected_result_additional_instructions}         ${AH1}
        ${AJ1}    Read Cell Data By Name   TC2.3    AJ${index}
        Set Global Variable    ${expected_result_certification_standards}         ${AJ1}
        Check Product In Home & Cate & Detail    ${check_position}    ${product_title}    ${no_product}    ${position_product_title_row}    ${position_product_title_col}    ${desciption}    ${price}    ${title}    ${administration}    ${additional_instructions}    ${certification standards}    ${index}    ${excel_type}
    END      

Check Product In Home & Cate & Detail
    [Arguments]    ${check_position}    ${product_title}    ${no_product}    ${position_product_title_row}    ${position_product_title_col}    ${desciption}    ${price}    ${title}    ${administration}    ${additional_instructions}    ${certification standards}    ${index}    ${excel_type}
    Run Keyword If     '${language}' == 'EN'
    ...    Language

    execute javascript    window.scrollTo(0, 2780)
    Click Element    xpath=/html/body/div/div/div/div[1]/div[6]/div[1]/div/div[${check_position}]

    Click Element    xpath=/html/body/div/div/div/div[1]/div[6]/div[2]/div/div/div[1]/div${position_product_title_row}/div/div${position_product_title_col}/div/img
    ${check1}   Run Keyword And Return Status    Element Should Be Visible    xpath=/html/body/div/div/div/div[1]/div[6]/div[2]/div/div/div[2]/div
    Run Keyword If   ${check1} == False
    ...    Click Element    xpath=/html/body/div/div/div/div[1]/div[6]/div[2]/div/div/div[1]/div${position_product_title_row}/div/div${position_product_title_col}/div/img
    
    #ส่งค่าเข้าไปเพื่อทำการตัดบรรทัด > home
    ${product_title_in_home_ws}=   Replace String     ${product_title}    ${SPACE}        ${EMPTY}
    ${product_desciption_ws}=      Replace String     ${desciption}       ${SPACE}        ${EMPTY}
    ${product_price_ws}=           Replace String     ${price}            ${SPACE}        ${EMPTY}

    #ดึงข้อมูลข้อความออกมาเพื่อนดูว่าตรงกับที่ส่งเข้าไปหรือไม่ > home
    ${position_product_title_row_in_home}=  Get Text     xpath=/html/body/div/div/div/div[1]/div[6]/div[2]/div/div/div[2]/div/div/h3
    ${position_product_des_in_home}=        Get Text     xpath=/html/body/div/div/div/div[1]/div[6]/div[2]/div/div/div[2]/div/div/span
    ${position_product_price_in_home}=      Get Text     xpath=/html/body/div/div/div/div[1]/div[6]/div[2]/div/div/div[2]/div/h6

    #ส่งค่าเข้าไปเพื่อเป็นการตัดช่องว่างของข้อความ > home
    ${position_product_title_row_in_home_cursting}=   Replace String    ${position_product_title_row_in_home}    ${SPACE}        ${EMPTY}
    ${position_product_des_in_home_cursting}=         Replace String    ${position_product_des_in_home}          ${SPACE}        ${EMPTY}
    ${position_product_price_in_home_cursting}=       Replace String    ${position_product_price_in_home}        ${SPACE}        ${EMPTY}
    
    
    #title > home
    ${Compate_product_title}    Run Keyword And Return Status    Should Be Equal    ${product_title_in_home_ws}    ${position_product_title_row_in_home_cursting}
    Run Keyword If      '${Compate_product_title}' == 'True' or '${Compate_product_title}' == 'False'
    ...    Change Boolean to Text    ${Compate_product_title}
    ${Compate_product_title_2}    Run Keyword And Return Status    Should Be Equal    ${result}    ${expected_result_title_home}
    Run Keyword If      '${Compate_product_title_2}' == 'True' or '${Compate_product_title_2}' == 'False'
    ...    Change Text     ${Compate_product_title_2}    O${index}    ${excel}    ${excel_type}


    #desciption > home
    ${Compate_product_des}    Run Keyword And Return Status    Should Be Equal    ${product_desciption_ws}    ${position_product_des_in_home_cursting}
    Run Keyword If      '${Compate_product_des}' == 'True' or '${Compate_product_des}' == 'False'
    ...    Change Boolean to Text    ${Compate_product_des}
    ${Compate_product_des_2}    Run Keyword And Return Status    Should Be Equal    ${result}    ${expected_result_desciption_home}
    Run Keyword If      '${Compate_product_des_2}' == 'True' or '${Compate_product_des_2}' == 'False'
    ...    Change Text     ${Compate_product_des_2}    Q${index}    ${excel}    ${excel_type}


    #price > home
    ${Compate_product_price}    Run Keyword And Return Status    Should Be Equal    ${product_price_ws}    ${position_product_price_in_home_cursting}
    Run Keyword If      '${Compate_product_price}' == 'True' or '${Compate_product_price}' == 'False'
    ...    Change Boolean to Text    ${Compate_product_price}
    ${Compate_product_price_2}    Run Keyword And Return Status    Should Be Equal    ${result}    ${expected_result_price_home}
    Run Keyword If      '${Compate_product_price_2}' == 'True' or '${Compate_product_price_2}' == 'False'
    ...    Change Text     ${Compate_product_price_2}    S${index}    ${excel}    ${excel_type}

    #menu
    Check Tab       ${check_position}    ${no_product}    ${product_title_in_home_ws}    ${product_desciption_ws}    ${product_price_ws}    ${index}                      ${excel_type}
    Open Product    ${no_product}        ${title}         ${product_price_ws}            ${product_desciption_ws}    ${administration}      ${additional_instructions}    ${certification standards}    ${index}    ${excel_type}


Check Tab

    [Arguments]    ${check_position}    ${no_product}    ${product_title_in_home_ws}    ${product_desciption_ws}    ${product_price_ws}    ${index}    ${excel_type}
    Sleep    3s
    Mouse Over     xpath=/html/body/div[1]/div/div/div[1]/div[2]/div[1]/div/div[2]/div/div/span[2]

    ${check_menu} =    Evaluate    ${check_position}+1

    Click Element      xpath=/html/body/div[1]/div/div/div[1]/div[2]/div[1]/div/div[2]/div/div/span[2]/div/a[${check_menu}]

    execute javascript    window.scrollTo(0, 500)   
    

    #ดึงข้อมูลข้อความออกมาเพื่อนดูว่าตรงกับที่ส่งเข้าไปหรือไม่ >    home
    ${Check_Product_Position_Title_Text}=    Get Text     xpath=/html/body/div[1]/div/div/div[1]/div[3]/div/div/div[2]/div/div[${no_product}]/div[2]/h2
    ${Check_Product_Position_Des_Text}=      Get Text     xpath=/html/body/div[1]/div/div/div[1]/div[3]/div/div/div[2]/div/div[${no_product}]/div[2]/div/p
    ${Check_Product_Position_Price_Text}=    Get Text     xpath=/html/body/div[1]/div/div/div[1]/div[3]/div/div/div[2]/div/div[${no_product}]/div[3]/div/div/h4
    
    #ส่งค่าเข้าไปเพื่อเป็นการตัดช่องว่างของข้อความ    >    home
    ${product_title_in_category_cursting}=   Replace String    ${Check_Product_Position_Title_Text}    ${SPACE}        ${EMPTY}
    ${product_des_in_category_cursting}=     Replace String    ${Check_Product_Position_Des_Text}      ${SPACE}        ${EMPTY}
    ${product_price_in_category_cursting}=   Replace String    ${Check_Product_Position_Price_Text}    ${SPACE}        ${EMPTY}


    #title > home
    ${Compate_product_title_in_category}    Run Keyword And Return Status    Should Be Equal    ${product_title_in_home_ws}    ${product_title_in_category_cursting}
    Run Keyword If      '${Compate_product_title_in_category}' == 'True' or '${Compate_product_title_in_category}' == 'False'
    ...    Change Boolean to Text    ${Compate_product_title_in_category}
    ${Compate_product_title_in_category_2}    Run Keyword And Return Status    Should Be Equal    ${result}    ${expected_result_title_category}
    Run Keyword If      '${Compate_product_title_in_category_2}' == 'True' or '${Compate_product_title_in_category_2}' == 'False'
    ...    Change Text     ${Compate_product_title_in_category_2}    U${index}    ${excel}    ${excel_type}

    #desciption > home   
    ${Compate_product_des_in_category}    Run Keyword And Return Status    Should Be Equal    ${product_desciption_ws}    ${product_des_in_category_cursting}
    Run Keyword If      '${Compate_product_des_in_category}' == 'True' or '${Compate_product_des_in_category}' == 'False'
    ...    Change Boolean to Text    ${Compate_product_des_in_category}
    ${Compate_product_des_in_category_2}    Run Keyword And Return Status    Should Be Equal    ${result}    ${expected_result_desciption_category}
    Run Keyword If      '${Compate_product_des_in_category_2}' == 'True' or '${Compate_product_des_in_category_2}' == 'False'
    ...    Change Text     ${Compate_product_des_in_category_2}    W${index}    ${excel}    ${excel_type}

    #price > home
    ${Compate_product_price_in_category}    Run Keyword And Return Status    Should Be Equal    ${product_price_ws}    ${product_price_in_category_cursting}
    Run Keyword If      '${Compate_product_price_in_category}' == 'True' or '${Compate_product_price_in_category}' == 'False'
    ...    Change Boolean to Text    ${Compate_product_price_in_category}
    ${Compate_product_price_in_category_2}    Run Keyword And Return Status    Should Be Equal    ${result}    ${expected_result_price_category}
    Run Keyword If      '${Compate_product_price_in_category_2}' == 'True' or '${Compate_product_price_in_category_2}' == 'False'
    ...    Change Text     ${Compate_product_price_in_category_2}    Y${index}    ${excel}    ${excel_type}


Open Product
    [Arguments]    ${no_product}    ${title}    ${product_price_ws}    ${product_desciption_ws}    ${administration}    ${additional_instructions}    ${certification standards}    ${index}    ${excel_type}

      Click Element    xpath=/html/body/div[1]/div/div/div[1]/div[3]/div/div[2]/div[2]/div/div[${no_product}]/div[1]
     
      Sleep    3s

     #title    > product detail
    ${product_title_in_product_detail_ws}=    Replace String    ${title}    ${SPACE}        ${EMPTY}
    ${title_product_in_product_detail}=     Get Text    xpath=/html/body/div[1]/div/div/div[1]/div[4]/div/div/div[2]/h1
    ${title_product_in_product_detail_cursting}=    Replace String    ${title_product_in_product_detail}    ${SPACE}        ${EMPTY}
    ${Compare_product_title_in_product_detail}    Run Keyword And Return Status    Should Be Equal    ${product_title_in_product_detail_ws}    ${title_product_in_product_detail_cursting}
    Run Keyword If      '${Compare_product_title_in_product_detail}' == 'True' or '${Compare_product_title_in_product_detail}' == 'False'
    ...    Change Boolean to Text    ${Compare_product_title_in_product_detail}
    ${Compare_product_title_in_product_detail_2}    Run Keyword And Return Status    Should Be Equal    ${result}    ${expected_result_title_detail}
    Run Keyword If      '${Compare_product_title_in_product_detail_2}' == 'True' or '${Compare_product_title_in_product_detail_2}' == 'False'
    ...    Change Text     ${Compare_product_title_in_product_detail_2}    AA${index}    ${excel}    ${excel_type}

    #price    > product detail
    ${price_product_in_product_detail}=    Get Text    xpath=/html/body/div[1]/div/div/div[1]/div[4]/div/div/div[2]/h4
    ${price_product_in_product_detail_cursting}=    Replace String    ${price_product_in_product_detail}    ${SPACE}        ${EMPTY}
    ${Compare_product_price_in_product_detail}    Run Keyword And Return Status    Should Be Equal    ${product_price_ws}    ${price_product_in_product_detail_cursting}
    Run Keyword If      '${Compare_product_price_in_product_detail}' == 'True' or '${Compare_product_price_in_product_detail}' == 'False'
    ...    Change Boolean to Text    ${Compare_product_price_in_product_detail}
    ${Compare_product_price_in_product_detail_2}    Run Keyword And Return Status    Should Be Equal    ${result}    ${expected_result_price_detail}
    Run Keyword If      '${Compare_product_price_in_product_detail_2}' == 'True' or '${Compare_product_price_in_product_detail_2}' == 'False'
    ...    Change Text     ${Compare_product_price_in_product_detail_2}    AE${index}    ${excel}    ${excel_type}

    #desciption    > product detail
    ${des_product_in_product_detail}=    Get Text    xpath=/html/body/div[1]/div/div/div[1]/div[4]/div/div/div[2]/div[4]
    ${des_product_in_product_detail_cursting}=    Replace String    ${des_product_in_product_detail}    ${SPACE}        ${EMPTY}
    ${Compare_product_des_in_product_detail}    Run Keyword And Return Status    Should Be Equal    ${product_desciption_ws}    ${des_product_in_product_detail_cursting}
    Run Keyword If      '${Compare_product_des_in_product_detail}' == 'True' or '${Compare_product_des_in_product_detail}' == 'False'
    ...    Change Boolean to Text    ${Compare_product_des_in_product_detail}
    ${Compare_product_des_in_product_detail_2}    Run Keyword And Return Status    Should Be Equal    ${result}    ${expected_result_desciption_detail}
    Run Keyword If      '${Compare_product_des_in_product_detail_2}' == 'True' or '${Compare_product_des_in_product_detail_2}' == 'False'
    ...    Change Text     ${Compare_product_des_in_product_detail_2}    AC${index}    ${excel}    ${excel_type}

    #Administration    > product detail
    ${product_solution_ws}=    Replace String    ${administration}    ${SPACE}        ${EMPTY}
    ${product_solution_in_product_detail}=    Get Text    xpath=/html/body/div[1]/div/div/div[1]/div[4]/div/div/div[1]/div[3]/div/div[2]/div[2]
    ${product_solution_in_product_detail_clearline}=  Replace String    ${product_solution_in_product_detail}    \n    ${SPACE}
    ${product_solution_in_product_detail_cursting}=    Replace String    ${product_solution_in_product_detail_clearline}    ${SPACE}        ${EMPTY}
    ${Compare_product_solution_in_product_detail}    Run Keyword And Return Status    Should Be Equal    ${product_solution_ws}    ${product_solution_in_product_detail_cursting}
    Run Keyword If      '${Compare_product_solution_in_product_detail}' == 'True' or '${Compare_product_solution_in_product_detail}' == 'False'
    ...    Change Boolean to Text    ${Compare_product_solution_in_product_detail}
    ${Compare_product_solution_in_product_detail_2}    Run Keyword And Return Status    Should Be Equal    ${result}    ${expected_result_administration}
    Run Keyword If      '${Compare_product_solution_in_product_detail_2}' == 'True' or '${Compare_product_solution_in_product_detail_2}' == 'False'
    ...    Change Text     ${Compare_product_solution_in_product_detail_2}    AG${index}    ${excel}    ${excel_type}
    ${product_additional_instructions_ws}=    Replace String    ${additional_instructions}    ${SPACE}        ${EMPTY}

    #Additional instructions    > product detail
    ${product_additional_instructions_in_product_detail}=    Get Text    xpath=/html/body/div[1]/div/div/div[1]/div[4]/div/div/div[1]/div[3]/div/div[3]/div[2]
    ${product_additional_instructions_in_product_detail_clearline}=  Replace String    ${product_additional_instructions_in_product_detail}    \n    ${SPACE}
    ${product_additional_instructions_in_product_detail_cursting}=    Replace String    ${product_additional_instructions_in_product_detail_clearline}    ${SPACE}        ${EMPTY}
        ${Compare_product_additional_instructions_in_product_detail}    Run Keyword And Return Status    Should Be Equal    ${product_additional_instructions_ws}    ${product_additional_instructions_in_product_detail_cursting}
    Run Keyword If      '${Compare_product_additional_instructions_in_product_detail}' == 'True' or '${Compare_product_additional_instructions_in_product_detail}' == 'False'
    ...    Change Boolean to Text    ${Compare_product_additional_instructions_in_product_detail}
    ${Compare_product_additional_instructions_in_product_detail_2}    Run Keyword And Return Status    Should Be Equal    ${result}    ${expected_result_additional_instructions}
    Run Keyword If      '${Compare_product_additional_instructions_in_product_detail_2}' == 'True' or '${Compare_product_additional_instructions_in_product_detail_2}' == 'False'
    ...    Change Text     ${Compare_product_additional_instructions_in_product_detail_2}    AI${index}    ${excel}    ${excel_type}
    ${product_cer_standard_ws}=    Replace String    ${certification standards}    ${SPACE}        ${EMPTY}

    #Certification Standards    > product detail
    ${product_cer_standard_in_product_detail}=    Get Text    xpath=/html/body/div[1]/div/div/div[1]/div[4]/div/div/div[1]/div[3]/div/div[3]/div[2]
    ${product_cer_standard_in_product_detail_clearline}=  Replace String    ${product_cer_standard_in_product_detail}    \n    ${SPACE}
    ${product_cer_standard_in_product_detail_cursting}=    Replace String    ${product_cer_standard_in_product_detail_clearline}    ${SPACE}        ${EMPTY}
    ${Compare_product_cer_standard_in_product_detail}    Run Keyword And Return Status    Should Be Equal    ${product_cer_standard_ws}    ${product_cer_standard_in_product_detail_cursting}
    Run Keyword If      '${Compare_product_cer_standard_in_product_detail}' == 'True' or '${Compare_product_cer_standard_in_product_detail}' == 'False'
    ...    Change Boolean to Text    ${Compare_product_cer_standard_in_product_detail}
    ${Compare_product_cer_standard_in_product_detail_2}    Run Keyword And Return Status    Should Be Equal    ${result}    ${expected_result_certification_standards}
    Run Keyword If      '${Compare_product_cer_standard_in_product_detail_2}' == 'True' or '${Compare_product_cer_standard_in_product_detail_2}' == 'False'
    ...    Change Text     ${Compare_product_cer_standard_in_product_detail_2}    AK${index}    ${excel}    ${excel_type}




Check ScrollTo
    execute javascript    window.scrollTo(0, 1100)
    Click Element         xpath=/html/body/div[1]/div/div/div[1]/div[3]/div/div[2]/div[2]/div/div[${no_product}]/div[1]
    

Change Text   
    [Arguments]    ${Check}    ${index}    ${excel}    ${excel_type}
    Run Keyword If  '${Check}' == 'True'
    ...            editFile    ${index}    TC2.3    ${excel}.${excel_type}    PASSED    alignment_value=center     font_color=1E8449      
    ...    ELSE    editFile    ${index}    TC2.3    ${excel}.${excel_type}    FAILED    alignment_value=center     font_color=FF0000    
Change Boolean to Text
    [Arguments]    ${Check}
    Run Keyword If  '${Check}' == 'True'
    ...    Set Global Variable    ${result}    PASSED
    ...    ELSE    Set Global Variable    ${result}    FAILED

*** Test Cases ***
Case - Organic Tea
    Read Excel    excel_type=xlsx