*** Settings ***
Library    SeleniumLibrary
Library    OperatingSystem
Library    RPA.Excel.Files
Library    ./libraries/Captcha.py

*** Variables ***
${URL}    https://www.cnas.org.cn/english/Auxiliary/SearchSystem/index.html?tpl=/LAS_FQ/publish/externalQueryL1En.jsp&name=Testing&CalibrationLaboratories
${CAPTCHA_DIR}    ${CURDIR}/capcaptcha
${EXCEL_FILE}    ${CURDIR}/resultados.xlsx
${EXCEL_FILE}    ${EXCEL_FILE}
${START_LIC}    ${START_LIC}
${END_LIC}    ${END_LIC}

*** Keywords ***
Open Chrome
    Open Browser    ${URL}    chrome    options=add_argument("--disable-blink-features=AutomationControlled")
    Maximize Browser Window
    Wait Until Element Is Visible    id=textIframe    30s
    Select Frame    id=textIframe

Search Laboratory
    [Arguments]    ${lic_id}
    ${search_success} =    Set Variable    ${False}
    
    WHILE    not ${search_success}
        Input License ID    ${lic_id}
        Click Search Button
        ${captcha_ok} =    Solve Captcha If Present
        IF    ${captcha_ok}
            ${search_success}    ${lic_id} =    Check And Process Search Results    ${lic_id}
        END
    END

Input License ID
    [Arguments]    ${lic_id}
    Wait Until Element Is Visible    id=licNo    30s
    Input Text    id=licNo    ${lic_id}

Click Search Button
    Wait Until Element Is Visible    xpath=//input[@value='Search']    30s
    Click Element    xpath=//input[@value='Search']
    Remove Files    ${CURDIR}/output/*.png

Solve Captcha If Present
    ${captcha_ok} =    Set Variable    ${False}
    Wait Until Element Is Visible    id=pirlAuthInterceptIframe    30s
    Select Frame    id=pirlAuthInterceptIframe
    Wait Until Element Is Visible    css=img    30s
    Sleep    10s
    
    WHILE    not ${captcha_ok}
        Remove Files    ${CURDIR}/output/*.png
        Capture Element Screenshot    css=img    ${CAPTCHA_DIR}/captcha1.png
        ${captcha_solution} =    Solve Captcha    ${CAPTCHA_DIR}/captcha1.png
        Log To Console    ✅ Captcha resuelto: ${captcha_solution}
        Sleep    10s
        Remove Files    ${CURDIR}/*.png
        Unselect Frame
        Wait Until Element Is Visible    id=textIframe    30s
        Select Frame    id=textIframe
        Wait Until Element Is Visible    id=authInterceptVName    30s
        Click Element    id=authInterceptVName
        Input Text    id=authInterceptVName    ${captcha_solution}
        Click Element    id=pirlbutton1
        Sleep    10s
        ${captcha_ok} =    Run Keyword And Return Status    Element Should Not Be Visible    id=pirlAuthInterceptIframe    timeout=30s
        Select Frame    id=pirlAuthInterceptIframe
    END
    Unselect Frame
    RETURN    ${captcha_ok}

Check And Process Search Results
    [Arguments]    ${lic_id}
    Wait Until Element Is Visible    id=textIframe    timeout=30s
    Select Frame    id=textIframe
    ${element_exists} =    Run Keyword And Return Status    Wait Until Element Is Visible    xpath=//a[contains(@onclick, "_showTop")]    timeout=30s

    IF    ${element_exists}
        ${element_text} =    Get Text    xpath=//a[contains(@onclick, "_showTop")]
        Log To Console    ✅ Texto del onclick encontrado: ${element_text}
        Save Data To Excel    ${element_text}    ${lic_id}
        RETURN    ${True}    ${lic_id}
    ELSE
        ${new_lic_id} =    Evaluate    'L' + str(int(${lic_id}[1:]) + 1)
        Log To Console    ❌ Elemento no encontrado, probando con otro ID
        RETURN    ${False}    ${new_lic_id}
    END

Save Data To Excel
    [Arguments]    ${data}    ${lic_id}
    ${excel_path}=    Set Variable    ${EXCEL_FILE}
    
    # Create workbook if it doesn't exist
    ${workbook_exists}=    Run Keyword And Return Status    File Should Exist    ${excel_path}
    
    IF    not ${workbook_exists}
        Create Workbook    ${excel_path}
        Set Worksheet Value    1    1    License ID
        Set Worksheet Value    1    2    Data
        Save Workbook
    END
    
    Open Workbook    ${excel_path}
    ${row}=    Evaluate    1
    ${next_empty}=    Set Variable    ${True}
    
    WHILE    ${next_empty}
        ${cell_value}=    Get Worksheet Value    ${row}    1
        IF    $cell_value is None or $cell_value == ''
            Set Worksheet Value    ${row}    1    ${lic_id}
            Set Worksheet Value    ${row}    2    ${data}
            ${next_empty}=    Set Variable    ${False}
        ELSE
            ${row}=    Evaluate    ${row} + 1
        END
    END
    
    Save Workbook
    Close Workbook

*** Tasks ***
Extract Labs and Save in a Excel
    ${START_LIC} =    Get Variable Value    ${START_LIC}
    ${END_LIC} =    Get Variable Value    ${END_LIC}
    Open Chrome
    FOR    ${lic_num}    IN RANGE    ${START_LIC}    ${END_LIC}+1
        ${lic_id} =    Set Variable    L${lic_num}
        Log To Console    Buscando laboratorio con ID: ${lic_id}
        Search Laboratory    ${lic_id}
    END
    Close Browser
