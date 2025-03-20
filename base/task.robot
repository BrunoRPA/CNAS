*** Settings ***
Library    SeleniumLibrary
Library    OperatingSystem
Library    RPA.Excel.Files
Library    ./libraries/Captcha.py

*** Variables ***
${URL}    https://www.cnas.org.cn/english/Auxiliary/SearchSystem/index.html?tpl=/LAS_FQ/publish/externalQueryL1En.jsp&name=Testing&CalibrationLaboratories
${EXCEL_FILE}    ${CURDIR}/laboratorios.xlsx
${START_LIC}   1061
${END_LIC}     30000
${CAPTCHA_DIR}    ${CURDIR}/capcaptcha

*** Keywords ***
Open Chrome
    Open Browser    ${URL}    chrome    options=add_argument("--disable-blink-features=AutomationControlled")
    Maximize Browser Window
    Wait Until Element Is Visible    id=textIframe    10s
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
    Wait Until Element Is Visible    id=licNo    6s
    Input Text    id=licNo    ${lic_id}

Click Search Button
    Wait Until Element Is Visible    xpath=//input[@value='Search']    6s
    Click Element    xpath=//input[@value='Search']
    Remove Files    ${CAPTCHA_DIR}/*.png

Solve Captcha If Present
    ${captcha_ok} =    Set Variable    ${False}
    Wait Until Element Is Visible    id=pirlAuthInterceptIframe    6s
    Select Frame    id=pirlAuthInterceptIframe
    Wait Until Element Is Visible    css=img    30s
    Sleep    5s
    
    WHILE    not ${captcha_ok}
        Capture Element Screenshot    css=img    ${CAPTCHA_DIR}/captcha1.png
        ${captcha_solution} =    Solve Captcha    ${CAPTCHA_DIR}/captcha1.png
        Log To Console    ✅ Captcha resuelto: ${captcha_solution}
        Sleep    3s
        Remove Files    ${CURDIR}/*.png
        Unselect Frame
        Wait Until Element Is Visible    id=textIframe    6s
        Select Frame    id=textIframe
        Wait Until Element Is Visible    id=authInterceptVName    6s
        Click Element    id=authInterceptVName
        Input Text    id=authInterceptVName    ${captcha_solution}
        Click Element    id=pirlbutton1
        Sleep    3s
        ${captcha_ok} =    Run Keyword And Return Status    Element Should Not Be Visible    id=pirlAuthInterceptIframe    timeout=5s
        Select Frame    id=pirlAuthInterceptIframe
    END
    Unselect Frame
    RETURN    ${captcha_ok}

Check And Process Search Results
    [Arguments]    ${lic_id}
    Wait Until Element Is Visible    id=textIframe    timeout=5s
    Select Frame    id=textIframe
    ${element_exists} =    Run Keyword And Return Status    Wait Until Element Is Visible    xpath=//a[contains(@onclick, "_showTop")]    timeout=5s

    IF    ${element_exists}
        ${element_text} =    Get Text    xpath=//a[contains(@onclick, "_showTop")]
        Log To Console    ✅ Texto del onclick encontrado: ${element_text}
        Save Data To Excel    ${element_text}    ${lic_id}
        RETURN    ${True}    ${lic_id}
    ELSE
        ${new_lic_id} =    Evaluate    'L' + str(int(${lic_id}[1:]) + 1)
        Log To Console    ❌ Elemento no encontrado, probando con otro ID
        Log To Console    ✅ Intentando con nuevo ID: ${new_lic_id}
        RETURN    ${False}    ${new_lic_id}
    END

Save Data To Excel
    [Arguments]    ${data}    ${lic_id}

    ${file_exists} =    Run Keyword And Return Status    File Should Exist    ${EXCEL_FILE}

    IF    not ${file_exists}
        Create Workbook    ${EXCEL_FILE}
        Create Worksheet    Sheet1
        Set Cell Value    1    1    Numero de Registro
        Set Cell Value    1    2    Nombre de la Organizacion
        Save Workbook    ${EXCEL_FILE}
    END

    Open Workbook    ${EXCEL_FILE}

    ${rows} =    Read Worksheet As Table    header=True
    ${last_row} =    Get Length    ${rows}

    ${row_index} =    Evaluate    ${last_row} + 2

    Set Cell Value    ${row_index}    1    ${lic_id}
    Set Cell Value    ${row_index}    2    ${data}

    Save Workbook
    Close Workbook



*** Tasks ***
Extract Labs and Save in a Excel
    Open Chrome
    FOR    ${lic_num}    IN RANGE    ${START_LIC}       ${END_LIC}+1
        ${lic_id} =    Set Variable    L${lic_num}
        Log To Console    Buscando laboratorio con ID: ${lic_id}
        Search Laboratory    ${lic_id}
    END
    Close Browser