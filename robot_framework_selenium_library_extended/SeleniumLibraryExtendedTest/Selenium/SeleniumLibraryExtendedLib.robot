*** Settings ***
Library    OperatingSystem
Library    SeleniumLibraryExtended
Library    String 
Test Teardown    Close All Browsers
*** Variables ***
${BROWSER}  chrome
${timeout}  30s
${search_box_field}    xpath=//input[@name="q"]
${other_options}    //*[text()="Finastra - Wikipedia"]
${filepath}    ${CURDIR}\\SampleData\\IE\\test13.pdf
${DownloadDir}    ${CURDIR}\\SampleData\\Downloads
${xlsx_download_link}    https://www.exceltip.com/excel-data/remove-data-validation-dropdown-list-in-cell.html/attachment/download-sample-file-xlsx
${Download-Sample-File-xlsx}    xpath=//a[text()="Download-Sample File-xlsx"]
*** Keywords ***
Open Chrome Browser by setting download path
    [Documentation]    This keyword opens the chrome browser by setting the download folder location mentioned in the argument ${FolderLocation}.
    [Arguments]    ${URL}    ${FolderLocation}
    ${prefs}=    Create Dictionary    plugins.always_open_pdf_externally=${TRUE}       download.default_directory=${FolderLocation}
    ${chrome_options}=    Evaluate    sys.modules['selenium.webdriver'].ChromeOptions()    sys, selenium.webdriver
    Call Method    ${chrome_options}    add_experimental_option    prefs   ${prefs} 
    Create Webdriver    Chrome    chrome_options=${chrome options}
    Maximize Browser Window
	Go To    ${URL}
    
Open Firefox Browser by setting download path
    [Documentation]    This keyword opens the firefox browser by setting the download folder location mentioned in the argument ${FolderLocation}.
    [Arguments]    ${URL}    ${FolderLocation}
    ${var1}    Set Firefox Profile And Enable Download Directory    ff    ${FolderLocation}
    Open Browser    ${URL}     Firefox    ff_profile_dir=${var1}
    Maximize Browser Window
    
*** Test Cases ***
Activate And Input Text and Scroll Element Into View keyword
    Open Browser    http://www.google.com    chrome
    Maximize Browser Window
    Wait Until Page Contains Element    ${search_box_field}    10s
    Activate And Input Text    ${search_box_field}     Finastra
    Press Keys    //input[@name="q"]    ENTER
    Scroll Element Into View    ${other_options} 
    
Print Pdf -- PDF print ff
    Open Browser    https://www.google.com/    ff
    Maximize Browser Window
    Input Text    //input[@name="q"]    sample pdf
    Press Keys    //input[@name="q"]    ENTER
    Wait Until Element Is Visible    //a[@href="http://www.pdf995.com/samples/pdf.pdf"]      ${timeout}
    Click Link    //a[@href="http://www.pdf995.com/samples/pdf.pdf"]    
    Print Pdf
    Sleep     10
     
    
Print Pdf -- PDF print chrome
    Open Browser    https://www.google.com/    chrome
    Maximize Browser Window
    Input Text    //input[@name="q"]    sample pdf
    Press Keys    //input[@name="q"]    ENTER
    Wait Until Element Is Visible    //a[@href="http://www.pdf995.com/samples/pdf.pdf"]      ${timeout}
    Click Link    //a[@href="http://www.pdf995.com/samples/pdf.pdf"]    
    Print Pdf
    Sleep     10
    
Download Pdf -- PDF download ff
    Log        ${CURDIR}\\SampleData\\Firefox
    Open Firefox Browser by setting download path    https://www.google.com/   ${CURDIR}\\SampleData\\Firefox
    Maximize Browser Window
    Input Text    //input[@name="q"]    sample pdf
    Press Keys    //input[@name="q"]    ENTER
    Wait Until Element Is Visible    //a[@href="http://www.pdf995.com/samples/pdf.pdf"]      ${timeout}
    Click Link    //a[@href="http://www.pdf995.com/samples/pdf.pdf"]   
    Download Pdf 
    Sleep     2      
    
Download Pdf -- PDF download chrome
    Open Chrome Browser by setting download path    https://www.google.com/    ${CURDIR}\\SampleData\\Chrome  
    Maximize Browser Window
    Input Text    //input[@name="q"]    sample pdf
    Press Keys    //input[@name="q"]    ENTER
    Wait Until Element Is Visible    //a[@href="http://www.pdf995.com/samples/pdf.pdf"]      ${timeout}
    Click Link    //a[@href="http://www.pdf995.com/samples/pdf.pdf"]   
    Download Pdf
    Sleep     2
	
Chrome download test in UI mode
    Create Chrome Webdriver And Enable Download Directory    chrome    ${DownloadDir}
    Go To    ${xlsx_download_link}
    Sleep    2s
    Click Element    ${Download-Sample-File-xlsx}
    sleep    4s
Chrome download test in headless mode
    Create Chrome Webdriver And Enable Download Directory    chrome    ${DownloadDir}    headless_mode=Yes
    Go To    ${xlsx_download_link}
    Sleep    2s
    Click Element    ${Download-Sample-File-xlsx}
    sleep    4s
Firefox download test in UI mode
    ${var1}    Set Firefox Profile And Enable Download Directory    firefox    ${DownloadDir}
    Open Browser    ${xlsx_download_link}     Firefox    ff_profile_dir=${var1}
    Maximize Browser Window
    Sleep    2s
    Click Element    ${Download-Sample-File-xlsx}
    sleep    4s
Firefox download test in headless mode
    ${var1}    Set Firefox Profile And Enable Download Directory    firefox    ${DownloadDir}    headless_mode=Yes
    Open Browser    ${xlsx_download_link}     Firefox    ff_profile_dir=${var1}
    Set Window Size    1920    1080
    Sleep    2s
    Click Element    ${Download-Sample-File-xlsx}
    sleep    4s