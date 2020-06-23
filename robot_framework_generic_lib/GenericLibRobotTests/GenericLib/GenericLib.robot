*** Settings ***
Library    GenericLib 
Library    SeleniumLibraryExtended  
Library    Collections    
*** Variables ***
${file}    ${CURDIR}/SampleData/download.csv
${DatasetPath}    ${CURDIR}/SampleData/Datasheet.xlsx
${DatasetPath1}    ${CURDIR}/SampleData/Datasheet1.xlsx
${DatasetPath2}    ${CURDIR}/SampleData/Datasheet2.xlsx
${DatasetPath3}    ${CURDIR}/SampleData/Datasheet3.xlsx
${EncodedURLPath}    ${CURDIR}/SampleData/EncodedURL.txt
${csvfile}    ${CURDIR}/SampleData/download.csv
${xlsxfile}    ${CURDIR}/SampleData/download.xlsx
@{ColumnDict}    Browser    Username    User    User1    User2    Password
${ExpectedURL}    https://www.finastra.com/
${EncededURL}    https%3A%2F%2Fwww.finastra.com%2F
${SamplePropertyFile}    ${CURDIR}/SampleData/sample.property
${key}    z
${HashValue}    34234342
*** Keywords *** 
Launch Browser in headless
    [Arguments]    ${url}    ${Browser}
    [Return]    ${Browser_index}
    ${browser_options}    Evaluate    sys.modules['selenium.webdriver'].ChromeOptions()    sys, selenium.webdriver    
    Call Method    ${browser_options}    add_argument    --headless
    ${options}     Call Method     ${browser_options}    to_capabilities        
    ${Browser_index}    Open Browser    ${url}    ${Browser}    desired_capabilities=${options}  
    Maximize Browser Window 
Open browser Template
    [Arguments]    ${Alldata}
    Launch Browser in headless    https://www.google.com    &{Alldata}[Browser]
    Close Browser
    [Return]    &{Alldata}[Browser] 
*** Test Cases ***
Verify Csv - 1st and 2nd column not found -- Negitive Scenario
    ${IsPassed}    Run Keyword And Continue On Failure    Run Keyword And Return Status    Verify Csv    ${file}    Category Combination=Checking    Household=1,902    Deposit Balance=11,198,695
	Run Keyword If    ${IsPassed}==True    Fail     Verify Csv failed to handle negative scenario         

Verify Csv - both data not found -- Negitive Scenario
    ${IsPassed}    Run Keyword And Continue On Failure    Run Keyword And Return Status    Verify Csv    ${file}     Category Combinations=Checking    Households=1,903    Deposit Balance=11,198,6945
	Run Keyword If    ${IsPassed}==True    Fail     Verify Csv failed to handle negative scenario 
	
Show Report True - All Passed ++ Positive Scenario
    Execute Template With Multiple Data    Open browser Template    ${DatasetPath}    datarow=1,3,4    show_report=true     show_list_of_column_names=@{ColumnDict}    
    ${Browser}    Return Value From Template 
    
Show Report True - All Failed -- Negitive Scenario
    ${Status}    Run Keyword And Return Status     Execute Template With Multiple Data    Open browser Template    ${DatasetPath}    datarow=2,5,8    show_report=true     show_list_of_column_names=@{ColumnDict}    
    Run Keyword If    '${Status}'=='False'    Log    Execute Template With Multiple Data keyword failed for some data
    
Show Report True - Passed and Failed -- Negitive Scenario
    ${Status}    Run Keyword And Return Status     Execute Template With Multiple Data    Open browser Template    ${DatasetPath}    datarow=1,2,3    show_report=true     show_list_of_column_names=@{ColumnDict}    
    Run Keyword If    '${Status}'=='False'    Log    Execute Template With Multiple Data keyword failed for some data

Show Report False - All Passed ++ Positive Scenario
    Execute Template With Multiple Data    Open browser Template    ${DatasetPath}    datarow=1,3,4     show_list_of_column_names=@{ColumnDict}    
    ${Browser}    Return Value From Template
    
Show Report False - All Failed -- Negitive Scenario
    ${Status}    Run Keyword And Return Status     Execute Template With Multiple Data    Open browser Template    ${DatasetPath}    datarow=2,5,8    show_report=false     show_list_of_column_names=@{ColumnDict}    
    Run Keyword If    '${Status}'=='False'    Log    Execute Template With Multiple Data keyword failed for some data
    
Show Report False - Passed and Failed -- Negitive Scenario
    ${Status}    Run Keyword And Return Status     Execute Template With Multiple Data    Open browser Template    ${DatasetPath}    datarow=1,2,3     show_list_of_column_names=@{ColumnDict}    
    Run Keyword If    '${Status}'=='False'    Log    Execute Template With Multiple Data keyword failed for some data

Show Dict Values - True ++ Positive Scenario
    Execute Template With Multiple Data    Open browser Template    ${DatasetPath}    datarow=1,3,4     show_list_of_column_names=@{ColumnDict}    
    ${Browser}    Return Value From Template
    
Show Dict Values - False -- Negitive Scenario
    ${show_dict_var}    Set Variable    No
    Execute Template With Multiple Data    Open browser Template    ${DatasetPath}    datarow=1,3,4     show_list_of_column_names=@{ColumnDict}    
    
Compare Xls Or Xlsx ++ Positive Scenario
    Compare Xls Or Xlsx    ${DatasetPath1}    ${DatasetPath2}
    
Compare Xls Or Xlsx -- Negitive Scenario
    ${Status}    Run Keyword And Return Status    Compare Xls Or Xlsx    ${DatasetPath2}    ${DatasetPath3}
    Run Keyword If    '${Status}'=='False'    Log    Files are different
    
Convert Csv To Xlsx
    Convert Csv To Xlsx    ${csvfile}    ${xlsxfile} 
    
Delete File If Exist    
    Delete File If Exist    ${xlsxfile}
    
Decode url
    ${Decoded_url}    Url Decoder    url=${EncededURL}
    Should Be Equal As Strings    ${Decoded_url}    ${ExpectedURL}
    
Decode url from the Filepath    
    ${Decoded_url}    Url Decoder    ${EncodedURLPath}
    Should Be Equal As Strings    ${Decoded_url}    ${ExpectedURL}
    
Build Hash
    ${return_value}    Build Hash    ${HashValue}
    
Get Data From Property File
    ${dict}    Get Data From Property File    ${SamplePropertyFile}
    ${value}    Get From Dictionary    ${dict}    ${key}
    