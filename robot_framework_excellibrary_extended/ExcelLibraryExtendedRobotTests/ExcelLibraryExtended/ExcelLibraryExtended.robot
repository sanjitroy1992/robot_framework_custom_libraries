*** Settings ***          
Library    ExcelLibraryExtended     
*** Variables ***
${Xlsx_Path}    ${CURDIR}\\SampleData\\Sample_Excel.xlsx
*** Test Cases ***
Edit Data Xlsx File Without Rownumber
    [Documentation]    By default it selects the first row of the corrosponding columnheader and updates the value.
    Edit Data Xlsx File    ${Xlsx_Path}    Workflow1.1    Pin_Number    hello world
    
Edit Data Xlsx File With Valid Rownumber
    [Documentation]    It selects the fifth row of the corrosponding columnheader and updates the value.
    Edit Data Xlsx File    ${Xlsx_Path}    Workflow1.1    Pin_Number    hello world    rownumber=5
    
Edit Data Xlsx File With Valid Rownumber Multiple Times
    [Documentation]    It selects the fourth row of the corrosponding columnheader and updates the value for multiple times.
    Edit Data Xlsx File    ${Xlsx_Path}    Workflow1.1    Pin_Number    hello world    rownumber=4
    Edit Data Xlsx File    ${Xlsx_Path}    Workflow1.1    Pin_Number    ExcelLibraryExtended    rownumber=4
    
Edit Data Xlsx File With an Invalid Rownumber
    [Documentation]    It selects the invalid row of the corrosponding columnheader and updates the value. [-ve scenario]
    ${Status}    Run Keyword And Return Status     Edit Data Xlsx File    ${Xlsx_Path}    Workflow1.1    Pin_Number    hello world    rownumber=cc
    Run Keyword If    ${Status}==True    Fail    Please Provide a valid 'rownumber' as an integer        