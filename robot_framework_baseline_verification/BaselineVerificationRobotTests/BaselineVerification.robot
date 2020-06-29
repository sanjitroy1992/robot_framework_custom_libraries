*** Settings ***
Library    SeleniumLibraryExtended    
Library    BaselineVerification    
*** Variables ***
@{SkipColumns}    DM Status    Sell Receiving Agent
${html1}    ${CURDIR}\\SampleData\\report.HTML
${html2}    ${CURDIR}\\SampleData\\detail report.HTML
${baseline1}     ${CURDIR}\\SampleData\\baseline1.xlsx
${baseline2}     ${CURDIR}\\SampleData\\baseline2.xlsx
${html_table1_xpath}    //table//div[text()="Ours"]/../../..
${html_table2_xpath}    (//tr[2]/td[3]/table)[4]
${html_table3_xpath}    //div[contains(text(),"Deal Number")]/../../..
${primary_column_name1}    Reference Num
${primary_column_name2}    11
${primary_column_name3}    Sender's Reference
*** Test Cases ***
Html Single Table Data Verification
    Verify Html Single Table Data    html_file_path=${html1}    html_table_xpath=${html_table1_xpath}  
    ...    baseline_excel_file_path=${baseline1}    baseline_sheet_name=Sheet1    primary_column_name=${primary_column_name1}    skip_columns_names=${SkipColumns}

Html Single Table Data Verification Without Column Headers
    Verify Html Single Table Data    html_file_path=${html1}    html_table_xpath=${html_table2_xpath}  
    ...    baseline_excel_file_path=${baseline1}    baseline_sheet_name=Sheet2    primary_column_name=${primary_column_name2}

Html All Table Data
    Verify Html All Table Data    html_file_path=${html2}    html_table_xpath=${html_table3_xpath} 
    ...    baseline_excel_file_path=${baseline2}    primary_column_name=${primary_column_name3}

Html All Table Data Sheet2
    Verify Html All Table Data    html_file_path=${html2}    html_table_xpath=${html_table3_xpath} 
    ...    baseline_excel_file_path=${baseline2}    baseline_sheet_name=Sheet2    primary_column_name=${primary_column_name3}