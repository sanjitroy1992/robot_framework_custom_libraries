*** Settings ***
Library    PDFLibrary 

*** Variables ***
${filepath}    ${CURDIR}\\SampleData\\Account_Statement.pdf
${String}    1039531801
${String2}    14901
${String3}    22010010999
${String4}    username
${Prefix}    calculate_iban
${Suffix}     password
${Page_No}    1
${Line_No}    7
@{StringList}    ${String}    ${String2}    ${String3}    ${String4}
*** Test Cases ***
To Validate Text In PDF For First Occurrence
    Pdf Text Compare    ${filepath}    ${StringList} 
    
To Validate Text In PDF For All Occurrence
    Pdf Text Compare    ${filepath}    ${StringList}    all_occurrence=True
    
To Validate Text In PDF For First Occurrence With Prefix And Suffix
    Pdf Text Compare    ${filepath}    ${StringList}    prefix=${Prefix}    suffix=${Suffix}    
    
To Validate Text In PDF For All Occurrence With Prefix And Suffix
    Pdf Text Compare    ${filepath}    ${StringList}    prefix=${Prefix}    suffix=${Suffix}    all_occurrence=True
    
To Validate Text In PDF For All Occurrence With Prefix
    Pdf Text Compare    ${filepath}    ${StringList}    prefix=${Prefix}    all_occurrence=True
    
To Validate Text In PDF For All Occurrence With Suffix
    Pdf Text Compare    ${filepath}    ${StringList}    suffix=${Suffix}    all_occurrence=True

    
To Validate Text In PDF With Page No
    Pdf Text Compare By Page And Line Number    ${filepath}    ${String}    ${Page_No} 
    
To Validate Text In PDF With Page No For All Occurence
    Pdf Text Compare By Page And Line Number    ${filepath}    ${String}    page_no=${Page_No}    all_occurrence=True
    
To Validate list of Text In PDF With Page No
    Pdf Text Compare By Page And Line Number    ${filepath}    ${StringList}    ${Page_No}
    
To Validate list of Text In PDF With Page No For All Occurence
    Pdf Text Compare By Page And Line Number    ${filepath}    ${StringList}    ${Page_No}    all_occurrence=True
    
To Validate Text In PDF With Page No And Line No
    Pdf Text Compare By Page And Line Number    ${filepath}    ${String}    page_no=${Page_No}    line_no=${Line_No}
    


To Get PDF Text With Prefix And Suffix For First Occurrence 
    @{Listvalues}    Get Pdf Text    ${filepath}    prefix=${Prefix}    suffix=${Suffix}
    log    ${Listvalues}
    
To Get PDF Text With Prefix And Suffix For All Occurrence 
    Get Pdf Text    ${filepath}    prefix=${Prefix}    suffix=${Suffix}    all_occurrence=True
    
To Get PDF Text With Prefix For Firsr Occurrence 
    Get Pdf Text    ${filepath}    prefix=${Prefix}    
    
To Get PDF Text With Suffix For Firsr Occurrence 
    Get Pdf Text    ${filepath}    suffix=${Suffix}   
    
To Get PDF Text With Prefix For All Occurrence 
    Get Pdf Text    ${filepath}    prefix=${Prefix}    all_occurrence=True
    
To Get PDF Text With Suffix For All Occurrence 
    Get Pdf Text    ${filepath}    suffix=${Suffix}    all_occurrence=True