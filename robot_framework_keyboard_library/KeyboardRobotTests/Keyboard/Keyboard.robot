*** Settings ***
Library    KeyboardLibrary    
Library    SeleniumLibraryExtended 
Test Teardown    Run Keywords    Sleep    2s    AND    Close All Browsers    
*** Test Cases ***
Enter Text
    Open Browser    https://www.google.com/    chrome
    Click Element    //input[@class="gLFyf gsfi"]   
    Native Type    Hello World
    Press Combination    KEY.ENTER

Enter Text using {SPACE}
    Open Browser    https://www.google.com/    chrome
    Click Element    //input[@class="gLFyf gsfi"]   
    Native Type    Hello{SPACE}World
    Press Combination    KEY.ENTER

Enter Text using {TAB}
    Open Browser    https://www.google.com/    chrome
    Click Element    //input[@class="gLFyf gsfi"]   
    Native Type    Hello{TAB}World
    Select All
    Press Combination    KEY.ENTER