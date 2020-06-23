*** Settings ***
Library    XSDValidator     
*** Test Cases ***
Verify XSDValidator
    Validate XSD File    ${CURDIR}\\SampleData\\XSD\\S202SCTDHEUDHEUA4D91530LL9101E.I
    