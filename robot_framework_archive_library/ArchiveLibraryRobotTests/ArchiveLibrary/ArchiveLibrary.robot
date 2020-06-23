*** Settings ***
Library    ArchiveLibrary   
Library    OperatingSystem

*** Test Cases ***
Archive Library
    
    Empty Directory    ${CURDIR}/SampleData/compressed_files/folder/folder1
    Extract File     ${CURDIR}/SampleData/compressed_files/folder/archivelibrary.zip     ${CURDIR}/SampleData/compressed_files/folder/folder1/zip
    Extract File     ${CURDIR}/SampleData/compressed_files/folder/xlrd.tar     ${CURDIR}/SampleData/compressed_files/folder/folder1/tar
    Extract File     ${CURDIR}/SampleData/compressed_files/folder/BOHistoryDetails.jar     ${CURDIR}/SampleData/compressed_files/folder/folder1/jar
    Extract File     ${CURDIR}/SampleData/compressed_files/folder/BankFusion.ear     ${CURDIR}/SampleData/compressed_files/folder/folder1/ear
    Extract File     ${CURDIR}/SampleData/compressed_files/folder/Uxp.war     ${CURDIR}/SampleData/compressed_files/folder/folder1/war
    
    Create Compressed File From Files In Directory     ${CURDIR}/SampleData/compressed_files/folder/folder1/zip     ${CURDIR}/SampleData/compressed_files/folder1/archivelibrary.zip    True    
    Create Compressed File From Files In Directory     ${CURDIR}/SampleData/compressed_files/folder/folder1/tar     ${CURDIR}/SampleData/compressed_files/folder1/xlrd.tar    True
    Create Compressed File From Files In Directory     ${CURDIR}/SampleData/compressed_files/folder/folder1/jar     ${CURDIR}/SampleData/compressed_files/folder1/BOHistoryDetails.jar    
    Create Compressed File From Files In Directory     ${CURDIR}/SampleData/compressed_files/folder/folder1/ear     ${CURDIR}/SampleData/compressed_files/folder1/BankFusion.ear
    Create Compressed File From Files In Directory     ${CURDIR}/SampleData/compressed_files/folder/folder1/war     ${CURDIR}/SampleData/compressed_files/folder1/Uxp.war
	Empty Directory    ${CURDIR}/SampleData/compressed_files/folder/folder1
    