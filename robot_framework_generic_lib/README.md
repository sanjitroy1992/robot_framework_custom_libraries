GenericLib
Library version:	3.0
Library scope:	test case
Named arguments:	supported
Introduction
Documentation for library GenericLib.

Shortcuts
Build Hash · Compare Xls Or Xlsx · Convert Csv To Xlsx · Convert To Business Day · Delete File If Exist · Execute Template With Multiple Data · Fin Update Properties File · FindTag ReplaceValue Xml · Get Data From Property File · Get Differences Of Two Lists · Get Link In Log · Get Matching Patterns · Jolokia Request · Kill All Process · Return Value From Template · Update Queue Message · Url Decoder · Verify Csv
Keywords
Keyword	Arguments	Documentation
Build Hash	input	
Usage
It is used to get the hash value for the input of any string and returns the encoded value

Arguments
'input' = Data input whose hash value is required

Returns
Returns the encoded data

Example:

|***TestCases*** |

${return_value} Build Hash 34234342

Compare Xls Or Xlsx	file1, file2	
Used to compare xls and xlsx file.

Convert Csv To Xlsx	csvfile, xlsxfile	
This keyword converts a file format from csv to xlsx

Convert To Business Day	Date, date_format=%d/%m/%Y, return_dateformat=%d/%m/%Y	
Used to convert any date to a business day.

It takes "Date" as an argument. If "Date" format is not equals to "%d/%m/%Y" then pass the new date format in "date_format" parameter

If the format of the return date is not equals to "%d/%m/%Y" then pass the return date format in "return_dateformat" parameter

|Example|

${d1} Convert To Business Day 22/04/2017 return_dateformat=%m/%d/%Y

# ${d1} 04/24/2017

${d2} Convert To Business Day 04/22/2017 %m/%d/%Y

# ${d2} 24/04/2017

Delete File If Exist	filename	
This keyword is used to delete a file with "filename". It deletes the file only if it exist in the location

Execute Template With Multiple Data	templateName, dataSet, datarow=None, sheetName=Sheet1, continue_on_failure=true, show_report=false, show_list_of_column_names=[]	
Usage
This keyword is used to run template with multiple data.

Data to template is passed either from an external file or directly using data dictionary or list of data dictionary from testcase to this keyword.

It supports different format for an external file like: csv, txt, xlsx, xls.

To run particular row data : Add 'rowid' column in dataSet file that contains ID(1,2,3..) for row data.

By default, this keyword will take all row data present in the given 'dataSet' file.To run particular row, set 'datarow' parameter.

Note: Please avoid using "." in column names as this keyword will ignore the text that comes after "." in column names

Arguments
'templateName' : Keyword name.

'dataSet' : Data file or data dictionary or list of data dictionary.

'datarow'[Optional] :

1. To pass single row data : datarow=RowID [Example: datarow=2]

2. To pass multiple random row data : datarow=RowID1,RowID2.. [Example: datarow=2,6,8,3]

3. To pass row present in some range : datarow=RowID1-RowID4 [Example: datarow=10-15]

'sheetName'[Optional] : If 'dataSet' is an xlsx or xls file with different sheet name then pass data sheet name.

'continue_on_failure' [Optional] : If keyword fails for some row and you do not want to continue with remaining rows then set continue_on_failure to false. By default, it is set to 'true'

'show_report' [Optional] : By setting this argument to 'true', it allows user to see the Execution Report for 'rowid' and 'status' columns in html format on the log.html file. By default, it is set to 'false'. Usage is shown in example 7 and 8.

'show_list_of_column_names' [Optional] : This option allows users to pass the columns names from dataSet, which are to be added in the Execution Report in the form of a list. Once passed as a list, the column names are shown after 'rowid' and 'status' columns in the Execution Report. Usage is shown in example 8.

* External file data format :

1. CSV file : data are separated by comma. First row has parameter name stored like 'ID','Branch'...etc and from second row start storing rows of data for respective parameter.

2. Text file : data can be separated by comma,semicolon or tab.

First line should have 'sep' parameter like sep=; or sep=, or sep=tab. Second line has parameter name stored like 'ID','Branch'...etc and from third line start storing data separated by 'sep' parameter for respective parameter.

3. Excel file(xls or xlsx): First row has parameter name.Start storing data from second row.

Note: In Order to skip any error related to Column not found in the data sheet and continue the execution, declare a variable named skip_error in a python file containing the substring of the error string that's common to all the fields and import that file in the "Variables" section of the test case. e.g skip_error = "Dictionary variable '&{Alldata}' has no key" in GenericConfig.py file

|Example|

1. To parse all row data present in external file : |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | sheetName=${sheetname} |

2. To parse single row data present in external file, pass particular row id to 'datarow' parameter: |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | datarow=23| sheetName=${sheetname} |

3. To parse multiple row data present in external file, pass all row id to 'datarow' parameter separated by ',' |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | datarow=23,25,78| sheetName=${sheetname} |

4. To parse range of row data present in external file, pass starting row id and end row id to 'datarow' parameter separated by '-' |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | datarow=20-25| sheetName=${sheetname} |

5. To parse range of row data present in external file and do not continue if any datarow fails. |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | datarow=20-25| sheetName=${sheetname} | continue_on_failure=false |

6. To parse dictionary data :

* Variable * |&{dataDict}| name=xyz | ID=234 |

* TestCases * |Execute Template With Multiple Data| Template_name | ${dataDict} |

7. To show Execution Report in log file. |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | sheetName=${sheetname} | show_report=true |

8. To Show Execution Report with column names and column data in log file.

* Variable * |@{ColumnDict}| URL | Username | Password |

* TestCases * |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | sheetName=${sheetname} | show_report=true | show_list_of_column_names=@{ColumnDict} |

Fin Update Properties File	file_name, *string_to_replace	
Usage
It updates the values for the existing keys in the .properties file

Arguments
'file_name' = .properties file location

'string_to_replace' = list of the elements that needs to be updated in the format "key=value".

Example:

|***Variable*** | |@{string_to_replace} | loan.repricing.limitvalidation.flag = noterror | loan.repricing.autoRollOver.validation.enabled=null | |***TestCases*** | |Fin Update Properties File | C:/loan.properties | @{string_to_replace}

FindTag ReplaceValue Xml	xml_file_name, xml_tag_name, text_to_replace	
|Usage| To replace tag value in xml file.

|Argument|

'xml_file_name' : xml file path

'xml_tag_name' : xml tag name

'text_to_replace' : tag value to replace

Get Data From Property File	file_name	
Usage
Gets data from .property file and returns dictionary. For more information on dictionary: https://github.com/robotframework/robotframework/blob/master/atest/testdata/standard_libraries/collections/dictionary.robot

Arguents
'file_name' = .property file location Example: |***TestCases*** | ${dict} | Get Data From Property File | ${filename} ${value} | Get From Dictionary | ${dict} | ${key}

Get Differences Of Two Lists	list1, list2	
Usage
Return the difference of 2 lists.

Arguments
list1 & list2

Get Link In Log	report_name, report_path	
Used to add any link in Robot Log File. Link points to any html page.

|Arguments|

'report_name' is used to display custom name for report.

'report_path' is the location of Report, which is given as href attribute of html anchor tag '<a>'

Get Matching Patterns	data, patterns	
|Usage| Used to get the matching patterns using regular expression. Multiple patterns can be also passed. It returns a list of all the matches.

|Argument| 'data' : Pass list of data 'pattern' : Pass list of patterns to be searched.

Jolokia Request	url, mbean=java.lang:type=Memory, attribute=HeapMemoryUsage	
|Usage| To get memory usage using Jolokia request.

|Argument| 'url' : Jolokia url

'mbean' : By default, java.lang:type=Memory

'attribute' : By default,HeapMemoryUsage

Kill All Process	*processList	
Used to Kill all running process by passing list of processname.

Arguments: '*processList' contains variable number of processname

Example:

**Variable**		
@{AllProcessToKill}	chrome.exe	chromedriver.exe
**TestCases**		
KillAllProcess	@{AllProcessToKill}	
Return Value From Template		
|Usage| Used to return value from 'Execute Template With Multiple Data' keyword if the passed template or keyword is returning any value

|Example| 1. If 'Execute Template With Multiple Data' keyword is running for more than one row data present in dataset file then dictionary will be return with 'rowid' as the dictionary key.

|Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | datarow=120-124 | sheetName=${sheetname} | |${v} | Return Value From Template |

2. If 'Execute Template With Multiple Data' keyword is running for single row data present in dataset file or dictionary variable passed then this keyword will return single value

|Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | datarow=120 | sheetName=${sheetname} | |${v} | Return Value From Template |

Update Queue Message	MQPropertiesFile, MisysUtilBatfile, *params	
Usage
Used to put or delete MQ queue messages.

PreRequisite : Security feature should be disable in MQ before updating any queue messages. Follow below steps to disable it.

1.Login to system where MQ is installed.

2.Make sure the user with which you start your WebSphere profile should be in mq server and user should be added to mqm group

3.Run below commands :

* runmqsc <QManager name> [Example: runmqsc MM453]

* alter qmgr chlauth(disabled)

* restart qmanager

Arguments
'MQPropertiesFile' = mq.properties file path.

'MisysUtilBatfile' = 'misys-mq-util' bat file location present under misys-mq-util folder.

'*params' : Can pass variable number of arguments in mq.properties file to configure Queue.

Input format for * params : key=value (example: MQCTestCase=C:\ITL)

Example:

create list variable:

* Variables *
@{QueueDetails} QueueManager=xyz HostName=pqr QueueName=qm Channel=SYSTEM.DEF.SVRCONN Port=1416 InputfilePath=D:\MQ_Queue_Messages\message1.txt
|Update Queue Message | ${mqpropertiesfile} | ${MisysUtilBatfile} | @{QueueDetails} |

Url Decoder	url	
Usage
It takes encoded url directly or location of file where url is present and returns the decoded url as an output.

Arguments
'url' = Pass the encoded url directly or pass location of file where url is present

Returns
Returns the decoded url

Example:

|***TestCases*** | 1. Decode url

${Decoded_url} | Url Decoder | url=https%3A%2F%2Fwww.finastra.com%2F

Log	${Decoded_url}
2. Decode url from the Filepath

${Decoded_url} | Url Decoder | D:/Projects/endoded_url.xml

Log	${Decoded_url}
Verify Csv	filename, *params	
Used to verify the data present in csv file(comma separated).

If all the expected values are equal to actual values then the keyword returns PASS else will raise an error and display all the differences in comparision. |Arguments|

'filename' = csv file path

'*params' = list of data to be verified.

Input Format for *param : Create a list with items present in key=value format. Example TradeID=23

NOTE: First item in *params should contain an unique element, that will be use to identify the correct row and all other values will be verified from that row.

Example:

**Variable**			
@{DataToVerify}	TradeID=Cre:123	NotionalAmount=1223.56	Collateral=1.234
**TestCases**			
${verifyStatus}	Verify Csv	${csvFilename}	@{DataToVerify}
Should Be Equal As Strings	${verifyStatus}	PASS	
Altogether 18 keywords.
Generated by Libdoc on 2020-02-10 17:34:14.

