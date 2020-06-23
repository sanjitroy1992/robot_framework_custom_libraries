import os
from robot.libraries.BuiltIn import BuiltIn
from robot.libraries.Process import Process
import datetime
import collections
from robot.api import logger
from itertools import zip_longest
import xlrd
import re
import shutil
import xml.etree.ElementTree
from pyjolokia import Jolokia
import csv
from openpyxl import Workbook
from HTML_external import HTML_external
import sys
import urllib.request, urllib.parse, urllib.error
import pandas as pd
import numpy as np
import math
import hashlib
from pyjavaproperties import Properties
import importlib
import subprocess
from subprocess import Popen, PIPE

class GenericLib(object):


    ROBOT_LIBRARY_VERSION = 3.0

    def convert_to_business_day(self, Date, date_format="%d/%m/%Y", return_dateformat='%d/%m/%Y'):
        """ Used to convert any date to a business day.

         It takes "Date" as an argument. If "Date" format is not equals to "%d/%m/%Y" then pass the new date format in "date_format" parameter

         If the format of the return date is not equals to "%d/%m/%Y" then pass the return date format in "return_dateformat" parameter

         |Example|

         ${d1}    Convert To Business Day    22/04/2017    return_dateformat=%m/%d/%Y

         # ${d1}    04/24/2017

         ${d2}    Convert To Business Day    04/22/2017    %m/%d/%Y

         # ${d2}    24/04/2017

        """

        dateobj = datetime.datetime.strptime(Date, date_format)
        weekDayNum = dateobj.isoweekday()
        if weekDayNum == 6:
            daysToAdd = datetime.timedelta(days=2)
            dateobj = dateobj + daysToAdd
        if weekDayNum == 7:
            daysToAdd = datetime.timedelta(days=1)
            dateobj = dateobj + daysToAdd

        datefromObj = (str(dateobj).split(" "))[0].split("-")
        dateformatList = date_format.split("/")
        day = datefromObj[2]
        month = datefromObj[1]
        year = datefromObj[0]
        returndateformatList = return_dateformat.split("/")
        if returndateformatList[0] == "%d":
            var1 = day
        elif returndateformatList[0] == "%m":
            var1 = month
        else:
            var1 = year
        if returndateformatList[1] == "%d":
            var2 = day
        elif returndateformatList[1] == "%m":
            var2 = month
        else:
            var2 = year
        if returndateformatList[2] == "%d":
            var3 = day
        elif returndateformatList[2] == "%m":
            var3 = month
        else:
            var3 = year
        BusinessDay = "{}/{}/{}".format(var1, var2, var3)
        return BusinessDay

    def killAllProcess(self, *processList):
        """ Used to Kill all running process by passing list of processname.

         Arguments: '*processList' contains variable number of processname

         Example:

        | ***Variable*** |
        | @{AllProcessToKill} | chrome.exe | chromedriver.exe |
        | ***TestCases*** |
        | KillAllProcess | @{AllProcessToKill} |

        """
        tasklist = subprocess.Popen(("tasklist"), stdout=PIPE, stderr=PIPE)
        stdout, stderr = tasklist.communicate()
        for process in processList:
            if process in str(stdout):
                killCommand = "Taskkill /IM" + " " + process + " /F"
                try:
                    subprocess.Popen(killCommand, stdout=PIPE, stderr=PIPE)
                except:
                    pass
        filesintemp = [name for name in os.listdir(os.environ['TEMP'])]
        for i in filesintemp:
            if "scoped_dir" in i:
                dir = os.path.join(os.environ['TEMP'], i)
                try:
                    shutil.rmtree(dir)
                except:
                    pass

    def execute_template_with_multiple_data(self, templateName, dataSet, datarow="None", sheetName="Sheet1",
                                               continue_on_failure='true', show_report='false',
                                               show_list_of_column_names=[]):
        """| Usage |
         This keyword is used to run template with multiple data.

         Data to template is passed either from an external file or directly using data dictionary or list of data dictionary from testcase to this keyword.

         It supports different format for an external file like: csv, txt, xlsx, xls.

         To run particular row data : Add 'rowid' column in dataSet file that contains ID(1,2,3..) for row data.

         By default, this keyword will take all row data present in the given 'dataSet' file.To run particular row, set 'datarow' parameter.

         Note: Please avoid using "." in column names as this keyword will ignore the text that comes after "." in column names

         | Arguments |

         'templateName' : Keyword name.

         'dataSet'      : Data file or data dictionary or list of data dictionary.

         'datarow'[Optional]      :

                                 1. To pass single row data : datarow=RowID   [Example: datarow=2]

                                 2. To pass multiple random row data : datarow=RowID1,RowID2..   [Example: datarow=2,6,8,3]

                                 3. To pass row present in some range : datarow=RowID1-RowID4   [Example: datarow=10-15]

         'sheetName'[Optional] : If 'dataSet' is an xlsx or xls file with different sheet name then pass data sheet name.

         'continue_on_failure' [Optional] : If keyword fails for some row and you do not want to continue with remaining rows then set continue_on_failure to false.
                                            By default, it is set to 'true'

         'show_report' [Optional] : By setting this argument to 'true', it allows user to see the Execution Report for 'rowid' and 'status' columns in html format on the log.html file.
                                    By default, it is set to 'false'. Usage is shown in example 7 and 8.

         'show_list_of_column_names' [Optional] : This option allows users to pass the columns names from dataSet, which are to be added in the Execution Report in the form of a list.
                                                  Once passed as a list, the column names are shown after 'rowid' and 'status' columns in the Execution Report.
                                                  Usage is shown in example 8.
        
         * External file data format :

         1. CSV file : data are separated by comma. First row has parameter name stored like 'ID','Branch'...etc and from second row start storing rows of data for respective parameter.

         2. Text file : data can be separated by comma,semicolon or tab.

          First line should have 'sep' parameter like sep=; or sep=, or sep=tab. Second line has parameter name stored like 'ID','Branch'...etc and from third line start storing data separated by 'sep' parameter for respective parameter.

         3. Excel file(xls or xlsx): First row has parameter name.Start storing data from second row.

         Note: In Order to skip any error related to Column not found in the data sheet and continue the execution, declare a variable
                named skip_error in a python file containing the substring of the error string that's common to all the fields
                 and import that file in the "Variables" section of the test case.
                e.g
                skip_error = "Dictionary variable '&{Alldata}' has no key"
                    in GenericConfig.py file

        |Example|

         1. To parse all row data present in external file :
         |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | sheetName=${sheetname} |

         2. To parse single row data present in external file, pass particular row id to 'datarow' parameter:
         |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | datarow=23| sheetName=${sheetname} |

         3. To parse multiple row data present in external file, pass all row id to 'datarow' parameter separated by ','
         |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | datarow=23,25,78| sheetName=${sheetname} |

         4. To parse range of row data present in external file, pass starting row id and end row id to 'datarow' parameter separated by '-'
         |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | datarow=20-25| sheetName=${sheetname} |

         5. To parse range of row data present in external file and do not continue if any datarow fails.
         |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | datarow=20-25| sheetName=${sheetname} | continue_on_failure=false |

         6. To parse dictionary data :

         *** Variable ***
         |&{dataDict}| name=xyz | ID=234 |

         *** TestCases ***
         |Execute Template With Multiple Data| Template_name | ${dataDict} |

         7. To show Execution Report in log file.
         |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | sheetName=${sheetname} | show_report=true |

         8. To Show Execution Report with column names and column data in log file.

         *** Variable ***
         |@{ColumnDict}| URL | Username | Password |

         *** TestCases ***
         |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | sheetName=${sheetname} | show_report=true | show_list_of_column_names=@{ColumnDict} |
        """
        try:
            template_parameters = {}
            table_header = """*HTML*
                           <div><div><table>
                           <caption><b>Execution Report</b></caption>
                           <tr><td style="background:#5CBFDE;text-align:center">rowid</td>
                           <td style="background:#5CBFDE;text-align:center">Status</td>
                           """
            for item in show_list_of_column_names:
                table_header = table_header+'<td style="background:#5CBFDE;text-align:center">' +item +'</td>'
            table_header = table_header + '</tr>'
            table_rows = ''
            error_found = ""
            global template_return_values
            continue_on_failure = str(continue_on_failure)
            data_type = type(dataSet)
            if data_type is collections.OrderedDict or data_type is dict or 'robot.utils.dotdict.DotDict' in str(data_type):
                try:
                    return_values = BuiltIn().run_keyword(templateName, dataSet)
                    status = "PASS"
                except Exception as e:
                    error_found = "{}\n".format(e)
            else:
                if data_type is list:
                    template_parameters = dataSet
                    status = "PASS"
                else:
                    if not os.path.isfile(dataSet):
                        raise AssertionError("Invalid Input Error ! \nFile {} does not exist.".format(dataSet))
                    else:
                        template_parameters, status = self._get_all_data_from_file(dataSet, str(datarow), sheetName) #------- Get all data from dataSet file and store in template_parameters
                return_values = collections.OrderedDict()
                r = 1
                for parameter in template_parameters:
                    flag_exc=0
                    if datarow == "None":
                        rowId = r
                        r = r+1

                    else:
                        rowId = int(parameter['rowid'])
                    try:
                        if continue_on_failure.lower() == 'false':
                            status1, value = BuiltIn().run_keyword_and_ignore_error(templateName, parameter)
                            if status1 == "FAIL":
                                raise AssertionError("FAIL")
                        else:
                            status1, value = BuiltIn().run_keyword_and_ignore_error(templateName, parameter)
                            if status1 == "FAIL":
                                logger.fail("failed")
                    except Exception as err:
                        flag_exc = 1
                        return_values[str(rowId)] = "Fails"
                        error_found = "{}\n".format(err)
                        table_rows = table_rows + """
                                     <tr><td style="background:red;text-align:center">""" + str(rowId) + '</td>'\
                                     '<td style="background-color:red;text-align:center">' + str(status1) + '</td>'

                        for item in show_list_of_column_names:
                            if item in parameter:
                                value1 = str(parameter[item])
                                table_rows = table_rows + """<td style="background-color:red;text-align:center">""" + str(value1) + '</td>'
                            else:
                                value1 = "DATA NOT FOUND!!"
                                table_rows = table_rows + """<td style="background-color:orange;text-align:center">""" + str(value1) + '</td>'

                        table_rows = table_rows + '</tr>'
                        if value:
                            return_values[str(rowId)] = value
                        if continue_on_failure.lower() == 'false':
                            break
                    if flag_exc == 0:
                        table_rows = table_rows + """<tr><td style="background-color:green;text-align:center">""" + str(rowId) + '</td>'\
                                         '<td style="background-color:green;text-align:center">' + str(status1) + '</td>'


                        for item in show_list_of_column_names:
                            if item in parameter:
                                value1 = str(parameter[item])

                                table_rows = table_rows + """<td style="background-color:green;text-align:center">""" + str(value1) + '</td>'
                            else:
                                value1 = "DATA NOT FOUND!!"
                                table_rows = table_rows + """<td style="background-color:orange;text-align:center">""" + str(value1) + '</td>'
                        table_rows = table_rows + '</tr>'
                        if value:
                            return_values[str(rowId)] = value
            del template_parameters
            template_return_values = ""
            if type(return_values) is collections.OrderedDict:
                if len(list(return_values.values())) == 1:
                    template_return_values = list(return_values.values())[0]
                else:
                    template_return_values = return_values
            elif type(return_values) is str or type(return_values) is str:
                template_return_values = return_values
            else:
                template_return_values = None
            if status != "PASS" and error_found == "":
                raise AssertionError(status)
            elif status == "PASS" and error_found != "":
                raise AssertionError("Execute Template With Multiple Data keyword failed for some data")
            elif status != "PASS" and error_found != "":
                raise AssertionError(
                    "Error: \n{}\n Execute Template With Multiple Data keyword failed for some data".format(status))
            else:
                logger.info("Execute Template With Multiple Data keyword passed for all the data")
        except Exception as err:
            raise AssertionError(err)
        finally:
            if show_report.lower() == 'true':
                print(table_header+table_rows)

    def return_value_from_template(self):
        """|Usage|
         Used to return value from 'Execute Template With Multiple Data' keyword if the passed template or keyword is returning any value

        |Example|
         1. If 'Execute Template With Multiple Data' keyword is running for more than one row data present in dataset file then dictionary will be return with 'rowid' as the dictionary key.

         |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | datarow=120-124 | sheetName=${sheetname} |
         |${v} | Return Value From Template |

         2. If 'Execute Template With Multiple Data' keyword is running for single row data present in dataset file or dictionary variable passed then this keyword will return single value

         |Execute Template With Multiple Data| Template_name | ${XLSXexcel_file} | datarow=120 | sheetName=${sheetname} |
         |${v} | Return Value From Template |


        """
        try:
            return template_return_values
        except:
            raise AssertionError("No Value returned from 'Execute Template With Multiple Data' keyword !!!")

    def _get_all_data_from_file(self, filename, datarow="None", sheetName='Sheet1'):
        """Usage: To read the file and return all row data in dictionary and then store each row dictionary data inside a list and returns list of dictionary.
         Type of file supported : csv , txt, xls, xlsx type of files.
        """
        if datarow != "None":
            rowNumbers = []  # get row numbers to fetch row data from file.
            if ',' in datarow:
                rowNumbers = datarow.split(",")
            elif '-' in datarow:
                rowRange = datarow.split("-")
                startRow = int(rowRange[0])
                EndRow = int(rowRange[1])
                for i in range(startRow, EndRow + 1):
                    rowNumbers.append(str(i))
            else:
                rowNumbers.append(datarow)
        name, fileFormat = os.path.splitext(filename)  # ----- Get file format
        fileFormat = fileFormat.replace(".", "")
        AllData = []  # ------- Store file data in AllData list
        RowNotFound = []
        if fileFormat in ['xls', 'xlsx', 'csv']:
            if fileFormat == 'csv':
                df_csv = pd.read_csv(filename, dtype=str, na_filter=False)
            elif fileFormat in ['xls', 'xlsx']:
                df_csv = pd.read_excel(filename, sheetName, dtype=str, na_filter=False)
            if datarow != "None":
                for item in rowNumbers:
                    if not np.any(df_csv['rowid'].values == str(item)):
                        RowNotFound.append(item)
                list_rows = df_csv[df_csv['rowid'].isin(rowNumbers)].values
            else:
                list_rows = df_csv.values
            # keyparam = df_csv.columns.values
            # ----------- uncomment above step to avoid splitting col names by "."
            keyparam = []
            for item in df_csv.columns.values:
                if ("$.." in str(item)):
                    keyparam.append(str(item))
                else:
                    keyparam.append(str(item).split('.')[0])
            # commenting for future reference
            # keyparam = df_csv.columns.str.split('.').str[0]
            AllData = []
            for r, row in enumerate(list_rows):
                fileRowData = collections.OrderedDict()
                for c, col in enumerate(row):
                    if isinstance(col, str):
                        val = str(col)
                    elif isinstance(col, float):
                        if math.isnan(col):
                            val = ""
                        elif col == int(col):
                            val = str(int(col))
                        else:
                            val = str(col)
                    elif isinstance(col, bool):
                        val = str(bool(col))
                    elif isinstance(col, int):
                        val = str(int(col))
                    else:
                        val = str(col)
                    if "${" in str(val):
                        try:
                            val = self._get_global_parameter(val)
                        except:
                            pass
                    fileRowData[keyparam[c]] = val.strip()
                AllData.append(fileRowData)
                del fileRowData

        elif fileFormat == 'txt':
            with open(filename, 'r') as fileobj:  # ----- Open ,read and close the file
                fileData = fileobj.readlines()
            if fileData[0].split('=')[
                0] != 'sep':  # ----check file separator (, or ; or tab. If file separator is tab then split by \t
                raise AssertionError("No separator is present in the given data source text file !!!")
            else:
                fileSeparator = (fileData[0].split('=')[-1]).replace("\n", "")
                if fileSeparator == 'tab':
                    keyparam = (fileData[1].replace("\n", "")).split('\t')
                else:
                    keyparam = (fileData[1].replace("\n", "")).split(fileSeparator)
                    # --------- Get all file row data, split it by file separator and then store each row data as a value and parameter-name as a key of row dictionary
                for item in range(2, len(fileData)):
                    fileRowData = collections.OrderedDict()
                    if fileSeparator == 'tab':
                        dataList = (fileData[item].replace("\n", "")).split('\t')
                    else:
                        dataList = (fileData[item].replace("\n", "")).split(fileSeparator)
                    for data in range(0, len(dataList)):
                        if 'rowid' in keyparam[data].lower():
                            keyparam[data] = 'rowid'
                        if "${" in str(dataList[data]):
                            try:
                                dataList[data] = self._get_global_parameter(str(dataList[data]))
                            except:
                                pass
                        fileRowData[keyparam[data]] = dataList[data]
                    AllData.append(fileRowData)
                    del fileRowData
                del fileData
        if len(RowNotFound) == 0:
            return AllData, "PASS"
        else:
            DatafileError = ""
            if RowNotFound != []:
                DatafileError = "Row ID: "
                for i in range(len(RowNotFound)):
                    if i == len(RowNotFound) - 1:
                        DatafileError += RowNotFound[i] + " not found in data file: " + filename
                    else:
                        DatafileError += RowNotFound[i] + ","
            if DatafileError != "":
                return AllData, DatafileError
            else:
                return AllData, "PASS"

    def _get_global_parameter(self, value):
        if "${" in value:
            value = value.replace("${", "<<").replace("}", ">>")
        if "<<" in value:
            paramCount = value.count('<<')
            if paramCount >= 1:
                varlist = []
                varValue = {}
                start = [pos for pos, char in enumerate(value) if char == "<"]
                end = [pos for pos, char in enumerate(value) if char == ">"]
                start = [start[i] for i in range(0, len(start)) if i % 2 != 0]
                end = [end[i] for i in range(0, len(end)) if i % 2 != 0]
                [varlist.append(value[start[i] + 1:end[i] - 1]) for i in range(paramCount)]
                for l in varlist:
                    varValue[l] = BuiltIn().get_variable_value("${}".format(l))
                for k in list(varValue.keys()):
                    if varValue[k] != None:
                        value = value.replace("<<" + k + ">>", varValue[k])
                value = value.replace("<<", "${").replace(">>", "}")
        return value

    def update_queue_message(self, MQPropertiesFile, MisysUtilBatfile, *params):
        """| Usage |
         Used to put or delete MQ queue messages.

         PreRequisite : Security feature should be disable in MQ before updating any queue messages. Follow below steps to disable it.

         1.Login to system where MQ is installed.

         2.Make sure the user with which you start your WebSphere  profile should be in mq server and  user should be added to mqm group

         3.Run below commands :

         *  runmqsc <QManager name> [Example: runmqsc MM453]

         *  alter qmgr chlauth(disabled)

         *  restart qmanager

         | Arguments |

         'MQPropertiesFile' = mq.properties file path.

         'MisysUtilBatfile' = 'misys-mq-util' bat file location present under misys-mq-util folder.

         '*params' : Can pass variable number of arguments in mq.properties file to configure Queue.

         Input format for * params :  key=value (example: MQCTestCase=C:\\ITL)

         Example:

         create list variable:

         | *** Variables *** |

         | @{QueueDetails}    QueueManager=xyz    HostName=pqr    QueueName=qm    Channel=SYSTEM.DEF.SVRCONN    Port=1416    InputfilePath=D:\\MQ_Queue_Messages\\message1.txt  |

         |Update Queue Message | ${mqpropertiesfile} | ${MisysUtilBatfile} | @{QueueDetails} |

         """
        # --- check if MQPropertiesFile exist
        if not os.path.exists(MQPropertiesFile):
            raise AssertionError("File not found error :" + MQPropertiesFile + " doesnot exist")

        # --- check if MisysUtilBatfile exist
        if not os.path.exists(MisysUtilBatfile):
            raise AssertionError("File not found error :" + MisysUtilBatfile + " doesnot exist")

        # ----- param will hold variable number of parameters that needs to be changed
        param_dict = collections.OrderedDict()
        for param in params:
            parameter = str(param).split("=")
            parameter_name = parameter[0]
            if parameter_name.lower() == 'inputfilepath':
                parameter_value = parameter[1].replace('\\', '\\\\')
            else:
                parameter_value = parameter[1]
            param_dict[parameter_name] = parameter_value

            # --- check if that parameter_name is present in file , if present then update it
            with open(MQPropertiesFile, 'r') as f:
                Filedata = f.read().split("\n")
                columns = list(param_dict.keys())
                for column in columns:

                    for i in range(0, len(Filedata) - 1):
                        if str(column).lower() == str(Filedata[i].split("=")[0]).lower():
                            parametername = Filedata[i].split("=")[0]
                            Filedata[i] = parametername + "=" + param_dict[column]

            with open(MQPropertiesFile, 'w') as fr:
                for data in Filedata:
                    fr.write(data)
                    fr.write("\n")
                fr.close()

        # ------ create cwd for to run bat process
        filedir = MisysUtilBatfile.split("\\")
        batcwd = ""
        for i in range(0, len(filedir) - 1):
            if i < len(filedir) - 2:
                batcwd = batcwd + filedir[i] + "\\"
            else:
                batcwd = batcwd + filedir[i]
        status = Process().run_process(MisysUtilBatfile, cwd=str(batcwd).strip())
        stdoutput = status.stdout

        if "failed" in stdoutput:
            print(" Following data passed in mq.properties file: ------------------------")
            for param in params:
                print(param)
            print("\n---------------Error Details --------------")
            raise AssertionError(stdoutput)

        else:
            print(" Following Data updated in mq.properties file: -----------------")
            for param in params:
                print(param)
            return "PASS"

    def get_link_in_log(self, report_name, report_path):
        """ Used to add any link in Robot Log File. Link points to any html page.

         |Arguments|

         'report_name' is used to display custom  name for report.

         'report_path' is the location of Report, which is given as href attribute of html anchor tag '<a>'
        """
        print("*HTML* Click to view Report <a href=" + report_path + ">" + report_name + "</a>")

    def verify_csv(self, filename, *params):
        """ Used to verify the data present in csv file(comma separated).

         If all the expected values are equal to actual values then the keyword returns PASS
         else will raise an error and display all the differences in comparision.
         |Arguments|

         'filename' = csv file path

         '*params' = list of data to be verified.

         Input Format for *param : Create a list with items present in key=value format. Example TradeID=23

         NOTE: First item in *params should contain an unique element, that will be use to identify the correct row and all other values will be verified from that row.

        Example:
        | ***Variable*** |
        | @{DataToVerify} | TradeID=Cre:123 | NotionalAmount=1223.56 | Collateral=1.234 |
        | ***TestCases*** |
        | ${verifyStatus} | Verify Csv | ${csvFilename} | @{DataToVerify} |
        | Should Be Equal As Strings | ${verifyStatus} | PASS |

        """
        if not os.path.exists(filename):
            raise AssertionError("FileNotFound Error!\n '{}' file does not exist.".format(filename))
        param_dict = collections.OrderedDict()  # split params and get parameters and their values to be verified and store in dictionary
        for param in params:
            parameter = str(param).split("=")
            param_dict[parameter[0]] = parameter[1]
            del parameter
        with open(filename, 'r') as f:  # Read data from csv file and store in fileData list
            reader = csv.reader(f)
            fileData = []
            for r, row in enumerate(reader):
                rowdata = [col for c, col in enumerate(row)]
                fileData.append(rowdata)
                del rowdata
            del reader
        headers = fileData[0]
        col_index_dict = collections.OrderedDict()
        colNotFound = []
        for column in list(param_dict.keys()):  # check if colname is present in headers list or not
            index = self._get_column_index_from_header(headers, column)
            if index != -1:
                col_index_dict[column] = index
            else:
                colNotFound.append(column)
        del headers
        if colNotFound:
            raise AssertionError(
                "ColumnNotFound Error...\nBelow Columns are not present in the file:\n{}".format(
                    '\n'.join(colNotFound)))
        # Get the column index of first column-name present in param_dict ,  so that we can search for the correct row , in which that column is present..then in that row only we will search and verify all other parameters
        find_in_other_row = True
        first_column_name = list(param_dict.keys())[0]
        first_column_value = param_dict[first_column_name]
        first_column_index = col_index_dict[first_column_name]
        valueNotFound = []
        for row in range(1, len(fileData)):
            if not find_in_other_row:
                break
            row_data = fileData[row]
            actualValue = row_data[first_column_index]
            if actualValue.strip() == first_column_value.strip():  # correct row is found , now don't search for the parameters in some other rows
                find_in_other_row = False  # now check for other column values present in this row
                for i in range(1, len(param_dict)):
                    column = list(param_dict.keys())[i]
                    value = param_dict[column]
                    col_index = col_index_dict[column]
                    actual_value = row_data[col_index]
                    if actual_value.strip() != value.strip():
                        valueNotFound.append([column, value, actual_value])
            else:
                find_in_other_row = True
        del fileData
        if find_in_other_row:
            raise AssertionError("No unique row found with key {} = {} ".format(first_column_name, first_column_value))
        elif valueNotFound:
            raise AssertionError("Values Not Found !! \n".format(self._failed_report(valueNotFound)))
        else:
            return 'PASS'

    def _get_column_index_from_header(self, headers, column):
        """ Used to return position of the columns present in headers.If column is not present in table headers then it will return -1.
            It is a private method and called inside verify_csv method.
            Argument: 'headers' is list of headers present in csv file
                      'column' is the column name.
        """
        index = -1
        indexFound = False
        for header in headers:
            index = index + 1
            if header.lower().strip() == column.lower().strip():
                indexFound = True
                break
        if indexFound:  # check for indexFound flag
            return index
        else:
            return -1

    def _failed_report(self, data):
        TableHeaderColor = 'red'
        cellcolor = 'Linen'
        Headers = [HTML_external.TableCell('Column', bgcolor=TableHeaderColor),
                   HTML_external.TableCell('Expected Value', bgcolor=TableHeaderColor),
                   HTML_external.TableCell('Actual Value', bgcolor=TableHeaderColor)
                   ]
        TableData = HTML_external.Table(header_row=Headers)
        for d in data:
            TableData.rows.append([HTML_external.TableCell(d[0], bgcolor=cellcolor),
                                   HTML_external.TableCell(d[1], bgcolor=cellcolor),
                                   HTML_external.TableCell(d[2], bgcolor=cellcolor)])
        logger.info("{}\n".format(TableData), True)

    def compare_xls_or_xlsx(self, file1, file2):
        """ Used to compare xls and xlsx file.
        """
        rb1 = xlrd.open_workbook(file1)
        rb2 = xlrd.open_workbook(file2)
        sheet1 = rb1.sheet_by_index(0)
        sheet2 = rb2.sheet_by_index(0)
        rowNotFound = ""
        for rownum in range(max(sheet1.nrows, sheet2.nrows)):
            if rownum < sheet1.nrows:
                row_rb1 = sheet1.row_values(rownum)
                row_rb2 = sheet2.row_values(rownum)
                for colnum, (c1, c2) in enumerate(zip_longest(row_rb1, row_rb2)):
                    if c1 != c2:
                        rowNotFound += "Row {} Col {} - {} != {}".format(rownum + 1, colnum + 1, c1, c2) + "\n"
            else:
                print("Row {} missing".format(rownum+1))
        if rowNotFound != "":
            raise AssertionError("Files are different.\n...................................\n" + str(rowNotFound))
        rb1.release_resources()
        rb2.release_resources()

    def FindTag_ReplaceValue_xml(self, xml_file_name, xml_tag_name, text_to_replace):
        """ |Usage|
         To replace tag value in xml file.

         |Argument|

         'xml_file_name' : xml file path

         'xml_tag_name' : xml tag name

         'text_to_replace' : tag value to replace
        """
        xml_tree = xml.etree.ElementTree.parse(xml_file_name)
        occurence_list = xml_tree.findall('.//' + xml_tag_name)
        for tag in occurence_list:
            tag.text = text_to_replace
        xml_tree.write(xml_file_name)

    def jolokia_request(self, url, mbean="java.lang:type=Memory", attribute="HeapMemoryUsage"):
        """ |Usage|
         To get memory usage using Jolokia request.

         |Argument|
         'url' : Jolokia url

         'mbean' : By default, java.lang:type=Memory

         'attribute' : By default,HeapMemoryUsage


        """
        j4p = Jolokia(url)
        data = j4p.request(type='read', mbean=mbean, attribute=attribute)
        logger.debug("Response {}".format(data['value']))
        return data['value']

    def get_matching_patterns(self, data, patterns):
        """ |Usage|
            Used to get the matching patterns using regular expression.
            Multiple patterns can be also passed.
            It returns a list of all the matches.

            |Argument|
            'data' : Pass list of data
            'pattern' : Pass list of patterns to be searched.

            """
        all_pattern_matches = []
        match_pattern = []
        for line in data:
            for pattern in patterns:
                matchResult = re.findall(pattern, line)
                if matchResult:
                    if pattern not in match_pattern:
                        match_pattern.append(pattern)
                    for m in matchResult:
                        all_pattern_matches.append(m)
            for p in patterns:
                if p not in match_pattern:
                    logger.warn("Pattern : {} not found !!".format(p))
            return all_pattern_matches

    def delete_file_if_exist(self, filename):
        """ This keyword is used to delete a file with "filename". It deletes the file only if it exist in the location """
        if os.path.exists(filename):
            os.remove(filename)
        else:
            print(("{} file does not exist".format(filename)))

    def convert_csv_to_xlsx(self, csvfile, xlsxfile):
        """ This keyword converts a file format from csv to xlsx"""
        if os.path.exists(csvfile):
            workbook = Workbook()
            worksheet = workbook.active
            with open(csvfile, 'r') as f:
                reader = csv.reader(f)
                for r, row in enumerate(reader):
                    for c, col in enumerate(row):
                        for idx, val in enumerate(col.split(',')):
                            cell = worksheet.cell(row=r + 1, column=c + 1)
                            cell.value = val
                    workbook.save(xlsxfile)
        else:
            print(("{} file does not exist".format(csvfile)))

    def fin_update_properties_file(self, file_name, *string_to_replace):
        """| Usage |
        It updates the values for the existing keys in the .properties file

        | Arguments |

        'file_name' = .properties file location

        'string_to_replace' = list of the elements that needs to be updated in the format "key=value".

        Example:

        |***Variable*** |
        |@{string_to_replace} | loan.repricing.limitvalidation.flag = noterror | loan.repricing.autoRollOver.validation.enabled=null |
        |***TestCases*** |
        |Fin Update Properties File | C:/loan.properties | @{string_to_replace}

       """

        key_list = []
        element_dict = {}

        if not os.path.exists(file_name):
            raise AssertionError("File not found error :" + file_name + " doesn't exist")
        else:
            with open(file_name, 'r') as f:
                filedata = f.readlines()
        for element in string_to_replace:
            element_dict[element.split("=")[0].strip()] = element.split("=")[1].strip()
        for key, v in list(element_dict.items()):
            found = False
            for index in range(len(filedata)):
                if not filedata[index].startswith('#' or '!') or filedata[index] == "":
                    k = filedata[index].split("=")[0].strip()
                    if k.lower() == key.lower():
                        filedata[index] = "{}={} \n".format(k, v)
                        found = True
                        break
            if not found:
                key_list.append(key)
        with open(file_name, "w")as f:
            for d in filedata:
                f.writelines(d)
        if not key_list:
            print("File is modified successfully")
        else:
            error_info = "Following keys are not found in the file :"
            for i in key_list:
                error_info = error_info + "\n" + i
            raise AssertionError(error_info)

    def url_decoder(self, url):

        """| Usage |
        It takes encoded url directly or location of file where url is present and returns the decoded url as an output.

        | Arguments |

        'url' = Pass the encoded url directly or pass location of file where url is present

        | Returns |
        Returns the decoded url

        Example:

        |***TestCases*** |
        1. Decode url \n
        ${Decoded_url} | Url Decoder | url=https%3A%2F%2Fwww.finastra.com%2F
        | Log | ${Decoded_url} |

        2. Decode url from the Filepath \n
        ${Decoded_url} | Url Decoder | D:/Projects/endoded_url.xml
        | Log | ${Decoded_url} |
       """
        if os.path.isfile(url):
            list_data = []
            with open(url, 'r') as f:
                filedata = f.readlines()
                for url in filedata:
                    list_data.append(urllib.parse.unquote(url))
            decoded_url = ''.join(list_data)
        else:
            decoded_url = urllib.parse.unquote(url)
        return decoded_url

    def build_hash(self, input):

        """| Usage |
        It is used to get the hash value for the input of any string and returns the encoded value

        | Arguments |

        'input' = Data input whose hash value is required

        | Returns |
        Returns the encoded data

        Example:

        |***TestCases*** |

        ${return_value}    Build Hash    34234342
       """
        try:
            obj = hashlib.sha256()
        except Exception as e:
            raise AssertionError("cryptographic algorithm is requested but is not available. {}\n".format(e))
        try:
            obj.update(input.encode("UTF-8"))
            hash_data = obj.hexdigest()
        except Exception as e:
            raise AssertionError("Unsupported Encoding Exception: {}\n".format(e))
        generateEsignField = re.sub("\D", "", hash_data)
        hashvalue = generateEsignField[::-1]
        if len(hashvalue) < 8:
            return hashvalue
        else:
            return hashvalue[0:8]

    def get_data_from_property_file(self, file_name):
        """| Usage |
            Gets data from .property file and returns dictionary.
            For more information on dictionary: https://github.com/robotframework/robotframework/blob/master/atest/testdata/standard_libraries/collections/dictionary.robot
            | Arguents |
            'file_name' = .property file location
                Example:
                    |***TestCases*** |
                    ${dict} | Get Data From Property File | ${filename}
                    ${value} | Get From Dictionary | ${dict} | ${key}
        """
        if not os.path.exists(file_name):
            raise AssertionError("File not found error :" + file_name + " doesn't exist")
        else:
            try:
                p = Properties()
                p.load(open(file_name))
                return p
            except:
                raise AssertionError("Error in Properties file. Please pass standard properties files")

    def get_differences_of_two_lists(self, list1, list2):
        """| Usage |
         Return the difference of 2 lists.

          | Arguments |

         list1 & list2
         """
        return list(set(list1) - set(list2))
