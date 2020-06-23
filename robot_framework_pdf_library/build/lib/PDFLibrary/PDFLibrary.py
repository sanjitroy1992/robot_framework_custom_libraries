import os
from pdfminer3.pdfparser import PDFParser
from pdfminer3.pdfdocument import PDFDocument
from pdfminer3.pdfpage import PDFPage
from pdfminer3.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer3.layout import LAParams, LTTextBox, LTTextLine
from pdfminer3.converter import PDFPageAggregator
from pdfminer3.pdfdocument import PDFPasswordIncorrect


class PDFLibrary(object):
    """
        PDFLibrary validates data in a PDF file and display the report in html table format.

        """

    def _createPDFDoc(self, fpath, password):
        fp = open(fpath, 'rb')
        parser = PDFParser(fp)
        try:
            document = PDFDocument(parser, password)
        except PDFPasswordIncorrect:
            raise AssertionError("Password '{}' is incorrect.".format(password))
        except TypeError:
            raise AssertionError("Unable to extract the pdf. Please check the password.")
        return fp, document

    def _createDeviceInterpreter(self):
        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        return device, interpreter

    def pdf_text_compare(self, filepath, string=None, prefix=None, suffix=None, all_occurrence="False", password=None):

        """ Validates text in a PDF file.


        == Arguments ==

            - filepath: Filepath for PDF file.\n
            - string: String to be validated in the PDF. In case of Multiple strings validation, strings could be passed in the form of a list.\n
            - prefix [Optional]: It is a string prior to "string that needs to be validated" of the same line.\n
            - suffix [Optional]: It is a string subsequent to "string that needs to be validated" of the same line.\n
            - all_occurrence [Optional]: If "True" it searches for all the occurrences of the string that needs to be validated and if found it displays all the occurences in html table. Default value is "False" which displays only the first occurrence of the string.\n
            - password [Optional]: Password for the password protected pdf file.\n

        == Example Test Cases ==

        1. To Validate Text In PDF For First Occurrence :
        |Pdf Text Compare| ${filepath} | ${string} |

        2. To Validate Text In PDF For All Occurrence :
        |Pdf Text Compare| ${filepath} | ${string} | all_occurrence=True |

        3. To Validate Text In PDF For First Occurrence With Prefix And Suffix :
        |Pdf Text Compare| ${filepath} | ${string} | prefix=${Prefix} | suffix=${Suffix} |

        4. To Validate Text In PDF For All Occurrence With Prefix And Suffix :
        |Pdf Text Compare| ${filepath} | ${string} | prefix=${Prefix} | suffix=${Suffix} | all_occurrence=True |

        5. To Validate Text In PDF For All Occurrence With Prefix :
        |Pdf Text Compare| ${filepath} | ${string} | prefix=${Prefix} | all_occurrence=True |

        6. To Validate Text In PDF For All Occurrence With Suffix :
        |Pdf Text Compare| ${filepath} | ${string} | suffix=${Suffix} | all_occurrence=True |

        7. To Validate Text For a PDF Password Protected File With Suffix For All Occurrence :
        |Pdf Text Compare| ${filepath} | ${string} | suffix=${Suffix} | all_occurrence=True | password=123456 |

        """
        if not os.path.exists(filepath):
            raise AssertionError('File {} does not exists'.format(filepath))

        if not string:
            raise AssertionError("Could not find a string to validate")

        str1 = """*HTML*
               <div><div><table>
               <caption><b>PDF Validation</b></caption>
               <tr><td style="background:#5CBFDE;color:#FFFFFF;text-align:center">Text</td>
               <td style="background:#5CBFDE;color:#FFFFFF;text-align:center">Page No.</td>
               <td style="background:#5CBFDE;color:#FFFFFF;text-align:center">Line No.</td>
               <td style="background:#5CBFDE;color:#FFFFFF;text-align:center">Status</td>
               </tr>
               """
        str2 = ""
        if type(string) is not list:
            stringlist = list()
            stringlist.append(string)
        else:
            stringlist = string
        stringnotfound = list()
        for stringToSearch in stringlist:
            stringToSearch = str(stringToSearch)
            fp, document = self._createPDFDoc(filepath, password)
            device, interpreter = self._createDeviceInterpreter()
            i = 0
            flag = 0
            for page in PDFPage.create_pages(document):
                i += 1
                # As the interpreter processes the page stored in PDFDocument object
                interpreter.process_page(page)
                # The device renders the layout from interpreter
                layout = device.get_result()
                # Out of the many LT objects within layout, we are interested in LTTextBox (A set of text objects that are grouped within a certain rectangular area) and LTTextLine
                extracted_text = ""
                for lt_obj in layout:
                    if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
                        extracted_text += lt_obj.get_text()
                list_temp = extracted_text.split("\n")
                j = 0
                for line in list_temp:
                    j += 1
                    if str(line).strip() == "":
                        # Checking for any empty line or images. and removing from count in J.
                        j -= 1
                    if stringToSearch and (stringToSearch in line):
                        if prefix and suffix:
                            if prefix in line and suffix in line:
                                indexPrefix = line.find(prefix)
                                indexSuffix = line.find(suffix)
                                indexStringToSearch = line.find(stringToSearch)
                                if(indexPrefix < indexStringToSearch) and (indexSuffix > indexStringToSearch):
                                    flag = 1
                                    str2 = str2 + """
                                                    <tr><td style="background:#5BB75B;color:#FFFFFF;text-align:center">""" + stringToSearch + """
                                                    </td><td style="background:#5BB75B;text-align:center">""" + str(i) + """
                                                    </td><td style="background:#5BB75B;text-align:center">""" + str(j) + """
                                                    </td><td style="background:#5BB75B;text-align:center">Passed
                                                    </tr>
                                                    """
                                    if all_occurrence.upper() == "FALSE":
                                        break
                        elif not prefix and suffix:
                            if suffix in line:
                                indexSuffix = line.find(suffix)
                                indexStringToSearch = line.find(stringToSearch)
                                if indexSuffix > indexStringToSearch:
                                    flag = 1
                                    str2 = str2 + """
                                                    <tr><td style="background:#5BB75B;color:#FFFFFF;text-align:center">""" + stringToSearch + """
                                                    </td><td style="background:#5BB75B;text-align:center">""" + str(i) + """
                                                    </td><td style="background:#5BB75B;text-align:center">""" + str(j) + """
                                                    </td><td style="background:#5BB75B;text-align:center">Passed
                                                    </tr>
                                                    """
                                    if all_occurrence.upper() == "FALSE":
                                        break

                        elif prefix and not suffix:
                            if prefix in line:
                                indexPrefix = line.find(prefix)
                                indexStringToSearch = line.find(stringToSearch)
                                if indexPrefix < indexStringToSearch:
                                    flag = 1
                                    str2 = str2 + """
                                                    <tr><td style="background:#5BB75B;color:#FFFFFF;text-align:center">""" + stringToSearch + """
                                                    </td><td style="background:#5BB75B;text-align:center">""" + str(i) + """
                                                    </td><td style="background:#5BB75B;text-align:center">""" + str(j) + """
                                                    </td><td style="background:#5BB75B;text-align:center">Passed
                                                    </tr>
                                                    """
                                    if all_occurrence.upper() == "FALSE":
                                        break
                        elif not(prefix and suffix):
                            flag = 1
                            str2 = str2 + """
                                                    <tr><td style="background:#5BB75B;color:#FFFFFF;text-align:center">""" + stringToSearch + """
                                                    </td><td style="background:#5BB75B;text-align:center">""" + str(i) + """
                                                    </td><td style="background:#5BB75B;text-align:center">""" + str(j) + """
                                                    </td><td style="background:#5BB75B;text-align:center">Passed
                                                    </tr>
                                                    """
                            if all_occurrence.upper() == "FALSE":
                                break
                if all_occurrence.upper() == "FALSE":
                    break
            if flag == 0:
                str2 = str2 + """
                                    <tr><td style="background:#D9534F;color:#FFFFFF;text-align:center">""" + stringToSearch + """
                                    </td><td style="background:#D9534F;text-align:center">Not Found
                                    </td><td style="background:#D9534F;text-align:center">Not Found
                                    </td><td style="background:#D9534F;text-align:center">Failed
                                    </tr>
                                    """

                stringnotfound.append(stringToSearch)
        print(str1 + str2)
        fp.close()
        if (len(stringnotfound)>0):
            raise AssertionError("{0} text Not Found.".format(stringnotfound))


    def get_pdf_text(self, filepath, prefix=None, suffix=None, all_occurrence="False", password=None):

        """ Gets text from a PDF file in the form of a list.


        == Arguments ==

            - filepath: Filepath for PDF file.\n
            - prefix [Optional]: It is a string prior to "string that needs to be validated" of the same line.\n
            - suffix [Optional]: It is a string subsequent to "string that needs to be validated" of the same line.\n
            - all_occurrence [Optional]: If "True" it searches for all the occurrences of the string that needs to be validated and if found it displays all the occurences in html table. Default value is "False" which displays only the first occurrence of the string.\n
            - password [Optional]: Password for the password protected pdf file.\n

        == Returns ==
        List of found texts else raises "Text Not Found" if no texts were found.

        == Example Test Cases ==

        1. To Get PDF Text With Prefix And Suffix For First Occurrence :
        | Get Pdf Text | ${filepath} | prefix=${Prefix} | suffix=${Suffix}

        2. To Get PDF Text With Prefix And Suffix For All Occurrence :
        | Get Pdf Text | ${filepath} | prefix=${Prefix} | suffix=${Suffix} | all_occurrence=True

        3. To Get PDF Text With Prefix For All Occurrence :
        | Get Pdf Text | ${filepath} | prefix=${Prefix} | all_occurrence=True

        4. To Get PDF Text With Suffix For All Occurrence :
        | Get Pdf Text | ${filepath} | suffix=${Suffix} | all_occurrence=True

        5. To Get PDF Text In The Form Of a List With Suffix For All Occurrence :
        | @{Stringfound} | Get Pdf Text | ${filepath} | suffix=${Suffix} | all_occurrence=True

        6. To Get PDF Text In The Form Of a List For a Password Protected File With Suffix For All Occurrence :
        | @{Stringfound} | Get Pdf Text | ${filepath} | suffix=${Suffix} | all_occurrence=True | password=123456 |

        """

        if not os.path.exists(filepath):
            raise AssertionError('File {} does not exists'.format(filepath))

        str1 = """*HTML*
               <div><div><table>
               <caption><b>PDF Validation</b></caption>
               <tr><td style="background:#5CBFDE;color:#FFFFFF;text-align:center">Text</td>
               <td style="background:#5CBFDE;color:#FFFFFF;text-align:center">Page No.</td>
               <td style="background:#5CBFDE;color:#FFFFFF;text-align:center">Line No.</td>
               </tr>
               """
        str2 = ""
        stringfound = list()
        fp, document = self._createPDFDoc(filepath, password)
        device, interpreter = self._createDeviceInterpreter()
        i = 0
        for page in PDFPage.create_pages(document):
            i += 1
            # As the interpreter processes the page stored in PDFDocument object
            interpreter.process_page(page)
            # The device renders the layout from interpreter
            layout = device.get_result()
            # Out of the many LT objects within layout, we are interested in LTTextBox (A set of text objects that are grouped within a certain rectangular area) and LTTextLine
            extracted_text = ""
            for lt_obj in layout:
                if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
                    extracted_text += lt_obj.get_text()
            list_temp = extracted_text.split("\n")
            j = 0
            for line in list_temp:
                j += 1
                if str(line).strip() == "":
                    # Checking for any empty line or images. and removing from count in J.
                    j -= 1
                if prefix and suffix:
                    if (prefix and suffix) in line:
                        lenPrefix = len(prefix)
                        indexSuffix = line.find(suffix)
                        indexPrefix = line.find(prefix)
                        strText = line[indexPrefix + lenPrefix:indexSuffix].strip()
                        flag = 1
                        str2 = str2 + """
                                            <tr><td style="background:#5BB75B;color:#FFFFFF;text-align:center">""" + strText + """
                                            </td><td style="background:#5BB75B;text-align:center">""" + str(i) + """
                                            </td><td style="background:#5BB75B;text-align:center">""" + str(j) + """
                                            </tr>
                                            """
                        stringfound.append(strText)
                        if all_occurrence.upper() == "FALSE":
                            break

                elif not suffix and prefix:

                    if prefix in line:
                        lenPrefix = len(prefix)
                        indexPrefix = line.find(prefix)
                        strText = line[indexPrefix + lenPrefix:-1].strip()
                        str2 = str2 + """
                                            <tr><td style="background:#5BB75B;color:#FFFFFF;text-align:center">""" + strText + """
                                            </td><td style="background:#5BB75B;text-align:center">""" + str(i) + """
                                            </td><td style="background:#5BB75B;text-align:center">""" + str(j) + """
                                            </tr>
                                            """
                        stringfound.append(strText)
                        if all_occurrence.upper() == "FALSE":
                            break

                elif not prefix and suffix:
                    if suffix in line:
                        indexSuffix = line.find(suffix)
                        strText = line[0:indexSuffix].strip()
                        str2 = str2 + """
                                            <tr><td style="background:#5BB75B;color:#FFFFFF;text-align:center">""" + strText + """
                                            </td><td style="background:#5BB75B;text-align:center">""" + str(i) + """
                                            </td><td style="background:#5BB75B;text-align:center">""" + str(j) + """
                                            </tr>
                                            """
                        stringfound.append(strText)
                        if all_occurrence.upper() == "FALSE":
                            break
                else:
                    raise AssertionError("Unable to get text. Please provide Prefix and Suffix")
            if all_occurrence.upper() == "FALSE":
                break
        if len(stringfound) == 0:
            raise AssertionError("Text Not Found.")
        print(str1 + str2)
        fp.close()
        return stringfound

    def pdf_text_compare_by_page_and_line_number(self, filepath, string=None, page_no=None, line_no=None, all_occurrence=False, password=None):

        """ Validates Text by page No and Line No in a PDF file.

                == Arguments ==
                    - filepath: Filepath for PDF file.\n
                    - string: String to be validated in the PDF. In case of Multiple strings validation, strings could be passed in the form of a list.\n
                    - page_no : page number from which string to be validated in the PDF\n
                    - line-no [Optional] : line number from which string to be validated in the PDF\n
                    - all_occurrence [Optional]: If "True" it searches for all the occurrences of the string that needs to be validated and if found it displays all the occurences in html table. Default value is "False" which displays only the first occurrence of the string.\n
                    - password [Optional]: Password for the password protected pdf file.\n

                == Example Test Cases ==

                1. To Validate Text In PDF With Page No :
                |Pdf Text Compare| ${filepath} | ${string} | page_no=${Page_No} |

                2. To Validate Text In PDF With Page No For All Occurence :
                |Pdf Text Compare| ${filepath} | ${string} | page_no=${Page_No} | all_occurrence=True |

                3. To Validate Text In PDF With Page No And Line No :
                |Pdf Text Compare| ${filepath} | ${string} | page_no=${Page_No} | line_no=${Line_No} |

                4. To Validate Text For a PDF Password Protected File With Page No For All Occurence :
                |Pdf Text Compare| ${filepath} | ${string} | page_no=${Page_No} | all_occurrence=True | | password=123456 |

                """
        if not os.path.exists(filepath):
            raise AssertionError('File {} does not exists'.format(filepath))

        if not string:
            raise AssertionError("Could not find a string to validate")

        str1 = """*HTML*
                   <div><div><table>
                   <caption><b>PDF Validation</b></caption>
                   <tr><td style="background:#5CBFDE;color:#FFFFFF;text-align:center">Text</td>
                   <td style="background:#5CBFDE;color:#FFFFFF;text-align:center">Page No.</td>
                   <td style="background:#5CBFDE;color:#FFFFFF;text-align:center">Line No.</td>
                   <td style="background:#5CBFDE;color:#FFFFFF;text-align:center">Status</td>
                   </tr>
                   """
        str2 = ""
        if type(string) is not list:
            stringlist = list()
            stringlist.append(string)
        else:
            stringlist = string
        stringnotfound = list()
        for stringToSearch in stringlist:
            stringToSearch = str(stringToSearch)
            fp, document = self._createPDFDoc(filepath, password)
            device, interpreter = self._createDeviceInterpreter()
            i = 0
            flag = 0
            for page in PDFPage.create_pages(document):
                i = i + 1
                # As the interpreter processes the page stored in PDFDocument object
                interpreter.process_page(page)
                # The device renders the layout from interpreter
                layout = device.get_result()
                # Out of the many LT objects within layout, we are interested in LTTextBox (A set of text objects that are grouped within a certain rectangular area) and LTTextLine
                extracted_text = ""
                for lt_obj in layout:
                    if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
                        extracted_text += lt_obj.get_text()
                list_temp = extracted_text.split("\n")
                if i == int(page_no):
                    j = 0
                    for line in list_temp:
                        j = j + 1
                        if str(line).strip() == "":
                            j = j - 1
                        if not line_no and (stringToSearch in line):
                            flag = 1
                            str2 = str2 + """
                                            <tr><td style="background:#5BB75B;color:#FFFFFF;text-align:center">""" + stringToSearch + """
                                            </td><td style="background:#5BB75B;text-align:center">""" + str(i) + """
                                            </td><td style="background:#5BB75B;text-align:center">""" + str(j) + """
                                            </td><td style="background:#5BB75B;text-align:center">Passed
                                            </tr>
                                            """
                            if not all_occurrence:
                                break
                        elif line_no:
                            if (j == int(line_no)) and (stringToSearch in line):
                                flag = 1
                                str2 = str2 + """
                                                <tr><td style="background:#5BB75B;color:#FFFFFF;text-align:center">""" + stringToSearch + """
                                                </td><td style="background:#5BB75B;text-align:center">""" + str(i) + """
                                                </td><td style="background:#5BB75B;text-align:center">""" + str(j) + """
                                                </td><td style="background:#5BB75B;text-align:center">Passed
                                                </tr>
                                                """
                                if not all_occurrence:
                                    break
            if flag == 0:
                str2 = str2 + """
                                    <tr><td style="background:#D9534F;color:#FFFFFF;text-align:center">""" + stringToSearch + """
                                    </td><td style="background:#D9534F;text-align:center">Not Found
                                    </td><td style="background:#D9534F;text-align:center">Not Found
                                    </td><td style="background:#D9534F;text-align:center">Failed
                                    </tr>
                                    """
                stringnotfound.append(stringToSearch)
        print(str1 + str2)
        fp.close()
        if (len(stringnotfound)>0):
            raise AssertionError("{0} text Not Found.".format(stringnotfound))