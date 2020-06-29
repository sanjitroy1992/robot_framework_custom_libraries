from Libraries.Selenium import SeleniumGroup
import os
import pandas as pd
from random import random
import lxml.html as html
from collections import defaultdict
import xlrd
from Libraries.Common.BaselineComparator import BaselineComparator
from HttpLibrary import logger
import HTML_external


class BaselineVerification(SeleniumGroup):
    def _report_html_update(self):
        TableHeaderColor = "LightCoral"
        global cellcolor, ReportHtml
        cellcolor = "white"
        ReportHtml = HTML_external.HTML_external.Table()
        ReportHtml.rows.append(
            [HTML_external.HTML_external.TableCell('Field Name', bgcolor=TableHeaderColor, header=True),
             HTML_external.HTML_external.TableCell('Expected Value', bgcolor=TableHeaderColor, header=True),
             HTML_external.HTML_external.TableCell('Actual Value', bgcolor=TableHeaderColor, header=True)])
        return ReportHtml

    def _update_table(self,key=None, expected=None, actual=None):
        ReportHtml.rows.append(
            [HTML_external.HTML_external.TableCell(str(key), bgcolor=cellcolor),
             HTML_external.HTML_external.TableCell(str(expected), bgcolor=cellcolor),
             HTML_external.HTML_external.TableCell(str(actual), bgcolor=cellcolor)])

    def _update_dict_if_same_found(self, key, dict):
        if key in dict.keys():
            k1 = key + "_1"
            flag = True
            while flag:
                if k1 in dict.keys():
                    value = k1.split("_")
                    k1 = value[0] + "_" + str(int(value[1]) + 1)
                else:
                    fn_key = k1
                    flag = False
        else:
            fn_key = key
        return fn_key



    def verify_view_page_data(self, table_id, input_datafile, rowid):
        """| Usage |
                It compares the application View page table.

                  | Arguments |

                 'locator' = locator of the table. It's needed to provide id of table.

                 'input_datafile' = Baseline excel file location

                 'rowid' = rowid of the baseline file

            Example :
                Open Browser | URL  | Chrome
                Verify View Page Data | //table[@id="tbl_UIMsg"] | E:\\_UTF_demo\\VERIFYVIEWPAGE_Data.xls | rowid=1

                """
        if not os.path.exists(input_datafile):
            raise AssertionError("File not found error: {} file does not exist".format(input_datafile))
        dict_UI = {}
        theirs_value = ""
        seq_value = ""
        header_list = self.seleniumlib.get_webelements(table_id + "//tr[1]//td")
        all_header = []
        for k in header_list:
            header_text = self.seleniumlib.get_text(k)
            all_header.append(header_text)
        ours_index = int(all_header.index("Ours")) + 1

        fieldname_index = int(all_header.index("Field Name")) + 1
        if "Theirs" in all_header:
            theirs_value = "Found"
            j = int(all_header.index("Theirs")) + 1
        row_list = self.seleniumlib.get_webelements(table_id + '//tr')
        for i in range(2, len(row_list) - 1):
            try:
                class_value = self.seleniumlib.get_element_attribute(table_id + '//tr[' + str(i) + ']', "class")
            except:
                continue
            try:
                self.seleniumlib._element_find('//table[@id="tbl_UIMsg"]//tr[' + str(i) + ']//td[contains(@class,"clsKeyName")]')
                continue
            except:
                if class_value == "Subsequence" or class_value == "Sequence":
                    sequence_text = self.seleniumlib.get_text(table_id + '//tr[' + str(i) + ']//td['+str(fieldname_index)+']')
                    seq_value = sequence_text
                    continue
            if seq_value != "":
                fieldname = table_id + '//tr[' + str(i) + ']//td[' + str(fieldname_index) + ']'
                fieldname_val = self.seleniumlib.get_text(fieldname)

                cell_ours = table_id + '//tr[' + str(i) + ']//td[' + str(ours_index) + ']'
                ours_text = self.seleniumlib.get_text(cell_ours)
                k1 = seq_value.strip() + '%' + fieldname_val.strip()
                fn_key = self._update_dict_if_same_found(k1, dict_UI)
                if ours_text.strip() != '':
                    if theirs_value != '':
                        theirs_text = self.seleniumlib.get_text(table_id + '//tr[' + str(i) + ']//td[' + str(j) + ']')
                        dict_value = ours_text.strip() + '%' + theirs_text.strip()
                        dict_UI[fn_key] = dict_value
                        continue
                    else:
                        dict_UI[fn_key] = ours_text
                        continue
                else:
                    if theirs_value != '':
                        theirs_text = self.seleniumlib.get_text(table_id + '//tr[' + str(i) + ']//td[' + str(j) + ']')
                        if theirs_text.strip() != "":
                            dict_value = ours_text.strip() + '%' + theirs_text.strip()
                            dict_UI[fn_key] = dict_value
                        else:

                            for a in range(1,len(row_list)):
                                try:
                                    self.seleniumlib._element_find('//table[@id="tbl_UIMsg"]//tr[' + str(i + a) + ']//td[contains(@class,"clsKeyName")]')
                                    cell_ours = table_id + '//tr[' + str(i + a) + ']//td[' + str(ours_index) + ']'
                                    ours_text = self.seleniumlib.get_text(cell_ours)
                                    k1 = seq_value.strip() + '%' + fieldname_val.strip() + '%' + str(a)
                                    fn_key = self._update_dict_if_same_found(k1, dict_UI)
                                    if theirs_value != '':
                                        theirs_text = self.seleniumlib.get_text(table_id + '//tr[' + str(i + a) + ']//td[' + str(j) + ']')
                                        dict_value = ours_text.strip() + '%' + theirs_text.strip()
                                        dict_UI[fn_key] = repr(dict_value).replace("'", "")

                                    else:
                                        dict_UI[fn_key] = ours_text
                                except:break
                    else:
                        for a in range(1, len(row_list)):
                            try:
                                self.seleniumlib._element_find('//table[@id="tbl_UIMsg"]//tr[' + str(i + a) + ']//td[contains(@class,"clsKeyName")]')
                                cell_ours = table_id + '//tr[' + str(i + a) + ']//td[' + str(ours_index) + ']'
                                ours_text = self.seleniumlib.get_text(cell_ours)
                                k1 = seq_value.strip() + '%' + fieldname_val.strip() + '%' + str(a)
                                fn_key = self._update_dict_if_same_found(k1, dict_UI)
                                dict_UI[fn_key] = ours_text
                            except:break


            elif class_value == "EVEN" or class_value == "ODD":
                cell_fieldname = table_id + '//tr[' + str(i) + ']//td[' + str(fieldname_index) + ']'
                cell_ours = table_id + '//tr[' + str(i) + ']//td[' + str(ours_index) + ']'
                fieldname_value = self.seleniumlib.get_text(cell_fieldname)
                ours_text = self.seleniumlib.get_text(cell_ours)
                if theirs_value != '':
                    theirs_text = self.seleniumlib.get_text(table_id + '//tr[' + str(i) + ']//td[' + str(j) + ']')
                    dict_value = ours_text.strip() + '%' + theirs_text.strip()
                    dict_UI[fieldname_value.strip()] = dict_value.strip()
                else:
                    dict_UI[fieldname_value.strip()] = ours_text.strip()
        print (dict_UI)
        excel_data = pd.read_excel(input_datafile, dtype=str, keep_default_na=False)
        data_list = excel_data.dropna(axis=1, how='all').to_dict(orient='records')
        excelfile_rowid = ""
        for datafile_dict in data_list:
            if (datafile_dict['rowid']) == rowid:
                excelfile_rowid = "found"
                break
        if excelfile_rowid == "":
            raise AssertionError("Rowid {} not found".format(rowid))
        del datafile_dict['rowid']
        d2 = datafile_dict
        for key in list(d2):
            if not d2[key]:
                d2.pop(key)

        for key, value in d2.items():
            if "\\\\n" in value:
                d2[key] = value.replace("\\\\n", "\\n")
        d1_keys = set(dict_UI.keys())
        d2_keys = set(d2.keys())
        intersect_keys = d1_keys.intersection(d2_keys)
        removed = d2_keys - d1_keys
        res = dict.fromkeys(removed, 0)
        if list(res) != []:
            raise AssertionError("Missing header in table: " + str(list(res)))
        modified = {o: (dict_UI[o], d2[o]) for o in intersect_keys if dict_UI[o] != d2[o]}
        keys = modified.keys()
        ReportHtml = self._report_html_update()
        if keys != set():
            for i in keys:
                base_value = (modified.get(i))[1]
                actual_value = (modified.get(i))[0]
                self._update_table(key=str(i), expected=str(base_value), actual=str(actual_value))
            logger.info("Below are the mismatched value for the keys from UI")
            logger.info(ReportHtml, html=True)
            raise AssertionError("Differences in baseline data ")
        else:
            print ("No differences in the baseline data")

    def _read_xls_data(self, excel_file_path=None, sheet_name=None):
        if not os.path.isfile(excel_file_path):
            raise AssertionError(
                "Invalid Input Error ! \nExcel File Path: {} does not exist.".format(excel_file_path))
        excel_data = pd.read_excel(excel_file_path, sheet_name=sheet_name, dtype=str)
        data_list = excel_data.to_dict(orient='record')
        data_dict = self._create_defaultdict_with_list_values(data_list)
        return data_dict

    def _create_defaultdict_with_list_values(self, data_list):
        data_dict = defaultdict(list)
        for i in data_list:
            for key, value in i.items():
                if type(key) == int:
                    data_dict[key].append(value)
                else:
                    data_dict[key.strip()].append(value)
        return dict(data_dict)

    def _create_app_xlsx_file(self, app_data_dict, baseline_excel_file_path, baseline_sheet_name, skip_columns_names,
                              sort_data_in_file):

        app_excel_export_filepath = os.getcwd() + "\\" + "Test_" + str(random()) + ".xlsx"
        a = pd.DataFrame(app_data_dict)
        writer = pd.ExcelWriter(app_excel_export_filepath, engine='xlsxwriter')
        a.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()
        BaselineComparator().mx_compare_xls_files(app_file=app_excel_export_filepath,
                                                  baseline_file=baseline_excel_file_path,
                                                  app_sheet=baseline_sheet_name, baseline_sheet='Sheet1',
                                                  skip_columns_names=skip_columns_names,
                                                  sort_data_in_file=sort_data_in_file)

        if os.path.exists(app_excel_export_filepath):
            os.remove(app_excel_export_filepath)

    def _check_prerequisite(self, html_file_path, html_table_xpath, primary_column_name):
        if not os.path.isfile(html_file_path):
            raise AssertionError(
                "Invalid Input Error ! \nHTML File Path: {} does not exist.".format(html_file_path))
        if not html_table_xpath:
            raise AssertionError("Either Invalid Xpath: {} or not passed".format(html_table_xpath))
        if not primary_column_name:
            raise AssertionError("Either Invalid primary column name: {} or not passed".format(primary_column_name))

    def verify_html_single_table_data(self, html_file_path=None, html_table_xpath=None,
                                          baseline_excel_file_path=None, baseline_sheet_name='Sheet1',
                                          primary_column_name=None, skip_columns_names=None, rowid=None,
                                          sort_data_in_file=False):
        """
        | Usage |
                To Verify HTML data from a single HTML table in application.

                  | Arguments |

                 ''html_file_path' = HTML file path.

                 'html_table_xpath' = XPATH of the html single table. TThis xpath should contain only one table in the entire html page.

                 'baseline_excel_file_path' = Baseline excel file path.

                 'baseline_sheet_name'[Optional] = Baseline file sheet name. By default the value is set to 'Sheet1'.

                 'primary_column_name' = Data compared with respect to this column name present in baseline excel file.
                                Values present in this column are compared with html file data and if not present
                                in html page then values are reported in the log file in the form of a list.

                 'rowid'[Optional] = To verify specific row of baseline excel.By default the value is set to None
                                         to verify the entire rows present in baseline

                 'skip_columns_names'[Optional] = Column names to be skipped while comparision.
                                Needs to be passed in the form of a list, see in example. By default the value is set to None.

                 'sort_data_in_file'[Optional] = Data for both files to be sorted for comparision.
                                Needs to be passed as "True". By default the value is set to "False".

            Example 1 :
                Verify Html Single Table Data | html_file_path=${EXECDIR}\\DatasetFiles\\report.HTML | html_table_xpath=//table//div[text()="Ours"]/../../.. | baseline_excel_file_path=${EXECDIR}\\DatasetFiles\\baseline1.xlsx | primary_column_name=Reference Num

            Example 2 (Passing specific sheet from baseline. eg:'Sheet2'):
                Verify Html Single Table Data | html_file_path=${EXECDIR}\\DatasetFiles\\report.HTML | html_table_xpath=//table//div[text()="Ours"]/../../.. | baseline_excel_file_path=${EXECDIR}\\DatasetFiles\\baseline1.xlsx | baseline_sheet_name=Sheet2 | primary_column_name=Reference Num

            Example 3 (Without baseline column headers):
                Verify Html Single Table Data | html_file_path=${EXECDIR}\\DatasetFiles\\report.HTML | html_table_xpath=//table//div[text()="Ours"]/../../.. | baseline_excel_file_path=${EXECDIR}\\DatasetFiles\\baseline1.xlsx | primary_column_name=10

            Example 4 (Verify with only one row from baseline excel.):
                Verify Html Single Table Data | html_file_path=${EXECDIR}\\DatasetFiles\\report.HTML | html_table_xpath=//table//div[text()="Ours"]/../../.. | baseline_excel_file_path=${EXECDIR}\\DatasetFiles\\baseline1.xlsx | primary_column_name=10  |  rowid=3

            Example 5 (Verify with selected multiple rows from baseline excel.):
                Verify Html Single Table Data | html_file_path=${EXECDIR}\\DatasetFiles\\report.HTML | html_table_xpath=//table//div[text()="Ours"]/../../.. | baseline_excel_file_path=${EXECDIR}\\DatasetFiles\\baseline1.xlsx | primary_column_name=10  |  rowid=2,5,7

            Example 6 (Verify with specified range of rows from baseline excel.):
                Verify Html Single Table Data | html_file_path=${EXECDIR}\\DatasetFiles\\report.HTML | html_table_xpath=//table//div[text()="Ours"]/../../.. | baseline_excel_file_path=${EXECDIR}\\DatasetFiles\\baseline1.xlsx | primary_column_name=10  |  rowid=3-6

            Note: For table without the headers. Start naming the column headers name from value 1 in baseline file.
        """
        self._check_prerequisite(html_file_path, html_table_xpath, primary_column_name)
        excel_dict = self._read_xls_data(baseline_excel_file_path, baseline_sheet_name)
        b = pd.DataFrame(excel_dict)
        if rowid:
            excel_dict = self._select_specified_baseline_rows(b, rowid)

        with open(html_file_path, 'r') as f:
            html_content = f.read()
            root = html.fromstring(html_content)

        table_rows = html_table_xpath
        try:
            rows = root.xpath(table_rows + '//tr')
        except:
            raise AssertionError("Invalid Xpath: {}".format(html_table_xpath))
        headers = []
        all_data = []
        if str(primary_column_name).replace(" ", "").strip().isalpha():
            for td in rows[2].xpath('td/div'):
                headers.append(td.text)
            for row in rows[3:]:
                d = {}
                for index, td in enumerate(row.xpath('td')):
                    d[headers[index]] = str(td.text_content()).replace(u'\xa0', u'')
                all_data.append(d)
            data_dict = self._create_defaultdict_with_list_values(all_data)
            try:
                excel_dict[primary_column_name]
            except KeyError:
                raise AssertionError('Primary key name: {} not valid'.format(primary_column_name))
        elif int(primary_column_name):
            for row in rows[1:]:
                d = {}
                for index, td in enumerate(row.xpath('td'), start=1):
                    d[index] = str(td.text_content()).replace(u'\xa0', u'')
                all_data.append(d)
            data_dict = self._create_defaultdict_with_list_values(all_data)
            try:
                primary_column_name = int(primary_column_name)
                excel_dict[primary_column_name]
            except KeyError:
                raise AssertionError('Primary key name: {} not valid'.format(primary_column_name))
        else:
            raise AssertionError('Primary key name: {} not valid'.format(primary_column_name))

        baseline_key = excel_dict[primary_column_name]
        html_dict = dict(data_dict)
        index_list = []
        not_found = []
        if primary_column_name in html_dict.keys():
            for key in baseline_key:
                list1 = html_dict[primary_column_name]
                if key not in list1:
                    not_found.append(key)
                else:
                    index_list.append(list1.index(key))
        data_dict = defaultdict(list)
        for key, value in html_dict.items():
            list_values = []
            for index in index_list:
                list_values.append(html_dict[key][index])
                data_dict[key] = list_values
        for key, value in excel_dict.items():
            indices = [i for i, x in enumerate(value) if str(x) == "nan"]
            for i in indices:
                data_dict[key][i] = ''
        app_data_dict = dict(data_dict)
        if len(not_found) != 0:
            raise AssertionError("{} primary key values not found in html page".format(','.join(not_found)))

        if rowid:
            baseline_excel_export_filepath = os.getcwd() + "\\" + "Test_Baseline" + str(random()) + ".xlsx"
            a = pd.DataFrame(excel_dict)
            writer = pd.ExcelWriter(baseline_excel_export_filepath, engine='xlsxwriter')
            a.to_excel(writer, sheet_name='Sheet1', index=False)
            writer.save()
            self._create_app_xlsx_file(app_data_dict, baseline_excel_export_filepath, 'Sheet1', skip_columns_names,
                                       sort_data_in_file)
            if os.path.exists(baseline_excel_export_filepath):
                os.remove(baseline_excel_export_filepath)
        else:
            self._create_app_xlsx_file(app_data_dict, baseline_excel_file_path, baseline_sheet_name, skip_columns_names,
                                       sort_data_in_file)

    def verify_html_all_table_data(self, html_file_path=None, html_table_xpath=None,
                                       baseline_excel_file_path=None,
                                       baseline_sheet_name='Sheet1',
                                       primary_column_name=None,
                                       skip_columns_names=None, rowid=None,
                                       sort_data_in_file=False):
        """
                | Usage |
                        To Verify HTML data from all the HTML tables in the application.

                          | Arguments |

                         ''html_file_path' = HTML file path.

                         'html_table_xpath' = XPATH of all the html tables. This xpath should contain all the tables present in the html page.

                         'baseline_excel_file_path' = Baseline excel file path.

                         'baseline_sheet_name'[Optional] = Baseline file sheet name. By default the value is set to 'Sheet1'.

                         'primary_column_name' = Data compared with respect to this column name present in baseline excel file.
                                        Values present in this column are compared with html file data and if not present
                                        in html page then values are reported in the log file in the form of a list.

                         'rowid'[Optional] = To verify specific row of baseline excel.By default the value is set to None
                                         to verify the entire rows present in baseline

                         'skip_columns_names'[Optional] = Column names to be skipped while comparision.
                                        Needs to be passed in the form of a list, see in example. By default the value is set to None.

                         'sort_data_in_file'[Optional] = Data for both files to be sorted for comparision.
                                        Needs to be passed as "True". By default the value is set to "False".

                    Example 1 :
                        Verify Html All Table Data | html_file_path=${EXECDIR}\\DatasetFiles\\report.HTML | html_table_xpath=//div[contains(text(),"Deal Number")]/../../.. | baseline_excel_file_path=${EXECDIR}\\DatasetFiles\\baseline2.xlsx | primary_column_name=Sender's Reference

                    Example 2 (Passing specific sheet from baseline. eg:'Sheet2'):
                        Verify Html All Table Data | html_file_path=${EXECDIR}\\DatasetFiles\\report.HTML | html_table_xpath=//div[contains(text(),"Deal Number")]/../../.. | baseline_excel_file_path=${EXECDIR}\\DatasetFiles\\baseline2.xlsx | baseline_sheet_name=Sheet2 | primary_column_name=Sender's Reference

                    Example 3 (Verify with only one row from baseline excel.) :
                        Verify Html All Table Data | html_file_path=${EXECDIR}\\DatasetFiles\\report.HTML | html_table_xpath=//div[contains(text(),"Deal Number")]/../../.. | baseline_excel_file_path=${EXECDIR}\\DatasetFiles\\baseline2.xlsx | primary_column_name=Sender's Reference  |  rowid=3

                    Example 4 (Verify with selected multiple rows from baseline excel.):
                        Verify Html All Table Data | html_file_path=${EXECDIR}\\DatasetFiles\\report.HTML | html_table_xpath=//div[contains(text(),"Deal Number")]/../../.. | baseline_excel_file_path=${EXECDIR}\\DatasetFiles\\baseline2.xlsx | primary_column_name=Sender's Reference  |  rowid=2,5,7

                    Example 5 (Verify with specified range of rows from baseline excel.) :
                        Verify Html All Table Data | html_file_path=${EXECDIR}\\DatasetFiles\\report.HTML | html_table_xpath=//div[contains(text(),"Deal Number")]/../../.. | baseline_excel_file_path=${EXECDIR}\\DatasetFiles\\baseline2.xlsx | primary_column_name=Sender's Reference  |  rowid=3-6

                """
        self._check_prerequisite(html_file_path, html_table_xpath, primary_column_name)
        excel_dict = self._read_xls_data(baseline_excel_file_path, baseline_sheet_name)
        b = pd.DataFrame(excel_dict)
        if rowid:
            excel_dict = self._select_specified_baseline_rows(b, rowid)

        with open(html_file_path, 'r') as f:
            html_content = f.read()
            root = html.fromstring(html_content)
        tables = root.xpath(html_table_xpath)
        headers = {}

        all_data_list = []
        for i in range(1, len(tables)):
            try:
                rows = root.xpath('(' + html_table_xpath + ')[{}]//tr'.format(i))
            except:
                raise AssertionError("Invalid Xpath: {}".format(html_table_xpath))
            for row in rows[4:-1]:
                td = row.xpath('td')
                for index in range(1, len(td) - 1):
                    if td[1].text_content() != '\xa0':
                        if td[1].text_content() not in headers:
                            headers[td[1].text_content()] = ''
        for i in range(1, len(tables)):
            try:
                rows = root.xpath('(' + html_table_xpath + ')[{}]//tr'.format(i))
            except:
                raise AssertionError("Invalid Xpath: {}".format(html_table_xpath))
            data = []
            for row in rows[4:-1]:
                td = row.xpath('td')
                list1 = []
                for index in range(1, len(td) - 1):
                    if td[index].text_content() != '\xa0':
                        list1.append(td[index].text_content())
                data.append(list1)

            all_data = [x for x in data if x != []]
            dict1 = {}
            for list_item in all_data:
                if list_item[0] not in dict1:
                    dict1[list_item[0]] = '%'.join(list_item[1:])
                else:
                    value = str('%'.join(list_item[1:]))
                    dict1[list_item[0]] = str(dict1[list_item[0]]) + '{}'.format('&') + value

            key_value = {}
            for key, value in headers.items():
                if key in dict1:
                    key_value[key] = dict1[key]
                else:
                    key_value[key] = ''
            all_data_list.append(key_value)
        html_dict = self._create_defaultdict_with_list_values(all_data_list)

        for key, value in excel_dict.items():
            if key not in html_dict:
                del html_dict[key]
        not_found = []
        index_list = []
        baseline_key = excel_dict[primary_column_name]
        html_key = html_dict[primary_column_name]

        if primary_column_name in html_dict.keys():
            for key in baseline_key:
                if key not in html_key:
                    not_found.append(key)
                else:
                    index_list.append(html_key.index(key))

        final_dict = {}
        for key, value in html_dict.items():
            list_values = []
            for index in index_list:
                list_values.append(html_dict[key][index])
                final_dict[key] = list_values

        for key, value in excel_dict.items():
            indices = [i for i, x in enumerate(value) if str(x) == "nan"]
            for i in indices:
                final_dict[key][i] = ''

        if len(not_found) != 0:
            raise AssertionError("{} primary key values not found in html page".format(','.join(not_found)))

        if rowid:
            baseline_excel_export_filepath = os.getcwd() + "\\" + "Test_Baseline" + str(random()) + ".xlsx"
            a = pd.DataFrame(excel_dict)
            writer = pd.ExcelWriter(baseline_excel_export_filepath, engine='xlsxwriter')
            a.to_excel(writer, sheet_name='Sheet1', index=False)
            writer.save()
            self._create_app_xlsx_file(final_dict, baseline_excel_export_filepath, 'Sheet1', skip_columns_names,
                                       sort_data_in_file)
            if os.path.exists(baseline_excel_export_filepath):
                os.remove(baseline_excel_export_filepath)
        else:
            self._create_app_xlsx_file(final_dict, baseline_excel_file_path, baseline_sheet_name, skip_columns_names,
                                       sort_data_in_file)

    def _find_start_row(self, filename, table_name, primary_column_name):
        wb = xlrd.open_workbook(filename)
        sheet = wb.sheet_by_index(0)
        if table_name != None:
            startrow = table_name
            for i in range(sheet.nrows):
                ptrow = i
                for j in range(sheet.ncols):
                    if sheet.cell_value(i, j) == startrow:
                        ptrow += 1
                        return ptrow
                    else:
                        continue
        else:
            for i in range(sheet.nrows):
                ptrow = i
                for j in range(sheet.ncols):
                    if sheet.cell_value(i, j) == primary_column_name:
                        return ptrow
                    else:
                        continue

    def _select_specified_baseline_rows(self, dataframe_baseline_excel_dict, rowid):
        if rowid:
            if ',' in rowid:
                rowNumbers = rowid.split(",")
                for i in range(len(rowNumbers)):
                    if int(rowNumbers[i]) - 2 >= len(dataframe_baseline_excel_dict.axes[0]):
                        raise AssertionError('Invalid rowid: {} '.format(int(rowNumbers[i])))
                    else:
                        rowNumbers[i] = int(rowNumbers[i]) - 2
                b1 = dataframe_baseline_excel_dict.iloc[rowNumbers]
            elif '-' in rowid:
                rowRange = rowid.split("-")
                startRow = int(rowRange[0])
                EndRow = int(rowRange[1])
                if startRow >= 2 and EndRow <= len(dataframe_baseline_excel_dict.axes[0]) or EndRow == len(
                        dataframe_baseline_excel_dict.axes[0]) + 1:
                    b1 = dataframe_baseline_excel_dict.iloc[startRow - 2:EndRow - 1]
                else:
                    raise AssertionError('Invalid rowid: {} '.format(rowid))
            else:
                rowid = int(rowid) - 2
                if rowid >= len(dataframe_baseline_excel_dict.axes[0]):
                    raise AssertionError('Invalid rowid: {} '.format(rowid + 2))
                elif rowid >= 1:
                    b1 = dataframe_baseline_excel_dict.append(dataframe_baseline_excel_dict, ignore_index=True)
                    b1 = b1.iloc[[0, int(rowid)]]
                    b1 = b1.drop(b1.index[0])
                else:
                    b1 = dataframe_baseline_excel_dict.append(dataframe_baseline_excel_dict, ignore_index=True)
                    b1 = b1.iloc[[0, int(rowid)]]
                    b1 = b1.drop_duplicates()
        baseline_excel_dict = b1.to_dict('list')
        return baseline_excel_dict

    def verify_excel_data(self, app_file, primary_column_name, baseline_file, baseline_sheet=0, table_name=None,
                              rowid=None,
                              skip_columns_names=None, sort_data_in_file=False):
        """
            | Usage |
                    To Verify Excel data from the generated excel report in the application.

                      | Arguments |

                     'app_file' = Generated excel report file path.

                     'primary_column_name' = Data compared with respect to this column name present in baseline excel file.
                                    Values present in this column are compared with generated excel file data and if not present
                                    in generated excel file then values are reported in the log file in the form of a list.

                     'baseline_file' = Baseline excel file path.

                     'baseline_sheet'[Optional] = Baseline file sheet name. By default the value is set to '0', first sheet of workbook.

                     'table_name'[Optional] = To verify specific table from the excel , need to specify table header name
                                            By default the value is set to None if there are no table header name in the file

                     'rowid'[Optional] = To verify specific row of baseline excel.By default the value is set to None
                                         to verify the entire rows present in baseline

                     'skip_columns_names'[Optional] = Column names to be skipped while comparision.
                                    Needs to be passed in the form of a list, see in example. By default the value is set to None.

                     'sort_data_in_file'[Optional] = Data for both files to be sorted for comparision.
                                    Needs to be passed as "True". By default the value is set to "False".

                Example 1 (By default, first sheet of baseline file will be considered for verification if baseline_sheet is not passed):
                    *** Variables ***
                    @{skip_columns}    Create Date    Trade Date
                    | ***TestCases*** |
                    | Verify Excel Data | app_file=${EXECDIR}\\DatasetFiles\\Summary Report.xls |  primary_column_name=Reference Num | baseline_file=${EXECDIR}\\DatasetFiles\\baseline2.xlsx | table_name=Theirs | skip_columns_names=${skip_columns} |

                Example 2 (Passing specific sheet from baseline. eg:'Ours'):
                    *** Variables ***
                    @{skip_columns}    Create Date
                    | ***TestCases*** |
                    | Verify Excel Data | app_file=${EXECDIR}\\DatasetFiles\\Summary Report.xls |  primary_column_name=Reference Num  | baseline_file=${EXECDIR}\\DatasetFiles\\baseline2.xlsx | baseline_sheet=Ours | table_name=Ours | skip_columns_names=${skip_columns} |

                Example 3 (Verify without table_name in generated excel report.):
                    *** Variables ***
                    @{skip_columns}    Create Date
                    | ***TestCases*** |
                    | Verify Excel Data | app_file=${EXECDIR}\\DatasetFiles\\Summary Report.xls |  primary_column_name=Reference Num  | baseline_file=${EXECDIR}\\DatasetFiles\\baseline2.xlsx | baseline_sheet=NoTableNameData | skip_columns_names=${skip_columns} |

                Example 4 (Verify with only one row from baseline excel.):
                    *** Variables ***
                    @{skip_columns}    Create Date
                    | ***TestCases*** |
                    | Verify Excel Data | app_file=${EXECDIR}\\DatasetFiles\\Summary Report.xls |  primary_column_name=Reference Num  | baseline_file=${EXECDIR}\\DatasetFiles\\baseline2.xlsx | rowid=2 | skip_columns_names=${skip_columns} |

                Example 5 (Verify with selected multiple rows from baseline excel.):
                    | Verify Excel Data | app_file=${EXECDIR}\\DatasetFiles\\Summary Report.xls |  primary_column_name=Reference Num  | baseline_file=${EXECDIR}\\DatasetFiles\\baseline2.xlsx | rowid=2,5,7 | skip_columns_names=${skip_columns} |

                Example 6 (Verify with specified range of rows from baseline excel.):
                    | Verify Excel Data | app_file=${EXECDIR}\\DatasetFiles\\Summary Report.xls |  primary_column_name=Reference Num  | baseline_file=${EXECDIR}\\DatasetFiles\\baseline2.xlsx | rowid=2-7 | skip_columns_names=${skip_columns} |
            """

        ################################ Pre-check#################################################################
        if not os.path.isfile(app_file):
            raise AssertionError(
                "Invalid Input Error ! \nActual excel file path: {} does not exist.".format(app_file))
        if not primary_column_name:
            raise AssertionError("Either Invalid primary column name: {} or not passed".format(primary_column_name))
        if not os.path.isfile(baseline_file):
            raise AssertionError(
                "Invalid Input Error ! \nBaseline excel file path: {} does not exist.".format(baseline_file))

        ############################### To select starting point of the excel #####################################
        start = self._find_start_row(app_file, table_name=table_name, primary_column_name=primary_column_name)

        ####################################### actual data #######################################################
        actual_data_dict = pd.read_excel(app_file, header=start, dtype=str)
        actual_data_list = actual_data_dict.to_dict(orient='record')
        actual_data_dict = self._create_defaultdict_with_list_values(actual_data_list)
        a = pd.DataFrame(actual_data_dict)

        ####################################### drop empty columns ###############################################
        a1 = a.dropna(how='all', axis=1)

        ###################################### To consider rows until empty row is found #########################
        try:
            first_row_with_all_NaN = a1[a1.isnull().all(axis=1) == True].index.tolist()[0]
            a1 = a1.loc[:first_row_with_all_NaN - 1]
        except:
            pass
        a1 = a1.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        ###################################### Actual data #######################################################
        actual_data_dict = a1.to_dict('list')

        ###################################### baseline data #####################################################
        baseline_excel_dict = self._read_xls_data(baseline_file, baseline_sheet)
        b = pd.DataFrame(baseline_excel_dict)
        b1 = b.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        if rowid:
            baseline_excel_dict = self._select_specified_baseline_rows(b1, rowid)
        else:
            baseline_excel_dict = b1.to_dict('list')
        ###########################################################################################################
        try:
            baseline_excel_dict[primary_column_name]
        except KeyError:
            raise AssertionError('Invalid Primary key name: {} '.format(primary_column_name))

        try:
            actual_data_dict[primary_column_name]
        except KeyError:
            raise AssertionError('Primary key : {} not present in actual report'.format(primary_column_name))

        not_found_columns = []
        for key, value in baseline_excel_dict.items():
            if key not in actual_data_dict:
                not_found_columns.append(key)

        if len(not_found_columns) != 0:
            raise AssertionError("{} are the list of columns not found in actual excel report".format(
                ','.join(map(str, not_found_columns))))

        ################################# Get only matched data from generated excel report ######################
        not_found = []
        index_list = []
        baseline_key = baseline_excel_dict[primary_column_name]
        baseline_key1 = [str(item) for item in baseline_key]
        html_key = actual_data_dict[primary_column_name]
        html_key1 = [str(item) for item in html_key]

        if primary_column_name in actual_data_dict.keys():
            for key in baseline_key1:
                if key not in html_key1:
                    not_found.append(key)
                else:
                    index_list.append(html_key1.index(key))

        if len(not_found) != 0:
            raise AssertionError("{} are the list of primary key values not found in actual excel report".format(
                ','.join(map(str, not_found))))

        final_dict = {}
        for key, value in actual_data_dict.items():
            list_values = []
            for index in index_list:
                list_values.append(actual_data_dict[key][index])
                final_dict[key] = list_values
        for key, value in baseline_excel_dict.items():
            indices = [i for i, x in enumerate(value) if str(x) == "nan"]
            for i in indices:
                final_dict[key][i] = ''
        ############################## Write to temporary excel and compare it with baseline excel file ############################
        baseline_excel_export_filepath = os.getcwd() + "\\" + "Test_Baseline" + str(random()) + ".xlsx"
        a = pd.DataFrame(baseline_excel_dict)
        writer = pd.ExcelWriter(baseline_excel_export_filepath, engine='xlsxwriter')
        a.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()
        self._create_app_xlsx_file(final_dict, baseline_excel_export_filepath, 'Sheet1', skip_columns_names,
                                   sort_data_in_file)
        if os.path.exists(baseline_excel_export_filepath):
            os.remove(baseline_excel_export_filepath)
