PDFLibrary
Library scope:	test case
Named arguments:	supported
Introduction
PDFLibrary validates data in a PDF file and display the report in html table format.

Shortcuts
Get Pdf Text · Pdf Text Compare · Pdf Text Compare By Page And Line Number
Keywords
Keyword	Arguments	Documentation
Get Pdf Text	filepath, prefix=None, suffix=None, all_occurrence=False, password=None	
Gets text from a PDF file in the form of a list.

Arguments
filepath: Filepath for PDF file.
prefix [Optional]: It is a string prior to "string that needs to be validated" of the same line.
suffix [Optional]: It is a string subsequent to "string that needs to be validated" of the same line.
all_occurrence [Optional]: If "True" it searches for all the occurrences of the string that needs to be validated and if found it displays all the occurences in html table. Default value is "False" which displays only the first occurrence of the string.
password [Optional]: Password for the password protected pdf file.
Returns
List of found texts else raises "Text Not Found" if no texts were found.

Example Test Cases
1. To Get PDF Text With Prefix And Suffix For First Occurrence :

Get Pdf Text | ${filepath} | prefix=${Prefix} | suffix=${Suffix}
2. To Get PDF Text With Prefix And Suffix For All Occurrence :

Get Pdf Text | ${filepath} | prefix=${Prefix} | suffix=${Suffix} | all_occurrence=True
3. To Get PDF Text With Prefix For All Occurrence :

Get Pdf Text | ${filepath} | prefix=${Prefix} | all_occurrence=True
4. To Get PDF Text With Suffix For All Occurrence :

Get Pdf Text | ${filepath} | suffix=${Suffix} | all_occurrence=True
5. To Get PDF Text In The Form Of a List With Suffix For All Occurrence :

@{Stringfound} | Get Pdf Text | ${filepath} | suffix=${Suffix} | all_occurrence=True
6. To Get PDF Text In The Form Of a List For a Password Protected File With Suffix For All Occurrence :

@{Stringfound}	Get Pdf Text	${filepath}	suffix=${Suffix}	all_occurrence=True	password=123456
Pdf Text Compare	filepath, string=None, prefix=None, suffix=None, all_occurrence=False, password=None	
Validates text in a PDF file.

Arguments
filepath: Filepath for PDF file.
string: String to be validated in the PDF. In case of Multiple strings validation, strings could be passed in the form of a list.
prefix [Optional]: It is a string prior to "string that needs to be validated" of the same line.
suffix [Optional]: It is a string subsequent to "string that needs to be validated" of the same line.
all_occurrence [Optional]: If "True" it searches for all the occurrences of the string that needs to be validated and if found it displays all the occurences in html table. Default value is "False" which displays only the first occurrence of the string.
password [Optional]: Password for the password protected pdf file.
Example Test Cases
1. To Validate Text In PDF For First Occurrence : |Pdf Text Compare| ${filepath} | ${string} |

2. To Validate Text In PDF For All Occurrence : |Pdf Text Compare| ${filepath} | ${string} | all_occurrence=True |

3. To Validate Text In PDF For First Occurrence With Prefix And Suffix : |Pdf Text Compare| ${filepath} | ${string} | prefix=${Prefix} | suffix=${Suffix} |

4. To Validate Text In PDF For All Occurrence With Prefix And Suffix : |Pdf Text Compare| ${filepath} | ${string} | prefix=${Prefix} | suffix=${Suffix} | all_occurrence=True |

5. To Validate Text In PDF For All Occurrence With Prefix : |Pdf Text Compare| ${filepath} | ${string} | prefix=${Prefix} | all_occurrence=True |

6. To Validate Text In PDF For All Occurrence With Suffix : |Pdf Text Compare| ${filepath} | ${string} | suffix=${Suffix} | all_occurrence=True |

7. To Validate Text For a PDF Password Protected File With Suffix For All Occurrence : |Pdf Text Compare| ${filepath} | ${string} | suffix=${Suffix} | all_occurrence=True | password=123456 |

Pdf Text Compare By Page And Line Number	filepath, string=None, page_no=None, line_no=None, all_occurrence=False, password=None	
Validates Text by page No and Line No in a PDF file.

Arguments
filepath: Filepath for PDF file.
string: String to be validated in the PDF. In case of Multiple strings validation, strings could be passed in the form of a list.
page_no : page number from which string to be validated in the PDF
line-no [Optional] : line number from which string to be validated in the PDF
all_occurrence [Optional]: If "True" it searches for all the occurrences of the string that needs to be validated and if found it displays all the occurences in html table. Default value is "False" which displays only the first occurrence of the string.
password [Optional]: Password for the password protected pdf file.
Example Test Cases
1. To Validate Text In PDF With Page No : |Pdf Text Compare| ${filepath} | ${string} | page_no=${Page_No} |

2. To Validate Text In PDF With Page No For All Occurence : |Pdf Text Compare| ${filepath} | ${string} | page_no=${Page_No} | all_occurrence=True |

3. To Validate Text In PDF With Page No And Line No : |Pdf Text Compare| ${filepath} | ${string} | page_no=${Page_No} | line_no=${Line_No} |

4. To Validate Text For a PDF Password Protected File With Page No For All Occurence : |Pdf Text Compare| ${filepath} | ${string} | page_no=${Page_No} | all_occurrence=True | | password=123456 |