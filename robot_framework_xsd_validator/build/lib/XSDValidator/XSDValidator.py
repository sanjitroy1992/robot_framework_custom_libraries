import xmlschema
from pprint import pprint
import os

class XSDValidator(object):
    ROBOT_LIBRARY_SCOPE = 'TEST SUITE'

    def validate_XSD_file(self, xml_fileinput):
        """| Usage |
        This keyword is used to validate am xml file against the XSD
        
        | Arguments |
        
         'xml_fileinput' = Output xml file path.
         
         |XSD Validator | ${xml_fileinput}
        """
        if not os.path.exists(xml_fileinput):
            raise AssertionError("File not found error :" + xml_fileinput + " doesnot exist")
        pprint(xmlschema.to_dict(xml_fileinput))
        print(xmlschema.validate(xml_fileinput))

