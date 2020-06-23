from robot.api.deco import keyword
import os
import json
import jsonschema
from JsonLibrary import jsonselect
import pprint
import objectpath
from jsonpath_rw_ext import parse
from jsonpath_rw.parser import DatumInContext, Index, Fields
from collections import OrderedDict
import codecs


class JsonLibrary(object):
    """JsonLibrary is used for manipulating JSON object(dictionary)

        You can get, add, update and delete your json object using JSONPath.

        == JSONPath Syntax ==
        | JSONPath | Description |
        | $        | the root object/element |
        | @        | the current object/element |
        | . or []  | child operator |
        | ..       | recursive descent. JSONPath borrows this syntax from E4X |
        | *        | wildcard. All objects/element regardless their names. |
        | []       | subscript operator. XPath uses it to iterate over element collections and for predicates. In Javascript and JSON it is the native array operator. |
        | [,]      | Union operator in XPath results in a combination of node sets. JSONPath allows alternate names or array indices as a set. |
        | [start:end:step] | array slice operator borrowed from ES4 |
        | ?()      | applies a filter (script) expression. |
        | ()       | script expression, using the underlying script engine. |

        == Example of use ==
        json_example.json

        | {
        |    "firstName": "John",
        |    "lastName": "doe",
        |    "age": 26,
        |    "gender": "male",
        |    "favoriteColor": [
        |        "blue", "black"
        |    ],
        |    "isMarried": false,
        |    "address": {
        |        "streetAddress": "naist street",
        |        "city": "Nara",
        |        "postalCode": "630-0192"
        |    },
        |    "officeaddress": {
        |        "streetAddress": "brooklyn street",
        |        "city": "newyork",
        |        "postalCode": "994-3332"
        |    },
        |    "phoneNumbers": [{
        |            "type": "iPhone",
        |            "number": "90-4567-8888"
        |        },
        |        {
        |            "type": "home",
        |            "number": "0123-4567-8910"
        |        }
        |    ]
        | }

        == Example Test Cases ==
        | *** Settings ***     |
        | Library              | JsonLibrary |
        |                      |
        | *** Test Cases ***   |
        | Add Json Object By JsonPath |

        | ${json_data}  |  Load JSON From File  |  ../JSON/DataSetFiles/OpenAPIData/JSonFiles/example.json
        | ${object_to_add}  |  Create Dictionary  |  latitude=13.1234  |  longitude=130.1234
        | ${json_data}  |  Add Object To Json  |  ${json_data}  |  $..address  |  ${object_to_add}
        | Dictionary Should Contain Sub Dictionary  |  ${json_data['address']}  |  ${object_to_add}
        | ${json_data}  |  Pretty Print Json  |  ${json_data}
        |  log  | ${json_data}  |

        == Result ==

        |    {
        |      "firstName": "John",
        |      "lastName": "doe",
        |      "age": 26,
        |      "gender": "male",
        |      "favoriteColor": [
        |        "blue",
        |        "black"
        |      ],
        |      "isMarried": false,
        |      "address": {
        |        "streetAddress": "naist street",
        |        "city": "Nara",
        |        "postalCode": "630-0192",
        |        "latitude": "13.1234",
        |        "longitude": "130.1234"
        |      },
        |      "officeaddress": {
        |        "streetAddress": "brooklyn street",
        |        "city": "newyork",
        |        "postalCode": "994-3332"
        |      },
        |      "phoneNumbers": [
        |        {
        |          "type": "iPhone",
        |          "number": "90-4567-8888"
        |        },
        |        {
        |          "type": "home",
        |          "number": "0123-4567-8910"
        |        }
        |      ]
        |    }
        """

    ROBOT_EXIT_ON_FAILURE = True
    ROBOT_LIBRARY_SCOPE = "GLOBAL"
    ROBOT_LIBRARY_DOC_FORMAT = "ROBOT"

    @keyword('Load Json From File')
    def load_json_from_file(self, file_name):
        """Loads Json data from file.

        == Arguments ==
            - file_name: Absolute json file name.

        == Returns ==
        Returns json object as dictionary

        == Example Test Cases ==

        | ${result}  |  Load Json From File  | path/to/file.json |
        """
        if not os.path.isfile(file_name):
            raise AssertionError('File {} does not exist'.format(file_name))
        with codecs.open(file_name, 'r', encoding='utf-8', errors='ignore') as json_file:
            data = json.load(json_file, object_pairs_hook=OrderedDict)
        return data

    @keyword('Add Object To Json')
    def add_object_to_json(self, json_data, json_path, object_to_add):
        """ Adds a dictionary object to json_data using json_path

            == Arguments ==
                - json_data: json data as a dictionary object.\n
                - json_path: jsonpath expression.\n
                - object_to_add: dictionary object to add to json_data which is matched by json_path.\n

            == Returns ==
            Returns a new json data object.

            == Example Test Cases ==

            | ${json_data}  |  Load JSON From File  |  ../JSON/DataSetFiles/OpenAPIData/JSonFiles/example.json
            | &{dict} | Create Dictionary | latitude=13.1234 | longitude=130.1234 |
            | ${json} | Add Object To Json | ${json_data} | $..address | ${dict} |
            """

        json_path_expr = parse(json_path)
        if not json_path_expr.find(json_data):
            raise JsonValidatorError("Invalid Input Error ! \nPath {} doesn't exist".format(json_path))
        for match in json_path_expr.find(json_data):
            if type(match.value) in [dict, OrderedDict]:
                match.value.update(object_to_add)
            else:
                raise JsonValidatorError("Invalid Object type. Object is neither Dictionary nor List")
        if isinstance(json_data, list):
            return json_data[0]
        else:
            return json_data


    @keyword('Get Value From Json')
    def get_value_from_json(self, json_data, json_path):
        """Gets Value From Json data using JSONPath

        == Arguments ==
            - json_data: json data as a dictionary object.\n
            - json_path: jsonpath expression.\n

        == Returns ==
        Returns matched value from json in string format or ``None`` if no value were found

        == Example Test Cases ==
        | ${json_data}  |  Load JSON From File  |  ../JSON/DataSetFiles/OpenAPIData/JSonFiles/example.json
        | ${values}  |  Get Value From Json  | ${json_data} |  $..address.postalCode |
        | Should Be Equal As Strings  |  ${values}  |  "630-0192" |

        """
        json_path_expr = parse(json_path)
        if not json_path_expr.find(json_data):
            raise JsonValidatorError("Invalid Input Error ! \n Either json_path {} doesn't exist \n Or Input is not a json object".format(json_path))
        for match in json_path_expr.find(json_data):
            value = self.json_to_string(match.value)
        return value

    @keyword('Update Values To Json')
    def update_values_to_json(self, json_data, *params):
        """ Updates value to json data. params needs to be passed using format jsonpath_expression=value.\n

            == Arguments ==

            - json_data: json data as a dictionary object.\n
            -*params: Pass the data in the format of jsonpath_expression=value. It takes multiple params separated by Tab spaces. For example: "$..address.city=Bangkok" \n

            == Returns ==
            Returns new json_object

            == Example Test Cases ==
            1. To Update Json data with single value :
            | ${json_data}  |  Load JSON From File | ../JSON/DataSetFiles/OpenAPIData/JSonFiles/json_example.json |
            | ${Updated_json_data}  |  Update Values To Json  |  ${json_data}  |  $..address.city=Bangkok |

            2. To Update Json data with multiple values :
            | ${json_data}  |  Load JSON From File  |  ../JSON/DataSetFiles/OpenAPIData/JSonFiles/json_example.json |
            | ${Updated_json_data}  |  Update Values To Json  |  ${json_data}  |  $..address.city=Bangkok  |  $..officeaddress.city=california  |  $..phoneNumbers[0].number=00001111 |

                """
        for value in params:
            json_data = self._update_value_to_json(json_data, value.split("=")[0], value.split("=")[1])
        return json_data

    def _update_value_to_json(self, json_data, json_path, new_value):
        """Update value to JSON using JSONPath

        == Arguments ==

            - json_data: json data as a dictionary object.\n
            - json_path: jsonpath expression\n
            - new_value: value to update\n

        == Returns ==
        Returns new json_object
        """
        json_path_expr = parse(json_path)
        if not json_path_expr.find(json_data):
            raise JsonValidatorError("Invalid Input Error ! \nPath {} doesn't exist".format(json_path))
        for match in json_path_expr.find(json_data):
            path = match.path
            if isinstance(path, Fields):
                fields_type = match.context.value[match.path.fields[0]]
                if isinstance(fields_type, str):
                    match.context.value[match.path.fields[0]] = new_value
                elif isinstance(fields_type, bool):
                    if new_value.lower() == "true":
                        match.context.value[match.path.fields[0]] = True
                    elif new_value.lower() == "false":
                        match.context.value[match.path.fields[0]] = False
                    else:
                        raise AssertionError("Please provide a valid boolean value as true or false. The provided value to be updated is: {}".format(new_value))
                elif isinstance(fields_type, int):
                    match.context.value[match.path.fields[0]] = int(new_value)
                elif isinstance(fields_type, float):
                    match.context.value[match.path.fields[0]] = float(new_value)
            elif isinstance(path, Index):
                path_type = match.context.value[match.path.index]
                if isinstance(path_type, str):
                    match.context.value[match.path.index] = new_value
                elif isinstance(path_type, bool):
                    if new_value.lower() == "true":
                        match.context.value[match.path.index] = True
                    elif new_value.lower() == "false":
                        match.context.value[match.path.index] = False
                    else:
                        raise AssertionError("Please provide a valid boolean value as true or false. The provided value to be updated is: {}".format(new_value))
                elif isinstance(path_type, int):
                    match.context.value[match.path.index] = int(new_value)
                elif isinstance(path_type, float):
                    match.context.value[match.path.index] = float(new_value)

        return json_data

    @keyword('Delete Object From Json')
    def delete_object_from_json(self, json_data, json_path):
        """Deletes Object From Json data using json_path expression

        == Arguments ==
            - json_data: json data as a dictionary object.\n
            - json_path: jsonpath expression\n

        == Returns ==
        Returns a new json_data

        == Example Test Cases ==
        | ${json_data}  |  Load JSON From File | ../JSON/DataSetFiles/OpenAPIData/JSonFiles/json_example.json |
        | ${new_json_data}  |  Delete Object From Json | ${json_data} |  $..address.streetAddress  |
        """
        json_path_expr = parse(json_path)
        if not json_path_expr.find(json_data):
            raise JsonValidatorError("Invalid Input Error ! \nPath {} doesn't exist".format(json_path))
        for match in json_path_expr.find(json_data):
            path = match.path
            if isinstance(path, Index):
                del(match.context.value[match.path.index])
            elif isinstance(path, Fields):
                del(match.context.value[match.path.fields[0]])
        return json_data

    def _validate_json(self, checked_json, schema):
        """ Validates JSON according to JSONSchema

        == Arguments ==
            - checked_json: validated JSON.
            - schema: schema that used for validation.
        """
        try:
            jsonschema.validate(checked_json, schema)
        except jsonschema.ValidationError as e:
            print("""Failed validating '{0}'in schema {1}:{2}On instance {3}:{4}""".format(e.validator,
                                                                                           list(e.relative_schema_path)[
                                                                                           :-1],
                                                                                           pprint.pformat(e.schema),
                                                                                           "[%s]" % "][".join(
                                                                                               repr(index) for index in
                                                                                               e.absolute_path),
                                                                                           pprint.pformat(
                                                                                               e.instance)))
            raise JsonValidatorError("Failed validating json by schema")
        except jsonschema.SchemaError as e:
            raise JsonValidatorError('Json-schema error')

    def _check_if_json_is_filepath_or_string(self, argument):
        """
        Checks whether argument is a filepath or OrderedDict or json sting data.

        == Arguments ==
        - argument: argument could be filepath or OrderedDict or json sting data

        == Returns ==
        Returns data in string type
        """
        flag = False
        try:
           flag = os.path.isfile(argument)
           if flag:
                argument = open(argument).read()
        except Exception as e:
           flag = False
        if not flag:
            if (str(argument).startswith("{") and str(argument).endswith("}")) or (str(argument).startswith("[") and str(argument).endswith("]")):
                argument = str(argument)
            elif str(argument).startswith("OrderedDict"):
                argument = self.json_to_string(argument)
            else:
                raise AssertionError("Invalid input format! Please check the input and pass either a filepath or a json sting data")
        return argument


    @keyword('Validate Jsonschema')
    def validate_jsonschema(self, json_data, json_schema):
        """
        Validate JSON data according to schema.

        == Arguments ==
        - json_data: json data as a dictionary object.\n
        - input_schema: Pass the JSON Schema, which will be used to validates the structure of JSON data. This could be passed as string or filepath.\n

        == Example of use ==

        - Schema.json (Below is an example of json Schema, it defines the Structure for the json data)

        |    {
        |        "properties": {
        |            "name": {
        |                "type": "string"
        |            },
        |            "price": {
        |                "type": "number",
        |                "minimum": 0
        |            },
        |            "sku": {
        |                "description": "Stock Keeping Unit",
        |                "type": "integer"
        |            }
        |        },
        |        "required": ["name", "price"]
        |    }

        |
        - The below example is a valid json data because according to the above Structure of "Schema.json", here "name" value type is a "string" and "price" value type is a "number"
        |    {
        |        "name": "eggs",
        |        "price": 80
        |    }

        |
        - The below example is an invalid json data because according to the above Structure of "Schema.json", here the "price" value type is not a "number" but a "string"
        |    {
        |        "name": "eggs",
        |        "price": "80"
        |    }

        == Example Test Cases ==
        1. To Validate JsonSchema:
        | Validate Jsonschema  | ${json_data} | ${json_schema} |

        2. To Validate JsonSchema From File:
        | Validate Jsonschema  | ../JSON/DataSetFiles/OpenAPIData/JSonFiles/json_data.json | ../JSON/DataSetFiles/OpenAPIData/JSonFiles/json_schema.json |
        """
        json_data = self._check_if_json_is_filepath_or_string(json_data)
        json_schema = self._check_if_json_is_filepath_or_string(json_schema)
        try:
            load_schema = json.loads(json_schema, object_pairs_hook=OrderedDict)
        except ValueError as e:
            raise JsonValidatorError('Error in schema: {}'.format(e))
        load_input_json = self.string_to_json(str(json_data))
        self._validate_json(load_input_json, load_schema)

    @keyword('String To Json')
    def string_to_json(self, source):
        """
        Deserialize string into JSON structure.

        == Arguments ==
        - source : JSON string

        == Returns ==
        JSON structure

        == Example Test Cases ==

        1. To convert string to json from a file by using "Get File" Builtin keyword to get json data.
        | *Settings* | *Value* |
        | Library    | OperatingSystem |
        | ${Json_String}  |  OperatingSystem.Get File | ../JSON/DataSetFiles/OpenAPIData/JSonFiles/json_example.json |
        | ${json_data}  |  String To Json | ${Json_String} |

        """
        try:
            load_input_json = json.loads(source, object_pairs_hook=OrderedDict)
        except ValueError as e:
            raise JsonValidatorError("Could not parse '%s' as JSON: %s" % (source, e))
        except TypeError as e:
            raise JsonValidatorError("Input %s type" % e)
        return load_input_json

    @keyword('Json To String')
    def json_to_string(self, source):
        """
        Serialize JSON structure into string.

        == Arguments ==
        - source : JSON structure

        == Returns ==
        JSON string

        == Example Test Cases ==

        1. To convert json to string from a file using "Load JSON From File" keyword:
        | ${json_data}  |  Load JSON From File | ../JSON/DataSetFiles/OpenAPIData/JSonFiles/json_example.json |
        | ${Json_String}    Json To String    ${json_data}

        """
        try:
            load_input_json = json.dumps(source)
        except ValueError as e:
            raise JsonValidatorError("Could serialize '%s' to JSON: %s" % (source, e))
        return load_input_json

    @keyword('Get Json Elements')
    def get_json_elements(self, json_string, expr):
        """
        Get list of elements from json_string for matching expression.

        == Arguments ==
        - json_string : JSON string;\n
        - expr : JSONPath expression;\n

        == Returns ==
        List of found elements or ``None`` if no elements were found

        == Example Test Cases ==

        1. To Get Json Elements from a file using "Load JSON From File" keyword:
        | ${json_data}  |  Load JSON From File | ../JSON/DataSetFiles/OpenAPIData/JSonFiles/json_example.json |
        | ${json_String_data}  |  Json To String | ${json_data} |
        | @{json_elements}  |  Get Json Elements | ${json_String_data} |  $..number |
        =>
        | [ 90-4567-8888 | 0123-4567-8910 ]

        """

        jsonpath_expr = parse(expr)
        load_input_json = self.string_to_json(json_string)
        value_list = []
        if not jsonpath_expr.find(load_input_json):
            raise JsonValidatorError("Invalid Input Error ! \nPath {} doesn't exist".format(expr))
        for match in jsonpath_expr.find(load_input_json):
            value_list.append(match.value)
        if not value_list:
            return None
        else:
            return value_list

    @keyword('Select Json Elements')
    def select_json_elements(self, json_string, expr):
        """
        Returns list of elements from _json_string_ with matching expression.

        == Arguments ==
        - json_string : JSON string;\n
        - expr : JSONPath expression;\n

        == Returns ==
        List of found elements or ``None`` if no elements were found

        == Example Test Cases ==

        1. To Select Json Elements from a file using "Load JSON From File" keyword:
        | ${json_data}  |  Load JSON From File | ../JSON/DataSetFiles/OpenAPIData/JSonFiles/json_example.json |
        | ${json_String_data}  |  Json To String | ${json_data} |
        | @{json_elements}  |  Select Json Elements | ${json_String_data} |  $..address |
        =>
        | [ "streetAddress" : "naist street" | "city" : "Nara" | "postalCode" : "630-0192" ]
        """

        load_input_json = self.string_to_json(json_string)
        match = jsonselect.select(expr, load_input_json)
        if not match:
            raise JsonValidatorError('Elements %s does not exist' % expr)
        list1 = []
        if isinstance(match, dict):
            for key, value in list(match.items()):
                list1.append('"{0}" : "{1}"'.format(key, value))
        elif isinstance(match, list):
            for item in list(match):
                list1.append(str(item))
        else:
            list1.append('"{}"'.format(match))
        return list1

    @keyword('Select Json Objects')
    def select_json_objects(self, json_string, expr):
        """
        Return list of elements from json_string, matching expression.

        == Arguments ==
        - json_string : JSON string;\n
        - expr : ObjectPath expression;\n

        == Returns ==
        List of found elements. If no elements were found, empty list will be returned

        1. To Select Json Object from a file using "Load JSON From File" keyword:
        | ${json_data}  |  Load JSON From File | ../JSON/DataSetFiles/OpenAPIData/JSonFiles/json_example.json |
        | ${json_String_data}  |  Json To String | ${json_data} |
        | @{json_elements}  |  Select Json Objects | ${json_String_data} |  $..number |
        | Should Be Equal As Strings  |  @{Json_Objects}[1]  |  0123-4567-8910 |

        """
        jsonpath_expr = parse(expr)
        load_input_json = self.string_to_json(json_string)
        load_input_json1 = json.loads(json.dumps(load_input_json))
        value_list = []
        if not jsonpath_expr.find(load_input_json1):
            raise JsonValidatorError("Invalid Input Error ! \nPath {} doesn't exist".format(expr))
        for match in jsonpath_expr.find(load_input_json1):
            value_list.append(match.value)
        if not value_list:
            return None
        else:
            return value_list

    @keyword('Json Element Should Exist')
    def json_element_should_exist(self, json_string, expr):
        """
        Check the existence of one or more elements, matching expression.

        == Arguments ==
        - json_string : JSON string;\n
        - expr : JSONPath expression;\n

        1. To Check Json Element Should Exist from a file using "Load JSON From File" keyword:
        | ${json_data}  |  Load JSON From File | ../JSON/DataSetFiles/OpenAPIData/JSonFiles/json_example.json |
        | ${json_String_data}  |  Json To String | ${json_data} |
        | Json Element Should Exist  |  ${json_String_data} |  .address.streetAddress:contains("naist street") |

        """
        value = self.select_json_elements(json_string, expr)
        if value is None:
            raise JsonValidatorError('Elements %s does not exist' % expr)
        else:
            print('Elements %s exist' % expr)

    @keyword('Json Element Should Not Exist')
    def json_element_should_not_exist(self, json_string, expr):
        """
        Check that one or more elements, matching expression, don't exist.

        == Arguments ==
        - json_string : JSON string;\n
        - expr : JSONPath expression;\n

        1. To Check Json Element Should Exist from a file using "Load JSON From File" keyword:
        | ${json_data}  |  Load JSON From File | ../JSON/DataSetFiles/OpenAPIData/JSonFiles/json_example.json |
        | ${json_String_data}  |  Json To String | ${json_data} |
        | Json Element Should Exist  |  ${json_String_data} |  .address.streetAddress:contains("new york") |

        """

        load_input_json = self.string_to_json(json_string)
        match = jsonselect.select(expr, load_input_json)
        if match is not None:
            raise JsonValidatorError('Elements %s exist but should not' % expr)
        else:
            print('Elements %s does not exist' % expr)

    def _json_path_search(self, json_dict, expr):
        """
        Scan JSON dictionary with using json-path passed sting of the format of $.element..element1[index] etc.

        *Args:*\n
        _json_dict_ - JSON dictionary;\n
        _expr_ - string of fuzzy search for items within the directory;\n

        *Returns:*\n
        List of DatumInContext objects:
        ``[DatumInContext(value=..., path=..., context=[DatumInContext])]``
        - value - found value
        - path  - value selector inside context.value (in implementation of jsonpath-rw: class Index or Fields)

        *Raises:*\n
        JsonValidatorError
        """
        path = parse(expr)
        results = path.find(json_dict)

        if len(results) is 0:
            raise JsonValidatorError("Nothing found in the dictionary {0} using the given path {1}".format(
                str(json_dict), str(expr)))
        return results

    @keyword('Update Json')
    def update_json(self, json_string, expr, value, index=0):
        """
        Replace the value in the JSON string.

        == Arguments ==
        - json_string : JSON string;\n
        - expr : JSONPath expression for determining the value to be replaced;\n
        - value : the value to be replaced with;\n
        - index : index for selecting item within a match list, default value is 0;\n

        == Returns ==
        Changed JSON in dictionary format.

        == Example Test Cases ==
        1. To update value in json from a file using "Load JSON From File" keyword:
        | ${json_data}  |  Load JSON From File | ../JSON/DataSetFiles/OpenAPIData/JSonFiles/json_example.json |
        | ${Json_String}    Json To String    ${json_data}
        | ${json_update}  |  Update Json  |  ${Json_String}  |  $..address.city  |  Bangkok |
        """
        load_input_json = self.string_to_json(json_string)
        matches = self._json_path_search(load_input_json, expr)

        datum_object = matches[int(index)]

        if not isinstance(datum_object, DatumInContext):
            raise JsonValidatorError("Nothing found by the given json-path")

        path = datum_object.path

        # Edit the directory using the received data
        # If the user specified a list
        if isinstance(path, Index):
            datum_object.context.value[datum_object.path.index] = value
        # If the user specified a value of type (string, bool, integer or complex)
        elif isinstance(path, Fields):
            datum_object.context.value[datum_object.path.fields[0]] = value

        return load_input_json

    @keyword('Pretty Print Json')
    def pretty_print_json(self, json_string):
        """
        Returns formatted JSON string.\n

        == Arguments ==
        - json_string: JSON string.

        == Example Test Cases ==
        ${JsonString}    Json To String    {a:1,foo:[{b:2,c:3},{d:"baz",e:4}]}
        |${pretty_json}  | Pretty Print Json |  ${JsonString} |
        | Log  |  ${pretty_json}  |
        =>\n
        | {
        |    "a": 1,
        |    "foo": [
        |      {
        |        "c": 3,
        |        "b": 2
        |      },
        |      {
        |        "e": 4,
        |        "d": "baz"
        |      }
        |    ]
        | }
        """
        if type(json_string) != str:
            json_string = self.json_to_string(json_string)
        return json.dumps(self.string_to_json(json_string), indent=2, ensure_ascii=False)


class JsonValidatorError(Exception):
    pass
