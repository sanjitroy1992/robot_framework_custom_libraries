*** Settings ***
Library         JsonLibrary      
Library         Collections
Library         String
Test Setup      SetUp Test

*** Keywords ***
SetUp Test
    ${json_data}    Load JSON From File    ${CURDIR}\\SampleData\\example.json
    ${JsonSchema}    Load JSON From File    ${CURDIR}\\SampleData\\SchemaString.json
    ${Json_Schema_Data}    Load JSON From File    ${CURDIR}\\SampleData\\JsonSchemaExample.json
    Set Test Variable    ${json_data}    ${json_data}
    Set Test Variable    ${JsonSchema}    ${JsonSchema}
    Set Test Variable    ${Json_Schema_Data}    ${Json_Schema_Data}

*** Variables ***
${json_data}
${JsonSchema}
${Json_Schema_Data}

*** Test Cases ***
Add Object To Json
    [Documentation]  Adding a dictionary object to json_data using json_path
    ${object_to_add}    Create Dictionary    latitude=13.1234    longitude=130.1234
    ${json_data}    Add Object To Json     ${json_data}    $..address    ${object_to_add}
    Dictionary Should Contain Sub Dictionary    ${json_data['address']}    ${object_to_add}
    ${json_data}    Pretty Print Json    ${json_data}
    log    ${json_data}
    
String to Json
    [Documentation]    Deserialize string into JSON structure.
    ${Json_String}    Json To String    ${json_data}
    ${String}    String To Json    ${Json_String}
    ${print}    Pretty Print Json    ${String}
    log    ${print}
    
Json to String
    [Documentation]    Serialize JSON structure into string.
    ${Json_String}    Json To String    ${json_data}
    log    ${Json_String}
    
Get Value From Json
    [Documentation]    Gets Value From Json data using JSONPath
    ${values}    Get Value From Json    ${json_data}    $..phoneNumbers[0].number
    Should Be Equal As Strings    ${values}    9045678888
    log    ${values}
    
Get Json Elements
    [Documentation]    Gets Value From Json data using JSONPath
    ${json_data}    Json To String    ${json_data}
    @{values}    Get Json Elements    ${json_data}    $..number
    log    ${values}
    
Select Json Objects
    [Documentation]    Returns list of elements from json_string, matching expression.
    ${Json_String}    Json To String    ${json_data}
    @{Json_Objects}    Select Json Objects    ${Json_String}    $..number
    Should Be Equal As Strings    @{Json_Objects}[1]    9045678888.0
    
Update Values To Json
    [Documentation]  Updates value to json data. params needs to be passed using format jsonpath_expression=value.
    ${json_data}    Update Values To Json    ${json_data}     $..address.city=New York    $..isMarried=True    $..phoneNumbers[1].number=8.9    $..phoneNumbers[0].number=1234    $..address.city=New York    $..officeaddress.city=Los Angeles    $..address.Boolean=False
    ${json_data}    Pretty Print Json    ${json_data}
    log    ${json_data}

Update One Value To Json
    [Documentation]  Updates one value to json data. params needs to be passed using format jsonpath_expression=value.
    ${json_data}    Update Values To Json    ${json_data}    $..address.Boolean=False
    ${json_data}    Pretty Print Json    ${json_data}
    Log    ${json_data}

Update Json
    [Documentation]    Replaces the value in the JSON string.
    ${json_String}    Json To String    ${json_data}
    ${json_data}    Update Json    ${json_String}     $..age    77
    log    ${json_data}
    
Delete Object From Json
    [Documentation]    Deletes Object From Json data using json_path expression
    ${json_data}    Delete Object From Json    ${json_data}    $..address.streetAddress
    Dictionary Should Not Contain Key    ${json_data}    streetAddress
    ${json_data}    Pretty Print Json    ${json_data}
    log    ${json_data}
    
PrettyPrintJson
    [Documentation]    Returns formatted JSON string.
    log    ${json_data}
    ${json_data}    Pretty Print Json    ${json_data}
    log    ${json_data}
    
JsonSchema Validation from file
    [Documentation]  Validate JSON data according to schema. where both are taken from filepath
    Validate Jsonschema    ${CURDIR}\\SampleData\\JsonSchemaExample.json    ${CURDIR}\\SampleData\\SchemaString.json
    
JsonSchema Validation for String
    [Documentation]  Validate JSON data according to schema. where type of both are string
    ${JsonSchema}    Json To String    ${JsonSchema}
    ${Json_Schema_Data}    Json To String    ${Json_Schema_Data}
    Validate Jsonschema    ${Json_Schema_Data}    ${JsonSchema}
    
JsonSchema Validation for OrderedDict
    [Documentation]  Validate JSON data according to schema. where type of both are OrderedDict
    Validate Jsonschema    ${Json_Schema_Data}    ${JsonSchema}
    
Select Json Elements
    [Documentation]    Returns list of elements from _json_string_ with matching expression.
    ${json_String}    Json To String    ${json_data}
    @{json_String}    Select Json Elements    ${json_String}    .address
    
Json Element Should Exist
    [Documentation]    Check the existence of one or more elements, matching expression.
    ${json_String}    Json To String    ${json_data}
    Json Element Should Exist    ${json_String}    .address.streetAddress:contains("naist street")
    
Json Element Should Not Exist
    [Documentation]    Check that one or more elements, matching expression, don't exist.
    ${json_String}    Json To String    ${json_data}
    Json Element Should Not Exist    ${json_String}    .address.streetAddress:contains("new york")