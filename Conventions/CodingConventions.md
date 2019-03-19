# Coding Conventions

## Class Naming Conventions

A class name should be a class name followed by a word;"Class" with upper camel cases.
Here is examples:

* A class name of **ExcelFirst** is `ExcelFirstClass`.
* A class name of **Logger** is `LoggerClass`.

## Constant and Variable Naming Conventions

At public scope, a constant should be **upper-snake**-case and a variable should be **upper-camel**-case. Here is examples:
```vb
Public const Global_Constant_String As String = "Global constant string"
Public GlobalVariableString As String = "Global variable string"
```

At private scope, a constant should be **lower-snake**-case and a variable should be **lower-camel**-case. Here is examples:
```vb
Private const private_Constant_String As String = "Private constant string"
Private privateVariableString As String = "Private variable string"
```
<class name>Class

'
' Standard modules
'   "<class name>Module"
'   "ExamplesFor<class name>Class"
'
' Public scope:
'   Constants - upper snake case
'     "Global_Constants"
'   Variables - upper camel case
'     "GlobalVariables"
'     "PublicMethod"
'     "PublicProperty"
'     "Argument"
'
' Private scope:
'   Constants - lower snake case
'     "private_Constant"
'   Variables - lower camel case
'     "privateMethod"
'     "privateProperty"
'
' Private and lifecycle-limited:
'   Constants - lower snake case plus underscore
'     "constant_In_Method_"
'   Variables - lower camel case plus underscore
'     "variableInMethod_"
'     "variableInProperty_"
'
' Note) Constants or variables followed by two underscores are used to avoid conflictions with reserved words.
'   "debug__"
'   "error__"



## Object Naming Conventions
Object name should be:
* upper camel case for private, and
* Object name 

<class name> aaa

`test`




## Structured Coding Conventions
