# Coding Conventions

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

## Class Naming Conventions

A class name should be a class name followed by a word;"Class" with **upper-camel**-case or **lower-snake**-case depending on public scope or private scope respecyively. Here is examples:

```vb
Dim foo As ExcelFirstClass
Dim bar As LoggerClass
```

A public member should be **upper-camel**-case. Here is examples:

```vb
Public ClassName As String
PUblic NumberOfInstances As Long
```

A private member should be **lower-camel**-case followed by "my". Here is examples:

```vb
Private myScriptingDictionary As Object
Private myWallet As Long
```

A property should be **upper-camel**-case with a corresponding private member. Here is examples:

```vb
Private myWallet As Long

...

Public Property Let Wallet()
  myWallet = Wallet
End Property

Public Property Get Wallet()
  Wallet = myWallet
End Property
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
