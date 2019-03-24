# Coding Conventions

## Constant and Variable Naming Conventions

At public scope, a constant or a variable should be **upper-snake**-case or **upper-camel**-case, respectively. Here is examples where **`Hours_A_Day`** as a constant, and **`WorkingHours`** and **`SleepingHours`** as variables are defined.

```vb
Public const Hours_A_Day As Integer = 24

Public WorkingHours As Integer
Public SleepingHours As Integer

WorkingHours = 8
SleepingHours = Hours_A_Day - WorkingHours
```

At private scope, a constant or a variable should be **lower-snake**-case or **lower-camel**-case, respectively. Here is examples where **`munching_Hours`** as a constant and **`actualWorkingHours`** as a variable are defined.

```vb
Private const munching_Hours As Integer = 2
Private actualWorkingHours As Integer

actualWorkingHours = WorkingHours - munching_Hours
```

At private scope with limited lifecycle (ie. constants/variables in functions/subroutines), a constant or a variable should be **lower-snake**-case or **lower-camel**-case followed by one underscore; "_", respectively. Here is examples:

```vb
Private const nap_Hours_ As Integer = 1
Private truthWorkingHours_ As Integer

truthWorkingHours_ = actualWorkingHours - nap_Hours_
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
