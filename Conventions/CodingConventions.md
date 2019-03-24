# Coding Conventions

## Constant and Variable Naming Conventions

At public scope, a constant or a variable should be **upper-snake**-case or **upper-camel**-case, respectively. Here is examples where **`Hours_A_Day`** as a constant, and **`WorkingHours`** and **`SleepingHours`** as variables are defined.

```vb
Public const Hours_A_Day As Integer = 24 ' Public constant

Public WorkingHours As Integer ' Public variable
Public SleepingHours As Integer ' Public variable

WorkingHours = 8
SleepingHours = Hours_A_Day - WorkingHours
```

At private scope, a constant or a variable should be **lower-snake**-case or **lower-camel**-case, respectively. Here is examples where **`munching_Hours`** as a constant and **`actualWorkingHours`** as a variable are defined.

```vb
Private const munching_Hours As Integer = 2
Private actualWorkingHours As Integer

actualWorkingHours = WorkingHours - munching_Hours
```

At private scope with limited lifecycle (ie. constants/variables in functions/subroutines), a constant or a variable should be **lower-snake**-case or **lower-camel**-case followed by one underscore; "_", respectively. Here is examples where **`nap_Hours_`** as a constant and **`truthWorkingHours_`** as a variable are defined.

```vb
Private const nap_Hours_ As Integer = 1
Private truthWorkingHours_ As Integer

truthWorkingHours_ = actualWorkingHours - nap_Hours_
```

Any name that conflicts with reserved words should the name followed by two underscores; **"__"**. Here is examples where two variables; **`error__`** and **`debug__`** are defined.

```vb
Public Enum Logger_Level '* Logger levels.
  Off
  Fatal
  Error__
  Warn
  Info
  Debug__
  Trace
  All
End Enum
```

## Class Naming Conventions

A class name should be a class name followed by a word; **"Class"** with **upper-camel**-case or **lower-camel**-case depending on public scope or private scope respecyively. Here is examples where **`DogClass`** at public scope and **`CatClass`** at private scope are defined.

```vb
Dim Dog As DogClass
Dim cat As catClass
```

A public method should be **upper-camel**-case. Here is examples where two methods; **`BiteBurglers`** and **`EatFood`** are defined.

```vb
Public BiteBurglers
Public EatFood
```

A private method should be **lower-camel**-case . Here is examples where two methods; **`biteOwner`** and **`stealFood`** are defined.

```vb
Private biteOwner
Private stealFood
```

A public property should be **upper-camel**-case and a corresponding member should be **lower-camel**-case followed by a word; **"my"**. Here is examples where a public property; **`Owner`** and a corresponding member; **`myOwner`** are defined.

```vb
Private myOwner As String

Public Property Let Owner()
  myOwner = Owner
End Property

Public Property Get Owner()
  Owner = myOwner
End Property
```

A private property should be **lower-camel**-case and a corresponding member should be **lower-camel**-case followed by a word; **"my"**. Here is examples where a public property; **`stolenFoods`** and a corresponding member; **`myStolenFoods`** are defined.

```vb
Private myStolenFoods As String

Public stolenFoods Let Owner()
  myStolenFoods = stolenFoods
End Property

Public Property Get stolenFoods()
  stolenFoods = myStolenFoods
End Property
```

## Object Naming Conventions

A public object should be **upper-camel**-case. Here is examples where two objects; **`Jack`** and **`Tom`** are defined.

```vb
Public Jack As Dog
Public Tom As Dog
```

A private object should be **lower-camel**-case. Here is examples where two objects; **`Jack`** and **`Tom`** are defined.

```vb
Private biteOwner
Private stealFood
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



Object name should be:
* upper camel case for private, and
* Object name 

<class name> aaa

`test`




## Structured Coding Conventions
