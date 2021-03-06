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
Private const munching_Hours As Integer = 2 ' Private constant
Private actualWorkingHours As Integer ' Private variable

actualWorkingHours = WorkingHours - munching_Hours
```

At private scope with limited lifecycle (ie. constants/variables in functions/subroutines), a constant or a variable should be **lower-snake**-case or **lower-camel**-case followed by one underscore; "_", respectively. Here is examples where **`nap_Hours_`** as a constant and **`truthWorkingHours_`** as a variable are defined.

```vb
Private const nap_Hours_ As Integer = 1 ' Private constant with limited lifecycle
Private truthWorkingHours_ As Integer ' Private variable with limited lifecycle

truthWorkingHours_ = actualWorkingHours - nap_Hours_
```

Any name that conflicts with reserved words should the name followed by two underscores; **"__"**. Here is examples where two variables; **`error__`** and **`debug__`** are defined.

```vb
Public Enum Logger_Level ' Logger levels.
  Off
  Fatal
  Error__ ' Avoid conflict with reserved word; "Error"
  Warn
  Info
  Debug__ ' Avoid conflict with reserved word; "Debug"
  Trace
  All
End Enum
```

## Class Naming Conventions

A class name should be a class name followed by a word; **"Class"** with **upper-camel**-case or **lower-camel**-case depending on public scope or private scope respecyively. Here is examples where a public class; **`DogClass`** and a private class; **`CatClass`** are defined.

```vb
Dim Dog As DogClass ' Public class
Dim cat As catClass ' Private class
```

A public method should be **upper-camel**-case. Here is examples where two methods; **`BiteBurglers`** and **`EatFood`** are defined.

```vb
Public BiteBurglers ' Public method
Public EatFood ' Public method
```

A private method should be **lower-camel**-case . Here is examples where two methods; **`biteOwner`** and **`stealFood`** are defined.

```vb
Private biteOwner ' Private method
Private stealFood ' Private method
```

A public property should be **upper-camel**-case and a corresponding member should be **lower-camel**-case followed by a word; **"my"**. Here is examples where a public property; **`Owner`** and a corresponding member; **`myOwner`** are defined.

```vb
Private myOwner As String ' Private member

Public Property Let Owner() ' Public property (Setter)
  myOwner = Owner
End Property

Public Property Get Owner() ' Public property (Getter)
  Owner = myOwner
End Property
```

A private property should be **lower-camel**-case and a corresponding member should be **lower-camel**-case followed by a word; **"my"**. Here is examples where a public property; **`stolenFoods`** and a corresponding member; **`myStolenFoods`** are defined.

```vb
Private myStolenFoods As String ' Private member

Private Property Let stolenFoods() ' Private property (Setter)
  myStolenFoods = stolenFoods
End Property

Private Property Get stolenFoods() ' Private property (Getter)
  stolenFoods = myStolenFoods
End Property
```

## Object Naming Conventions

A public object should be **upper-camel**-case. Here is examples where two objects; **`Jack`** and **`Tom`** are defined.

```vb
Public Jack As DogClass ' Public object
Public Jill As DogClass ' Public object
```

A private object should be **lower-camel**-case. Here is examples where two objects; **`jim`** and **`chris`** are defined.

```vb
Private tom As DogClass ' Private object
Private mary As DogClass ' Private object
```

## Example Naming Conventions

An example name should consists of:
- **"ExamplesFor"** and a class name,
- an underscore; **"_"**, and
- description with **upper-camel**-case.
Here is examples where a method; **"BiteBurglers"** of an object; **"DogClass"** is used.

```vb
Sub ExamplesForDogClass_BiteBurglers()
' The above name consists of: 
' - "ExamplesFor" and a class name; "DogClass",
' - an underscore; "_", and
' - description; "BiteBurglers".

Public Jack As DogClass
  set Jack = New DogClass
  Jack.BiteBurglers()
End Sub
```

The **"ExamplesFor" and a class name** is also used as an object name of standard module. Here is sampel for `LoggerClass` class.
* A name of Standard Module that implementes `LoggerClass` class is **LoggerClass**.
* A name of Standard Module that teaches how to use `LoggerClass` class is **ExamplesForLoggerClass**.
* A name of Standard Module that defines common constants including `LoggerClass` class and declaration of windows libraries is **CommonModule**.
