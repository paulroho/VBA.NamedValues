# VBA.NamedValues
Demo for serializing/deserializing multiple named values in a string in VBA.
This is mainly useful for passing multiple values to a form, eg. in Microsoft Access via `OpenArgs`.

## Basic Usage

Say you want to pass the values of the variables `strFirstName` and `strLastName` to a form.

### Calling side
On the calling side use an instance of the class `NamedValues`, stuff it with the values and get the serialized result via the property 'AsString':

```vbnet
With New NamedValues
  .Add "FirstName", strFirstName
  .Add "LastName", strLastName

  DoCmd.OpenForm "TargetForm", OpenArgs:=.AsString   
End With
```

Remark: This is just the most dense version of getting things done. For real code I would recommend to invest more in maintainability.

### Receiving side
On the receiving side (say a form) you use another instance of `NamedValues`, pass the serialized string to its property `AsString` and ask the class to give you back the parts:

```vbnet
With New NamedValues
  .AsString = Nz(Me.OpenArgs, vbNullString)
  
  Me.txtFirstName.Value = .Item("FirstName")
  Me.txtLastName.Value = .Item("LastName")
End With
```

## Getting Started

To use `NamedValues` in your class, you need `NamedValue` along with the class `KeyValuePair`.
Both can be found in the [Code folder](https://github.com/paulroho/VBA.NamedValues/tree/master/NamedValues.accdb.Content/Code) in this repository.

### Sample Database 
For an easy kick start there is a sample as a Microsoft Access Database from this repository: [NamedValues.accdb](https://github.com/paulroho/VBA.NamedValues/raw/master/NamedValues.accdb).

**Hint** - This sample has no dependency on any specific VBA host. So you can use it without any modification in other applications such as *Microsoft Excel*.


## API Guide

### Instantiating `NamedValues`
You can instantiate `NamedValues` as every other class in VBA via the `New` keyword and assign it to a variable:

```vbnet
Dim Nv As NamedValues

Set Nv = New NamedValues
' Use it here, e.g.:
Nv.Add "TheKey", TheValue
Set Nv = Nothing 
```
As a matter of good practice, don't forget to set your variable to `Nothing` as soon as you don't need it anymore.

#### Shorter version leveraging the keyword `With`
For simple local uses of a class instance as this, I highly recommend using the way shorter syntax using a `With`-block:

```vbnet
With New NamedValues
   ' Use it here, e.g.:
   .Add "TheKey", TheValue
End With
```
This way of using the class has several advantages:
* It makes the use of a local variable obsolete.
* Each line is shorter, especially to access the members, you just start with the dot (`.`).
* The indentation provides a nice visual representation for the scope of the instance.
* It provides no way of forgetting setting a variable to `Nothing`.

In case you need to get the reference to the instance, use the property [Self](#self). But you should never need to do this. 

You will require this if you need to pass the instance of a `With`-block to a method:
```vbnet
With New NamedValues
   ' ...
   AnotherMethod .Self
   ' ...
End With
```
This can happen if you have to pass lots of values via `NamedValues` and want to refactor that out into several methods:
```vbnet
With New NamedValues
   ' ...
   SetSomeValues .Self
   SetSomeMoreValues .Self
   AndHereEvenMore .Self
   ' ...
End With

Private Sub SetSomeValues(ByVal Nv As NamedValues)
   With Nv
      ' ...
      .Add "AKey", "AValue"
      ' ...
   End With
End Sub

' ...
```
**Remark** - In the context of passing data to a form, I would see such a requirement as a code smell that you might want to do too much that could better be done in other ways. Just sayin'. 


### Properties

#### Self
The property `Self` gives you a reference to the current instance. This property is read only.
```vbnet
Property Get Self() As NamedValues
```
For a use of the property `Self` refer to the section on [instantiation](#shorter-version-leveraging-the-keyword-with)

#### ValueSeparator
The property `ValueSeparator` lets you specify the string that is used as the separator between two key/value-pairs within the serialized string. The default is the semicolon (`;`).
```vbnet
Property Let ValueSeparator(ByVal Value As String)
Property Get ValueSeparator() As String
```
You might want to change the default (semicolon `;`) if you have to pass a semicolon in your data.
For `NamedValues` to work properly, you have to make sure,
* that the string specified as the `ValueSeparator` never appears in your data.
* that `ValueSeparator` is always different from (and no subset of) [KeyValueSeparator](#keyvalueseparator).

In the current version there is no check for all this.

#### KeyValueSeparator
The property `KeyValueSeparator` lets you specify the string that is used as the separator between the key and its associated value. The default is the assignment character (`=`).

```vbnet
Property Let KeyValueSeparator(ByVal Value As String)
Property Get KeyValueSeparator() As String
```
You might want to change the default (assignment character `=`) if you have to pass a semicolon in your data.
For `NamedValues` to work properly, you have to make sure,
* that the string specified as the `KeyValueSeparator` never appears in your data.
* that `KeyValueSeparator` is always different from (and no subset of) [ValueSeparator](#valueseparator).
In the current version there is no check for all this.

#### Item
The parametrized property `Item` lets you return a value by its key. 
```vbnet
Property Get Item(ByVal Key As String) As String
```
The value must have been added to the instance of `NamedValues` via the method [Add](#add). If it does not exist, an error is raised.

#### ItemOrDefault
The parametrized property `ItemOrDefault` lets you return a value by its key and provide a default value in case the value does not exist. 
```vbnet
Property Get ItemOrDefault(ByVal Key As String, Optional ByVal Default As Variant = Null) As Variant
```

#### AsString
The property `AsString` provides you read/write access to the serialized data.
```vbnet
Property Let AsString(ByVal NewString As String)
```

### Methods

#### Add
The method `Add` lets you add a value along with its key to the instance of `NamedValues`.
```vbnet
Sub Add(ByVal Key As String, ByVal Value As String)
```

#### Exists
The method `Exists` tells you if a entry with a specific key already exists.
```vbnet
Function Exists(ByVal Key As String) As Boolean
``` 


## Contributing

If you find any issues with this code (and there are several!), feel free to [file an issue](https://github.com/paulroho/VBA.NamedValues/issues) and/or send a [pull request](https://github.com/paulroho/VBA.NamedValues/pulls).


## License
This sample is licensed under the [MIT License](LICENSE).


## CNamedValues
The first version of this class appeared back in 2004 in my talk "Praxiseinsatz von benutzerdefinierten Klassen in Microsoft Access" at the [AEK7](http://donkarl.com/AEK/Archiv/AEK7.htm) ([Access Entwickler Konferenz](http://www.donkarl.com?AEK) - Access Developer's Conference) in Nurnberg, Germany under the name `CNamedValues`. Since in the meantime I refrain from using prefixes depending on the type of modules, it now has the name `NamedValues`. If I just would be clever, it maybe would have the name `SerializableDictionary` or alike.
The full blown text for this talk (in German) can be downloaded from the [conference website](http://donkarl.com/Downloads/AEK/) (scroll down to AEK 7).
