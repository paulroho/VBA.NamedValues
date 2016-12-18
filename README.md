# VBA.NamedValues
Demo for serializing/deserializing multiple named values in a string.
This is mainly useful for passing multiple values to a form, eg. in Microsoft Access via `OpenArgs`.

## Basic Usage

Say you want to pass the values of the variables `strFirstName` and `strLastName` to a form.

### Calling side
On the calling side use an instance of the class `NamedValues`, stuff it with the values and get the serialized result via the property 'AsString':

```javascript
var s = "JavaScript syntax highlighting";
alert(s);
```

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

    With New NamedValues
      .AsString = Nz(Me.OpenArgs, vbNullString)
      
      Me.txtFirstName.Value = .Item("FirstName")
      Me.txtLastName.Value = .Item("LastName")
    End With

## Getting Started

To use `NamedValues` in your class, you need `NamedValue` along with the class `KeyValuePair`.
Both can be found in the [Code folder](https://github.com/paulroho/VBA.NamedValues/tree/master/NamedValues.accdb.Content/Code) in this repository.

### Sample Database 
For an easy kick start there is a sample as a Microsoft Access Database from this repository: [NamedValues.accdb](https://github.com/paulroho/VBA.NamedValues/raw/master/NamedValues.accdb).

**Hint** - This sample has no dependency on any specific VBA host. So you can use it without any modification in other applications such as *Microsoft Excel*.


## Contributing

If you find any issues with this code (and there are many!), feel free to [file an issue](https://github.com/paulroho/VBA.NamedValues/issues) and/or send a [pull request](https://github.com/paulroho/VBA.NamedValues/pulls).


## License
This sample is licensed under the [MIT License](LICENSE).