<div align="center">

## transferring data


</div>

### Description

This VB code shows how you can use ADO and Microsoft Scripting Runtime Library to transfer data from a database table to a text file.
 
### More Info
 
This code uses the Microsoft Scripting Runtime Library and Microsoft ActiveX Data Objects Library, so to use it make sure you check off the Microsoft Scripting Runtime and ADO in Visual Basics References.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Syrob](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/syrob.md)
**Level**          |Intermediate
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/syrob-transferring-data__1-31807/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
' declare a constant
Const conType As Integer = 8209
'declare an object variable
Dim rs As ADODB.Recordset
' create a recordset object and
' assign it to the object variable
Set rs = New ADODB.Recordset
' execute the open method
With rs
.ActiveConnection = _
"provider=microsoft.jet.oledb.4.0;" _
& "data source = c:yourdb.mdb"
 .Source = "select * from yourtable"
 .Open
 End With
' declare a dynamic array
Dim str()
' execute the getrows method
str() = rs.GetRows
' we do not need the recordset object
Set rs = Nothing
' declare the object variables
Dim fso As FileSystemObject
Dim txtfile As TextStream
' create a filesystemobject
Set fso = _
CreateObject("Scripting.FileSystemObject")
' execute the createtextfile function
Set txtfile = _
fso.CreateTextFile("c:\testfile.txt", True)
' we do not need the object
Set fso = Nothing
' declare the variables that
' will hold the values from the array
Dim v, strv
' declare the variables
Dim j As Integer
Dim x As Integer
' loop through the array
For x = LBound(str, 2) To UBound(str, 2)
 For j = LBound(str, 1) To UBound(str, 1)
' dump a value from the array into
' the strv variable and do
' with it whatever you want
 strv = str(j, x)
' it is an example what you can do
 If VarType(strv) = conType Then
 strv = CStr(strv)
 strv = "Picture"
 ElseIf IsNull(strv) Then
 strv = "Null"
 ElseIf strv = "" Then
 strv = """"
 ElseIf j > 0 Then
 strv = "'" & strv & "'"
 End If
' build a text line
 v = v & strv & ", "
' check if the loop reached the end of a row
 If j = UBound(str, 1) Then
' since the whole record was dumped into
' the v variable, the text line can
' be written into the file
 txtfile.WriteLine v
' reset the variable
 v = ""
 End If
 Next j
 Next x
' close the object
txtfile.Close
End Sub
```

