<div align="center">

## Get an Array of string from Command$


</div>

### Description

This function returns an array with the command line arguments,

contained in the command$, like cmd.exe does (with %1 , %2 ...)

If the function does not succeed, the returned array has Ubound=-1 , like split function in VB.
 
### More Info
 
INPUT: comSTR = the string with the arguments (command$)

Just copy and paste the code in your project.

OUTPUT: Array of string with the arguments


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Nik Keso](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nik-keso.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/nik-keso-get-an-array-of-string-from-command__1-72701/archive/master.zip)





### Source Code

```
'===================================================================
' GetCommandArgs - © Nik Keso 2009
'----------------------------------
'The function returns an array with the command line arguments,
'contained in the command$, like cmd.exe does (with %1 , %2 ...)
'If the function does not succeed, the returned array has Ubound=-1 ,
'like split function in VB.
'----------------------------------
'INPUT: comSTR = the string with the arguments (command$)
'OUTPUT: Array of string with the arguments
'===================================================================
Function GetCommandArgs(ByVal comSTR As String) As String()
Dim CountQ As Integer 'chr(34) counter
Dim OpenQ As Boolean ' left open string indicator (ex:  "c:\bbb ccc.bat ) OpenQ=true, (ex:  "c:\bbb ccc.bat" ) OpenQ=false
Dim ArgIndex As Integer
Dim tmpSTR As String
Dim strIndx As Integer
Dim TmpArr() As String
GetCommandArgs = Split("", " ") 'trick to return uninitialized array like split, if function does NOT succeed!
TmpArr = Split("", " ")
comSTR = Trim$(comSTR) 'remove front and back spaces
If Len(comSTR) = 0 Then Exit Function
CountQ = UBound(Split(comSTR, """"))
If CountQ Mod 2 = 1 Then Exit Function 'like cmd.exe , command$ must contain even number of chr(34)=(")
strIndx = 1
Do
  If Mid$(comSTR, strIndx, 1) = """" Then OpenQ = Not OpenQ
  If Mid$(comSTR, strIndx, 1) = " " And OpenQ = False Then
    If tmpSTR <> "" Then 'don't include the spaces between args as args!!!!!
      ReDim Preserve TmpArr(ArgIndex)
      TmpArr(ArgIndex) = tmpSTR
      ArgIndex = ArgIndex + 1
    End If
    tmpSTR = ""
  Else
    tmpSTR = tmpSTR & Mid$(comSTR, strIndx, 1)
  End If
  strIndx = strIndx + 1
Loop Until strIndx = Len(comSTR) + 1
ReDim Preserve TmpArr(ArgIndex)
TmpArr(ArgIndex) = tmpSTR
GetCommandArgs = TmpArr
End Function
```

