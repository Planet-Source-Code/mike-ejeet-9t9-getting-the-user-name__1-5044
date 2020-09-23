<div align="center">

## Getting the User Name


</div>

### Description

By calling this function you can retrieve the current user name on that computer. Enjoy!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike\-Ejeet 9t9](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-ejeet-9t9.md)
**Level**          |Intermediate
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-ejeet-9t9-getting-the-user-name__1-5044/archive/master.zip)





### Source Code

```
Place the following code into a module:
Private Declare Function GetUserName Lib "advapi32.dll" _
      Alias "GetUserNameA" (ByVal lpBuffer As String, _
      nSize As Long) As Long
Public Function UserName() As String
  Dim llReturn As Long
  Dim lsUserName As String
  Dim lsBuffer As String
  lsUserName = ""
  lsBuffer = Space$(255)
  llReturn = GetUserName(lsBuffer, 255)
  If llReturn Then
    lsUserName = Left$(lsBuffer, InStr(lsBuffer, Chr(0)) - 1)
  End If
  UserName = lsUserName
End Function
```

