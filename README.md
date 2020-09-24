<div align="center">

## Create A Shortcut On The Desktop/Program File/StartMenu etc\.\.\.


</div>

### Description

This code enables you to create shortcuts on the desktop or startmenu or in the program file. This is especially useful when you need your setup program to create a program shorcut to the desktop which is not provided in the usual setup kit that ships with visual basic.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-05-15 21:17:42
**By**             |[urbano da gama](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/urbano-da-gama.md)
**Level**          |Intermediate
**User Rating**    |4.8 (67 globes from 14 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD58245152000\.zip](https://github.com/Planet-Source-Code/urbano-da-gama-create-a-shortcut-on-the-desktop-program-file-startmenu-etc__1-8120/archive/master.zip)

### API Declarations

```
Public Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" _
     (ByVal lpstrFolderName As String, _
     ByVal lpstrLinkName As String, _
     ByVal lpstrLinkPath As String, _
     ByVal lpstrLinkArguments As String, _
     ByVal fPrivate As Long, _
     ByVal sParent As String) As Long
```





