<div align="center">

## Best Shell & Wait \(No API's\!\)


</div>

### Description

Makes it easy to perform a clean "Shell & Wait" where your applicatoin kicks off an external application and waits for it to return before continuing. Many shell & wait examples I have found tend to overdrive the proccessor in a loop or require you to make API calls. This one uses the Windows Script object to take advantage of it's built-in wait parameter on the .Run method...scripting's version of Shell.
 
### More Info
 
FileName - The name of the file you wish to run with any required switches included.

Should work on any Windows 98 machine. Others may need to get the newest VB service pack or install Windows Scripting Host (http://msdn.microsoft.com/scripting/jscript/download/55beta.htm). This is also included in Internet Explorer 5. If you already have IE5, this will work and it will be included when you build your setup file for distribution.

True if the file was run and returned.

False if there was a file open or save error.

EXAMPLE: ShellAndWait ("notepad.exe c:\temp\teset.txt)

None - Will not block other applications or overdrive the proccessor.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-roberts.md)
**Level**          |Intermediate
**User Rating**    |4.9 (78 globes from 16 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-roberts-best-shell-wait-no-api-s__1-8349/archive/master.zip)





### Source Code

```
Function ShellAndWait(FileName As String)
Dim objScript
On Error GoTo ERR_OpenForEdit
Set objScript = CreateObject("WScript.Shell")
' Open a file for editing in Notepad and wait for return.
'The second parameter (after the FileName) is the Display Mode (normal w/focus,
'minimized...even hidden. For more info visit:
'http://msdn.microsoft.com/scripting/windowshost/doc/wsMthRun.htm
' The third parameter is the "Wait for return" parameter. This should be set to
' True for the Wait.
ShellApp = objScript.Run(FileName, 1, True)
ShellAndWait = True
EXIT_OpenForEdit:
 Exit Function
ERR_OpenForEdit:
 MsgBox Err.Description
 GoTo EXIT_OpenForEdit
End Function
```

