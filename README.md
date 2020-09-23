<div align="center">

## DelTree


</div>

### Description

DelTree function using FileSystemObject. Removes folder regardless of files/folders/system/hidden contained within. Couldn't find any deltree code here that worked, most used the kill statement and such in a rather large sub, found this in MSDN...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[michael schmidt](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-schmidt.md)
**Level**          |Intermediate
**User Rating**    |5.0 (45 globes from 9 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michael-schmidt-deltree__1-10399/archive/master.zip)





### Source Code

```
<font face="Verdana" size="2" color="#000000">
<b>Public Sub DelTree(ByVal vDir As Variant)<br>
Dim FSO, FS<p>
Set FSO = CreateObject("Scripting.FileSystemObject")<br>
FS = FSO.deletefolder(vDir, True)<p>
End Sub
```

