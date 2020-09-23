<div align="center">

## Open Folders As Windows from VB5


</div>

### Description

This code is for those people who want to open up folders/directories as separate windows, as compared to the alternative fileboxes.
 
### More Info
 
Opens a folder/directory in a separate window


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ted R\. Smith](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ted-r-smith.md)
**Level**          |Unknown
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ted-r-smith-open-folders-as-windows-from-vb5__1-1051/archive/master.zip)





### Source Code

```
Directory = "C:\"
Shell "Explorer " + Directory, vbNormalFocus
' The above code opens the C:\ directory as a new window
```

