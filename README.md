<div align="center">

## Find and Replace in multiple files


</div>

### Description

Find and replace text strings in multiple files within a directory. This allows for wilcards. Please Vote if you like it! :)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tom Bruinsma](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tom-bruinsma.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tom-bruinsma-find-and-replace-in-multiple-files__1-23197/archive/master.zip)





### Source Code

```
Dim strFind As String
 Dim strReplace As String
 Dim strDestination As String
 Dim strSource As String
 Dim strFilter
 strFind = "Hello" 'What to find
 strReplace = "Goodbye" 'What to replace it with
 strDestination = "c:\temp" 'Where to put the files once they have been modified
 strSource = "c:\output" 'Where to get the files
 strFilter = "*.txt" 'wildcards
 'verification complete
 Dim parse As String
 Dim hold As String
 'FIND AND REPLACE
 sdir = Dir(strSource & "\" & fMainForm.txtIncludeFilter)
 Do While sdir <> ""
 Open fMainForm.txtSource & "\" & sdir For Input As #1
 Do While Not EOF(1)
 Line Input #1, parse
 hold = hold & Replace(parse, fMainForm.txtReplace, fMainForm.txtFind)
 Loop
 Loop
 Open fMainForm.txtSource & "\" & sdir For Output As #1
 Print #1, hold
 Close #1
 hold = ""
 parse = ""
 sdir = Dir()
 Loop
```

