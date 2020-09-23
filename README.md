<div align="center">

## Get a chunk from a file, but specify start and finish \(usefull for splitting files  / P2P program\)


</div>

### Description

This code can be made to take a certain chunk of your choice from a file. Mainly for making resumable downloads etc, or working on a patch for a program... I made this cos im working on a file sharing program that downloads from multiple sources. SERVERAL UPDATES HAVE TAKEN PLACE (thanks to guys who posted comments improving the code)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jon Barker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jon-barker.md)
**Level**          |Intermediate
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jon-barker-get-a-chunk-from-a-file-but-specify-start-and-finish-usefull-for-splitting-file__1-35844/archive/master.zip)





### Source Code

```
Option Explicit
Sub GetFilePart(FilenameFROM As String, FilenameTO As String, StartPos As Long, EndPos As Long)
 Dim bData() As Byte
 Dim sBuf As String
 Dim i As Long
 Dim l As Long
 Dim File1 as Integer
 Dim File2 as Integer
 File2 = Freefile
 Open FilenameTO For Binary As #File2
 File1 = Freefile
 Open FilenameFROM For Binary As #File1
 Get #File1, StartPos, bData
 ReDim bData(1 To 2048) As Byte
 For l = 1 To (EndPos - StartPos) \ 2048
 Get #File1, 2048 * (l - 1) + StartPos, bData
 Put #File2, , bData
 Next
 i = (EndPos - StartPos) Mod 2048
 If i <> 0 Then
 ReDim bData(1 To i) As Byte
 Get #File1, , bData
 Put #File2, , bData
 End If
 Close #File1
 Close #File2
End Sub
Private Sub cmdStart_Click()
 Call GetFilePart("c:\filename.exe", "C:\filename2.exe", 100, 3000)
 'this example above would get bytes 100 to 3000 in the program 'c:\filename.exe', and put them into
 'the file 'C:\filename2.exe'. The size of 'C:\filename2.exe' is now 2900 byes.
End Sub
```

