<div align="center">

## Change form backround with common dialog


</div>

### Description

Kind of like skinning your program
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Demian Net](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/demian-net.md)
**Level**          |Intermediate
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/demian-net-change-form-backround-with-common-dialog__1-6975/archive/master.zip)





### Source Code

```
' Name your form Form1
 ' Load comdlg32.ocx
 ' Make a Command1
 ' Make a common dialog named CDialog
 Private Sub Command1_Click()
   On Error GoTo fileOpenErrr
    CDialog.CancelError = True
    CDialog.FLAGS = &H4& Or &H100&
    CDialog.DefaultExt = ".jpg"
    CDialog.DialogTitle = "Select File To Open"
    CDialog.Filter = "JPEG (*.jpg)|*.jpg|GIF (*.gif)|*.gif|BITMAP (*.bmp)|*.bmp"
    CDialog.ShowOpen
 Set Form1.Picture = LoadPicture(CDialog.filename)
 fileOpenErrr:
    Exit Sub
 End Sub
 ' This is what I use for a sort of skin effect on my programs.
```

