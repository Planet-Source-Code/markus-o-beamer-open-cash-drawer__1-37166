<div align="center">

## Open Cash Drawer


</div>

### Description

This piece of code will pop open an MMF ED 2000 Cash Register. It sends the ascii value of A down com port 1. I put this code on planetsource code cause I had a hard time finding out how to do this. I tried the MSComm object and had a lot of difficulty with it. Hope this helps all the coders out there.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Markus O Beamer](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/markus-o-beamer.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/markus-o-beamer-open-cash-drawer__1-37166/archive/master.zip)





### Source Code

```
Sub popDrawer()
On Error GoTo err
 If hasCashDrawer Then
  Open "COM1" For Output Access Write As #1
    Print #1, Chr$(65); "A";
  Close #1
 End If
 Exit Sub
err:
 MsgBox ("Problems opening drawer")
End Sub
```

