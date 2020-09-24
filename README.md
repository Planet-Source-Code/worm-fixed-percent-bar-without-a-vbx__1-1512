<div align="center">

## \*FIXED\* Percent Bar Without a VBX


</div>

### Description

This function will let you make a percent bar without a vbx. Now you don't have to add a vbx to your program making it a bigger hassle, plus you can now pick any color percent bar you want
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Worm](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/worm.md)
**Level**          |Unknown
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/worm-fixed-percent-bar-without-a-vbx__1-1512/archive/master.zip)

### API Declarations

```
Sub precentworm (picture As Control, ByVal percent)
Dim num$
If Not picture.AutoRedraw Then
picture.AutoRedraw = -1
End If
picture.Cls
picture.ScaleWidth = 100
picture.DrawMode = 10
num$ = Format$(percent, "###") + "%"
picture.CurrentX = 50 - picture.TextWidth(num$) / 2
picture.CurrentY = (picture.ScaleHeight - picture.TextHeight(num$)) / 2
picture.Print num$
picture.Line (0, 0)-(percent, picture.ScaleHeight), , BF
picture.Refresh
End Sub
```


### Source Code

```
'Put this in a command button or any control you want
picture1.ForeColor = RGB(0, 0, 255) 'blue bar but you can change it
For i = 0 To 100 Step 2
precentworm picture1, i
Next
picture.Cls
```

