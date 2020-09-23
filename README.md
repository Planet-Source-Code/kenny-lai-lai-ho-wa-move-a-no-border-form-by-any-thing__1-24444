<div align="center">

## Move a No\-Border form by any thing\!


</div>

### Description

To have a better UI, programmers usually use a borderless form. But there's no caption bar and the form cannot be move. This code help you to move a form by clicking anything that support mousemove, mouseup, mousedown events. See it , it's cool!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kenny Lai, Lai Ho Wa](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kenny-lai-lai-ho-wa.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kenny-lai-lai-ho-wa-move-a-no-border-form-by-any-thing__1-24444/archive/master.zip)





### Source Code

<p><b><<<General_Declaration>>></b><br>
Dim ism as boolean</p>
<p><br>
Public Type POINTAPI<br>
  x As Long<br>
  y As Long<br>
End Type</p>
<p><br>
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long</p>
<p><br>
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long</p>
<hr>
<p><br>
<br>
Private Sub lbltitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)<br>
ism = True<br>
x1 = x + lblTitle.Left<br>
y1 = y + lblTitle.Top<br>
End Sub</p>
<hr>
<p><br>
Private Sub lbltitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)<br>
If ism = True Then<br>
i = GetCursorPos(Pos)<br>
x2 = Pos.x * Screen.TwipsPerPixelX<br>
y2 = Pos.y * Screen.TwipsPerPixelY<br>
Me.Move (x2 - x1), (y2 - y1)<br>
End If<br>
End Sub</p>
<hr>
<p><br>
Private Sub lbltitle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)<br>
If Me.Top &lt; 0 Then Me.Top = 0<br>
If Me.Left &lt; 0 Then Me.Left = 0<br>
If Me.Top > (Screen.Height - (Me.Height / 10)) Then Me.Top = Screen.Height * 9 / 10<br>
If Me.Left > (Screen.Width - (Me.Width / 10)) Then Me.Left = Screen.Width * 9 / 10<br>
ism = False<br>
End Sub</p>
</body>
</html>

