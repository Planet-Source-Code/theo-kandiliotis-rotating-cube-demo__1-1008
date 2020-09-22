<div align="center">

## Rotating Cube Demo


</div>

### Description

The HELLO WORLD of 3D programming :)

A rotating cube demo. You can adjust the rotation (clockwise or anticlockwise) and the speed by moving the mouse cursor towards the right and left edges of the form.
 
### More Info
 
This is Math stuff mostly,has little to do with VB's methods. I used the Line method to do all the drawing on the form.

To run this project ,you have to paste the text indicated, into a blank text file and save it in ASCII format with an FRM suffix.

Notice that it WILL NOT work if you paste it directly on a VB5 code window...Also make sure that no lines are wrapped before you save it as FRM.

Have fun!

No Side Effects


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ Theo Kandiliotis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/theo-kandiliotis.md)
**Level**          |Unknown
**User Rating**    |5.9 (609 globes from 103 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/theo-kandiliotis-rotating-cube-demo__1-1008/archive/master.zip)

### API Declarations

No API's


### Source Code

```

------======== start copying AFTER this line ======---------
VERSION 5.00
Begin VB.Form frmMain
  AutoRedraw   =  -1 'True
  BackColor    =  &H00C0C0C0&
  Caption     =  "Rotating Cube DEMO"
  ClientHeight  =  3195
  ClientLeft   =  60
  ClientTop    =  345
  ClientWidth   =  4680
  FillColor    =  &H00C0C0C0&
  ForeColor    =  &H00FF0000&
  LinkTopic    =  "Form1"
  ScaleHeight   =  213
  ScaleMode    =  3 'Pixel
  ScaleWidth   =  312
  StartUpPosition =  3 'Windows Default
  WindowState   =  2 'Maximized
  Begin VB.PictureBox Picture1
   BackColor    =  &H00FFFFFF&
   BorderStyle   =  0 'None
   Height     =  1140
   Left      =  -1035
   ScaleHeight   =  76
   ScaleMode    =  3 'Pixel
   ScaleWidth   =  772
   TabIndex    =  0
   Top       =  1440
   Width      =  11580
   Begin VB.Label Label1
     AutoSize    =  -1 'True
     Caption     =  "Move the mouse towards the edges of the form to adjust rotation and speed"
     BeginProperty Font
      Name      =  "MS Sans Serif"
      Size      =  12
      Charset     =  161
      Weight     =  700
      Underline    =  0  'False
      Italic     =  0  'False
      Strikethrough  =  0  'False
     EndProperty
     Height     =  300
     Left      =  0
     TabIndex    =  1
     Top       =  0
     Width      =  9135
   End
  End
  Begin VB.Timer Timer1
   Interval    =  1
   Left      =  3825
   Top       =  2835
  End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private X(8) As Integer
Private y(8) As Integer
Private Const Pi = 3.14159265358979
Private CenterX As Integer
Private CenterY As Integer
Private Const SIZE = 250
Private Radius As Integer
Private Angle As Integer
Private CurX As Integer
Private CurY As Integer
Private CubeCorners(1 To 8, 1 To 3) As Integer
Private Sub Form_Load()
Show
With Picture1
.Width = Label1.Width
.Height = Label1.Height
End With
Picture1.Move ScaleWidth / 2 - Picture1.ScaleWidth / 2, Picture1.Height
CenterX = ScaleWidth / 2
CenterY = ScaleHeight / 2
Angle = 0
Radius = Sqr(2 * (SIZE / 2) ^ 2)
CubeCorners(1, 2) = SIZE / 2
CubeCorners(2, 2) = SIZE / 2
CubeCorners(3, 2) = -SIZE / 2
CubeCorners(4, 2) = -SIZE / 2
CubeCorners(5, 2) = SIZE / 2
CubeCorners(6, 2) = SIZE / 2
CubeCorners(7, 2) = -SIZE / 2
CubeCorners(8, 2) = -SIZE / 2
End Sub
Private Sub DrawCube()
Cls
For i = 1 To 8
X(i) = CenterX + CubeCorners(i, 1) - CubeCorners(i, 3) / 8
y(i) = CenterY + CubeCorners(i, 2) + CubeCorners(i, 3) / 8
Next
Line (X(3), y(3))-(X(4), y(4))
Line (X(4), y(4))-(X(8), y(8))
Line (X(3), y(3))-(X(7), y(7))
Line (X(7), y(7))-(X(8), y(8))
Line (X(1), y(1))-(X(3), y(3))
Line (X(1), y(1))-(X(2), y(2))
Line (X(5), y(5))-(X(6), y(6))
Line (X(5), y(5))-(X(1), y(1))
Line (X(5), y(5))-(X(7), y(7))
Line (X(6), y(6))-(X(8), y(8))
Line (X(2), y(2))-(X(4), y(4))
Line (X(2), y(2))-(X(6), y(6))
Line (X(1), y(1))-(X(4), y(4))
Line (X(2), y(2))-(X(3), y(3))
Line (X(4), y(4))-(X(8), y(8))
Line (X(3), y(3))-(X(7), y(7))
DoEvents
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
CurX = X
CurY = y
End Sub
Private Sub Timer1_Timer()
Select Case CurX
Case Is > ScaleWidth / 2
Angle = Angle + Abs(CurX - ScaleWidth / 2) / 20
If Angle > 360 Then Angle = 0
Case Else
Angle = Angle - Abs(CurX - ScaleWidth / 2) / 20
If Angle < 0 Then Angle = 360
End Select
For i = 1 To 3 Step 2
CubeCorners(i, 3) = Radius * Cos((Angle) * Pi / 180)
CubeCorners(i, 1) = Radius * Sin((Angle) * Pi / 180)
Next
For i = 2 To 4 Step 2
CubeCorners(i, 3) = Radius * Cos((Angle + 2 * 45) * Pi / 180)
CubeCorners(i, 1) = Radius * Sin((Angle + 2 * 45) * Pi / 180)
Next
For i = 5 To 7 Step 2
CubeCorners(i, 3) = Radius * Cos((Angle + 6 * 45) * Pi / 180)
CubeCorners(i, 1) = Radius * Sin((Angle + 6 * 45) * Pi / 180)
Next
For i = 6 To 8 Step 2
CubeCorners(i, 3) = Radius * Cos((Angle + 4 * 45) * Pi / 180)
CubeCorners(i, 1) = Radius * Sin((Angle + 4 * 45) * Pi / 180)
Next
DrawCube
End Sub
-----==== paste the above into a text file and save it with
an FRM suffix in ASCII format.Then just load the FRM file
in the VB5 enviroment  =========-------------------------
```

