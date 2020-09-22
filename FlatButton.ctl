VERSION 5.00
Begin VB.UserControl FlatButton 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   990
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   MouseIcon       =   "FlatButton.ctx":0000
   MousePointer    =   99  'Benutzerdefiniert
   PropertyPages   =   "FlatButton.ctx":08CA
   ScaleHeight     =   960
   ScaleWidth      =   990
   ToolboxBitmap   =   "FlatButton.ctx":0912
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   0
      MouseIcon       =   "FlatButton.ctx":0C24
      MousePointer    =   99  'Benutzerdefiniert
      ScaleHeight     =   945
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   0
      Width           =   975
      Begin VB.PictureBox PictureDraw 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   375
         Left            =   540
         ScaleHeight     =   315
         ScaleWidth      =   345
         TabIndex        =   3
         Top             =   90
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.PictureBox PictureOrg 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   375
         Left            =   60
         ScaleHeight     =   315
         ScaleWidth      =   345
         TabIndex        =   2
         Top             =   90
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FlatButton"
         Height          =   210
         Left            =   105
         MouseIcon       =   "FlatButton.ctx":14EE
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   1
         Top             =   660
         Width           =   735
      End
      Begin VB.Image Picture2 
         Height          =   480
         Left            =   240
         MouseIcon       =   "FlatButton.ctx":1DB8
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "FlatButton.ctx":2682
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Timer Timer1 
      Left            =   210
      Top             =   390
   End
End
Attribute VB_Name = "FlatButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
        
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
        
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
   
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private WithEvents eForm As Form
Attribute eForm.VB_VarHelpID = -1
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_ALIAS = &H10000
Const M_Def_Font = "FlatButton"

Private Type POINTAPI
x As Long
y As Long
End Type

Dim Status As String
Dim PictureLeft As Double
Dim PictureTop As Double
Dim LabelLeft As Double
Dim Labeltop As Double
Dim M_Font As String
Dim M_Hit As Boolean
Dim M_Default As Boolean
Dim M_Cancel As Boolean
Dim M_Sound As Boolean
Dim M_Cursor As Integer
Dim M_PicAlign As Integer
Dim M_Style As Integer
Dim M_ButtonType As Integer
Dim M_GreyIcon As Boolean

Event Click()
Event DblClick()
Event MouseMove()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseOver()
Event MouseOut()

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Dim rRed As Integer, rGreen As Integer, rBlue As Integer


Sub GreyPic()
Dim AveCol As Integer, a As Integer, Total As Long
Dim x As Double
Dim y As Double
On Error Resume Next

Total = (PictureOrg.Height * PictureOrg.Width)
For y = 0 To PictureOrg.Height Step 15
For x = 0 To PictureOrg.Width Step 15
AveCol = 0
a = 0
RGBfromLONG (GetPixel(PictureOrg.hdc, x / 15, y / 15))
Rem AveCol = AveCol + rRed: a = a + 1
Rem AveCol = AveCol + rBlue: a = a + 1
AveCol = AveCol + rGreen: a = a + 1
If AveCol <= 0 Then AveCol = 0
AveCol = (AveCol / a)

If (GetPixel(PictureOrg.hdc, x / 15, y / 15)) <> PictureOrg.BackColor Then
SetPixel PictureDraw.hdc, x / 15, y / 15, RGB(AveCol, AveCol, AveCol) ' set pixel
Else
SetPixel PictureDraw.hdc, x / 15, y / 15, PictureOrg.BackColor ' set pixel
End If

Next x
PictureDraw.Refresh
Next y
On Error GoTo 0
End Sub

Private Function RGBfromLONG(LongCol As Long)
Dim Blue As Double, Green As Double, Red As Double, GreenS As Double, BlueS As Double
On Error Resume Next
Blue = Fix((LongCol / 256) / 256)
Green = Fix((LongCol - ((Blue * 256) * 256)) / 256)
Red = Fix(LongCol - ((Blue * 256) * 256) - (Green * 256))
rRed = Red: rBlue = Blue: rGreen = Green
On Error GoTo 0
End Function


Sub About()
On Error Resume Next
FrmAbout.Show vbModal
On Error GoTo 0
End Sub

Sub ChangeStyle()
Dim Result&, Handle&, Parent, P As POINTAPI, AA$, x&
Dim Fl&, Ft&, Fw&, Fh&
Dim Snd As Variant
On Error Resume Next

If ButtonType = 9 Then Exit Sub

Result = GetCursorPos(P)
Handle = WindowFromPoint(P.x, P.y)
Parent = GetParent(Handle)

Fl = ScaleX(Picture1.Left, vbTwips, vbPixels)
Ft = ScaleY(Picture1.Top, vbTwips, vbPixels)
Fw = ScaleX(Picture1.Width, vbTwips, vbPixels)
Fh = ScaleY(Picture1.Height, vbTwips, vbPixels)


Label1.MousePointer = M_Cursor
Picture1.MousePointer = M_Cursor
Picture2.MousePointer = M_Cursor
UserControl.MousePointer = M_Cursor

If (P.x < Fl) Or (P.y < Ft) Or (P.x > Fl + Fw) Or (P.y > Ft + Fh) Then
If Parent <> UserControl.hwnd Then
If Status = "OFF" Then Exit Sub
Picture1.Cls
Label1.FontUnderline = False
Status = "OFF"
Timer1.Interval = 0
If GreyIcon = True Then
Picture2.Picture = PictureDraw.Image
Else
Picture2.Picture = PictureOrg.Image
End If
Label1.Left = LabelLeft
Label1.Top = Labeltop
Picture2.Left = PictureLeft
Picture2.Top = PictureTop
RaiseEvent MouseOut
Else
If Status = "ON" Then Exit Sub
Call FlatButtonOver
If Hit = True Then Label1.FontUnderline = True
Status = "ON"
Snd = "MenuCommand"
If Sound = True Then Result = PlaySound(Snd, 0&, SND_ALIAS Or SND_ASYNC Or SND_NODEFAULT)
RaiseEvent MouseOver
End If
Else
End If

On Error GoTo 0
End Sub

Sub FlatButtonOver()
Dim Länge As Double
On Error Resume Next

If ButtonType = 9 Then Exit Sub

Select Case Style

Case Is = 0
Picture1.Cls

Case Is = 1
Picture1.Line (Picture1.Width, 1)-(1, 1), RGB(255, 255, 255), BF
Picture1.Line (1, Picture1.Height)-(1, 1), RGB(255, 255, 255), BF
Picture1.Line (Picture1.Width, Picture1.Height - 10)-(1, Picture1.Height - 10), RGB(50, 50, 50), BF
Picture1.Line (Picture1.Width - 10, Picture1.Height)-(Picture1.Width - 10, 1), RGB(50, 50, 50), BF

Case Is = 2
Picture1.Line (Picture1.Width - (Picture1.Height * 0.2), 1)-(Picture1.Height * 0.2, 1), RGB(255, 255, 255), BF
Picture1.Line (1, Picture1.Height - (Picture1.Height * 0.2))-(1, Picture1.Height * 0.2), RGB(255, 255, 255), BF
Picture1.Circle (Picture1.Height * 0.2, Picture1.Height * 0.2), Picture1.Height * 0.2, RGB(255, 255, 255), 1.5, 3.25
Picture1.Line (Picture1.Width - (Picture1.Height * 0.2), Picture1.Height - 10)-(Picture1.Height * 0.2, Picture1.Height - 10), RGB(50, 50, 50), BF
Picture1.Line (Picture1.Width - 10, Picture1.Height - (Picture1.Height * 0.2))-(Picture1.Width - 10, (Picture1.Height * 0.2) - 75), RGB(50, 50, 50), BF
Picture1.Circle (Picture1.Width - 10 - (Picture1.Height * 0.2), Picture1.Height - 10 - (Picture1.Height * 0.2)), Picture1.Height * 0.2, RGB(50, 50, 50), 4.6, 0
Picture1.Circle (Picture1.Height * 0.2, Picture1.Height - 10 - (Picture1.Height * 0.2)), Picture1.Height * 0.2, RGB(255, 255, 255), 3.1, 4
Picture1.Circle (Picture1.Height * 0.2, Picture1.Height - 10 - (Picture1.Height * 0.2)), Picture1.Height * 0.2, RGB(0, 0, 0), 4, 4.8
Picture1.Circle (Picture1.Width - (Picture1.Height * 0.2), Picture1.Height * 0.2), Picture1.Height * 0.2, RGB(255, 255, 255), 0.75, 1.65
Picture1.Circle (Picture1.Width - (Picture1.Height * 0.2), Picture1.Height * 0.2), Picture1.Height * 0.2, RGB(50, 50, 50), 0, 0.75

Case Is = 3
If Picture1.Height <= Picture1.Width Then
Länge = Picture1.Height
Else
Länge = Picture1.Width
End If
Picture1.Circle ((Picture1.Width / 2) - 25, (Picture1.Height / 2) - 25), Länge / 2.1, RGB(255, 255, 255), 0.8, 3.9
Picture1.Circle ((Picture1.Width / 2) - 25, (Picture1.Height / 2) - 25), Länge / 2.1, RGB(50, 50, 50), 3.9, 0.8

Case Is = 4
Picture1.Line (1, Picture1.Height)-(Picture1.Width / 2, 1), RGB(255, 255, 255)
Picture1.Line (Picture1.Width, Picture1.Height)-(Picture1.Width / 2, 1), RGB(50, 50, 50)
Picture1.Line (Picture1.Width, Picture1.Height - 10)-(1, Picture1.Height - 10), RGB(50, 50, 50), BF

Case Is = 5
Picture1.Line (Picture1.Width - (Picture1.Height * 0.2), 1)-(Picture1.Height * 0.2, 1), RGB(255, 255, 255), BF
Picture1.Line (1, Picture1.Height - (Picture1.Height * 0.2))-(1, Picture1.Height * 0.2), RGB(255, 255, 255), BF
Picture1.Line ((Picture1.Height * 0.2), 1)-(1, Picture1.Height * 0.2), RGB(255, 255, 255)
Picture1.Line (Picture1.Width - (Picture1.Height * 0.2), 1)-(Picture1.Width, Picture1.Height * 0.2), RGB(50, 50, 50)
Picture1.Line (Picture1.Width - (Picture1.Height * 0.2), Picture1.Height - 10)-(Picture1.Height * 0.2, Picture1.Height - 10), RGB(50, 50, 50), BF
Picture1.Line (Picture1.Width - 10, Picture1.Height - (Picture1.Height * 0.2))-(Picture1.Width - 10, (Picture1.Height * 0.2)), RGB(50, 50, 50), BF
Picture1.Line (1, Picture1.Height - (Picture1.Height * 0.2))-(Picture1.Height * 0.2, Picture1.Height), RGB(255, 255, 255)
Picture1.Line (Picture1.Width - (Picture1.Height * 0.2), Picture1.Height)-(Picture1.Width, Picture1.Height - (Picture1.Height * 0.2)), RGB(50, 50, 50)

End Select

Picture2.Picture = PictureOrg.Image

Label1.Left = LabelLeft - 15
Label1.Top = Labeltop - 15
Picture2.Left = PictureLeft - 15
Picture2.Top = PictureTop - 15

On Error GoTo 0
End Sub

Sub FlatButtonOn()
Dim Länge As Double
On Error Resume Next

If ButtonType = 9 Then Exit Sub

Select Case Style

Case Is = 0
Picture1.Cls

Case Is = 1
Picture1.Line (Picture1.Width, 1)-(1, 1), RGB(0, 0, 0), BF
Picture1.Line (1, Picture1.Height)-(1, 1), RGB(0, 0, 0), BF
Picture1.Line (Picture1.Width, Picture1.Height - 10)-(1, Picture1.Height - 10), RGB(255, 255, 255), BF
Picture1.Line (Picture1.Width - 10, Picture1.Height)-(Picture1.Width - 10, 1), RGB(255, 255, 255), BF

Case Is = 2
Picture1.Line (Picture1.Width - (Picture1.Height * 0.2), 1)-(Picture1.Height * 0.2, 1), RGB(50, 50, 50), BF
Picture1.Line (1, Picture1.Height - (Picture1.Height * 0.2))-(1, Picture1.Height * 0.2), RGB(50, 50, 50), BF
Picture1.Circle (Picture1.Height * 0.2, Picture1.Height * 0.2), Picture1.Height * 0.2, RGB(50, 50, 50), 1.5, 3.25
Picture1.Line (Picture1.Width - (Picture1.Height * 0.2), Picture1.Height - 10)-(Picture1.Height * 0.2, Picture1.Height - 10), RGB(255, 255, 255), BF
Picture1.Line (Picture1.Width - 10, Picture1.Height - (Picture1.Height * 0.2))-(Picture1.Width - 10, (Picture1.Height * 0.2) - 75), RGB(255, 255, 255), BF
Picture1.Circle (Picture1.Width - 10 - (Picture1.Height * 0.2), Picture1.Height - 10 - (Picture1.Height * 0.2)), Picture1.Height * 0.2, RGB(255, 255, 255), 4.6, 0
Picture1.Circle (Picture1.Height * 0.2, Picture1.Height - 10 - (Picture1.Height * 0.2)), Picture1.Height * 0.2, RGB(50, 50, 50), 3.1, 4
Picture1.Circle (Picture1.Height * 0.2, Picture1.Height - 10 - (Picture1.Height * 0.2)), Picture1.Height * 0.2, RGB(255, 255, 255), 4, 4.8
Picture1.Circle (Picture1.Width - (Picture1.Height * 0.2), Picture1.Height * 0.2), Picture1.Height * 0.2, RGB(50, 50, 50), 0.75, 1.65
Picture1.Circle (Picture1.Width - (Picture1.Height * 0.2), Picture1.Height * 0.2), Picture1.Height * 0.2, RGB(255, 255, 255), 0, 0.75

Case Is = 3
If Picture1.Height <= Picture1.Width Then
Länge = Picture1.Height
Else
Länge = Picture1.Width
End If
Picture1.Circle ((Picture1.Width / 2) - 25, (Picture1.Height / 2) - 25), Länge / 2.1, RGB(50, 50, 50), 0.8, 3.9
Picture1.Circle ((Picture1.Width / 2) - 25, (Picture1.Height / 2) - 25), Länge / 2.1, RGB(255, 255, 255), 3.9, 0.8

Case Is = 4
Picture1.Line (1, Picture1.Height)-(Picture1.Width / 2, 1), RGB(50, 50, 50)
Picture1.Line (Picture1.Width, Picture1.Height)-(Picture1.Width / 2, 1), RGB(255, 255, 255)
Picture1.Line (Picture1.Width, Picture1.Height - 10)-(1, Picture1.Height - 10), RGB(255, 255, 255), BF

Case Is = 5
Picture1.Line (Picture1.Width - (Picture1.Height * 0.2), 1)-(Picture1.Height * 0.2, 1), RGB(50, 50, 50), BF
Picture1.Line (1, Picture1.Height - (Picture1.Height * 0.2))-(1, Picture1.Height * 0.2), RGB(50, 50, 50), BF
Picture1.Line ((Picture1.Height * 0.2), 1)-(1, Picture1.Height * 0.2), RGB(50, 50, 50)
Picture1.Line (Picture1.Width - (Picture1.Height * 0.2), 1)-(Picture1.Width, Picture1.Height * 0.2), RGB(255, 255, 255)
Picture1.Line (Picture1.Width - (Picture1.Height * 0.2), Picture1.Height - 10)-(Picture1.Height * 0.2, Picture1.Height - 10), RGB(255, 255, 255), BF
Picture1.Line (Picture1.Width - 10, Picture1.Height - (Picture1.Height * 0.2))-(Picture1.Width - 10, (Picture1.Height * 0.2)), RGB(255, 255, 255), BF
Picture1.Line (1, Picture1.Height - (Picture1.Height * 0.2))-(Picture1.Height * 0.2, Picture1.Height), RGB(50, 50, 50)
Picture1.Line (Picture1.Width - (Picture1.Height * 0.2), Picture1.Height)-(Picture1.Width, Picture1.Height - (Picture1.Height * 0.2)), RGB(255, 255, 255)

End Select

Picture2.Picture = PictureOrg.Image

Label1.Left = LabelLeft + 15
Label1.Top = Labeltop + 15
Picture2.Left = PictureLeft + 15
Picture2.Top = PictureTop + 15

On Error GoTo 0
End Sub

Sub FlatButtonOff()
On Error Resume Next
If ButtonType = 9 Then Exit Sub
Label1.Left = LabelLeft
Label1.Top = Labeltop
Picture2.Left = PictureLeft
Picture2.Top = PictureTop
Call FlatButtonOver
On Error GoTo 0
End Sub

Private Sub eForm_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn And Default = True Then
UserControl.SetFocus
End If
If KeyCode = vbKeyEscape And Cancel = True Then
UserControl.SetFocus
End If
On Error GoTo 0
End Sub

Private Sub Label1_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Picture1_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Picture2_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Timer1_Timer()
Static done_before As Boolean
Static CurPosLast As POINTAPI
Dim CurPosAkt As POINTAPI
Call GetCursorPos(CurPosAkt)

If (CurPosAkt.x <> CurPosLast.x) Or (CurPosAkt.y <> CurPosLast.y) Then
Call ChangeStyle
Else

End If

CurPosLast = CurPosAkt
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
RaiseEvent MouseDown(Button, Shift, x, y)
If Button <> 1 Then Exit Sub
Status = "ON"
Call FlatButtonOn
On Error GoTo 0
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Status = ""
Call ChangeStyle
Call FlatButtonOff
If Status = "ON" And Button = 1 Then RaiseEvent Click
RaiseEvent MouseUp(Button, Shift, x, y)
On Error GoTo 0
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Status = "ON" Then Exit Sub
Status = "OFF"
Call ChangeStyle
Timer1.Interval = 50
RaiseEvent MouseMove
On Error GoTo 0
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
RaiseEvent MouseDown(Button, Shift, x, y)
If Button <> 1 Then Exit Sub
Status = "ON"
Call FlatButtonOn
On Error GoTo 0
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Status = ""
Call ChangeStyle
Call FlatButtonOff
If Status = "ON" And Button = 1 Then RaiseEvent Click
RaiseEvent MouseUp(Button, Shift, x, y)
On Error GoTo 0
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Status = "ON" Then Exit Sub
Status = "OFF"
Call ChangeStyle
Timer1.Interval = 50
RaiseEvent MouseMove
On Error GoTo 0
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
RaiseEvent MouseDown(Button, Shift, x, y)
If Button <> 1 Then Exit Sub
Status = "ON"
Call FlatButtonOn
On Error GoTo 0
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Status = ""
Call ChangeStyle
Call FlatButtonOff
If Status = "ON" And Button = 1 Then RaiseEvent Click
RaiseEvent MouseUp(Button, Shift, x, y)
On Error GoTo 0
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Status = "ON" Then Exit Sub
Status = "OFF"
Call ChangeStyle
Timer1.Interval = 50
RaiseEvent MouseMove
On Error GoTo 0
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
On Error Resume Next
UserControl.Width = UserControl.Width + 10
UserControl.Width = UserControl.Width - 10
On Error GoTo 0
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_InitProperties()
Label1.Caption = "FlatButton"
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If Default = True And KeyCode = vbKeyReturn Then
Call FlatButtonOn
RaiseEvent Click
Call FlatButtonOff
End If

If Cancel = True And KeyCode = vbKeyEscape Then
Call FlatButtonOn
DoEvents
Call FlatButtonOff
Unload eForm
End If

On Error GoTo 0
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Status = ""
Call ChangeStyle
Call FlatButtonOff
If Status = "ON" And Button = 1 Then RaiseEvent Click
RaiseEvent MouseUp(Button, Shift, x, y)
On Error GoTo 0
End Sub

Sub ButtonRePaint()
On Error Resume Next

If ButtonType = 9 Then
Picture2.Visible = False
Label1.Visible = False
UserControl.Width = 105
Picture1.Width = 105
Picture1.Line (45, Picture1.Height)-(50, 1), RGB(125, 125, 125), BF
Picture1.Line (55, Picture1.Height)-(60, 1), RGB(255, 255, 255), BF
UserControl.MousePointer = 0
Picture1.MousePointer = 0
Else
Picture1.Cls
Picture2.Visible = True
Label1.Visible = True
End If
Picture1.Left = 0
Picture1.Top = 0
Picture1.Width = UserControl.Width
Picture1.Height = UserControl.Height
DoEvents

Select Case PicAlign

Case Is = 0
Picture2.Top = 25
Picture2.Left = (UserControl.Width - Picture2.Width) / 2
Label1.Top = (UserControl.Height - Label1.Height) - 25
Label1.Left = (UserControl.Width - Label1.Width) / 2

Case Is = 1
Picture2.Top = 25
Picture2.Left = 25
Label1.Top = (UserControl.Height - Label1.Height) - 25
Label1.Left = (UserControl.Width - Label1.Width) / 2

Case Is = 2
Picture2.Top = 25
Picture2.Left = (UserControl.Width - Picture2.Width) - 25
Label1.Top = (UserControl.Height - Label1.Height) - 25
Label1.Left = (UserControl.Width - Label1.Width) / 2

Case Is = 3
Picture2.Top = (UserControl.Height - Picture2.Height) / 2
Picture2.Left = (UserControl.Width - Picture2.Width) / 2
Label1.Top = (UserControl.Height - Label1.Height) / 2
Label1.Left = (UserControl.Width - Label1.Width) / 2

Case Is = 4
Picture2.Top = (UserControl.Height - Picture2.Height) / 2
Picture2.Left = 50
Label1.Top = (UserControl.Height - Label1.Height) / 2
Label1.Left = Picture2.Width + Picture2.Left + (((UserControl.Width - (Picture2.Width + Picture2.Left)) - Label1.Width) / 2)

Case Is = 5
Picture2.Top = (UserControl.Height - Picture2.Height) / 2
Picture2.Left = (UserControl.Width - Picture2.Width) - 50
Label1.Top = (UserControl.Height - Label1.Height) / 2
Label1.Left = ((UserControl.Width - (Picture2.Width + 50)) - Label1.Width) / 2

Case Is = 6
Picture2.Top = (UserControl.Height - Picture2.Height) - 25
Picture2.Left = (UserControl.Width - Picture2.Width) / 2
Label1.Top = 25
Label1.Left = (UserControl.Width - Label1.Width) / 2

Case Is = 7
Picture2.Top = (UserControl.Height - Picture2.Height) - 25
Picture2.Left = 25
Label1.Top = 25
Label1.Left = (UserControl.Width - Label1.Width) / 2

Case Is = 8
Picture2.Top = (UserControl.Height - Picture2.Height) - 25
Picture2.Left = (UserControl.Width - Picture2.Width) - 25
Label1.Top = 25
Label1.Left = (UserControl.Width - Label1.Width) / 2


End Select

If GreyIcon = True Then
Picture2.Picture = PictureDraw.Image
Else
Picture2.Picture = PictureOrg.Image
End If

DoEvents
Labeltop = Label1.Top
LabelLeft = Label1.Left
PictureTop = Picture2.Top
PictureLeft = Picture2.Left
On Error GoTo 0
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Call ButtonRePaint
On Error GoTo 0
End Sub

Private Sub UserControl_Show()
On Error Resume Next
With UserControl
If Ambient.UserMode Then
If TypeOf .Parent Is MDIForm Then
ElseIf TypeOf .Parent Is Form Then
Set eForm = .Parent
End If
End If
End With
On Error GoTo 0
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next
Timer1.Interval = 0
On Error GoTo 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
Timer1.Interval = 50
Label1.Caption = PropBag.ReadProperty("Caption", "Label1")
Picture1.BackColor = PropBag.ReadProperty("Backcolor", &H8000000F)
PictureOrg.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
PictureDraw.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
Label1.Caption = PropBag.ReadProperty("Caption", "FlatButton")
Label1.FontBold = PropBag.ReadProperty("FontBold", 0)
Label1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
Label1.FontName = PropBag.ReadProperty("FontName", "")
Label1.FontSize = PropBag.ReadProperty("FontSize", 0)
Label1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
Label1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)

Hit = PropBag.ReadProperty("Hit", Nothing)
Default = PropBag.ReadProperty("Default", Nothing)
Cancel = PropBag.ReadProperty("Cancel", Nothing)
Sound = PropBag.ReadProperty("Sound", Nothing)
Cursor = PropBag.ReadProperty("Cursor", Nothing)
PicAlign = PropBag.ReadProperty("PicAlign", Nothing)
Style = PropBag.ReadProperty("Style", Nothing)
ButtonType = PropBag.ReadProperty("ButtonType", Nothing)
GreyIcon = PropBag.ReadProperty("GreyIcon", Nothing)

Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
Set Picture = PropBag.ReadProperty("Picture", Nothing)
Set Background = PropBag.ReadProperty("Background", Nothing)
Set Picture = PropBag.ReadProperty("Picture", Nothing)
Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)

UserControl.Width = UserControl.Width + 10
UserControl.Width = UserControl.Width - 10

On Error GoTo 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
Call PropBag.WriteProperty("Caption", Label1.Caption, "Label1")
Call PropBag.WriteProperty("Backcolor", Picture1.BackColor, &H8000000F)
Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
Call PropBag.WriteProperty("Picture", Picture, Nothing)
Call PropBag.WriteProperty("Background", Background, Nothing)
Call PropBag.WriteProperty("BackColor", Picture1.BackColor, &H8000000F)
Call PropBag.WriteProperty("Picture", Picture, Nothing)
Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
Call PropBag.WriteProperty("Caption", Label1.Caption, "FlatButton")
Call PropBag.WriteProperty("FontBold", Label1.FontBold, 0)
Call PropBag.WriteProperty("FontItalic", Label1.FontItalic, 0)
Call PropBag.WriteProperty("FontName", Label1.FontName, "")
Call PropBag.WriteProperty("FontSize", Label1.FontSize, 0)
Call PropBag.WriteProperty("FontStrikethru", Label1.FontStrikethru, 0)
Call PropBag.WriteProperty("FontUnderline", Label1.FontUnderline, 0)
Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)

Call PropBag.WriteProperty("Hit", Hit, Nothing)
Call PropBag.WriteProperty("Default", Default, Nothing)
Call PropBag.WriteProperty("Cancel", Cancel, Nothing)
Call PropBag.WriteProperty("Sound", Sound, Nothing)
Call PropBag.WriteProperty("Cursor", Cursor, Nothing)
Call PropBag.WriteProperty("PicAlign", PicAlign, Nothing)
Call PropBag.WriteProperty("Style", Style, Nothing)
Call PropBag.WriteProperty("ButtonType", ButtonType, Nothing)
Call PropBag.WriteProperty("GreyIcon", GreyIcon, Nothing)

On Error GoTo 0
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
    BackColor = Picture1.BackColor
    BackColor = PictureOrg.BackColor
    BackColor = PictureDraw.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Picture1.BackColor() = New_BackColor
    PictureOrg.BackColor() = New_BackColor
    PictureDraw.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Gibt eine Grafik zurück, die in einem Steuerelement angezeigt werden soll, oder legt diese fest."
    Set Picture = Picture2.Picture
    Set Picture = PictureOrg.Picture
    Set Picture = PictureDraw.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Picture2.Picture = New_Picture
    Set PictureOrg.Picture = New_Picture
    Set PictureDraw.Picture = New_Picture
    PropertyChanged "Picture"
DoEvents
Call GreyPic
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Gibt ein Font-Objekt zurück."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get Caption() As String
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Gibt Schriftstile für Fettschrift zurück oder legt diese fest."
Attribute FontBold.VB_ProcData.VB_Invoke_Property = "Standart"
    FontBold = Label1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Label1.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Gibt Schriftstile für Kursivschrift zurück oder legt diese fest."
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = "Standart"
    FontItalic = Label1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Label1.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property


Public Property Get FontName() As String
Attribute FontName.VB_Description = "Gibt den Namen der Schriftart an, die in jeder Zeile für die gegebene Ebene verwendet wird."
Attribute FontName.VB_ProcData.VB_Invoke_Property = "Standart"
    FontName = Label1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Label1.FontName() = New_FontName
    PropertyChanged "FontName"
End Property


Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Gibt die Größe der Schriftart (in Punkten) an, die in jeder Zeile für die gegebene Ebene verwendet wird."
Attribute FontSize.VB_ProcData.VB_Invoke_Property = "Standart"
    FontSize = Label1.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    Label1.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property


Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Gibt Schriftstile für durchgestrichene Schrift zurück oder legt diese fest."
Attribute FontStrikethru.VB_ProcData.VB_Invoke_Property = "Standart"
    FontStrikethru = Label1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    Label1.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Gibt Schriftstile für unterstrichene Schrift zurück oder legt diese fest."
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = "Standart"
    FontUnderline = Label1.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    Label1.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gibt die Vordergrundfarbe zurück, die zum Anzeigen von Text und Grafiken in einem Objekt verwendet wird, oder legt diese fest."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Background() As Picture
    Set Background = Picture1.Picture
End Property

Public Property Set Background(ByVal New_Background As Picture)
    Set Picture1.Picture = New_Background
    PropertyChanged "Background"
End Property

Public Property Get Hit() As Boolean
    Hit = M_Hit
End Property

Public Property Let Hit(ByVal New_Hit As Boolean)
    M_Hit = New_Hit
    PropertyChanged "Hit"
End Property

Public Property Get Default() As Boolean
    Default = M_Default
End Property

Public Property Let Default(ByVal New_Default As Boolean)
    M_Default = New_Default
    PropertyChanged "Default"
End Property

Public Property Get Cancel() As Boolean
    Cancel = M_Cancel
End Property

Public Property Let Cancel(ByVal New_Cancel As Boolean)
    M_Cancel = New_Cancel
    PropertyChanged "Cancel"
End Property

Public Property Get Sound() As Boolean
    Sound = M_Sound
End Property

Public Property Let Sound(ByVal New_Sound As Boolean)
    M_Sound = New_Sound
    PropertyChanged "Sound"
End Property

Public Property Get GreyIcon() As Boolean
    GreyIcon = M_GreyIcon
End Property

Public Property Let GreyIcon(ByVal New_GreyIcon As Boolean)
    M_GreyIcon = New_GreyIcon
    PropertyChanged "GreyIcon"
End Property

Public Property Get Cursor() As Integer
    Cursor = M_Cursor
End Property

Public Property Let Cursor(ByVal New_Cursor As Integer)
    M_Cursor = New_Cursor
    PropertyChanged "Cursor"
End Property

Public Property Get PicAlign() As Integer
    PicAlign = M_PicAlign
End Property

Public Property Let PicAlign(ByVal New_PicAlign As Integer)
    M_PicAlign = New_PicAlign
    PropertyChanged "PicAlign"
End Property

Public Property Get Style() As Integer
    Style = M_Style
End Property

Public Property Let Style(ByVal New_Style As Integer)
    M_Style = New_Style
    PropertyChanged "Style"
End Property

Public Property Get ButtonType() As Integer
    ButtonType = M_ButtonType
End Property

Public Property Let ButtonType(ByVal New_ButtonType As Integer)
    M_ButtonType = New_ButtonType
    PropertyChanged "ButtonType"
End Property

