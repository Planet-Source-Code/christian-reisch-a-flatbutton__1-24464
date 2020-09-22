VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   0  'Kein
   Caption         =   "Flatbutton"
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   30
      ScaleHeight     =   165
      ScaleWidth      =   4005
      TabIndex        =   1
      Top             =   30
      Width           =   4035
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "FlatButton"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   30
         TabIndex        =   2
         Top             =   -30
         Width           =   1275
      End
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   1140
      Y2              =   30
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   4080
      X2              =   4080
      Y1              =   1140
      Y2              =   30
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   0
      X2              =   4080
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   30
      X2              =   4080
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "About.frx":08CA
      Top             =   270
      Width           =   480
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long

Private Type RECT
       Left As Long
       Top As Long
       Right As Long
       Bottom As Long
End Type

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Const COLOR_BTNFACE = 15

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
       
Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_CHARSTREAM = 4
Private Const DT_DISPFILE = 6
Private Const DT_EXPANDTABS = &H40
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_INTERNAL = &H1000
Private Const DT_LEFT = &H0
Private Const DT_METAFILE = 5
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_PLOTTER = 0
Private Const DT_RASCAMERA = 3
Private Const DT_RASDISPLAY = 1
Private Const DT_RASPRINTER = 2
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10

Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
       
Private Const CLR_INVALID = -1

Private Sub TextEffect(obj As Object, ByVal sText As String, ByVal lX As Long, ByVal lY As Long, Optional ByVal bLoop As Boolean = False, Optional ByVal lStartSpacing As Long = 128, Optional ByVal lEndSpacing As Long = -1, Optional ByVal oColor As OLE_COLOR = vbWindowText)

Dim lhDC As Long
Dim i As Long
Dim x As Long
Dim lLen As Long
Dim hBrush As Long
Static tR As RECT
Dim iDir As Long
Dim bNotFirstTime As Boolean
Dim lTime As Long
Dim lIter As Long
Dim bSlowDown As Boolean
Dim lCOlor As Long
Dim bDoIt As Boolean
       
lhDC = obj.hdc
iDir = -1
i = lStartSpacing
tR.Left = lX: tR.Top = lY: tR.Right = lX: tR.Bottom = lY
       
OleTranslateColor oColor, 0, lCOlor
       
hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))
lLen = Len(sText)

SetTextColor lhDC, lCOlor
bDoIt = True

Do While bDoIt
    lTime = timeGetTime
    
    If (i < -3) And Not (bLoop) And Not (bSlowDown) Then
        bSlowDown = True
        iDir = 1
        lIter = (i + 4)
    End If
                      
    If (i > 128) Then iDir = -1
    
    If Not (bLoop) And iDir = 1 Then
            
        If (i = lEndSpacing) Then
        
            bDoIt = False
        
        Else
                
            lIter = lIter - 1
                            
            If (lIter <= 0) Then
                    
                i = i + iDir
                lIter = (i + 4)
            
            End If

        End If
                      
        Else
                
            i = i + iDir
        
        End If
                      
        FillRect lhDC, tR, hBrush
        x = 32 - (i * lLen)
        SetTextCharacterExtra lhDC, i
        DrawText lhDC, sText, lLen, tR, DT_CALCRECT
        tR.Right = tR.Right + 4
            
        If (tR.Right > obj.ScaleWidth \ Screen.TwipsPerPixelX) Then tR.Right = obj.ScaleWidth \ Screen.TwipsPerPixelX
                      
        DrawText lhDC, sText, lLen, tR, DT_LEFT
                      
        Do
            DoEvents
            If obj.Visible = False Then Exit Sub
                      
        Loop While (timeGetTime - lTime) < 20
Line3.Refresh
    Loop

    DeleteObject hBrush

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
Call TextEffect(Me, "FlatButton", 50, 17, , 200, 2, RGB(0, 0, 255))
Call TextEffect(Me, "Demo Version 1.0", 50, 32, , 100, 2, RGB(0, 0, 255))
Me.FontSize = 7
Call TextEffect(Me, "Copyright by Christian Reisch", 5, 60, , 35, 2, RGB(0, 0, 0))
Command1.Visible = True
On Error GoTo 0
End Sub

Private Sub Form_Load()
On Error Resume Next
Command1.Visible = False
On Error GoTo 0
End Sub
