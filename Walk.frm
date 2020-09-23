VERSION 5.00
Begin VB.Form WalkExample 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Walking Example"
   ClientHeight    =   5325
   ClientLeft      =   585
   ClientTop       =   510
   ClientWidth     =   8205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8205
   Begin VB.Timer WALK 
      Interval        =   20
      Left            =   4320
      Top             =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Esc to exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   270
      Left            =   3600
      TabIndex        =   1
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Copyright (C) 2003 Lam Ri Hui"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   300
      Left            =   2640
      TabIndex        =   0
      Top             =   4920
      Width           =   3120
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   360
      Top             =   480
      Width           =   1935
   End
   Begin VB.Image Char 
      Height          =   495
      Left            =   3600
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image Front2 
      Height          =   510
      Left            =   5280
      Picture         =   "Walk.frx":0000
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Right1 
      Height          =   510
      Left            =   4800
      Picture         =   "Walk.frx":0D8A
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Right2 
      Height          =   510
      Left            =   5280
      Picture         =   "Walk.frx":1B14
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Front1 
      Height          =   510
      Left            =   4800
      Picture         =   "Walk.frx":289E
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Left2 
      Height          =   510
      Left            =   5280
      Picture         =   "Walk.frx":3628
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Left1 
      Height          =   510
      Left            =   4800
      Picture         =   "Walk.frx":43B2
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Back2 
      Height          =   510
      Left            =   5280
      Picture         =   "Walk.frx":513C
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Back1 
      Height          =   510
      Left            =   4800
      Picture         =   "Walk.frx":5D2E
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "WalkExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DECLARE VARIABLES
Dim WalkL As Boolean 'Walk Left
Dim WalkR As Boolean 'Walk Right
Dim WalkU As Boolean 'Walk Up
Dim WalkD As Boolean 'Walk Down
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'UP ARROW PRESSED
If KeyCode = 38 Then WalkU = True

'DOWN ARROW PRESSED
If KeyCode = 40 Then WalkD = True

'LEFT ARROW PRESSED
If KeyCode = 37 Then WalkL = True

'RIGHT ARROW PRESSED
If KeyCode = 39 Then WalkR = True

'ESC PRESSED
If KeyCode = 27 Then End

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'UP ARROW RELEASED
If KeyCode = 38 Then WalkU = False

'DOWN ARROW RELEASED
If KeyCode = 40 Then WalkD = False

'LEFT ARROW RELEASED
If KeyCode = 37 Then WalkL = False

'RIGHT ARROW RELEASED
If KeyCode = 39 Then WalkR = False

End Sub

Private Sub Form_Load()
'LOAD PICTURE
Char.Picture = Front1.Picture
End Sub

Private Sub WALK_Timer()
'STOP WALKING INTO THE SIDE OF THE FORM
If Char.Left <= 0 Then
WalkL = False
Char.Left = 0
End If

If Char.Left >= 7680 Then
WalkR = False
Char.Left = 7680
End If

If Char.Top <= 0 Then
WalkU = False
Char.Top = 0
End If

'MOVE
If Char.Top >= 4800 Then
WalkD = False
Char.Top = 4800
End If

If WalkU = True Then
If Char.Picture = Back1.Picture Then
Char.Picture = Back2.Picture
Char.Top = Char.Top - 100
ElseIf Char.Picture = Back2.Picture Then
Char.Picture = Back1.Picture
Char.Top = Char.Top - 100
Else
Char.Picture = Back1.Picture
Char.Top = Char.Top - 100
End If
End If

If WalkD = True Then
If Char.Picture = Front1.Picture Then
Char.Picture = Front2.Picture
Char.Top = Char.Top + 100
ElseIf Char.Picture = Front2.Picture Then
Char.Picture = Front1.Picture
Char.Top = Char.Top + 100
Else
Char.Picture = Front1.Picture
Char.Top = Char.Top + 100
End If
End If

If WalkL = True Then
If Char.Picture = Right1.Picture Then
Char.Picture = Right2.Picture
Char.Left = Char.Left - 100
ElseIf Char.Picture = Right2.Picture Then
Char.Picture = Right1.Picture
Char.Left = Char.Left - 100
Else
Char.Picture = Right1.Picture
Char.Left = Char.Left - 100
End If
End If

If WalkR = True Then
If Char.Picture = Left1.Picture Then
Char.Picture = Left2.Picture
Char.Left = Char.Left + 100
ElseIf Char.Picture = Left2.Picture Then
Char.Picture = Left1.Picture
Char.Left = Char.Left + 100
Else
Char.Picture = Left1.Picture
Char.Left = Char.Left + 100
End If
End If
End Sub
