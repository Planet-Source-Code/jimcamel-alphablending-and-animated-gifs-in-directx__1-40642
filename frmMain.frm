VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alphablending in DirectX!"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   159
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   20
      Left            =   0
      Max             =   100
      TabIndex        =   2
      Top             =   2130
      Value           =   50
      Width           =   3015
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Use Alphablending"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1875
      Width           =   1695
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   0
      ScaleHeight     =   123
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   0
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Just a little api call so I can open my website
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Check1_Click()
alpha = (Check1.Value = 1)
End Sub

Private Sub Form_Activate()
'Call the main loop
MainLoop
End Sub

Private Sub Form_Load()
'Call the Initilisation of DX
Init
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'The game is over
GameOver = True
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If you click on the button
If Button = 1 And Y < 25 Then
    'Open my website
    Call ShellExecute(0, "open", "www.aeonlegend.com", "", App.Path, 1)
End If
End Sub
