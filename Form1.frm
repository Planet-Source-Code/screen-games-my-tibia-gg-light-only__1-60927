VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Tibia GG Light Hack"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3630
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   242
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Startat As Long

Private Sub Check1_Click()
End Sub

Private Sub Command1_Click()
If Startat = 0 Then
Startat = 1
Command1.Caption = "Stop"
Else
Startat = 0
Command1.Caption = "Start"
End If

If Startat = 1 Then Timer1.Enabled = True
If Startat = 0 Then Timer1.Enabled = False
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Call Window_Ontop(Me)
End Sub
Private Sub Timer1_Timer()
Call Hack_Light(80)
End Sub
