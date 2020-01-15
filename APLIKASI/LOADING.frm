VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000D&
   Caption         =   "Form3"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15720
   LinkTopic       =   "Form3"
   ScaleHeight     =   3030
   ScaleWidth      =   15720
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7680
      Top             =   720
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 5
Me.Label1.Caption = Me.ProgressBar1.Value & "%"
If Me.ProgressBar1.Value = Me.ProgressBar1.Max Then
Unload Me
Form1.Show
End If
End Sub
