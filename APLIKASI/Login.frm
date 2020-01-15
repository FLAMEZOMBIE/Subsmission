VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H8000000D&
   Caption         =   "Login"
   ClientHeight    =   3120
   ClientLeft      =   345
   ClientTop       =   540
   ClientWidth     =   5550
   LinkTopic       =   "Form2"
   ScaleHeight     =   3120
   ScaleWidth      =   5550
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1560
      Top             =   1920
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\APLIKASI\Login.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\APLIKASI\Login.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Cmdlogin 
      Caption         =   "&LOGIN"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox TxtPassword 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Txtusername 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim query As String
Private Sub Cmdlogin_Click()
Call konekdb
If Txtusername.Text = "" Then
MsgBox "Username Masih Kosong", vbCritical, "Warning"
Txtusername.SetFocus
ElseIf TxtPassword.Text = "" Then
MsgBox "Password Masih Kosong", vbCritical, "Warning"
TxtPassword.SetFocus
Else

querry = "select *From login where username= '" & Txtusername.Text & "' and password='" & TxtPassword.Text & "'"
RsAdmin.Open (querry), konek
    If RsAdmin.EOF Then
    MsgBox "Username Salah", vbExclamation, "Warning"
    Txtusername.Text = ""
    TxtPassword.Text = ""
    Txtusername.SetFocus
    Else
    Unload Me
    Form3.Show
    MsgBox "Login Success", vbInformation, "fyi"
    End If
End If
End Sub

