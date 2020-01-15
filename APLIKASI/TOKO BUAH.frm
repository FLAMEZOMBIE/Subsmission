VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "TOKO"
   ClientHeight    =   7590
   ClientLeft      =   4140
   ClientTop       =   3135
   ClientWidth     =   14370
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   14370
   Begin VB.CommandButton Command6 
      Caption         =   "AddChart"
      Height          =   375
      Left            =   10080
      TabIndex        =   28
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   8880
      TabIndex        =   27
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   7680
      TabIndex        =   26
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox TxtCari 
      Height          =   495
      Left            =   12480
      TabIndex        =   25
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CARI"
      Height          =   495
      Left            =   12480
      TabIndex        =   24
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6480
      TabIndex        =   21
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   495
      Left            =   5280
      TabIndex        =   20
      Top             =   2160
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "TOKO BUAH.frx":0000
      Height          =   2055
      Left            =   5280
      TabIndex        =   19
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3625
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   10080
      Top             =   2160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\APLIKASI\DatabaseBuah.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\APLIKASI\DatabaseBuah.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "BUAH"
      Caption         =   "NEXT/PERV"
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
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   0
      TabIndex        =   17
      Top             =   6840
      Width           =   5175
      Begin VB.TextBox TxtKembalian 
         Height          =   285
         Left            =   1680
         TabIndex        =   30
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Kembalian"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton CmdHitung 
      Caption         =   "&Hitung"
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   6360
      Width           =   5175
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   0
      TabIndex        =   13
      Top             =   5160
      Width           =   5175
      Begin VB.TextBox TxtBayar 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   450
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Bayar"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   0
      TabIndex        =   8
      Top             =   4080
      Width           =   5175
      Begin VB.TextBox TxtTotalBayar 
         Height          =   285
         Left            =   1680
         TabIndex        =   29
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Total Bayar"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   3000
      Width           =   5175
      Begin VB.CommandButton CmdBersih 
         Caption         =   "&ClearAll"
         Height          =   615
         Left            =   2880
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton CmdJumlah 
         Caption         =   "&Jumlah"
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.TextBox TxtKode 
         Height          =   495
         Left            =   1320
         TabIndex        =   22
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox TxtNama 
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox TxtJumlah 
         Height          =   495
         Left            =   1320
         TabIndex        =   2
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox TxtHarga 
         Height          =   495
         Left            =   1320
         TabIndex        =   1
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "KODE"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Nama Barang"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Jumlah Barang"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fruity Store"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Harga Barang"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBersih_Click()
TxtHarga.Text = ""
    TxtJumlah.Text = ""
    TxtNama.Text = ""
    TxtBayar.Text = ""
    TxtKode.Text = ""
    TxtKembali.Text = "0"
    TxtTotalBayar.Text = "0"
    TxtHarga.SetFocus
End Sub

Private Sub CmdHitung_Click()
Dim Bayar, Kembalian
    Bayar = TxtBayar.Text - TxtTotalBayar.Text
    TxtKembalian.Text = Bayar
End Sub

Private Sub CmdJumlah_Click()
Dim Bayar, Kembali
    Bayar = TxtHarga.Text * TxtJumlah.Text
    TxtTotalBayar.Text = Bayar
End Sub
Sub kosong()
TxtNama.Text = ""
TxtHarga.Text = ""
TxtKode.Text = ""
TxtNama.SetFocus
End Sub

Private Sub Command1_Click()
Call kosong
Command1.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset!NAMA_BUAH = TxtNama.Text
Adodc1.Recordset!HARGA = TxtHarga.Text
Adodc1.Recordset!KODE = TxtKode.Text
Adodc1.Recordset.Update
MsgBox "Save Complete", vbInformation, "fyi"
DataGrid1.Refresh
Call kosong
End Sub

Private Sub Command3_Click()
    CARI = "KODE='" & TxtCari.Text & "'"
    Adodc1.Recordset.Find CARI
    If Adodc1.Recordset.EOF Then
        MsgBox "DATA TIDAK ADA", vbCritical, "WARNING"
        Else
        TxtHarga.Text = Adodc1.Recordset!HARGA
        TxtKode.Text = Adodc1.Recordset!KODE
        TxtNama.Text = Adodc1.Recordset!NAMA_BUAH
    End If
End Sub

Private Sub Command4_Click()
Adodc1.Recordset!NAMA_BUAH = TxtNama.Text
Adodc1.Recordset!HARGA = TxtHarga.Text
Adodc1.Recordset!KODE = TxtKode.Text
Adodc1.Recordset.Update
MsgBox "Edit Complete", vbInformation, "fyi"
DataGrid1.Refresh
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Delete
MsgBox "Delete Complete", vbInformation, "fyi"
Call kosong
End Sub

Private Sub Command6_Click()
AddChart = " NEXT/PERV ' " & DataGrid1.Align & " ' "
        TxtHarga.Text = Adodc1.Recordset!HARGA
        TxtKode.Text = Adodc1.Recordset!KODE
        TxtNama.Text = Adodc1.Recordset!NAMA_BUAH
End Sub

Private Sub lblTotal_Click()
.Text = "Rp." & Val(lblTotal.Text)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TxtBayar_Change()

End Sub
