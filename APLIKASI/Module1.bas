Attribute VB_Name = "Module1"
Public konek As ADODB.Connection
Public RsAdmin As ADODB.Recordset

Public lokasidb As String

Public Sub konekdb()
Set konek = New ADODB.Connection
Set RsAdmin = New ADODB.Recordset

lokasidb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Login.mdb"
      konek.Open lokasidb
End Sub
