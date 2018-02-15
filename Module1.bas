Attribute VB_Name = "Module1"
Public Conn As New ADODB.Connection
Public N As New ADODB.Recordset
Public M As ADODB.Recordset
Public Sub db()
Set Conn = New ADODB.Connection
Set M = New ADODB.Recordset
Set N = New ADODB.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\nilai.mdb"
End Sub


