VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim conn As ADODB.Connection
Dim rst As ADODB.Recordset

Sub Run_Report()
Dim Server_Name As String
Dim DatabaseName As String
Dim SQL As String

Server_Name = "SERVER-NAME"
DatabaseName = "BikeStores"
SQL = "select * from production.products"

Call Connect_To_SQLServer(Server_Name, DatabaseName, SQL)

End Sub
Sub Connect_To_SQLServer(ByVal Server_Name As String, ByVal Database_Name As String, ByVal SQL_Statement As String)
Dim strConn As String
Dim wsReport As Worksheet
Dim col As Integer

strConn = "Provider=SQLOLEDB;"
strConn = strConn & "Server=" & Server_Name & ";"
strConn = strConn & "Database=" & Database_Name & ";"
strConn = strConn & "Trusted_Connection=yes;"
'strConn = strConn & "uid=dummyuser;pwd=<password>"

Set conn = New ADODB.Connection
With conn
    .Open ConnectionString:=strConn
    .CursorLocation = adUseClient
End With

Set rst = New ADODB.Recordset
With rst
    .ActiveConnection = conn
    .Open Source:=SQL_Statement
End With

Set wsReport = ThisWorkbook.Worksheets.Add
With wsReport
    For col = 0 To rst.Fields.Count - 1
        .Cells(1, col + 1).Value = rst.Fields(col).Name
    Next col
    
    .Range("A2").CopyFromRecordset Data:=rst
End With


Set wsReport = Nothing

Call Close_Objects

End Sub

Private Sub Close_Objects()

If rst.State <> 0 Then rst.Close
If conn.State <> 0 Then conn.Close

'//Release Memomory
Set rst = Nothing
Set conn = Nothing

End Sub

