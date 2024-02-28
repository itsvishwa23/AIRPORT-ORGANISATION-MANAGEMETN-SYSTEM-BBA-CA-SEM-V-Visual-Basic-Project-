Attribute VB_Name = "Module1"
Public conn As ADODB.Connection
Public rsGlb As ADODB.Recordset
Public Function NextId(strTableName As String, intFieldIndex As Integer) As Integer
Dim intNextId  As Integer
Dim r As ADODB.Recordset
Set r = New ADODB.Recordset
r.Open "Select * from " + strTableName, conn, adOpenForwardOnly, adLockOptimistic
If Not r.EOF Then
    r.MoveLast
    intNextId = r.Fields(intFieldIndex)
Else
    intNextId = 0
End If
intNextId = intNextId + 1
Set r = Nothing
NextId = intNextId
End Function

Public Sub Connect()

    Set conn = New ADODB.Connection
    Dim StrSql As String
    StrSql = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\aoms.mdb;Persist Security Info=False"
    conn.Open StrSql
End Sub


Public Sub DisConnect()
    conn.Close
End Sub
Public Sub QuerySelect(s1 As String)
    Set rsGlb = New ADODB.Recordset
   'MsgBox s1
    With rsGlb
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .SOURCE = s1
        .CursorLocation = adUseClient
        .Open
    End With
End Sub
Public Function NextIdNew(strTableName As String, intFieldIndex As Integer) As Integer
Dim intNextId  As Integer
Dim r As ADODB.Recordset
Set r = New ADODB.Recordset
r.Open "Select max(cno) from " + strTableName, conn, adOpenForwardOnly, adLockOptimistic
If Not r.EOF Then
    intNextId = r.Fields(0)
Else
    intNextId = 0
End If
intNextId = intNextId + 1
Set r = Nothing
NextIdNew = intNextId
End Function
