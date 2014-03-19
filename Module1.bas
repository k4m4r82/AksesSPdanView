Attribute VB_Name = "Module1"
Option Explicit

Public conn As ADODB.Connection

Public Function konekToServer(ByVal userName As String, ByVal userPass As String, _
                              ByVal serverName As String, ByVal dbPath As String) As Boolean
    Dim strCon As String
    
    On Error GoTo errHandle
    
    strCon = "DRIVER=Firebird/Interbase(r) Driver;UID=" & userName & ";PWD=" & userPass & ";" & _
             "DBNAME=" & serverName & ":" & dbPath & "\DBPOS.FDB"
    Set conn = New ADODB.Connection
    conn.ConnectionString = strCon
    conn.Open
    
    konekToServer = True
    
    Exit Function
errHandle:
    konekToServer = False
End Function

Public Sub closeRecordset(ByVal vRs As ADODB.Recordset)
    On Error Resume Next
    
    If Not (vRs Is Nothing) Then
        If vRs.State = adStateOpen Then vRs.Close
    End If
    
    Set vRs = Nothing
End Sub

Public Function openRecordset(ByVal query As String) As ADODB.Recordset
    Dim obj As ADODB.Recordset
    
    Set obj = New ADODB.Recordset
    obj.CursorLocation = adUseClient
    obj.Open query, conn, adOpenForwardOnly, adLockReadOnly
    Set openRecordset = obj
End Function

Public Function getRecordCount(ByVal vRs As ADODB.Recordset) As Long
    On Error Resume Next
    
    vRs.MoveLast
    getRecordCount = vRs.RecordCount
    vRs.MoveFirst
End Function
