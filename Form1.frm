VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAksesSP2 
      Caption         =   "Akses SP 2"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAksesView 
      Caption         =   "Akses View"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAksesSP1 
      Caption         =   "Akses SP 1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Option Explicit
    
Dim cmd     As ADODB.Command
Dim param   As ADODB.Parameter

Dim i       As Long
Dim strSql  As String

Private Function addSupplier(ByVal nama As String, ByVal alamat As String, ByVal telepon As String) As Boolean
    On Error GoTo errHandle
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
                    
    Set param = cmd.CreateParameter("nama", adVarChar, adParamInput, 30, nama)
    cmd.Parameters.Append param
    
    Set param = cmd.CreateParameter("alamat", adVarChar, adParamInput, 50, alamat)
    cmd.Parameters.Append param
    
    Set param = cmd.CreateParameter("telepon", adVarChar, adParamInput, 20, telepon)
    cmd.Parameters.Append param
    
    cmd.CommandText = "proc_add_supplier"
    cmd.CommandType = adCmdStoredProc
    cmd.Execute
            
    addSupplier = True
    
    For i = 0 To cmd.Parameters.Count - 1
        cmd.Parameters.Delete (0)
    Next
    
    Set cmd = Nothing
    Set param = Nothing
    
    Exit Function
errHandle:
    addSupplier = False
End Function

Private Function getStokBarang(ByVal kodeBarang As Long) As Long
    On Error GoTo errHandle
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    
    Set param = cmd.CreateParameter("kode_barang", adVarChar, adParamInput, 20, kodeBarang)
    cmd.Parameters.Append param
    
    Set param = cmd.CreateParameter("stok", adNumeric, adParamOutput)  'ini tambahan parameter yg harus didaftarkan untuk menampung return value
    param.Precision = 5
    param.NumericScale = 2
    cmd.Parameters.Append param

    cmd.CommandText = "func_get_stok"
    cmd.CommandType = adCmdStoredProc
    cmd.Execute
    
    getStokBarang = cmd.Parameters("stok").Value 'membaca nilai kembalian fungsi (return value)
    
    For i = 0 To cmd.Parameters.Count - 1
        cmd.Parameters.Delete (0)
    Next
    
    Set cmd = Nothing
    Set param = Nothing
    
    Exit Function
errHandle:
    getStokBarang = 1
End Function

Private Sub cmdAksesSP1_Click()
    'CARA 1
    'strSql = "EXECUTE PROCEDURE proc_add_supplier('KoKom Armagedon', 'Yogykarta', '0813 8176 xxxx')"
    'conn.Execute strSql 'conn -> variabel dengan tipe ADODB.Connection
    
    'CARA 2
    Dim result  As Boolean

    result = addSupplier("KoKom Armagedon", "Yogykarta", "0813 8176 xxxx")
End Sub

Private Sub cmdAksesSP2_Click()
    'store procedure yang dijadikan sebagai fungsi
    Debug.Print "jumlah stok : " & getStokBarang("12345")
End Sub

Private Sub cmdAksesView_Click()
    Dim rs      As ADODB.Recordset
    
    strSql = "SELECT * FROM v_info_pembelian" 'v_info_pembelian -> nama view
    Set rs = openRecordset(strSql)
    If Not rs.EOF Then
        For i = 1 To getRecordCount(rs)
            Debug.Print "Supplier : " & rs("nama").Value & vbCrLf & _
                        "Alamat : " & rs("alamat").Value & vbCrLf & _
                        "Nota : " & rs("nota").Value & vbCrLf & _
                        "Tanggal " & Format(rs("tanggal").Value, "dd/MM/yyyy")
            rs.MoveNext
        Next i
    End If
    Call closeRecordset(rs)
End Sub

Private Sub Form_Load()
    Dim ret As Boolean
    
    ret = konekToServer("SYSDBA", "masterkey", "127.0.0.1", App.Path)
End Sub
