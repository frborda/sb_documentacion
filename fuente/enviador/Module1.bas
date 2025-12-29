Attribute VB_Name = "Module1"
Public cn As ADODB.Connection
Public Declare Function GetTickCount Lib "kernel32" () As Long
Type TLegislador
     sId                      As String
     sNombre                  As String
     sApellido                As String
     sDNI                     As String
     sClase                   As String
     sIcono                   As String
     sTemplate                As String
     sTemplate11()            As Byte
End Type
Public Sub Log(texto As String)
Open "C:\logBanca.txt" For Append As #1
Print #1, texto & Now()
Close #1
End Sub
Public Function SetearRsW(pCadena As String, ByRef pRst As ADODB.Recordset) As Boolean
    On Error GoTo TrapError
    
    SetearRsW = False
    
    With pRst
        If .State = adStateOpen Then
            .Close
        End If
        .Source = pCadena
        .ActiveConnection = cn
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open
        DoEvents
        If Not .BOF And Not .EOF Then
            SetearRsW = True
        End If
    End With
Exit Function
TrapError:
    Select Case Err.Number
        Case Else
            MsgBox "Error N° " & Err.Number & Chr(10) & Err.Description & "Originado en " & Err.Source
            End
    End Select
Return

End Function

Public Sub AbrirConexionSQLServer()
On Error GoTo errSql

    Dim rsBanca As New ADODB.Recordset
    Dim strcad     As String
    Set cn = New ADODB.Connection
    'Cadena de Conexion de la base sqv_config
    If True Then 'vmGen
        strConexionSQL = "PROVIDER=SQLOLEDB.1;PASSWORD=hcdn11;PERSIST SECURITY INFO=TRUE;USER ID=SQV;INITIAL CATALOG=SQV_Config;DATA SOURCE=10.1.1.5"
    Else 'SBA
        strConexionSQL = "Provider=SQLOLEDB.1;Password=unipaas;Persist Security Info=True;" _
                              & "User ID=sqv;Initial Catalog=sqv_config;Data Source=siprevo"
    End If
    With cn
        .ConnectionString = strConexionSQL
        .CursorLocation = adUseServer
        .ConnectionTimeout = 1
        .Open
    End With
    'Cargo los Recorset
    Set RsSQV = New ADODB.Recordset
    Set RsSB = New ADODB.Recordset
    Set rsBanca = New ADODB.Recordset
    'Leo la base viegente de Uso y cambio la conexion
    strcad = "SELECT valor FROM configuracion WHERE variable = 'base_vigente'"
    Call SetearRsW(strcad, rsBanca)
    strConexionSQL = rsBanca.Fields(0).Value
    If InStr(strConexionSQL, "prueba") > 0 Then
        MsgBox "SQV esta en modo prueba"
    End If
    With cn
        .Close
        .ConnectionString = strConexionSQL
        .CursorLocation = adUseServer
        .ConnectionTimeout = 30
        .Open
    End With
    
Exit Sub
errSql:
    MsgBox "Error al conectar a la base de datos"
    End
End Sub
Public Function strString(xTam As Long, strValor As String, strRelleno As String, Optional strTipo As String = "D") As String
    If Len(Trim(strValor)) < xTam Then
        If strTipo = "I" Then
            strString = String(xTam - Len(Trim(strValor)), strRelleno) & strValor
        Else
            strString = strValor & String(xTam - Len(Trim(strValor)), strRelleno)
        End If
    Else
        strString = Left(strValor, xTam)
    End If

End Function

Public Function NullCadena(Optional strCadena As Variant) As String
    NullCadena = strCadena & ""
End Function

Public Function BinAHex(dataBin() As Byte) As String
Dim i As Long
    BinAHex = ""
    For i = LBound(dataBin) To (UBound(dataBin))
        BinAHex = BinAHex & CerosIzquierda(Hex(dataBin(i)), 2)
    Next
End Function

Public Function CerosIzquierda(strText As String, nLong As Long) As String
    If nLong > Len(strText) Then
        CerosIzquierda = Left(String(nLong - Len(strText), "0") & strText, nLong)
    Else
        CerosIzquierda = Right(strText, nLong)
    End If
End Function
