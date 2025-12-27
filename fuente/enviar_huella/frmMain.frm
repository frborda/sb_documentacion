VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EnviaHuellas SQV"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock WinSock 
      Left            =   600
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7000
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdEnviarHuellas 
      Caption         =   "Enviar huellas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1260
      TabIndex        =   2
      Top             =   660
      Width           =   1935
   End
   Begin VB.TextBox txtBanca 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   60
      Width           =   3555
   End
   Begin VB.Label Label1 
      Caption         =   "Banca:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rsbanca As ADODB.Recordset
Dim sLegisla As TLegislador
Dim rta As String
Dim ipBanca As String

Private Sub cmdEnviarHuellas_Click()
Dim rta As Integer
rta = MsgBox("¿Está seguro de que desea enviar las huellas a la banca " & txtBanca.Text & "?", vbYesNo)
If rta = vbYes Then
    Dim rs As New ADODB.Recordset
    SetearRsW "SELECT Ip FROM BancasIp WHERE BancaNumero = " & txtBanca.Text, rs
    If Not rs.EOF Then
        ipBanca = rs.Fields(0)
        WinSock.RemoteHost = ipBanca
        WinSock.Connect rs.Fields(0), 7000
    Else
        MsgBox "La banca es inválida"
    End If
End If
End Sub
Private Sub Form_Load()
rta = ""
Call AbrirConexionSQLServer
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
            Resume
    End Select
Return

End Function

Public Sub AbrirConexionSQLServer()
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
        .ConnectionTimeout = 30
        .Open
    End With
    'Cargo los Recorset
    Set RsSQV = New ADODB.Recordset
    Set RsSB = New ADODB.Recordset
    Set rsbanca = New ADODB.Recordset
    'Leo la base viegente de Uso y cambio la conexion
    strcad = "SELECT valor FROM configuracion WHERE variable = 'base_vigente'"
    Call SetearRsW(strcad, rsbanca)
    strConexionSQL = rsbanca.Fields(0).Value
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
End Sub

Public Sub EnviarHuellasHCDN(ip As String)
    Dim strcad                      As String
    Dim naUX                        As Long
    Dim nIndice                     As Long
    Dim nEnvio                      As Integer
    Dim strLegisladorSecuencia      As String
    Dim strMantenimientoSecuencia   As String
    Dim fLSec                       As Boolean
    Dim strCadenaHuella             As String
    Dim nHuella                     As Integer
    Dim nCantHuellas                As Integer
    Dim xHuella                     As String
    Dim xUltimaHuella               As String
    Dim i As Integer
    Dim lFin As Boolean
    Dim strSQLWhere                 As String
    Dim nCantLegisladoresConHuella  As Integer
    Dim strVersionDatosSQV          As String
    Dim nEspera As Long
    Dim strPrefijoVersionBanca  As String
    Dim nProcesados As Integer
    Dim nEnviarMaximoHuellas As Integer
    Dim nTicks As Long
    'Funcion para Enviar Tabla de Legisladores.
    '**************************************************************
    'Busca en config el valor de version de datos sqv
    Set rsbanca = New ADODB.Recordset
    'Leo la base viegente de Uso y cambio la conexion
    strcad = "SELECT TOP 1 version_datos_sqv FROM config"
    Call SetearRsW(strcad, rsbanca)
    strVersionDatosSQV = rsbanca.Fields(0).Value
    strPrefijoVersionBanca = Replace(rsbanca.Fields(0).Value, "_", "")
    strPrefijoVersionBanca = Left(strPrefijoVersionBanca, 4) & Mid(strPrefijoVersionBanca, 9, 4)
    strcad = "SELECT     COUNT(*) " & _
" FROM         (SELECT     a.id" & _
"                       FROM          legisladores_sb a INNER JOIN" & _
"                                              legisladores_activos ON legisladores_activos.id = a.id" & _
"                       Where a.tipo = 1  AND (legisladores_activos.descripcion <> 'Activo sin incorporar')" & _
"                       Union" & _
"                       SELECT     b.id" & _
"                       FROM         legisladores_sb b" & _
"                       WHERE     b.tipo = 0) U INNER JOIN" & _
"                      legisladores_sb sb ON U.id = sb.id"
    Call SetearRsW(strcad, rsbanca)
    If Not rsbanca.EOF Then ' si se encuentra al menos un legislador
        nCantLegisladoresConHuella = rsbanca.Fields(0)
        If nEnviarMaximoHuellas > 0 And nCantLegisladoresConHuella > nEnviarMaximoHuellas Then nCantLegisladoresConHuella = nEnviarMaximoHuellas
    End If
    rsbanca.Close
    If True Then
        strcad = "SELECT     sb.*" & _
" FROM         (SELECT     a.id" & _
"                       FROM          legisladores_sb a INNER JOIN" & _
"                                              legisladores_activos ON legisladores_activos.id = a.id" & _
"                       Where a.tipo = 1 AND (legisladores_activos.descripcion <> 'Activo sin incorporar')" & _
"                       Union" & _
"                       SELECT     b.id" & _
"                       FROM         legisladores_sb b" & _
"                       WHERE     b.tipo = 0) U INNER JOIN" & _
"                      legisladores_sb sb ON U.id = sb.id" & _
" ORDER BY sb.tipo DESC,U.id"
        Call SetearRsW(strcad, rsbanca)
        fLSec = True
        nIndice = 0
        nEnvio = 0
        strLegisladorSecuencia = strString(4, Hex(Fix(nEnvio)), "0", "I")
        Version = strPrefijoVersionBanca & Format(nCantLegisladoresConHuella, "0000") 'cantidad de personas a enviar en decimal de 4 bytes
        Dim lTick As Long
        lTick = GetTickCount
        Enviar "SCANCL"
        While rta = ""
            DoEvents
        Wend
        If Not rtaOk() Then
            MsgBox "Proceso abortado por SCANCL!"
            WinSock.Close
            Me.cmdEnviarHuellas.Enabled = True
            Exit Sub
        End If
        rta = ""
        Enviar "SNUVER " & Version
        While rta = ""
            DoEvents
        Wend
        If Not rtaOk() Then
            MsgBox "Error SNUVER!"
            WinSock.Close
            Exit Sub
        End If
        rta = ""
        Enviar "SNUVER " & Version
        While rta = ""
            DoEvents
        Wend
        If Not rtaOk() Then
            MsgBox "Error SNUVER!"
            WinSock.Close
            Exit Sub
        End If
        rta = ""
        frmMain.prgBar.Value = 0
        frmMain.prgBar.Max = rsbanca.RecordCount
        Dim FDelay As Long
        FDelay = GetTickCount
        While Not rsbanca.EOF ' si se encuentra al menos un legislador
            lTick = GetTickCount
            nProcesados = 0
            If Not rsbanca.EOF Then
                    With sLegisla
                        If Not IsNull(rsbanca!template) Then
                         .sTemplate11 = rsbanca!template
                         .sId = rsbanca!id
                         .sNombre = NullCadena(rsbanca!nombre)
                         .sApellido = NullCadena(rsbanca!apellido)
                         .sDNI = NullCadena(rsbanca!dni)
                         .sClase = IIf(rsbanca!es_legislador = 1, "S", "V") ' Considerar el caso de mantenimiento
                         .sIcono = "0"
                         nCantHuellas = 1
                         If nCantHuellas > 0 Then
                            If Val(rsbanca!tipo) = 0 Then
                               If fLSec = True Then
                                    'Ultima Secuencia de Legislador
                                    If nEnvio = 0 Then
                                          'Si no hay Legisladores para enviar.
                                          strLegisladorSecuencia = strLegisladorSecuencia & "0001"
                                    Else
                                          strLegisladorSecuencia = strLegisladorSecuencia & strString(4, Hex(Fix(nEnvio)), "0", "I")
                                    End If
                                    'Primera Secuencia de Mantenimiento
                                    strMantenimientoSecuencia = strString(4, Hex(Fix(nEnvio)), "0", "I")
                                    fLSec = False
                               End If
                               '.sNombre = .sNombre & " " & Now
                            End If
                            
                            nHuella = 0
                            i = 0
                            lFin = False
                            'Envio de datos del legislador con la ultima huella
                            xUltimaHuella = BinAHex(.sTemplate11)
                            strCadenaHuella = "SRLEGI "
                            strCadenaHuella = strCadenaHuella & xUltimaHuella
                            Enviar strCadenaHuella
                            While rta = ""
                                DoEvents
                            Wend
                            If Not rtaOk() Then
                                If InStr(1, rta, "TVERRE") > 0 Then
                                    MsgBox "Huellas enviadas exitosamente!"
                                    frmMain.prgBar = 0
                                    txtBanca.Text = ""
                                    Exit Sub
                                End If
                                MsgBox "Proceso abortado por error SRLEGI!"
                                WinSock.Close
                                Me.cmdEnviarHuellas.Enabled = True
                                Exit Sub
                            End If
                            rta = ""
                            frmMain.prgBar.Value = frmMain.prgBar.Value + 1
                            DoEvents
                         End If
                      End If
                    End With
                'nIndice = 0
                strCadenaHuella = ""
            End If
            rsbanca.MoveNext
        Wend
    Else
        MsgBox ("No hay huellas para enviar en la base de datos")
    End If
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

Private Sub WinSock_Close()
Me.cmdEnviarHuellas.Enabled = True
Me.prgBar.Value = 0
End Sub

Private Sub WinSock_Connect()
Me.cmdEnviarHuellas.Enabled = False
Me.EnviarHuellasHCDN (ipBanca)
End Sub

Private Sub WinSock_DataArrival(ByVal bytesTotal As Long)
Dim dato As String
WinSock.GetData dato, , bytesTotal
If InStr(1, dato, "TESTAD") > 0 Then
    dato = ""
End If
rta = dato
End Sub

Private Sub Enviar(pDato As String)
WinSock.SendData "f" & pDato
End Sub

Private Function rtaOk() As Boolean
Dim rt As Boolean
rt = False
If InStr(1, rta, "TACKNL") > 0 Then
    rt = True
End If
rtaOk = rt
End Function

