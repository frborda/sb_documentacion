VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "EnviadorPro"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   2955
   End
   Begin VB.TextBox txtBanca 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   60
      Width           =   2955
   End
   Begin MSWinsockLib.Winsock WinSock 
      Left            =   4080
      Top             =   1860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ipBanca As String
Dim version As String
Dim huellas() As String
Dim cantHuellas As Integer
Dim sLegisla As TLegislador
Dim huellaActual As Integer
Dim firstNuver As Boolean
Dim enviado As String

Private Sub cmdEnviar_Click()
prgBar.Value = 0
If Trim(txtBanca.Text) <> "" Then
    ObtieneDatosEnviar
End If
End Sub

Private Sub Form_Load()
enviado = ""
AbrirConexionSQLServer
Open "C:\logBanca.txt" For Output As #1
Print #1, Now()
Close #1
End Sub

Private Sub Procesar()
huellaActual = -1
prgBar.Value = 0
prgBar.Max = cantHuellas
cmdEnviar.Caption = "Conectando..."
firstNuver = False
WinSock.Connect ipBanca, 7000
End Sub

Private Sub WinSock_Connect()
Enviar "SCANCL"
End Sub

Private Sub WinSock_DataArrival(ByVal bytesTotal As Long)
Dim dato As String
WinSock.GetData dato, , bytesTotal
If InStr(1, dato, "TACKNL") > 0 Then
    If (InStr(1, enviado, "SCANCL") > 0) Then
        cmdEnviar.Caption = "Estado previo borrado."
        Enviar "SNUVER " & version
    ElseIf (InStr(1, enviado, "SNUVER")) Then
        If (firstNuver = False) Then
            firstNuver = True
            cmdEnviar.Caption = "Huellas borradas(1)."
            Enviar "SNUVER " & version
        Else
            cmdEnviar.Caption = "Huellas borradas(2). Enviando..."
            huellaActual = huellaActual + 1
            Enviar huellas(huellaActual)
        End If
    ElseIf (InStr(1, enviado, "SRLEGI")) Then
        cmdEnviar.Caption = "Enviadas " & huellaActual
        huellaActual = huellaActual + 1
        Enviar huellas(huellaActual)
    End If
    If (huellaActual > -1) Then
        prgBar.Value = huellaActual
    End If
ElseIf InStr(1, dato, "ERRE") > 0 Then
    MsgBox "Exito al enviar las huellas! Enviadas " & huellaActual
    txtBanca.Enabled = True
    cmdEnviar.Caption = "Enviar"
    cmdEnviar.Enabled = True
    prgBar.Value = 0
    WinSock.Close
ElseIf InStr(1, dato, "TNACK") > 0 Then
    If (InStr(1, enviado, "SRLEGI") > 0) Then
        cmdEnviar.Caption = "Reintentando " & huellaActual
        Enviar huellas(huellaActual)
    End If
Else
    MsgBox "Mensaje desconocido: " & dato & " con " & enviado
    End
End If
Call Log(dato)
End Sub

Private Sub Enviar(pDato As String)
WinSock.SendData "f" & pDato
enviado = pDato
End Sub

Private Sub ObtieneDatosEnviar()
Dim rs As New ADODB.Recordset
SetearRsW "SELECT Ip FROM BancasIp WHERE BancaNumero = " & txtBanca.Text, rs
If Not rs.EOF Then
    ipBanca = rs.Fields(0)
    cantHuellas = -1
    cmdEnviar.Enabled = False
    cmdEnviar.Caption = "Cargando huellas en memoria..."
    txtBanca.Enabled = False
    Call EnviarHuellasHCDN(ipBanca)
    'Ya tengo las huellas
    If (UBound(huellas) > 0) Then
        Call Procesar
    Else
        MsgBox "No se encontraron huellas"
        cmdEnviar.Enabled = True
        cmdEnviar.Caption = "Enviar"
        txtBanca.Enabled = True
    End If
Else
    MsgBox "La banca es inválida"
    txtBanca.Text = ""
    txtBanca.SetFocus
End If
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
    Dim rsBanca As New ADODB.Recordset
    'Funcion para Enviar Tabla de Legisladores.
    '**************************************************************
    'Busca en config el valor de version de datos sqv
    Set rsBanca = New ADODB.Recordset
    'Leo la base viegente de Uso y cambio la conexion
    strcad = "SELECT TOP 1 version_datos_sqv FROM config"
    Call SetearRsW(strcad, rsBanca)
    strVersionDatosSQV = rsBanca.Fields(0).Value
    strPrefijoVersionBanca = Replace(rsBanca.Fields(0).Value, "_", "")
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
    Call SetearRsW(strcad, rsBanca)
    If Not rsBanca.EOF Then ' si se encuentra al menos un legislador
        nCantLegisladoresConHuella = rsBanca.Fields(0)
        If nEnviarMaximoHuellas > 0 And nCantLegisladoresConHuella > nEnviarMaximoHuellas Then nCantLegisladoresConHuella = nEnviarMaximoHuellas
    End If
    rsBanca.Close
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
        Call SetearRsW(strcad, rsBanca)
        fLSec = True
        nIndice = 0
        nEnvio = 0
        strLegisladorSecuencia = strString(4, Hex(Fix(nEnvio)), "0", "I")
        version = strPrefijoVersionBanca & Format(nCantLegisladoresConHuella, "0000") 'cantidad de personas a enviar en decimal de 4 bytes
        Dim lTick As Long
        lTick = GetTickCount
        Dim FDelay As Long
        FDelay = GetTickCount
        While Not rsBanca.EOF ' si se encuentra al menos un legislador
            lTick = GetTickCount
            nProcesados = 0
            If Not rsBanca.EOF Then
                    With sLegisla
                        If Not IsNull(rsBanca!template) Then
                         .sTemplate11 = rsBanca!template
                         .sId = rsBanca!id
                         .sNombre = NullCadena(rsBanca!nombre)
                         .sApellido = NullCadena(rsBanca!apellido)
                         .sDNI = NullCadena(rsBanca!dni)
                         .sClase = IIf(rsBanca!es_legislador = 1, "S", "V") ' Considerar el caso de mantenimiento
                         .sIcono = "0"
                         nCantHuellas = 1
                         If nCantHuellas > 0 Then
                            If Val(rsBanca!tipo) = 0 Then
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
                            cantHuellas = cantHuellas + 1
                            ReDim Preserve huellas(0 To cantHuellas)
                            huellas(cantHuellas) = strCadenaHuella
                         End If
                      End If
                    End With
                'nIndice = 0
                strCadenaHuella = ""
            End If
            rsBanca.MoveNext
            DoEvents
        Wend
    Else
        MsgBox ("No hay huellas para enviar en la base de datos")
    End If
End Sub

Private Sub WinSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Error al conectar a " & ipBanca
txtBanca.Enabled = True
cmdEnviar.Caption = "Enviar"
cmdEnviar.Enabled = True
prgBar.Value = 0
WinSock.Close
End Sub
