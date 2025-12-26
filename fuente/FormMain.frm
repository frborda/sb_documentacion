VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servidor de Bancas  x MSA"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6765
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   1380
   End
   Begin MSComctlLib.ProgressBar prgBAR 
      Height          =   375
      Left            =   30
      TabIndex        =   35
      Top             =   6150
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Timer tmIps 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   13530
      Top             =   9780
   End
   Begin VB.CommandButton cmdUpdateIPBancas 
      Caption         =   "Actualización de IPs"
      Height          =   465
      Left            =   11130
      TabIndex        =   34
      Top             =   13800
      Width           =   2205
   End
   Begin VB.CommandButton cmdSimulaVotos 
      Caption         =   "Simular Votos"
      Enabled         =   0   'False
      Height          =   615
      Left            =   11190
      TabIndex        =   33
      Top             =   11010
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Habilitar Timer"
      Enabled         =   0   'False
      Height          =   435
      Left            =   11430
      TabIndex        =   32
      Top             =   12930
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10440
      Top             =   5100
   End
   Begin VB.CommandButton cmdFOrzar 
      Caption         =   "Mensajes Intensivos (Sentarse)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7500
      TabIndex        =   31
      Top             =   660
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reconectar 14"
      Enabled         =   0   'False
      Height          =   555
      Left            =   11520
      TabIndex        =   30
      Top             =   12030
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.TextBox txtLogEnvio 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   7530
      MultiLine       =   -1  'True
      TabIndex        =   20
      Text            =   "FormMain.frx":0000
      Top             =   4440
      Width           =   5415
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   7050
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "FormMain.frx":0111
      Top             =   4500
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   6960
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "FormMain.frx":0126
      Top             =   2220
      Width           =   435
   End
   Begin VB.TextBox txtEnviando 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   7440
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "FormMain.frx":013B
      Top             =   2160
      Width           =   5415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "STATE"
      Height          =   375
      Left            =   8670
      TabIndex        =   14
      Top             =   6600
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ENVIAR"
      Height          =   495
      Left            =   8070
      TabIndex        =   13
      Top             =   6000
      Width           =   4335
   End
   Begin VB.TextBox txtTemp 
      Height          =   915
      Left            =   7110
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "FormMain.frx":024C
      Top             =   7080
      Width           =   6135
   End
   Begin VB.TextBox TRsCola 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "[Botones]"
      Height          =   1215
      Left            =   2160
      TabIndex        =   6
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   855
         Left            =   3000
         Picture         =   "FormMain.frx":0252
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdHuellas 
         Caption         =   "&Enviar Huellas a las Terminales"
         Height          =   855
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdLimpiarLog 
         Caption         =   "&Limpiar log de Errores"
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtLog 
      Height          =   4335
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1800
      Width           =   6735
   End
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   4320
      Top             =   -120
   End
   Begin MSWinsockLib.Winsock WSocket 
      Index           =   0
      Left            =   3840
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DrawMode        =   1  'Blackness
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      Picture         =   "FormMain.frx":0784
      ScaleHeight     =   148.536
      ScaleMode       =   0  'User
      ScaleWidth      =   182.667
      TabIndex        =   10
      Top             =   120
      Width           =   2085
      Begin VB.Label lblSQV 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SQV 4.1"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Label lblBasePrueba 
      Caption         =   "BASE DE PRUEBA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   30
      TabIndex        =   36
      Top             =   1260
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Label lblContador 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   11250
      TabIndex        =   29
      Top             =   9630
      Width           =   1515
   End
   Begin VB.Label Label8 
      Caption         =   "MI"
      Height          =   255
      Left            =   12990
      TabIndex        =   28
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label ERWRWE 
      Caption         =   "FE"
      Height          =   255
      Left            =   12990
      TabIndex        =   27
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "RC"
      Height          =   255
      Left            =   12990
      TabIndex        =   26
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label lblRegistrosEnCola 
      Caption         =   "0"
      Height          =   495
      Left            =   12990
      TabIndex        =   25
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label lblFaltanEnviar 
      Caption         =   "0"
      Height          =   495
      Left            =   12990
      TabIndex        =   24
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lblMensajesInsertados 
      Caption         =   "Label7"
      Height          =   615
      Left            =   12990
      TabIndex        =   23
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "12345678901234567890123456789012345678901234567890"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7590
      TabIndex        =   22
      Top             =   4200
      Width           =   5355
   End
   Begin VB.Label Label5 
      Caption         =   "         1         2         3         4         5         6"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7590
      TabIndex        =   21
      Top             =   3960
      Width           =   5355
   End
   Begin VB.Label Label4 
      Caption         =   "         1         2         3         4         5         6"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7500
      TabIndex        =   18
      Top             =   1680
      Width           =   5355
   End
   Begin VB.Label Label2 
      Caption         =   "12345678901234567890123456789012345678901234567890"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   16
      Top             =   1980
      Width           =   5355
   End
   Begin VB.Label Label3 
      Caption         =   "<- Estados de las Bancas"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblVersion 
      Caption         =   "Monitor de actividad"
      Height          =   255
      Left            =   30
      TabIndex        =   4
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "CICLOS"
      Height          =   255
      Left            =   5190
      TabIndex        =   1
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label LabelTimer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5910
      TabIndex        =   0
      Top             =   1560
      Width           =   855
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'      Programa    : Servidor de Bancas
'
'
'
'      Principios a Tener en Cuenta:
'      A) Del SQV se Recibe Número de Banca que luego es Transformado en Socket para
'         el Tratamiento en el programa con el Type B2Skt y luego se devuelve con Skt2B
'
'      B) Nota los mensajes de que necesitan Velocidad de procesamiento no pasaran por la cola
'         Ya que la banca sera capas de separar varios mensajes y comenzaran con una f
'
'********************************************************************************************
Public Sub iniVariables()
     CountTimer = 0
     FlagTimerLeoCola = True
     CountTimerAux = 0
End Sub
Public Sub iniVariblesDim(ftam As Long, fUltimaBanca As Long)
    Dim NumRegistros   As Long
    Dim naUX           As Long
    NumRegistros = ftam
            
    ReDim Banca_ip(0 To NumRegistros - 1)
    
     
    ' Redimenciono ultimo estado de bancas
    '*************************************
    ReDim UltimoEstadodeBanca(0 To NumRegistros)
    ReDim UltimoEstadodePresencia(0 To NumRegistros)
    ReDim BanderaReset(0 To NumRegistros)
    ReDim FueConfigurada(0 To NumRegistros)
    For naUX = 0 To NumRegistros
         UltimoEstadodeBanca(naUX) = "off"
         UltimoEstadodePresencia(naUX) = "off"
         BanderaReset(naUX) = False
         FueConfigurada(naUX) = False
    Next
    
    ' Redimenciono y seteo en False los convertidores de Socket a Banca y viceversa.!
    '********************************************************************************
    ReDim Skt2B(0 To NumRegistros)
    For naUX = 0 To NumRegistros
        Skt2B(naUX).Estado = False
    Next
    
    ReDim B2Skt(0 To fUltimaBanca)
    For naUX = 0 To UltimaBanca
        B2Skt(naUX).Estado = False
    Next
    
    'Redimencion el Vector ultima secuencia para las bancas
    '***************************************************************
    ReDim UltimaSecuenciaBanca(0 To ConexionesAbiertas)
    For naUX = 0 To ConexionesAbiertas
        UltimaSecuenciaBanca(naUX) = 125
    Next
End Sub

Private Sub cmdFOrzar_Click()
Dim i As Integer
Dim x As Integer
For x = 1 To 1
    For i = 1 To 256
        Dim msgSQV As MensajeSQV
         With msgSQV
            .sTipo = "mevt"
            .sObjeto = Trim(Str(i))
            .sComponente = "term.seat"
            .sAtributo = "switch"
            .sValor = "open"
        End With
        Call InsertarMsgSQV(msgSQV)
         With msgSQV
            .sTipo = "mevt"
            .sObjeto = Trim(Str(i))
            .sComponente = "term.seat"
            .sAtributo = "switch"
            .sValor = "closed"
        End With
        Call InsertarMsgSQV(msgSQV)
    Next i
Next x
End Sub

Private Sub CmdHuellas_Click()
    Dim xRespuesta As Integer
    Dim nBanca     As String
    xRespuesta = MsgBox("Confirma la transmision de huellas ?", vbQuestion + vbYesNo)
    If xRespuesta = vbYes Then
        nBanca = InputBox("Ingrese el número de Banca para mandar las huellas", "SB")
        Call SincronizarBancas(nBanca)
    End If
End Sub
Public Sub SincronizarBancas(nBanca As String)
    Dim nVectorBancas As String
    Dim i As Integer
    Dim RsB As New ADODB.Recordset
    Dim cConsulta As String
    Dim rsTVersion As ADODB.Recordset
    Dim VersionSQV As String
    Dim PorHacer(0 To 256) As String
    Dim VectorCompleto As String
    Dim nTick As Long
    Dim nCantidadFallaron As Integer
    Dim rsTemp As ADODB.Recordset
    Dim BancasConError As String
        If Not Len(nBanca) = 0 Then
            If Not IsNumeric(nBanca) Then
                If LCase(nBanca) <> "brc" Then
                    If InStr(1, nBanca, ";") > 0 Then
                        'While RsCola.RecordCount > 10
                        '    DoEvents
                        'Wend
                        'Call EnviarHuellasHCDN(nBanca)
                    End If
                Else 'brc
'                    While RsCola.RecordCount > 10
'                        DoEvents
'                    Wend
                    Call EnviarHuellasHCDN(nBanca) 'version hcdn 2011 'banana
                    For i = LBound(PorHacer) To UBound(PorHacer)
                        PorHacer(i) = "0"
                    Next i
'                    nTick = GetTickCount
'                    While GetTickCount - nTick < 10000
'                        DoEvents
'                    Wend
'                    Set rsTVersion = New ADODB.Recordset
'                    SetearRsW "SELECT top 1 version_datos_sqv from config", rsTVersion
'                    If Not rsTVersion.EOF Then
'                        rsTVersion.MoveFirst
'                        VersionSQV = Trim(rsTVersion.Fields("version_datos_sqv"))
'                    Else
'                        MsgBox "Error al obtener la ultima version de la DB de SQV", vbCritical
'                        Exit Sub
'                    End If
'                    rsTVersion.Close
'                    Set rsTVersion = Nothing
'                    Set rsTemp = New ADODB.Recordset
'                    cConsulta = "SELECT BancaNumero FROM BancasIp WHERE (Version <> version_datos_banca OR version_datos_sqv <> '" & VersionSQV & "') AND BancaNumero > 0 ORDER BY BancaNumero"
'                    SetearRsW cConsulta, rsTemp
'                    While Not rsTemp.EOF
'                        BancasConError = BancasConError & rsTemp.Fields(0) & " "
'                        rsTemp.MoveNext
'                    Wend
'                    nCantidadFallaron = rsTemp.RecordCount
'                    rsTemp.Close
'                    Set rsTemp = Nothing
'                    MsgBox "Revisar las siguientes bancas (error al cargar huellas) : " & BancasConError, vbCritical
'                    If nCantidadFallaron > 0 Then
'                        VectorCompleto = Join(PorHacer, ";")
'                        Call ReconectarBancas
'                        MostrarErr ("Esperando reconexion de bancas pendientes " & nCantidadFallaron & " " & VectorCompleto)
'                        nTick = GetTickCount
'                        Do While GetTickCount - nTick < 10000 'considerar tambien un Not BancasReconectadas() And
'                            DoEvents
'                        Loop
'                        MostrarErr ("Sincronizacion de bancas pendientes " & " " & VectorCompleto)
'                        SincronizarBancas (VectorCompleto) 'se llama recursivamente, pero ya no entra por brc...
'                    Else
'                        MostrarErr ("*** Todas las bancas estan actualizadas ***")
'                    End If
                End If
            Else
                If Val(nBanca) < 0 Or Val(nBanca) > ConexionesAbiertas Then
                    Exit Sub
                Else
                    Call EnviarHuellasHCDN(nBanca) 'version hcdn 2011 'banana
                End If
            End If
        End If
End Sub

Private Sub CmdLimpiarLog_Click()
    txtLog.Text = ""
End Sub

Private Sub cmdSimulaVotos_Click()
Dim i As Integer
Dim msgSQV As MensajeSQV
For i = 1 To 256
    With msgSQV
    .sTipo = "mevt"
    .sObjeto = Trim(Str(i))
    .sComponente = "term.keyb.no"
    .sAtributo = "state"
    .sValor = "on"
    End With
    Call InsertarMsgSQV(msgSQV)
Next i
End Sub
Public Sub ActualizaIPs()
Dim i As Integer
Dim IPS(0 To 256) As String
Dim rsTemp As ADODB.Recordset
i = -1
Set rsTemp = New ADODB.Recordset
SetearRsW "SELECT Ip FROM BancasIP ORDER BY BancaNumero", rsTemp
While Not rsTemp.EOF
    i = i + 1
    If Trim(Banca_ip(i).IP) <> Trim(rsTemp.Fields(0)) Then
        WSocket(i).Close
        Banca_ip(i).IP = Trim(rsTemp.Fields(0))
        WSocket(i).RemoteHost = Banca_ip(i).IP
        WSocket(i).RemotePort = 7000
        WSocket(i).Connect
    End If
    rsTemp.MoveNext
Wend
rsTemp.Close
Set rsTemp = Nothing
End Sub
Private Sub Command1_Click()
Call EnviarSktxCola(Str(18), "SCANCL")
Call EnviarxSkt("X", 18, "STATUS")
End Sub
Private Sub Command2_Click()
FormMain.txtTemp.Text = WSocket(18).State
End Sub

Private Sub Command3_Click()
WSocket(14).Close
WSocket(14).Connect
End Sub

Private Sub Command4_Click()
If Timer1.Enabled = True Then
    Timer1.Enabled = False
    Command4.Caption = "Habilitar Timer"
Else
    Timer1.Enabled = True
    Command4.Caption = "Deshabilitar Timer"
End If
End Sub

Private Sub Form_Activate()
    If App.PrevInstance = True Then
        End
    End If
    'Call CargarGrilla
End Sub
Private Sub Form_Load()
    Dim i As Integer
    CantidadInserciones = 0
    ContadorTLEVER = 0
    TICK_LOG = GetTickCount
    For i = LBound(VectorSAUTOD) To UBound(VectorSAUTOD)
        VectorSAUTOD(i) = "0"
    Next i
    For i = LBound(VectorEnvio) To UBound(VectorEnvio)
        VectorEnvio(i) = "0"
    Next i
    For i = LBound(VectorTicks) To UBound(VectorTicks)
        VectorTicks(i) = 0
    Next i
    SecuenciaStatus = 0
    CantidadInsertados = 0
    Prefijo_Tick = Now()
    'nTicks = GetTickCount
    'If GetTickCount - nTicks > 5000 Then
    '    MostrarErr ("TIMEOUT DE ENVIO EXCEDIDO EN BANCA " & fSocket)
    'End If
    ReDim EnviandoHuellas(MAX_BANCA)
    For i = LBound(Control_Secuencia) To UBound(Control_Secuencia)
        Control_Secuencia(i) = ""
    Next i
    For i = LBound(EnvioCompletado) To UBound(EnvioCompletado)
        EnvioCompletado(i) = False
    Next i
    For i = LBound(LogEnvioCompletado) To UBound(LogEnvioCompletado)
        LogEnvioCompletado(i) = "@"
    Next i
    For i = LBound(LogBancasMuertas) To UBound(LogBancasMuertas)
        LogBancasMuertas(i) = "0"
    Next i
    FormMain.txtLogEnvio.Text = Join(LogEnvioCompletado, "")
    lblVersion.Caption = "Monitor de actividad " & setVersion
    FormMain.txtEnviando.Text = ""
    ' Abro conexion contra el SQL Server
    Call AbrirConexionSQLServer
    'Inicializo Variables globales
    Call iniVariables
    
    'Seteo Cola de Mensajes
    Call setearCola
    
    ' Elimino todos los mensajes de la cola del SQV,
    Call EliminarMensajesSQV(-1)
    
    ' Levanto todas las Bancas
    Call CargarBancas
    
    ' Cargo Legisladores en un  type
    Call CargarLegisladores
        
    'Minimizar el Formulario
    'Me.WindowState =1
    Call Cn.Execute("UPDATE BancasDeshabilitadas SET habilitada = 1")
    For i = 0 To 256
        BancasDeshabilitadas(i) = False 'Arrancan todas habilitadas
    Next i
    txtLog.Text = "CONFIGURANDO Y RECIBIENDO VERSIONES..."
End Sub
Private Sub ActualizaLogMuertas()
Dim x As String
For i = 0 To 256
    x = x & "Banca " & Str(i) & ": " & LogBancasMuertas(i) & vbCrLf
Next i
Call Log_Particular("Bancas_Muertas.txt", Now() & vbCrLf & vbCrLf & x)
End Sub
Private Sub CmdSalir_Click()
    Unload Me
End Sub

'***************************************************
' Aqui es donde Coloco el Socket y Numero de Banca *
'***************************************************
Public Sub CargarBancas()
    Dim strSql           As String
    Dim NumRegistros     As Long
    Dim conta            As Long
    Dim UltimaBanca      As Long
    Dim naUX             As Long
    
    conta = 0
    strSql = "SELECT * FROM bancasIP order by bancanumero"
    Call SetearRsBanca(strSql)
    If Not RsBanca.EOF Then
        RsBanca.MoveLast
        UltimaBanca = RsBanca!BancaNumero
        RsBanca.MoveFirst
        
        NumRegistros = RsBanca.RecordCount
        ConexionesAbiertas = NumRegistros
        
        Call iniVariblesDim(NumRegistros, UltimaBanca)
        
        'Cargo la primera conexion
        '*****************************************************
        WSocket(conta).RemoteHost = RsBanca!IP
        Banca_ip(conta).IP = RsBanca!IP
        
        WSocket(conta).RemotePort = RsBanca!Puerto
        Banca_ip(conta).Puerto = RsBanca!Puerto
        Banca_ip(conta).tVersion = RsBanca!Version
                
        WSocket(conta).Connect
        B2Skt(RsBanca!BancaNumero).Socket = conta
        B2Skt(RsBanca!BancaNumero).Estado = True
        Skt2B(conta).Banca = RsBanca!BancaNumero
        
        Skt2B(conta).Estado = True
        
        RsBanca.MoveNext
        'Cargo Conexiones Suesivas.
        '****************************************************
        Do While Not RsBanca.EOF
            conta = conta + 1
            Load WSocket(conta)
            WSocket(conta).RemoteHost = RsBanca!IP
            Banca_ip(conta).IP = RsBanca!IP
            
            WSocket(conta).RemotePort = RsBanca!Puerto
            Banca_ip(conta).Puerto = RsBanca!Puerto
            
            B2Skt(RsBanca!BancaNumero).Socket = conta
            B2Skt(RsBanca!BancaNumero).Estado = True
            
            Skt2B(conta).Banca = RsBanca!BancaNumero
            Skt2B(conta).Estado = True
            
            'Para el sistema de Huellas
            
            Banca_ip(conta).tBancaMinMax = RsBanca!secuencialegislador
            Banca_ip(conta).tBancaBusca = Mid(RsBanca!idstring, 1, 40)
            Banca_ip(conta).tBancaSecuencia = NullCadena(RsBanca!ultimolegislador)
            Banca_ip(conta).tbancaMinMaxMan = RsBanca!secuenciamantenimiento
            Banca_ip(conta).tVersion = RsBanca!Version
            
            WSocket(conta).Connect
            RsBanca.MoveNext
        Loop
    End If
End Sub

Public Sub GuardarCacheBanca()
    Dim naUX  As Integer
    Dim sCad  As String
    For naUX = 0 To UBound(Banca_ip)
        sCad = "UPDATE bancasIP set ultimolegislador = '" & Mid(Banca_ip(naUX).tBancaSecuencia, 1, 40) & "' where bancanumero =  " & Str(Skt2B(naUX).Banca)
        Cn.Execute (sCad)
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim naUX     As Integer
    Call GuardarCacheBanca
    For naUX = 0 To ConexionesAbiertas - 1
'        If Skt2B(naUX).Estado = True Then
'            With msgSQV
'                .sTipo = "mevt"
'                .sObjeto = Skt2B(naUX).Banca 'Convierto socket por banca
'                .sComponente = "term"
'                .sAtributo = "state"
'                .sValor = "off"
'            End With
'            Call InsertarMsgSQV(msgSQV)
'            Call EnviarxSkt("f", naUX, "SRESET")
'        End If
         WSocket(naUX).Close 'HCDN 2011
         'Se cierra el Socket formalmente para evitar complicaciones con las bancas
    Next
    End
End Sub

Private Sub Timer_Timer()
     CountTimer = CountTimer + 1
    If FlagTimerLeoCola = True Then
        FlagTimerLeoCola = False
        ' Call EnviarDatosSocket
        ' Nueva funcion que envia socket por Socket
        Call EnviarDatosSkt
        Call LeerMensajesSQV
        FlagTimerLeoCola = True
    End If
    If CountTimer = 100 Or CountTimer = 200 Then
        'Call CargarGrilla
        LabelTimer.Caption = CountTimerAux
        CountTimerAux = CountTimerAux + 1
    End If
    If CountTimer = 300 Then
        CountTimer = 0
        'Call CargarGrilla
        If EstadoEnviandoHuellas = False Then
            Call EnviarSktxCola("brc", "STATUS") 'banana
            'Call EnviarxSkt("f", 256, "STATUS") FUNCIONA!
            Log_DEBUG ("Se envio STATUS a todas las bancas")
            Call ReconectarBancas
        End If
    End If
    'lblRegistrosEnCola.Caption = RsCola.RecordCount
End Sub



'**********************************************************************************************
' Solo Manejo de Funciones para el Socket :) **************************************************
'**********************************************************************************************

'***************************************************
' Enviar los datos a las Bancas                    *
'***************************************************
Public Sub EnviarxSkt(fsecuencia As String, fSocket As Integer, fMensaje As String)

On Error GoTo errSocket
If InStr(fMensaje, "SRESET") And fSocket = 256 Then
    FormMain.txtTemp = "ASDAD"
End If
If InStr(fMensaje, "SCANCL") And fSocket = 0 Then
    fSocket = fSocket
End If
    Dim sEnviar As String
    Dim nTicks As Long
    Dim Registrado As Boolean
    sEnviar = Left(fsecuencia, 1) & fMensaje & vbCrLf

    'If DebugSecuencia = fsecuencia And fSocket = 256 Then 'InStr(1, fMensaje, "SAUTOD") > 0
    '    MsgBox DebugSecuencia
    'End If
    If InStr(fMensaje, "SFINVT") And fSocket = 33 Then
        FormMain.txtTemp.Text = "INTENTO DE ENVIO SLIMVT " & Now
    End If
    
    If Mid(fMensaje, 1, 6) = "SIDRXH" And Len(fMensaje) = 6 Then
        'CONFIGURACION DE IDENTIFICACION POR HUELLA
        'Aqui se indica para el comando de identificacion por huella
        'los parametros que aceleran la busqueda de huella en la banca
        'Implementacion banca Cordoba 03:
        ' Se envia:
        '  1. Numero de huella a partir del cual hacer la busqueda.
        '     Esto permite buscar solo huellas de Mantenimiento, enviando la primer huella de mantenimiento.
        '  2. Cache de busqueda rapida de huellas mas probables.
        '  2.1 Ultimas tres huellas que se identificaron en esa banca.
        '  2.2 Huellas mas probables definidas por configuracion.
        '
        '
        If False Then
            'Configuracion de bancas frecuentes cordoba 03
            sEnviar = fsecuencia & fMensaje & " ^" _
                & Banca_ip(fSocket).tBancaMinMax _
                & Mid(Banca_ip(fSocket).tBancaSecuencia, 1, 12) & Banca_ip(fSocket).tBancaBusca & vbCrLf
        Else
            sEnviar = fsecuencia & fMensaje & vbCrLf '& " ^" REVISAR no se entiende para que agrega este caret (ap 090908)
            ' antes 090908 sEnviar = fsecuencia & fMensaje & " ^" & vbCrLf 'REVISAR no se entiende para que agrega este caret (ap 090908)
        End If
        'If fSocket = 70 Then Stop 'solo depuracion
        'Repetir bloque en el otro else
        EnvioCompletado(fSocket) = False
        nTicks = GetTickCount
        WSocket(fSocket).SendData sEnviar
        Call Log_DEBUG("(SEND1) - SOCKET " & Str(fSocket) & " - " & sEnviar)
        If Not MODOLIGHT Then
            Call Log_Banca(Trim(Str(fSocket)) & ".txt", Now() & " -(1) SE ENVIO " & sEnviar)
        End If
        'DoEvents
        Do While Not EnvioCompletado(fSocket)
            If GetTickCount - nTicks > 5 Then
                If Not Registrado Then
                    Registrado = True
                    LogEnvioCompletado(fSocket) = Max("A", Min("Z", Chr(Asc(LogEnvioCompletado(fSocket)) + 1)))
                    FormMain.txtLogEnvio.Text = Join(LogEnvioCompletado, "")
                    Log_DEBUG (PanelDeControl(FormMain.txtLogEnvio.Text, "01 No se recibio completed " & Str(fSocket) & " Mensaje: " & fMensaje))
                End If
                If GetTickCount - nTicks > 15000 Then
                    MostrarErr ("TIMEOUT DE ENVIO EXCEDIDO EN BANCA " & fSocket)
                    LogEnvioCompletado(fSocket) = "*"
                    FormMain.txtLogEnvio.Text = Join(LogEnvioCompletado, "")
                    GoTo errSocket
                    Exit Do
                End If
            End If
            DoEvents
        Loop
    Else
        If (Left(fMensaje, 6) = "SRLEGI") Then
            'If (Mid(fMensaje, 11, 1) = "0") Or True Then
                'vrs cba Call MostrarErr("Enviando Huella : " & Str(Int("&H" & Mid(fMensaje, 9, 4)) + 1) & " a la banca : " & Skt2B(fSocket).Banca)
                Call MostrarErr("Enviando Huella : (" & Len(sEnviar) & ")" & IDLegisladorHuella(fMensaje) & " a la banca : " & Skt2B(fSocket).Banca & " Cola: " & RsCola.RecordCount, Str(Skt2B(fSocket).Banca))
                'Call GuardarLog(Str(fSocket), fMensaje)
                'srlx
            'End If
        End If
        EnvioCompletado(fSocket) = False
        nTicks = GetTickCount
        WSocket(fSocket).SendData sEnviar 'banana
        Call Log_DEBUG("(SEND2) - SOCKET " & Str(fSocket) & " - " & sEnviar)
        If Not MODOLIGHT Then
            Call Log_Banca(Trim(Str(fSocket)) & ".txt", Now() & " -(2) SE ENVIO " & sEnviar)
        End If
        'DoEvents NICO
        Do While Not EnvioCompletado(fSocket)
            If GetTickCount - nTicks > 5 Then
                If Not Registrado Then
                    Registrado = True
                    LogEnvioCompletado(fSocket) = Max("A", Min("Z", Chr(Asc(LogEnvioCompletado(fSocket)) + 1)))
                    FormMain.txtLogEnvio.Text = Join(LogEnvioCompletado, "")
                    Log_DEBUG (PanelDeControl(FormMain.txtLogEnvio.Text, "01 No se recibio completed " & Str(fSocket) & " Mensaje: " & fMensaje))
                    'MostrarErr ("--------------------> Envio no completo (" & Str(GetTickCount - nTicks) & "), se va a esperar. " & fSocket)
                End If
                If GetTickCount - nTicks > 5000 Then
                    If Not MODOLIGHT Then
                        Call Log_Banca(Trim(Str(fSocket)) & ".txt", Now() & " - TIMEOUT 5s de EnvioCompletado SKT2B : " & Skt2B(fSocket).Banca & " - " & sEnviar)
                    End If
                    LogEnvioCompletado(fSocket) = "*"
                    FormMain.txtLogEnvio.Text = Join(LogEnvioCompletado, "")
                    GoTo errSocket
                    Exit Do
                End If
            End If
            Exit Do ' para probar que no espere... poner arriba tambien
            DoEvents
        Loop
    End If
    DoEvents
    'Guardo log en log_banca
    'If Not fMensaje = "STATUS" Then
        'Call GuardarLog(Str(fSocket), (sEnviar))
    'End If
        'If fSocket = 128 Or fSocket = 159 Then
        'End If
        'End If
    Exit Sub
errSocket:
    If Err.Number = 40006 Then
        Call WSocketClose(fSocket, "ERROR SOCKET (" & Err.Number & " " & Err.Description & ")")
        Call Log_DEBUG("XSE CERRO LA CONEXION CON EL SOCKET " & Str(fSocket) & " | CAUSA: " & Err.Description)
        If Not MODOLIGHT Then
            Call Log_Banca(Trim(Str(fSocket)) & ".txt", Now() & " - ERROR 40006: " & Err.Description)
        End If
        Exit Sub
    Else
        'MostrarErr (Err.Number & " " & Err.Description & " FUNCION EnviarxSkt")
        If Not MODOLIGHT Then
            Call Log_Banca(Trim(Str(fSocket)) & ".txt", Now() & " - ERROR FATAL: " & Err.Description)
        End If
    End If
End Sub
Public Function PanelDeControl(panel As String, Optional id As String) As String
'Dim Buffer As String
'Dim x As String
'Dim i As Integer
'Buffer = "panel" & id & vbCrLf
'panel = Replace(panel, "@", ".")
'For i = 0 To 9
'    Buffer = Buffer & i & "         "
'Next i
'Buffer = Buffer & vbCrLf
'For i = 0 To 9
'    Buffer = Buffer & "0123456789"
'Next i
'Buffer = Buffer & vbCrLf & Mid(panel, 1, 100)
'Buffer = Buffer & vbCrLf & Mid(panel, 101, 100)
'Buffer = Buffer & vbCrLf & Mid(panel, 201, 57)
'PanelDeControl = Buffer
End Function
Public Function BancaEnDebug(fSocket As Integer) As Boolean
    BancaEnDebug = False
    If fSocket <= 20 Then BancaEnDebug = True
End Function

Public Function IDLegisladorHuella(fMensaje As String) As String
Dim xIDHexadecimalInvertido As String
Dim xIDHexadecimal As String
Dim xIDDecimal As Integer
Dim i As Integer
'srlegi 180800000f060204xxxxx
'123456789012345678901234
xIDHexadecimalInvertido = Mid(fMensaje, 8 + (8 * 2), 4 * 2)
xIDHexadecimal = "&H"
For i = 1 To 8 Step 2
    xIDHexadecimal = Trim(xIDHexadecimal) & Mid(xIDHexadecimalInvertido, 8 - i, 2)
Next i

IDLegisladorHuella = Str(Val(xIDHexadecimal))
End Function

Public Sub Configurar(Index As Integer)
    If LCase(Command) = "-on" Then
       'Modo Presencia
       Call EnviarxSkt("X", Index, "SCONFG 2010010")
    Else
       'Call EnviarxSkt("X", Index, "SCONFG 2010011")
       'Call EnviarxSkt("X", Index, "SCONFG 9010011")
       Call EnviarxSkt("X", Index, "SCONFG 9010061")
    End If
    'Modo Presencia :)
    'Call EnviarxSkt("f", index, "SCONFG ^1E1800003C03")
    'Call EnviarxSkt("f", index, "SCANCL")
    Call EnviarxSkt("X", Index, "STATUS")
    Call EnviarxSkt("X", Index, "SLEVER")
    txtLog.Text = "BANCAS PROCESADAS " & ContadorTLEVER
End Sub

Private Sub Timer1_Timer()
Dim c As Integer
For c = 1 To 10
    Call EnviarSktxCola("brc", "STATUS")
Next c
End Sub
Private Sub Timer2_Timer()
Transcurrido = FormMain.TiempoTranscurrido
End Sub
Private Sub tmIps_Timer()
ActualizaIPs
End Sub

Private Sub WSocket_Connect(Index As Integer)
    EnviandoHuellas(Index) = False
    If Skt2B(Index).Banca = 0 Then
        If EstadoEnviandoHuellas Then
            Call WSocketClose(Index, "BANCA CERO DURANTE ENVIO HUELLAS")
            Exit Sub
        End If
    End If
    If Not MODOLIGHT Then
        Call Log_Banca(Trim(Str(Skt2B(Index).Banca)) & ".txt", "-----------------------SE CONECTÓ LA BANCA (" & Now() & "-----------------------")
    End If
    Banca_ip(Index).Estado = True ' Pongo en el Estado en que se encuentra la conexion
    'Call EliminarMensajeCola(Index, "", True)
    
    Call EnviarSktxCola(Str(Index), "SCANCL")
    'Call EnviarSktxCola(Str(Index), "SRESET")
    'Hacer una especie de filter para hacer un while y que termine el mensaje de reset.
    'Call FormMain.EnviarxSkt("f", Index, "SRESET")
    'MostrarErr ("WARNING: Se conectó a la BANCA " & Index)
    'Call EnviarSktxCola(Str(Index), "SRESET")
    ''Call EnviarSktxCola(Str(index), "SCONFG ^1E1801003C03")
    ''Modo Presencia :)
    'Call EnviarSktxCola(Str(index), "SCONFG ^1E1800003C03")
    'Call EnviarSktxCola(Str(Index), "SCANCL") 'revisar este arranque 110211
    'Call EnviarSktxCola(Str(index), "STATUS")
    
    'Call Configurar(index) '091028 eliminado
    
    UltimoEstadodeBanca(Index) = "ok"
    UltimoEstadodePresencia(Index) = "nada"
    FueConfigurada(Index) = False

    'Envio Estado de Conectado al SQV
    '********************************
    With msgSQV
            .sTipo = "mevt"
            .sObjeto = Skt2B(Index).Banca
            .sComponente = "term"
            .sAtributo = "state"
            .sValor = "ok"
    End With
    Call InsertarMsgSQV(msgSQV)

End Sub

Private Sub WSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'On Error GoTo errorSkt

    Dim DatoLlegado       As String
    Dim strAux            As String
    Dim strPars()         As String
    Dim naUX              As Long
    
    'Si el sokcet esta cerrado por algo me voy
    '*****************************************
    If WSocket(Index).State = 0 Then
        If Not MODOLIGHT Then
            Call Log_Banca(Trim(Str(Index)) & ".txt", Now() & " - SE RECIBIERON DATOS CON EL SOCKET CERRADO | CANTIDAD BYTES: " & Str(bytesTotal))
        End If
       Exit Sub
    End If
    If Skt2B(Index).Banca = 0 Then
        If EstadoEnviandoHuellas Then
            Call WSocketClose(Index, "BANCA CERO DURANTE ENVIO HUELLAS")
            Exit Sub
        End If
    End If
    
    'Leo el dato del Socket ----> Mirar bien esta funcion Importante !!!!!!!!!!!!
    '**********************
    WSocket(Index).GetData DatoLlegado, , bytesTotal
    Log_DEBUG ("DATO RECIBIDO BANCA " & Str(Index) & " - " & DatoLlegado)
    If Not MODOLIGHT Then
        Call Log_Banca(Trim(Str(Index)) & ".txt", Now() & " - SE RECIBIO " & DatoLlegado & "(" & Str(bytesTotal) & ")")
    End If
    strDato(Index) = strDato(Index) & DatoLlegado
    If InStr(strDato(Index), Chr(13)) > 0 Then
       strDato(Index) = Replace(strDato(Index), vbCrLf, ";")
       strDato(Index) = Replace(strDato(Index), Chr(13), ";")
        'If (Index <= 2) Then
        '    Log_DEBUG (" - DATO MULTILINEA DE BANCA " & Index & " : " & strDato(Index) & " - ")
        'End If
       strPars = Split(strDato(Index), ";")
       StrProximo(Index) = strPars(UBound(strPars))
       For naUX = 0 To UBound(strPars) - 1
            'Aqui se llama a la funcion para interpretar los datos
            '*****************************************************
            'Ver si pongo que chequee "f" o no  !!!!!!!!!!!!!!
'            If Not (Mid(strPars(nAux), 1, 1) = "f" Or (Mid(strPars(nAux), 1, 1) > "A" And Mid(strPars(nAux), 1, 1) < "Z")) Then EstadoxEnviar(index) = True
            'Call MostrarErr("Mensaje llegado " & strPars(nAux) & " Socket " & Str(index))
            Call InterpretaDatosSkt(Index, strPars(naUX)) 'manzana
            'If (Index <= 2) Then
                'Log_DEBUG (" - VUELTA " & naUX & "DE BANCA " & Index & " : " & strPars(naUX) & " - ")
            'End If
       Next
       strDato(Index) = StrProximo(Index)
    End If
Exit Sub
errorSkt:
    'Paso Algo con el Socket Mandar a funcion de err
    '***********************************************
    'MsgBox "Enviar a la Funcion de ERROR SOCKET"
    Call MostrarErr("Se intento acceder a un socket con error: " & Str(Skt2B(Index).Banca), Str(Skt2B(Index).Banca))
    Exit Sub
End Sub

Private Sub WSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Index = 256 Then
    txtTemp.Text = Description
End If
End Sub
Public Sub ErrorSwitchBanca1(Index As Integer)
        With msgSQV
            .sTipo = "mevt"
            .sObjeto = Skt2B(Index).Banca 'Convierto socket por banca
            .sComponente = "term.ioc"
            .sAtributo = "state"
            .sValor = "error"
        End With
        Call InsertarMsgSQV(msgSQV)
        Log_DEBUG ("ERROR IOC SWITCH BANCA " & Index)
End Sub
Public Sub WSocketClose(Index As Integer, Optional sObservaciones As String)
    Dim msgSQV As MensajeSQV
    
    ' Enviar Estado de que la banca se deconecto
    '*******************************************
    With msgSQV
        .sTipo = "mevt"
        .sObjeto = Skt2B(Index).Banca 'Convierto socket por banca
        .sComponente = "term"
        .sAtributo = "state"
        .sComentario = "a05"
        .sValor = "off"
    End With
    Call InsertarMsgSQV(msgSQV)
    
    'Ponemos estadon en False
    
    UltimoEstadodeBanca(Index) = "off"
    UltimoEstadodePresencia(Index) = "nada"
    Banca_ip(Index).Estado = False
    BanderaReset(Index) = True
    If EnviandoHuellas(Index) Then
        EnviandoHuellas(Index) = False
        Call MostrarEnviandoHuellas
    End If
    
    WSocket(Index).Close
    'DoEvents
    
    'Limpio la cola de mensajes de esta banca
    Call EliminarMensajeCola(Index, "", True) ' True elimina todos los mensajes
    'Call MostrarErr(sObservaciones & "-02 Se cerro la conexion de la banca : " & Str(Skt2B(Index).Banca), Str(Skt2B(Index).Banca))
    Call Log_DEBUG(sObservaciones & "-SKTCLOSE Se cerro la conexion de la banca : " & Str(Skt2B(Index).Banca))
End Sub

Private Sub WSocket_Close(Index As Integer)
'    Dim msgSQV As MensajeSQV
'    Cerrar_Banca (Index)
'    If Banca_ip(Index).Estado = True Then
'        ' Enviar Estado de que la banca se deconecto
'        '*******************************************
'        With msgSQV
'            .sTipo = "mevt"
'            .sObjeto = Skt2B(Index).Banca 'Convierto socket por banca
'            .sComponente = "term"
'            .sAtributo = "state"
'            .sComentario = "Por cierre de socket"
'            .sValor = "off"
'        End With
'        Call InsertarMsgSQV(msgSQV)
'
'        'Ponemos estadon en False
'
'        UltimoEstadodeBanca(Index) = "off"
'        UltimoEstadodePresencia(Index) = "nada"
'        If EnviandoHuellas(Index) Then
'            Call MostrarEnviandoHuellas
'            EnviandoHuellas(Index) = False
'        End If
'        Banca_ip(Index).Estado = False ' Pongo en el Estado en que se encuentra la conexion
'        'Limpio la cola de mensajes de esta banca
'        Call EliminarMensajeCola(Index, "", True) ' True elimina todos los mensajes
'        BanderaReset(Index) = True
'    End If
    Call MostrarErr("03 Se cerro la conexion de la banca : " & Str(Skt2B(Index).Banca), Str(Skt2B(Index).Banca))
    Call Log_DEBUG("03 Se cerro la conexion de la banca : " & Str(Skt2B(Index).Banca))
    If Not MODOLIGHT Then
        Call Log_Banca(Trim(Str(Index)) & ".txt", Now() & " - 03 Se cerro la conexion de la banca.")
    End If
End Sub
Private Sub Cerrar_Banca(Index As Integer)
        If Val(LogBancasMuertas(Index)) > 50 Then
            LogBancasMuertas(Index) = "50"
        Else
            LogBancasMuertas(Index) = Str(Val(LogBancasMuertas(Index)) + 1)
        End If
        ActualizaLogMuertas
        ' Enviar Estado de que la banca se deconecto
        '*******************************************
        With msgSQV
            .sTipo = "mevt"
            .sObjeto = Skt2B(Index).Banca 'Convierto socket por banca
            .sComponente = "term"
            .sAtributo = "state"
            .sComentario = "01Por cierre de socket"
            .sValor = "off"
        End With
        Call InsertarMsgSQV(msgSQV)
        
        'Ponemos estadon en False
        
        UltimoEstadodeBanca(Index) = "off"
        UltimoEstadodePresencia(Index) = "nada"
        If EnviandoHuellas(Index) Then
            Call MostrarEnviandoHuellas
            EnviandoHuellas(Index) = False
        End If
        Banca_ip(Index).Estado = False ' Pongo en el Estado en que se encuentra la conexion
        'Limpio la cola de mensajes de esta banca
        Call EliminarMensajeCola(Index, "", True) ' True elimina todos los mensajes
        BanderaReset(Index) = True
End Sub
'Funcion para reconectar bancas.
'*********************************************************************
Public Sub ReconectarBancas()
    Dim naUX As Integer
        For naUX = 0 To ConexionesAbiertas - 1
            If Skt2B(naUX).Banca <= MAX_BANCA Then ' ATENCION
                If BancasDeshabilitadas(Skt2B(naUX).Banca) = False Then
                    If WSocket(naUX).State = sckClosed Or WSocket(naUX).State = sckConnecting Then
                        Cerrar_Banca (naUX)
                        If Not MODOLIGHT Then
                            Call Log_Banca(Trim(Str(Skt2B(naUX).Banca)) & ".txt", Now() & " - WARNING: FUNCION Cerrar_Banca por SktClosed")
                        End If
                        WSocket(naUX).Close
                        WSocket(naUX).Connect
                        DoEvents
                    Else
                        If WSocket(naUX).State = sckError Then
                            WSocket(naUX).Close
                            Cerrar_Banca (naUX)
                            If Not MODOLIGHT Then
                                Call Log_Banca(Trim(Str(Skt2B(naUX).Banca)) & ".txt", Now() & " - WARNING: FUNCION Cerrar_Banca por SktError")
                            End If
                            If naUX = 256 Then
                                FormMain.txtEnviando.Text = "OK 256 2"
                            End If
                            WSocket(naUX).Connect
                            DoEvents
                        End If
                    End If
                Else
                    With msgSQV
                        .sTipo = "mevt"
                        .sObjeto = Str(naUX)
                        .sComponente = "term"
                        .sComentario = "a01"
                        .sAtributo = "state"
                        .sValor = "off"
                    End With
                    Call InsertarMsgSQV(msgSQV)
                    UltimoEstadodeBanca(naUX) = "off"
                End If
            Else
                Call MostrarErr("No se reconecta banca " & Str(Skt2B(naUX).Banca), Str(Skt2B(naUX).Banca))
            End If
        Next
End Sub
Private Sub WSocket_SendComplete(Index As Integer)
If Not MODOLIGHT Then
    Call Log_Banca(Trim(Str(Index)) & ".txt", Now() & " - EVENTO SENDCOMPLETE")
End If
EnvioCompletado(Index) = True
End Sub
Public Function Min(a, b) As Variant
Min = IIf(a < b, a, b)
End Function
Public Function Max(a, b) As Variant
Max = IIf(a > b, a, b)
End Function
Public Sub Reconectar(Index As Integer)
WSocket(Index).Close
WSocket(Index).Connect
End Sub
Public Function TiempoTranscurrido() As Boolean
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
SetearRsW "SELECT TiempoTranscurrido FROM ComunicacionRapida", rsTemp
If Not rsTemp.EOF Then
    If rsTemp.Fields(0) = "1" Then
        TiempoTranscurrido = True
    Else
        TiempoTranscurrido = False
    End If
Else
    TiempoTranscurrido = False
End If
rsTemp.Close
Set rsTemp = Nothing
End Function
