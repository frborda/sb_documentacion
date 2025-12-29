VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmMenu 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menú Principal"
   ClientHeight    =   6045
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   8955
   ControlBox      =   0   'False
   ForeColor       =   &H00808080&
   Icon            =   "frmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdConsultas 
      Height          =   1005
      Left            =   450
      TabIndex        =   3
      Top             =   3360
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   1773
      BackColor       =   12230304
      Caption         =   "&Consultas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      PicOpacity      =   0
   End
   Begin VB.CheckBox chkScreens 
      BackColor       =   &H00404040&
      Caption         =   "Habilitar capturas de pantalla automáticas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   2340
      TabIndex        =   2
      Top             =   660
      Width           =   4080
   End
   Begin MSWinsockLib.Winsock WSock 
      Left            =   -180
      Top             =   4770
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Proyecto1.ButtonOffice cmdConsola 
      Height          =   525
      Left            =   240
      TabIndex        =   1
      Top             =   90
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   926
      BackColor       =   16744576
      Caption         =   "Consola de Operación"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdConfig 
      Height          =   1005
      Left            =   4470
      TabIndex        =   4
      Top             =   3360
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   1773
      BackColor       =   12230304
      Caption         =   "Con&figuraciones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdPresidente 
      Height          =   495
      Left            =   450
      TabIndex        =   5
      Top             =   4380
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "&Orden de Selección de Presidente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdPeriodos 
      Height          =   495
      Left            =   4470
      TabIndex        =   6
      Top             =   4380
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "&Períodos Legislativos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdModificacionDatos 
      Height          =   495
      Left            =   450
      TabIndex        =   7
      Top             =   4890
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "&Modificación de Datos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdLogIdentificaciones 
      Height          =   495
      Left            =   450
      TabIndex        =   8
      Top             =   5400
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "Log de Identificaciones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdSalir 
      Height          =   495
      Left            =   4470
      TabIndex        =   9
      Top             =   5400
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "&Salir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdEstadisticas 
      Height          =   495
      Left            =   4470
      TabIndex        =   13
      Top             =   4890
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "Estadísticas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      PicOpacity      =   0
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Versión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   450
      TabIndex        =   12
      Top             =   3060
      Width           =   8055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "de la Nación Argentina"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   3300
      TabIndex        =   11
      Top             =   1830
      Width           =   3465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Honorable Cámara de Diputados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   750
      TabIndex        =   10
      Top             =   1140
      Width           =   7515
   End
   Begin VB.Label lblModoPrueba 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consola en modo PRUEBA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   4740
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private xArchivoDBLeg    As Long
Const cLongRegDBLeg      As Long = 144
Private strArchivoHuella As String

Private Sub chkScreens_Click()
If chkScreens.Value = vbChecked Then
    AutoCaptura = True
Else
    AutoCaptura = False
End If
End Sub

Private Sub cmdConfig_Click()
    Me.Enabled = False
    FrmConfigurarConsola.Show vbModal
    Me.Enabled = True
End Sub
Private Sub hacerSplitVector(ByVal pCadena As String, ByRef pVector() As String)
    pVector = Split(pCadena, ";")
End Sub
Private Sub cmdConsola_Click()
Dim Rs As ADODB.Recordset
Dim mVectorIdentificacion() As String
cmdConsola.Enabled = False
En_Proceso_De_Carga = True
Set Rs = New ADODB.Recordset
SetearRs "SELECT * FROM vector", Rs
If Not Rs.EOF Then
    hacerSplitVector Trim(Rs!vector_identificacion), mVectorIdentificacion
End If
If mVectorIdentificacion(0) = "0" Then
    frmElegirPresidente.Show vbModal, Me
End If
Rs.Close
Set Rs = Nothing
En_Proceso_De_Carga = False
frmCargando.Show vbModal, Me
If Error_Carga = False Then
    Screen.MousePointer = 11
    If chkScreens.Value = vbChecked Then
        Screens_Habilitadas = True
    Else
        Screens_Habilitadas = False
    End If
    frmConsolaOperacion.Show 1, Me
Else
    cmdConsola.Enabled = True
End If
End Sub
Private Sub cmdconsultas_Click()
    Me.Enabled = False
    If PermisosTotales.ConsultaActas = 1 Then
        frmConsultas.Show vbModal
    Else
        MsgBox "El usuario no tiene permisos para realizar esta tarea", vbInformation
    End If
    Me.Enabled = True
End Sub
Private Sub cmdEstadisticas_Click()
frmPreEstadisticas.Show vbModal, Me
End Sub
Private Sub cmdLegisladores_Click()
    Me.Enabled = False
    If PermisosTotales.ABMLegisladores = 1 Then
        'frmABMLegisladores.Show vbModal ' -> Deshabilitado HCDN 11
        frmHistorico.Show vbModal
    Else
        MsgBox "El usuario no tiene permisos para realizar esta tarea", vbInformation + vbOKOnly
    End If
'    Dim strLeg               As String
'    Dim strlogleg            As String
'    Dim RsLegis              As ADODB.Recordset
'    Dim strCadena            As String
'    Dim strSql               As String
'    Dim strHuella            As String
'    Dim xId                  As String
'    Dim xNom                 As String
'    Dim xApe                 As String
'    Dim xBloqueControl       As String
'    Dim xIndice              As String
'    Dim xBloque              As String
'    Dim xRespuesta           As Integer
'    Dim xDesp                As Long
'    Dim x                    As Long
'    Dim i                    As Long
'    Dim xCadena              As Long
'
'    Set RsLegis = New ADODB.Recordset
'    FlHuellas.Path = App.Path
'    FlHuellas.Pattern = "*.huella"
'    FlHuellas.Refresh
'    DoEvents
'    If FlHuellas.ListCount = 0 Then
'        Exit Sub
'    End If
'    Call RecuperarBDLeg(20, strLeg, strlogleg)
'    Call AbrirDB
'    xDesp = 8
'    xId = "NINGUNO"
'    xBloqueControl = "09"
'    If Len(strLeg) Mod 144 <> 0 Then
'        Exit Sub 'error
'    End If
'    For x = 1 To Len(strLeg) Step 144 ' Leer archivo linea a linea
'        ' Esta es una nueva linea - obtener indice de linea y bloque
'        If Mid(strLeg, x, 8) <> "SRLEGI ^" Then
'            'error
'            MsgBox "Informacion no reconocida en posicion: " & Str(x) & " (linea: " & Int(Str(x / 144)) & ") " & vbCrLf & Mid(strLeg, x, 40) & "..."
'            Exit Sub
'        End If
'        xIndice = Mid(strLeg, x + xDesp, 4)
'        xBloque = Mid(strLeg, x + xDesp + 4, 2)
'        If ((Val(xBloqueControl) + 1) Mod 10) <> Val(xBloque) Then
'            Exit Sub 'error
'        End If
'        If xBloque = "00" Then
'            xBloqueControl = xBloque
'            If xId <> "NINGUNO" Then
'                'buscar y grabar
'                strSql = "SELECT * FROM Legisladores WHERE Id = '" & xId & "'"
'                SetearRsW strSql, RsLegis
'                If RsLegis.RecordCount = 0 Or RsLegis.EOF = True Or RsLegis.BOF = True Then
'                    ' Si no esta el legislador, debo insertarlo
'                    xRespuesta = MsgBox("¿Desea enrolar al legislador " & xApe & ", " & xNom & " ?", vbQuestion + vbYesNo)
'                    If xRespuesta = vbYes Then
'                        frmAltaLegislador.Id = xId
'                        frmAltaLegislador.Apellido = xApe
'                        frmAltaLegislador.Nombre = xNom
'                        frmAltaLegislador.Huella = strHuella
'                        frmAltaLegislador.Operacion = "addnew"
'                        frmAltaLegislador.Show 1, Me
'                        Call BackupFiles
'                    End If
'                    RsLegis.Close
'                Else
'                    ' Si encuentro el legislador, debo actualizar el campo Template1 despues de confirmar operacion
'                    xRespuesta = MsgBox("¿Desea actualizar la impresión dactilar del legislador " & xApe & ", " & xNom & "?", vbQuestion + vbYesNo)
'                    If xRespuesta = vbYes Then
'                        frmAltaLegislador.Id = xId
'                        frmAltaLegislador.Apellido = xApe
'                        frmAltaLegislador.Nombre = xNom
'                        frmAltaLegislador.Huella = strHuella
'                        frmAltaLegislador.Operacion = "update"
'                        frmAltaLegislador.Show 1, Me
'                        Call BackupFiles
'                    End If
'                    RsLegis.Close
'                End If
'            End If
'            xId = Mid(strLeg, x + xDesp + 4 + 2, 8)
'            xApe = Trim(HexATexto(Mid(strLeg, x + xDesp + 4 + 2 + 8, 60), 30))
'            xNom = Trim(HexATexto(Mid(strLeg, x + xDesp + 4 + 2 + 8 + 60, 60), 30))
'            strHuella = ""
'        Else
'            xBloqueControl = xBloque
'            strHuella = strHuella & Mid(strLeg, x + xDesp + 4 + 2, 128)
'        End If
'    Next x
    Me.Enabled = True
End Sub
Private Sub cmdLogIdentificaciones_Click()
Me.Enabled = False
frmLogIdentificaciones.Show vbModal, Me
Me.Enabled = True
End Sub

Private Sub cmdModificacionDatos_Click()
Me.Enabled = False
frmPreABM.Show vbModal, Me
Me.Enabled = True
End Sub
Private Sub cmdpresidente_Click()
Me.Enabled = False
    frmOrdenSeleccionPresidente.Show 1, Me
Me.Enabled = True
End Sub

Private Sub cmdimpresoras_Click()
Me.Enabled = False
    frmImpresoras.Show 1, Me
Me.Enabled = True
End Sub
Private Sub cmdperiodos_Click()
Me.Enabled = False
frmAltaPeriodo.Show vbModal
Me.Enabled = True
End Sub
Private Sub cmdSalir_Click()
Cerrar
End Sub
Private Sub Cerrar()
Dim xRta As Integer
xRta = MsgBox("¿Desea terminar la actual sesión de trabajo?", vbQuestion + vbYesNo, "Terminar sesión de consola SQV")
If xRta = vbYes Then
    frmLogin.SetDefault
    End
End If
End Sub
Private Sub CargarLegisladoresActivos()
    Dim X As Long
    Dim strSql As String
    Dim Cn As ADODB.Connection
    Set Cn = New ADODB.Connection
    With Cn
        .ConnectionString = strconexion
        .CommandTimeout = 15
        .CursorLocation = adUseClient
        .Open
    End With
    Cn.BeginTrans
        strSql = "delete from legisladores_activos "
        Cn.Execute (strSql)
        For X = 1 To 70
            strSql = "INSERT INTO legisladores_activos (Id, DeskId, OrdenPresidente) VALUES ('" & Trim(Str(X)) & "','" & Trim(Str(X)) & "'," & X + 1 & ")"
            Cn.Execute (strSql)
        Next X
    Cn.CommitTrans
End Sub

Private Sub FIXIT()
'Dim rsTemp As ADODB.Recordset
'Set rsTemp = New ADODB.Recordset
'SetearRs "SELECT id,apellido,nombre FROM Legisladores_activos ORDER BY apellido,nombre", rsTemp
'While Not rsTemp.EOF
'    Dim res As String
'    Dim temp As String
'    res = InputBox(rsTemp.Fields("apellido") & " " & rsTemp.Fields("nombre") & ": ", "DeskId")
'    temp = "UPDATE legisladores_estado SET numero_orden_activacion = " & res & " WHERE id_legislador = " & rsTemp.Fields("id")
'    EjecutarSQL (temp)
'    rsTemp.MoveNext
'Wend
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    If FlagBasePrueba Then
        lblModoPrueba.Visible = True
    Else
        lblModoPrueba.Visible = False
    End If
    EsSeleccionDeOrador = False
    lblVersion.Caption = "Versión " & Consola_Version
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Cerrar
    End If
End Sub
Private Sub ValidarVersionBancas()
    'control de consistencia de versiones de huellas
    Dim rsBancas As ADODB.Recordset
    Dim strSQLLegisladores As String
    
    Dim nError_datos_sqv, nError_datos_banca As Integer
    Dim strError As String
    
    Call SetVersion_datos_sqv
    
    
    
    Set rsBancas = New ADODB.Recordset
    'comparo los registros de bancas con la version de version_datos_sqv en config
    SetearRsW "SELECT count(BancaNumero) as cuantos from bancasip WHERE version_datos_sqv<>'" & strVersion_datos_sqv & "'", rsBancas
    If Not rsBancas.EOF Then
        rsBancas.MoveFirst
        nError_datos_sqv = rsBancas.Fields("cuantos")
    Else
        nError_datos_sqv = 0
    End If
    
    'comparo para cada registro de banca bancasip."version_datos_banca" = bancasip."Version"
    SetearRsW "SELECT BancaNumero,version_datos_banca,version from bancasip order by BancaNumero", rsBancas
    If Not rsBancas.EOF Then
        rsBancas.MoveFirst
        nError_datos_banca = 0
        While Not rsBancas.EOF
            'If rsBancas.Fields("bancanumero") > 0 And Trim(rsBancas.Fields("version_datos_banca")) <> Trim(rsBancas.Fields("version")) Then
            If rsBancas.Fields("bancanumero") > 0 Then
                If InStr(Trim(rsBancas.Fields("version_datos_banca")), "ERROR") > 0 Or Trim(rsBancas.Fields("version_datos_banca")) <> Trim(rsBancas.Fields("version")) Then
                    nError_datos_banca = nError_datos_banca + 1
                End If
            End If
            rsBancas.MoveNext
        Wend
    Else
        nError_datos_banca = 0
    End If
    rsBancas.Close
    Set rsBancas = Nothing
    If nError_datos_banca > 0 Or nError_datos_sqv > 0 Then
        strError = "Existen errores de integridad de datos de bancas."
        If nError_datos_banca > 0 Then
            strError = strError & Chr(13) & "Bancas desincronizadas: " & nError_datos_banca
        End If
        If nError_datos_sqv > 0 Then
            strError = strError & Chr(13) & "Datos obsoletos: " & nError_datos_sqv
        End If
        strError = strError & Chr(13) & Chr(13) & "¿Desea ver el listado de bancas con error?"
        If MsgBox(strError, vbQuestion + vbYesNo, "Error de inegridad de datos.") = vbYes Then
            'si elige que si abro el listado de bancas con error
            frmConfigurarUnidadBanca.Show vbModal
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo TrapError
    Dim strSql As String
    Dim RsTemp As ADODB.Recordset
    Dim strNewDir As String
    EntroAMenu = True
    ' ------------------------------------------------------------------------------------------------
    ' directorio de archivos e imagenes de legiladores y archivo de exportacion/importacion .MDB
    ' ------------------------------------------------------------------------------------------------
    Set RsTemp = New ADODB.Recordset
    strSql = "SELECT top 1 directorio_enrolamiento,archivo_enrolamiento From config "
    SetearRs strSql, RsTemp
    If Not RsTemp.EOF Then
        RsTemp.MoveFirst
        strDirectorioEnrolamiento = Trim(RsTemp.Fields("directorio_enrolamiento").Value)
        strArchivoEnrolamiento = Trim(RsTemp.Fields("archivo_enrolamiento").Value)
    Else
        strDirectorioEnrolamiento = "c:\temp\"
        strArchivoEnrolamiento = "bdExportEnrolamiento.mdb"
    End If
    'se fija si existen los directorios y sino los crea
    'crear directorios
    If False Then
        strNewDir = strDirectorioEnrolamiento
        If Trim(dir(strDirectorioEnrolamiento, vbDirectory)) = "" Then
            MkDir (strNewDir)
        End If
        strNewDir = strDirectorioEnrolamiento & "RESGUARDO"
        If Trim(dir(strNewDir, vbDirectory)) = "" Then
           MkDir (strNewDir)
        End If
        strNewDir = strDirectorioEnrolamiento & "IMAGENES"
        If Trim(dir(strNewDir, vbDirectory)) = "" Then
           MkDir (strNewDir)
        End If
    End If
    
    RsTemp.Close
    Set RsTemp = Nothing
    ' ------------------------------------------------------------------------------------------------
    
    getIP
    establecerPermisos
    Call CentrarTodo
    'control de consistencia de versiones de huellas
    Call ValidarVersionBancas
    '-----------------------------------------------
    DoEvents
    Dim f As New frmCargaImagenes
    f.Show vbModal
    Exit Sub
TrapError:
    Select Case err.Number
        Case 6
            MsgBox "Error N° " & err.Number & Chr(10) & err.Description & " Originado en " & err.Source
            End
        Case 52
            MsgBox "Error N° " & err.Number & Chr(10) & err.Description & " Originado en " & err.Source & vbCrLf & "Verifique que esté conectado a la red."
            End
        Case Else
            If MsgBox("Error N° " & err.Number & Chr(10) & err.Description & "Originado en " & err.Source & vbCrLf & "Verifique que haya ingresado a la red del servidor SQV. " & vbCrLf & "¿Desea editar la configuración?", vbQuestion + vbYesNo, "Acceso a servidor SQV") = vbYes Then
                frmSetearConfig.Show vbModal
            Else
                'If MsgBox("Error N° " & Err.Number & Chr(10) & Err.Description & "Originado en " & Err.Source & vbCrLf & "Verifique que haya ingresado a la red del servidor SQV. ¿Desea reintentar?", vbQuestion + vbYesNo, "Acceso a servidor SQV") Then Resume
                'End
                'REVISAR SI ESTO HACE FALTA 14FEB
            End If
    End Select
End Sub
Private Sub CentrarTodo()
'    With frmCenter
'        .Left = (Screen.Width - .Width) / 2
'        .Top = (Screen.Height - .Height) / 2
'    End With
End Sub
Private Sub establecerPermisos()
    '0 - Administrador
    '1 - Administrador bancas
    '2 - Operador avanzado
    '3 - Operador básico
    '4 - Operador consulta
    Select Case gTipoUsuario
        Case 4
            'cmdCartel.Enabled = False
            cmdPresidente.Enabled = False
            ' cmdTiron.Enabled = False
            cmdConfig.Enabled = False
            'cmdImpresoras.Enabled = False
            cmdPeriodos.Enabled = False
            cmdConsultas.Enabled = True
        Case 3, 2
            'cmdCartel.Enabled = False
            cmdPresidente.Enabled = True
            ' cmdTiron.Enabled = True
            cmdConfig.Enabled = False
            'cmdImpresoras.Enabled = False
            cmdPeriodos.Enabled = True
            cmdConsultas.Enabled = True
        Case 1
            'cmdCartel.Enabled = False
            cmdPresidente.Enabled = False
            ' cmdTiron.Enabled = False
            cmdConfig.Enabled = True
            'cmdImpresoras.Enabled = False
            cmdPeriodos.Enabled = False
            cmdConsultas.Enabled = True
        Case 0
            'cmdCartel.Enabled = True
            cmdPresidente.Enabled = True
            ' cmdTiron.Enabled = True
            cmdConfig.Enabled = True
            'cmdImpresoras.Enabled = True
            cmdPeriodos.Enabled = True
            cmdConsultas.Enabled = True
    End Select
End Sub
Private Sub getIP()
    Dim ip As String
    ip = WSock.LocalIP
    Datos.establecerIP ip
End Sub
Private Sub RecuperarBDLeg(ByVal nUltimoIndice As Long, ByRef strBDLegisladores As String, ByRef strlogleg As String)
'    ReDim ListaLegisladores((nUltimoIndice))
'    Dim i             As Long
'    Dim xIndice       As String
'    Dim xId           As String
'    Dim xApellido     As String
'    Dim xNombre       As String
'    Dim xValida       As String
'    Dim xIndiceALeer  As String
'    Dim strRegistro   As String
'    Dim nValidos      As Long
'    Dim nInvalidos    As Long
'    Dim nInexistentes As Long
'    Dim x             As Long
'    Dim strOrigen     As String
'    Dim strDestino    As String
'    Dim strPath       As String
'    Dim strNumeroArch As String
'    strBDLegisladores = ""
'
'
'    With FlHuellas
'        If .ListCount = 0 Then
'            Exit Sub
'        End If
'        .Path = App.Path
'        .Refresh
'        For x = 0 To .ListCount - 1
'            .ListIndex = x
'            xIndiceALeer = Mid(.FileName, 6, 4)
'            Call LeerUnLeg(xIndiceALeer, strRegistro)
'            If Left(strRegistro, 6) <> "SRLEGI" Then
'                xIndice = ""
'                xId = ""
'                xApellido = ""
'                xNombre = ""
'                xValida = "* No encontrado:" & xIndiceALeer
'                nInexistentes = nInexistentes + 1
'            Else
'                xIndice = Mid(strRegistro, 9, 4)
'                xId = Mid(strRegistro, 15, 8)
'                xApellido = HexATexto(Mid(strRegistro, 23, 60), 30)
'                'If LCase(Trim(xApellido)) = "gastaldi" Then
'                '    Stop
'                'End If
'                xNombre = HexATexto(Mid(strRegistro, 83, 60), 30)
'                If xIndice = xIndiceALeer Then
'                    xValida = "     "
'                    nValidos = nValidos + 1
'                Else
'                    xValida = "*" & xIndiceALeer
'                    nInvalidos = nInvalidos + 1
'                End If
'                strBDLegisladores = strBDLegisladores & strRegistro
'            End If
'            ListaLegisladores(i) = xIndice & ";" & xValida & ";" & _
'                                    xId & ";" & _
'                                    xApellido & ";" & _
'                                    xNombre
'            strlogleg = Trim(ListaLegisladores(i)) & vbCrLf & strlogleg
'        Next x
'        strBDLegisladores = strBDLegisladores & _
'        "SRLEGI ^" & CerosIzquierda((nUltimoIndice) + 1, 4) & "00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF" & vbCrLf
'        strlogleg = "Fin de lectura de base de datos: " & vbCrLf & " Invalidos (Indice no coincide con el nombre de archivo) " & Str(nInvalidos) & ", Inexistentes : " & Str(nInexistentes) & vbCrLf & ", Validos: " & Str(nValidos) & vbCrLf & " Total Procesado: " & Str(nValidos + nInvalidos + nInexistentes) & vbCrLf & strlogleg
'    End With
End Sub


Private Sub AbrirDBLeg(nLegislador As String)
   Dim strArchivo As String
   xArchivoDBLeg = FreeFile()
   strArchivo = App.Path & "\DBLeg" & CerosIzquierda(Trim((nLegislador)), 4) & ".huella"
   strArchivoHuella = strArchivo
   Open strArchivo For Binary As #xArchivoDBLeg
End Sub
Private Sub CerrarDBLeg()
    Close #xArchivoDBLeg
End Sub
Private Sub InsDBLeg(strContenido As String)
   Put #xArchivoDBLeg, , strContenido
End Sub
Private Sub LeeDBLeg(strContenido As String)
    strContenido = Space(cLongRegDBLeg * 10)
   Get #xArchivoDBLeg, , strContenido
End Sub

Private Function LongFija(strText As String, nLong As Long) As String
    If nLong > Len(strText) Then
        LongFija = Left(strText & Space(nLong - Len(strText)), nLong)
    Else
        LongFija = Left(strText, nLong)
    End If
End Function
Private Function CerosIzquierda(strText As String, nLong As Long) As String
    If nLong > Len(strText) Then
        CerosIzquierda = Left(String(nLong - Len(strText), "0") & strText, nLong)
    Else
        CerosIzquierda = Right(strText, nLong)
    End If
End Function

Private Function TextoAHex(strTexto As String, nLong As Long) As String
    Dim i As Long
    TextoAHex = ""
    For i = 1 To Len(strTexto)
        TextoAHex = TextoAHex & Hex(Asc(Mid(strTexto, i, 1)))
    Next
    TextoAHex = CerosIzquierda(TextoAHex, nLong)
End Function

Private Function HexATexto(strTexto As String, nLong As Long) As String
Dim i As Long
    'convierte un texto que contenga pares hexadecimales codificados como string en un string ascii
    HexATexto = ""
    For i = 1 To Len(strTexto) Step 2
        HexATexto = HexATexto & HexAChr(Mid(strTexto, i, 2))
    Next
    HexATexto = CerosIzquierda(HexATexto, nLong)
End Function
Private Function HexAChr(strHex) As String
    'convierte dos digitos hexadecimales codificados como string en un caracter ascii
    Dim nDecimal As Long
    nDecimal = 0
    nDecimal = DigitoHexADec(Mid(strHex, 1, 1)) * 16
    nDecimal = nDecimal + DigitoHexADec(Mid(strHex, 2, 1))
    HexAChr = Chr(nDecimal)
End Function
Private Function DigitoHexADec(charHex) As Long
    If charHex >= "0" And charHex <= "9" Then
        DigitoHexADec = Asc(charHex) - 48
    Else
        DigitoHexADec = Asc(charHex) - 55
    End If
End Function
Private Function ExisteLeg(xIndice As String) As Boolean
    ExisteLeg = Left(LeerUnLeg(xIndice), 6) = "SRLEGI"
End Function
Private Function LeerUnLeg(xIndice As String, Optional strRegistro As String)
    'Dim strRegistro As String
    
    'guarda todo lo que haya en la ventana dbleg menos la linea de fin
    Call AbrirDBLeg((xIndice))
    Call LeeDBLeg(strRegistro)
    Call CerrarDBLeg
    LeerUnLeg = strRegistro
End Function
Private Sub Form_Unload(Cancel As Integer)
frmLogin.SetDefault
End Sub
