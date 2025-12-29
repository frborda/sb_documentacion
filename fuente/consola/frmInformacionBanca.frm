VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmInformacionBanca 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información de la banca seleccionada"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7800
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdDeshabilitarBanca 
      Height          =   360
      Left            =   6150
      TabIndex        =   21
      Top             =   490
      Width           =   1600
      _ExtentX        =   2831
      _ExtentY        =   635
      BackColor       =   12230304
      Caption         =   "Deshabilitar Banca"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdPruebaScanNueva 
      Height          =   375
      Left            =   2300
      TabIndex        =   20
      Top             =   3360
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BackColor       =   49152
      Caption         =   "Prueba de Scan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin VB.CommandButton cmdLimpiarIdentificacion 
      Caption         =   "Limpiar Identificación"
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      Top             =   2880
      Width           =   1800
   End
   Begin VB.CommandButton cmdHistoricoBanca 
      Caption         =   "Ver Historico"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Muestra el historial de la presente banca"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdAsignarId 
      Caption         =   "A&signar identificac."
      Height          =   375
      Left            =   5940
      TabIndex        =   6
      Top             =   2880
      Width           =   1700
   End
   Begin VB.CommandButton cmdAsignarVoto 
      Caption         =   "&Asignar voto"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4230
      TabIndex        =   5
      Top             =   2880
      Width           =   1700
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar prueba"
      Height          =   375
      Left            =   4400
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "&Prueba de Scan"
      Height          =   375
      Left            =   2700
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdReinicio 
      Caption         =   "&Reinicio Banca"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4710
      TabIndex        =   2
      Top             =   480
      Width           =   1400
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "C&errar"
      Height          =   375
      Left            =   6060
      TabIndex        =   8
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Image imgFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   60
      Stretch         =   -1  'True
      Top             =   60
      Width           =   3015
   End
   Begin VB.Label lblNN 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1560
      TabIndex        =   17
      Top             =   2940
      Width           =   700
   End
   Begin VB.Label Label4 
      Caption         =   "Leg. sin identificar:"
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   2940
      Width           =   1395
   End
   Begin VB.Label lblLocalidad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4320
      TabIndex        =   15
      Top             =   2460
      Width           =   3315
   End
   Begin VB.Label lblLocalidadEtiqueta 
      Alignment       =   1  'Right Justify
      Caption         =   "Provincia : "
      Height          =   315
      Left            =   3210
      TabIndex        =   14
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Label lblBloque 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4320
      TabIndex        =   13
      Top             =   2100
      Width           =   3315
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Bloque : "
      Height          =   315
      Left            =   3210
      TabIndex        =   12
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label lblAgrupacion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4320
      TabIndex        =   11
      Top             =   1740
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label lblAgrupacionEtiqueta 
      Alignment       =   1  'Right Justify
      Caption         =   "Agrupación : "
      Height          =   315
      Left            =   3210
      TabIndex        =   10
      Top             =   1740
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblLegislador 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4320
      TabIndex        =   9
      Top             =   1380
      Width           =   3315
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Sr. legislador : "
      Height          =   315
      Left            =   3210
      TabIndex        =   7
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label lblBanca 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4320
      TabIndex        =   1
      Top             =   60
      Width           =   3315
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Banca : "
      Height          =   315
      Left            =   3180
      TabIndex        =   0
      Top             =   60
      Width           =   975
   End
End
Attribute VB_Name = "frmInformacionBanca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mNumeroBanca      As Long
Private mfoto             As ADODB.Stream
Private dejarCerrar       As Boolean
Private dejarProbar       As Boolean
Private dejarIdentificar  As Boolean
Private dejarCambiarVoto  As Boolean
Private dejarAbstener     As Boolean
Private permitirOperar    As Boolean
Private strEstadoVotacion As String
Public ModoPruebaScan   As Boolean
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Property Let EstadoVotacion(ByVal vNewValue As Variant)
    strEstadoVotacion = vNewValue
    strEstadoVotacion = LCase(Trim(strEstadoVotacion))
End Property
Private Sub cmdAsignarId_Click()
    Dim asignar As New frmAsignarLegislador
    asignar.mostrarLegisladores Val(lblBanca.Caption)
    asignar.Show vbModal
    Set asignar = Nothing
    Unload Me
End Sub
Private Sub cmdAsignarVoto_Click()
    Dim voto As New frmDefinirVoto
    voto.Banca = Val(lblBanca.Caption)
    voto.Show vbModal
    Set voto = Nothing
    Unload Me
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Public Property Let NroBanca(ByVal vNewValue As Variant)
    mNumeroBanca = vNewValue
    If mNumeroBanca <> 0 Then
        lblBanca.Caption = mNumeroBanca
    Else
        lblBanca.Caption = "Presidente"
    End If
End Property

Public Sub MostrarDatos(pNumeroBanca As Integer, Optional pIdLegislador As String, Optional pPermitirPrueba As Boolean = False, Optional pPermitirIdentificar As Boolean = True, Optional pCambiarVoto As Boolean = False, Optional pdejarAbstener As Boolean = False)
    NroBanca = pNumeroBanca
    If IsNull(pIdLegislador) = False Then
        Dim rstAux As New ADODB.Recordset
        SetearRs "SELECT nombre,apellido,grupo_politico,bloque_politico,PICTURE,distritos.distrito AS Provincia FROM Legisladores LEFT OUTER JOIN distritos ON Legisladores.distrito = distritos.id_distrito WHERE id='" & pIdLegislador & "'", rstAux
        If rstAux.EOF = False Then
            cargarDatos rstAux
        End If
    End If
    If pPermitirPrueba = False Then
        cmdScan.Enabled = False
        cmdCancelar.Enabled = False
        dejarProbar = False
    Else
        dejarProbar = True
    End If
    If pPermitirIdentificar = False Then
        cmdAsignarId.Enabled = False
        dejarIdentificar = False
    Else
        cmdAsignarId.Enabled = True
        dejarIdentificar = True
    End If
    If pdejarAbstener = False Then
        dejarAbstener = False
        'cmdAbstener.Enabled = False
    Else
        dejarAbstener = True
        'cmdAbstener.Enabled = True
    End If
    If pCambiarVoto = False Then
        dejarCambiarVoto = False
        cmdAsignarVoto.Enabled = False
    Else
        dejarCambiarVoto = True
        cmdAsignarVoto.Enabled = True
    End If
    'cmdVerificar.Enabled = True
    cmdReinicio.Enabled = True
    ' restricciones usuario
    ' Parche temporal... cambiar
    ' If (gTipoUsuario <> 3) And (gTipoUsuario <> 4) Then
    If (gTipoUsuario <> 3) And (gTipoUsuario <> 4) And (gTipoUsuario <> 0) Then
        permitirOperar = False
        'cmdVerificar.Enabled = False
        cmdReinicio.Enabled = False
        cmdScan.Enabled = False
        cmdAsignarVoto.Enabled = False
        cmdAsignarId.Enabled = False
        cmdCerrar.Enabled = True
        'cmdAbstener = False
        dejarCerrar = True
    Else
        permitirOperar = True
    End If
    
End Sub
Private Sub DeterminarLeyendaBotonAbstencion()
    Dim strSql As String
    Dim RsAbs  As ADODB.Recordset
    Set RsAbs = New ADODB.Recordset
    Dim strVectorResult() As String
    
    strSql = "SELECT Vector_resultado FROM Vector"
    SetearRs strSql, RsAbs
    strVectorResult = Split(RsAbs.Fields("Vector_resultado").Value, ";")
    RsAbs.Close
    Set RsAbs = Nothing
    If strVectorResult(mNumeroBanca) = ABSTENCION_AUTORIZADA Then
        'cmdAbstener.Caption = "Cancelar A&bstención"
    Else
        'cmdAbstener.Caption = "&Abstener"
    End If
    
End Sub

Private Sub cargarDatos(pRst As ADODB.Recordset)
    Dim strImagen As String
    If (IsNull(pRst!Apellido) = False) And (IsNull(pRst!Nombre) = False) Then
        lblLegislador.Caption = Trim(pRst!Apellido) & ", " & Trim(pRst!Nombre)
    End If
    If IsNull(pRst!grupo_politico) = False Then
        lblAgrupacion.Caption = Trim(pRst!grupo_politico)
    End If
    If IsNull(pRst!bloque_politico) = False Then
        lblBloque.Caption = Trim(pRst!bloque_politico)
    End If
    If IsNull(pRst!Provincia) = False Then
        lblLocalidad.Caption = Trim(pRst!Provincia)
    End If
    If IsNull(pRst!Picture) = False Then
        Set mfoto = New ADODB.Stream
        mfoto.Type = adTypeBinary
        mfoto.Open
        mfoto.Write pRst!Picture
        mfoto.SaveToFile App.Path & "\foto.gif", adSaveCreateOverWrite
        imgFoto.Picture = LoadPicture(App.Path & "\foto.gif")
    End If
End Sub
Private Sub cmdDeshabilitarBanca_Click()
Datos.GrabarMensaje "banca?deshabilitar", Trim(Str(mNumeroBanca)), , True
Unload Me
End Sub
Private Sub cmdHistoricoBanca_Click()
    frmHistoricoBanca.Banca = Trim(lblBanca.Caption)
    frmHistoricoBanca.Show 1, Me
End Sub
Private Sub cmdLimpiarIdentificacion_Click()
Datos.GrabarMensaje "limpieza_individual", Trim(Str(mNumeroBanca)), , True
frmConsolaOperacion.MensajeEsperado = MensajeVacio
Unload Me
End Sub
Private Sub cmdPruebaScanNueva_Click()
Datos.GrabarMensaje "scan?prueba", Trim(Str(mNumeroBanca)), , True
frmNuevoScan.Titulo = "Prueba de Scan | BANCA "
frmNuevoScan.Banca = Trim(Str(mNumeroBanca))
frmNuevoScan.Show vbModal, Me
End Sub
Private Sub cmdReinicio_Click()
    MensajesSQV.reiniciarBanca lblBanca.Caption
    Unload Me
End Sub

Private Sub cmdScan_Click()
Dim nTicks As Long
nTicks = GetTickCount
While GetTickCount - nTicks < 2000
    DoEvents
Wend
cmdScan.Enabled = False
MensajesSQV.PruebaScan Trim(Str(mNumeroBanca))
lblLegislador.Caption = ""
ControlesHabilitados = False
Unload Me
End Sub
Private Sub cmdCancelar_Click()
    MensajesSQV.PruebaScanFin Str(mNumeroBanca)
    ControlesHabilitados = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    SetearRs "SELECT habilitada FROM BancasDeshabilitadas WHERE banca = " & Trim(Str(mNumeroBanca)), RsTemp
    If RsTemp.EOF Then
        MsgBox ("Error de integridad respecto al vector de bancas deshabilitadas")
    Else
        If RsTemp.Fields(0) = 1 Then
            cmdDeshabilitarBanca.Caption = "Deshabilitar Banca"
        Else
            cmdDeshabilitarBanca.Caption = "Habilitar Banca"
        End If
    End If
    Me.ModoPruebaScan = False
    'cmdIdentificacionTeclado.Visible = False
    Call DeterminarLeyendaBotonAbstencion
    
    If frmConsolaOperacion.txtOcup.Caption <> "" Then
        lblNN.Caption = frmConsolaOperacion.txtOcup.Caption
    End If
    If flPresidenteLegislador = True Then     ' Si es el vicegobernador
        If (strEstadoVotacion = "votando" Or strEstadoVotacion = "larga") Then
            cmdAsignarVoto.Enabled = True
        Else
            cmdAsignarVoto.Enabled = False
        End If
    ElseIf flPresidenteLegislador = False Then ' Si el presidente es legislador
        If strEstadoVotacion = "empate" Then
            cmdAsignarVoto.Enabled = True      ' Vota en caso de empate
        Else
            cmdAsignarVoto.Enabled = False
        End If
    End If
    dejarCerrar = True
'    cmdHardReset.Enabled = True
    Call HabilitarControles
    If frmConsolaOperacion.dcTipoOperacion.BoundText = "quorum" And frmConsolaOperacion.cmdModoNominal.Caption = "Habilitar identificación" Then
        cmdScan.Enabled = True
    Else
        cmdScan.Enabled = False
    End If
End Sub
Public Sub LanzaScan()
MensajesSQV.PruebaScan Trim(Str(mNumeroBanca))
End Sub
Private Sub HabilitarControles()
    lblAgrupacionEtiqueta.Visible = AGRUPACION_POLITICA_HABILITADA
    lblAgrupacion.Visible = AGRUPACION_POLITICA_HABILITADA
    lblLocalidadEtiqueta.Visible = DISTRITO_HABILITADO
    lblLocalidad.Visible = DISTRITO_HABILITADO
End Sub

Public Property Let ControlesHabilitados(pModo As Boolean)
    If permitirOperar = True Then
        'cmdVerificar.Enabled = pModo
        cmdReinicio.Enabled = pModo
        'cmdScan.Enabled = dejarProbar
        cmdAsignarVoto.Enabled = dejarCambiarVoto
        cmdAsignarId.Enabled = dejarIdentificar
        cmdCerrar.Enabled = pModo
'        cmdAbstener = dejarAbstener
        dejarCerrar = pModo
 '       cmdHardReset.Enabled = pModo
    Else
        'cmdVerificar.Enabled = False
        cmdReinicio.Enabled = False
        cmdScan.Enabled = False
        cmdAsignarVoto.Enabled = False
        cmdAsignarId.Enabled = False
        cmdCerrar.Enabled = True
        'cmdAbstener = False
        dejarCerrar = True
'        cmdHardReset.Enabled = True
    End If
End Property
Public Sub CancelaAuth()
Datos.GrabarMensaje "limpiaridpruebascan", Trim(Str(mNumeroBanca)), , True
Datos.GrabarMensaje "pruebascanlimpiar", frmConsolaOperacion.info.lblBanca.Caption, "", True
frmConsolaOperacion.MensajeEsperado = MensajeVacio
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If dejarCerrar = False Then
        Cancel = True
    Else
        MensajesSQV.PruebaScanFin Str(mNumeroBanca)
        ControlesHabilitados = True
        If ModoPruebaScan = True Then
            Datos.GrabarMensaje "limpiaridpruebascan", Trim(Str(mNumeroBanca)), , True
            frmConsolaOperacion.MensajeEsperado = MensajeVacio
        End If
    End If
End Sub
Private Sub lblLegislador_Change()
If lblLegislador.Caption = "Identificación negativa" Then
    Dim nTicks As Long
    nTicks = GetTickCount
    While GetTickCount - nTicks < 4000
        DoEvents
    Wend
    SendKeys "%" & Chr(vbKeyP)
End If
End Sub

