VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmMantenimiento 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menú de mantenimiento de la consola"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdReinicio 
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "&Reiniciar bancas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Impresión automática"
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
      Height          =   855
      Left            =   150
      TabIndex        =   11
      Top             =   3600
      Width           =   7575
      Begin VB.CheckBox chkVistaPrevia 
         BackColor       =   &H00404040&
         Caption         =   "Habilitar Vista Previa"
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
         Height          =   255
         Left            =   3360
         TabIndex        =   12
         Top             =   360
         Width           =   2535
      End
      Begin Proyecto1.ButtonOffice cmdImpresionAutomatica 
         Height          =   435
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   767
         BackColor       =   12230304
         Caption         =   "Habilitar Impresión Automática"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
   End
   Begin MSDataListLib.DataCombo dcAbstencion 
      Height          =   345
      Left            =   2850
      TabIndex        =   4
      Top             =   5220
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   609
      _Version        =   393216
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcTipoQuorum 
      Height          =   345
      Left            =   2850
      TabIndex        =   0
      Top             =   4800
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   609
      _Version        =   393216
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frameParametros 
      BackColor       =   &H00404040&
      Caption         =   "Parametros de operación"
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
      Height          =   1215
      Left            =   150
      TabIndex        =   5
      Top             =   4500
      Width           =   7635
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Abstención"
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
         Height          =   285
         Left            =   1500
         TabIndex        =   7
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Quorum Tipo"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin Proyecto1.ButtonOffice cmdModoMantenimiento 
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   660
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "&Modo Mantenimiento"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdCancelarId 
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "Limpiar identificaciones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdConfiguracionBancas 
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   1740
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "Configuracion Bancas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdTiempo 
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "Tiemp&o de Votación"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdVolver 
      Height          =   495
      Left            =   150
      TabIndex        =   19
      Top             =   5820
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "&Ventana anterior"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdProgreso 
      Height          =   705
      Left            =   120
      TabIndex        =   20
      Top             =   2820
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1244
      BackColor       =   12230304
      Caption         =   "Mostrar Progreso de Votación en el Cartel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Permite configurar el modo de visualización del progreso de la votación"
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
      Height          =   495
      Left            =   2940
      TabIndex        =   21
      Top             =   2940
      Width           =   4635
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Permite configurar el tiempo de Votación"
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
      Height          =   225
      Left            =   2940
      TabIndex        =   10
      Top             =   2400
      Width           =   4635
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Libera a TODAS las terminales de sus respectivas identificaciones"
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
      Height          =   405
      Left            =   2940
      TabIndex        =   9
      Top             =   1260
      Width           =   4755
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Permite configurar las bancas y sincronizarlas con los datos de Legisladores."
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
      Height          =   405
      Left            =   2940
      TabIndex        =   8
      Top             =   1800
      Width           =   4635
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Alterna entre modo Mantenimiento de bancas (secuencia de prueba) o modo normal"
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
      Height          =   435
      Left            =   2940
      TabIndex        =   3
      Top             =   660
      Width           =   4635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cierra esta ventana y vuelve a la ventana anterior."
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
      Height          =   225
      Left            =   2970
      TabIndex        =   2
      Top             =   5940
      Width           =   4635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Permite reiniciar todas las bancas del recinto."
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
      Index           =   0
      Left            =   2940
      TabIndex        =   1
      Top             =   210
      Width           =   4635
   End
End
Attribute VB_Name = "frmMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstAbstencion             As New ADODB.Recordset
Private rstTipoQuorum             As New ADODB.Recordset
Private xInicializando As Boolean
Private Ignorar As Boolean

Private Sub chkVistaPrevia_Click()
If chkVistaPrevia.Value = vbChecked Then
    VistaPrevia = True
Else
    VistaPrevia = False
End If
End Sub

Private Sub cmdCancelarId_Click()
Dim r As Integer
r = MsgBox("¿Está seguro de que desea limpiar la identificación de TODAS las bancas?", vbYesNo, "ALERTA")
If r = vbYes Then
    If (gTipoUsuario = 0) Or (gTipoUsuario = 2) Then
        Datos.GrabarMensaje "cambio?cancelarids", " ", , True
    Else
        MsgBox "Ud. no dispone de permisos para realizar esta acción.", vbInformation + vbOKOnly
    End If
    Unload Me
End If
End Sub

Private Sub cmdImpresionAutomatica_Click()
Datos.GrabarMensaje "cambio?listar", " ", , True
Unload Me
End Sub

Private Sub cmdIniciarSQV_Click()
     Call IniciarSQVporBandera
'    Dim strPathToProgram As String
'On Error GoTo TrapError:
'    strPathToProgram = Environ("sqvinicio")
'    If strPathToProgram = "" Then strPathToProgram = "\\Siprevo\Sistemas\siprevo\inicioSQV"
'    strPathToProgram = strPathToProgram & "\prender.bat"
'    Shell ("" & strPathToProgram & "")
'    Unload Me
'TrapError:
    Unload Me
End Sub

Private Sub cmdModoMantenimiento_Click()
    MensajesSQV.ModoMantenimiento
    Unload Me
End Sub
Private Sub cmdNormalMantenimiento_Click()
    MensajesSQV.ModoNormalMantenimiento
    Unload Me
End Sub
Private Sub cmdConfiguracionBancas_Click()
    frmConfigurarUnidadBanca.Show vbModal
    Unload Me
End Sub

Private Sub cmdPresenciaConIdentificacion_Click()
    MensajesSQV.PresenciaConIdentificacion
    Unload Me
End Sub
Private Sub cmdProgreso_Click()
If cmdProgreso.Caption = "Mostrar Progreso de Votación en el Cartel" Then
    EjecutarSQL ("UPDATE config SET Mostrar_Progreso_Votacion = 1")
    cmdProgreso.Caption = "Ocultar Progreso de Votación en el Cartel"
Else
    EjecutarSQL ("UPDATE config SET Mostrar_Progreso_Votacion = 0")
    cmdProgreso.Caption = "Mostrar Progreso de Votación en el Cartel"
End If
End Sub
Private Sub cmdReinicio_Click()
    Datos.GrabarMensaje "estadoioc", "", "brc", True
    Unload Me
End Sub

Private Sub cmdSalirSqv_Click()
    MensajesSQV.SalirSQV
    Unload Me
End Sub

Private Sub cmdTiempo_Click()
frmTiempoVotacion.Show vbModal
End Sub
Private Sub cmdVolver_Click()
    frmMantenimiento.Visible = False
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim Consulta As String
    Dim rsTemp As ADODB.Recordset
    Dim rsProgreso As ADODB.Recordset
    Set rsProgreso = New ADODB.Recordset
    SetearRs "SELECT Mostrar_Progreso_Votacion FROM config", rsProgreso
    If rsProgreso.EOF Then
        MsgBox "Error de integridad de tabla CONFIG!"
    Else
        If IsNull(rsProgreso.Fields(0)) Then
            cmdProgreso.Caption = "Mostrar Progreso de Votación en el Cartel"
        Else
            If rsProgreso.Fields(0) = 1 Then
                cmdProgreso.Caption = "Ocultar Progreso de Votación en el Cartel"
            Else
                cmdProgreso.Caption = "Mostrar Progreso de Votación en el Cartel"
            End If
        End If
    End If
    If VistaPrevia = False Then
        chkVistaPrevia.Value = vbUnchecked
    Else
        chkVistaPrevia.Value = vbChecked
    End If
    xInicializando = True
    
    establecerPermisos
    
    strSql = "SELECT rtrim(Tipo_de_Abstención) as tipo, Descripcion From modabs ORDER BY Descripcion"
    SetearRs strSql, rstAbstencion
    With dcAbstencion
        Set .RowSource = rstAbstencion
        .ListField = "Descripcion"
        .BoundColumn = "tipo"
    End With
    
    strSql = "SELECT rtrim(codigo) as codigo, Descripcion From TipoMayoriaQuorum ORDER BY Descripcion"
    SetearRs strSql, rstTipoQuorum
    With dcTipoQuorum
        Set .RowSource = rstTipoQuorum
        .ListField = "Descripcion"
        .BoundColumn = "Codigo"
    End With

    If Trim(dcAbstencion.Tag) <> Trim(frmConsolaOperacion.dcAbstencion.Tag) Then
        dcAbstencion.Tag = Trim(frmConsolaOperacion.dcAbstencion.Tag)
        dcAbstencion.BoundText = Trim(frmConsolaOperacion.dcAbstencion.Tag)
    End If
    If Trim(dcTipoQuorum.Tag) <> Trim(frmConsolaOperacion.dcTipoQuorum.Tag) Then
        dcTipoQuorum.Tag = Trim(frmConsolaOperacion.dcTipoQuorum.Tag)
        dcTipoQuorum.BoundText = Trim(frmConsolaOperacion.dcTipoQuorum.Tag)
    End If
    xInicializando = False
    Set rsTemp = New ADODB.Recordset
    SetearRs "SELECT Listar_automaticamente FROM vector", rsTemp
    Ignorar = False
    If rsTemp.Fields(0) = 0 Then
        cmdImpresionAutomatica.Caption = "Habilitar Impresión Automática"
    Else
        cmdImpresionAutomatica.Caption = "Deshabilitar Impresión Automática"
    End If
    If Trim(frmConsolaOperacion.txtTitulo.Text) = "MANTENIMIENTO DEL SISTEMA SQV" Then
        cmdModoMantenimiento.Caption = "Modo Normal"
    End If
End Sub
Private Sub establecerPermisos()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Libero rstTipo abstencion
    If rstAbstencion.State = adStateOpen Then
        rstAbstencion.Close
    End If
    Set rstAbstencion = Nothing
    'Libero rstTipo Quorum
    If rstTipoQuorum.State = adStateOpen Then
        rstTipoQuorum.Close
    End If
    Set rstTipoQuorum = Nothing
    PrimeraVezMantenimiento = True
End Sub
Private Sub ResetHard_Click()
    MensajesSQV.reiniciarBancaHard "brc"
    Unload Me
End Sub

Private Sub dcTipoQuorum_Change()
    If Not xInicializando Then
        MensajesSQV.cambiarTipoQuorum dcTipoQuorum.BoundText
        Unload Me
    End If

End Sub

Private Sub dcAbstencion_Change()
    If Not xInicializando Then
        MensajesSQV.ModoVotacion dcAbstencion.BoundText
        Unload Me
    End If
End Sub



