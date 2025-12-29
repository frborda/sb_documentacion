VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form FrmConfigurarConsola 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuraciones"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10260
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdBancas 
      Height          =   825
      Left            =   240
      TabIndex        =   11
      Top             =   150
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   1455
      BackColor       =   16576
      Caption         =   "Bancas"
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
      TabIndex        =   5
      Top             =   3000
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "&Volver al Menú"
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
   Begin Proyecto1.ButtonOffice cmdUsuarios 
      Height          =   825
      Left            =   3510
      TabIndex        =   12
      Top             =   150
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   1455
      BackColor       =   8421504
      Caption         =   "Usuarios"
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
   Begin Proyecto1.ButtonOffice cmdOtros 
      Height          =   825
      Left            =   6780
      TabIndex        =   13
      Top             =   150
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   1455
      BackColor       =   8421504
      Caption         =   "Otros"
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
   Begin VB.Frame frmBancas 
      BackColor       =   &H00404040&
      Height          =   2955
      Left            =   180
      TabIndex        =   6
      Top             =   -30
      Width           =   9915
      Begin Proyecto1.ButtonOffice cmdConfigurarUnidadBanca 
         Height          =   495
         Left            =   90
         TabIndex        =   7
         Top             =   1500
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   873
         BackColor       =   12230304
         Caption         =   "Cambiar IPs de Bancas"
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
      Begin Proyecto1.ButtonOffice cmdBancasProbables 
         Height          =   495
         Left            =   90
         TabIndex        =   8
         Top             =   2070
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   873
         BackColor       =   12230304
         Caption         =   "Banca Probable"
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
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   1740
         X2              =   7860
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Permite cambiar la IP de una o varias bancas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   2400
         TabIndex        =   10
         Top             =   1650
         UseMnemonic     =   0   'False
         Width           =   7005
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Asignacion de banca probable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2370
         TabIndex        =   9
         Top             =   2190
         Width           =   4635
      End
   End
   Begin VB.Frame frameUsuarios 
      BackColor       =   &H00404040&
      Height          =   2955
      Left            =   180
      TabIndex        =   14
      Top             =   -30
      Width           =   9915
      Begin Proyecto1.ButtonOffice Command3 
         Height          =   495
         Left            =   90
         TabIndex        =   15
         Top             =   2070
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   873
         BackColor       =   12230304
         Caption         =   "Usuarios"
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
      Begin Proyecto1.ButtonOffice Command4 
         Height          =   495
         Left            =   90
         TabIndex        =   16
         Top             =   1500
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   873
         BackColor       =   12230304
         Caption         =   "Cambiar Palabra Secreta"
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
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   1740
         X2              =   7860
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Administración de Usuarios del sistema"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2370
         TabIndex        =   18
         Top             =   1590
         Width           =   6885
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Actualizar palabra secreta de usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   2400
         TabIndex        =   17
         Top             =   2160
         Width           =   4635
      End
   End
   Begin VB.Frame frmOtros 
      BackColor       =   &H00404040&
      Height          =   2955
      Left            =   180
      TabIndex        =   19
      Top             =   -30
      Width           =   9915
      Begin Proyecto1.ButtonOffice cmdConexion 
         Height          =   495
         Left            =   90
         TabIndex        =   20
         Top             =   1500
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   873
         BackColor       =   12230304
         Caption         =   "&Acceso a datos"
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
      Begin Proyecto1.ButtonOffice Command2 
         Height          =   495
         Left            =   90
         TabIndex        =   21
         Top             =   2070
         Visible         =   0   'False
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   873
         BackColor       =   12230304
         Caption         =   "Valores de Configuración"
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
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         X1              =   1740
         X2              =   7860
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Permite definir los parámetros de la conexión con la Base de datos que uilizará la consola."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   2400
         TabIndex        =   23
         Top             =   1530
         Width           =   4635
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Permite setear valores de configuracion para operacion de SQV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   2400
         TabIndex        =   22
         Top             =   2190
         Visible         =   0   'False
         Width           =   7455
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Configuración de orden y cantidad de copias fuera del recinto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   555
      Left            =   11610
      TabIndex        =   4
      Top             =   7830
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Configuración de orden y cantidad de copias para el recinto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   555
      Left            =   11190
      TabIndex        =   3
      Top             =   7170
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.Label lblDistritos 
      BackStyle       =   0  'Transparent
      Caption         =   "Formulario de mantenimiento de datos de Distritos Electorales. Altas, Bajas y Modificaciones de Distritos Electorales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   11280
      TabIndex        =   2
      Top             =   6390
      Visible         =   0   'False
      Width           =   7515
   End
   Begin VB.Label lblSecciones 
      BackStyle       =   0  'Transparent
      Caption         =   "Formulario de mantenimiento de datos de Secciones Electorales. Altas, Bajas y Modificaciones de  Secciones Electorales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   525
      Left            =   11280
      TabIndex        =   1
      Top             =   5790
      Visible         =   0   'False
      Width           =   7485
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cierra esta ventana y vuelve al Menú Principal del Sistema."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   2430
      TabIndex        =   0
      Top             =   3120
      Width           =   6885
   End
End
Attribute VB_Name = "FrmConfigurarConsola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim optBancas As Boolean
Dim optOtros As Boolean
Dim optUsuarios As Boolean
Option Explicit

Private Sub cmdABM_Bloques_Click()
    frmABMBloques.Show 1
End Sub

Private Sub cmdABM_Secciones_Click()
    frmABMSecciones.Show vbModal
End Sub

Private Sub cmdABMDistritos_Click()
     frmABMDistritos.Show vbModal
End Sub

Private Sub cmdABMMandatos_Click()
    frmABMMandatos.Show vbModal
End Sub

Private Sub ButtonOffice1_Click()

End Sub
Private Sub cmdBancas_Click()
optBancas = True
optUsuarios = False
optOtros = False
ActualizarColores
End Sub

Private Sub cmdBancasProbables_Click()
frmBancasProbables.Show vbModal
End Sub
Private Sub cmdConexion_Click()
    frmConfig.Show vbModal
End Sub
Private Sub cmdConfigImpresion_Click()
Dim xForm As New frmConfigImpresionAutomatica
xForm.DentroRecinto = False
xForm.Show vbModal
Set xForm = Nothing
End Sub
Private Sub cmdConfigImpresionR_Click()
Dim xForm As New frmConfigImpresionAutomatica
xForm.DentroRecinto = True
xForm.Show vbModal
Set xForm = Nothing
End Sub
Private Sub cmdConfigurarUnidadBanca_Click()
    frmConfigurarUnidadBanca.Show vbModal
End Sub

Private Sub cmdOtros_Click()
optBancas = False
optUsuarios = False
optOtros = True
ActualizarColores
End Sub

Private Sub cmdUsuarios_Click()
optBancas = False
optUsuarios = True
optOtros = False
ActualizarColores
End Sub

Private Sub cmdVolver_Click()
    Unload Me
End Sub
Private Sub cmdAgrupacionPolitica_Click()
    frmABMPartidos.Show vbModal
End Sub

Private Sub Command2_Click()
    frmSetearConfig.Show vbModal
End Sub

Private Sub Command3_Click()
    If PermisosTotales.ABMUsuarios = 1 Then
        frmUsuarios.Show vbModal
    Else
        MsgBox "El usuario no tiene permisos para realizar esta operación", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Command4_Click()
    frrmCambiarPassword.Show vbModal, Me
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    establecerPermisos
    Call HabilitarControles
    frmOtros.Visible = False
    frameUsuarios.Visible = False
    frmBancas.Visible = True
    optBancas = True
    optOtros = False
    optUsuarios = False
End Sub
Private Sub HabilitarControles()
    If PROVINCIA_HABILITADA = True Then
        lblDistritos.Visible = False
        lblSecciones.Visible = False
    End If
End Sub
Private Sub establecerPermisos()
'?
End Sub
Private Sub ActualizarColores()
Dim colorOK As Variant
Dim colorNO As Variant
colorOK = &H40C0&
colorNO = &H808080
If optUsuarios = True Then
    cmdUsuarios.BackColor = colorOK
    frameUsuarios.Visible = True
    cmdOtros.BackColor = colorNO
    frmOtros.Visible = False
    cmdBancas.BackColor = colorNO
    frmBancas.Visible = False
ElseIf optBancas = True Then
    cmdUsuarios.BackColor = colorNO
    frameUsuarios.Visible = False
    cmdBancas.BackColor = colorOK
    frmBancas.Visible = True
    cmdOtros.BackColor = colorNO
    frmOtros.Visible = False
Else
    cmdOtros.BackColor = colorOK
    frmOtros.Visible = True
    cmdUsuarios.BackColor = colorNO
    frameUsuarios.Visible = False
    cmdBancas.BackColor = colorNO
    frmBancas.Visible = False
End If
cmdVolver.SetFocus
End Sub
