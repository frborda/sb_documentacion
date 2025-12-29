VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmConsultas 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menú de consultas"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdListadoLegisladores 
      Height          =   465
      Left            =   90
      TabIndex        =   3
      Top             =   120
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   820
      BackColor       =   12230304
      Caption         =   "Listados de &Legisladores"
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
   Begin Proyecto1.ButtonOffice cmdActas 
      Height          =   555
      Left            =   90
      TabIndex        =   4
      Top             =   1320
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   979
      BackColor       =   12230304
      Caption         =   "&Consultar y modificar actas"
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
      Height          =   555
      Left            =   90
      TabIndex        =   5
      Top             =   2640
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   979
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
   Begin Proyecto1.ButtonOffice cmdListadoDatosRecinto 
      Height          =   705
      Left            =   90
      TabIndex        =   6
      Top             =   600
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   1244
      BackColor       =   12230304
      Caption         =   "Listado de datos de Recinto"
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Permite listar las bancas probables y huellas de los diputados"
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
      Left            =   2490
      TabIndex        =   7
      Top             =   720
      Width           =   4635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Permite consultar los Legisladores."
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
      Left            =   2460
      TabIndex        =   2
      Top             =   240
      Width           =   4635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cierra esta ventana y vuelve al Menú Principal del Sistema."
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
      Left            =   2460
      TabIndex        =   1
      Top             =   2790
      Width           =   5085
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Permite consultar y modificar las actas de una sesión."
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
      Height          =   315
      Left            =   2460
      TabIndex        =   0
      Top             =   1470
      Width           =   4635
   End
End
Attribute VB_Name = "frmConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActas_Click()
    'Dim periodo As New frmElegirPeriodo
    'periodo.ActualizarDatos = False
    Dim periodo As New frmMostrarPeriodos
    ImpresionDeConsola = False
    cmdActas.Enabled = False
    cmdActas.Caption = "Cargando..."
    DoEvents
    periodo.Show vbModal
    Set periodo = Nothing
    cmdActas.Enabled = True
    cmdActas.Caption = "&Consultar y modificar actas"
End Sub

Private Sub cmdListadoDatosRecinto_Click()
    Dim frm As New frmListadoDatosRecinto
    frm.Show vbModal
End Sub

Private Sub cmdListadoLegisladores_Click()
   frmListados.Show vbModal, Me
End Sub
Private Sub cmdVolver_Click()
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
