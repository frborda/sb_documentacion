VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmConsolaOperacion 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   15105
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmConsolaOperacionAEB.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFecha 
      Height          =   285
      Left            =   1110
      Locked          =   -1  'True
      TabIndex        =   135
      Top             =   7515
      Width           =   975
   End
   Begin VB.TextBox txtNo 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   134
      ToolTipText     =   "Cantidad de votos Negativos"
      Top             =   7515
      Width           =   450
   End
   Begin VB.TextBox txtHora 
      Height          =   285
      Left            =   1110
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   7860
      Width           =   975
   End
   Begin VB.TextBox txtPresentes 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3375
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   " "
      Top             =   7515
      Width           =   435
   End
   Begin VB.TextBox txtAusentes 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4755
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   " "
      Top             =   7515
      Width           =   405
   End
   Begin VB.TextBox txtOcup 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14190
      Locked          =   -1  'True
      TabIndex        =   45
      ToolTipText     =   "Legisladores ocupadando banca sin identificarse"
      Top             =   7860
      Width           =   435
   End
   Begin VB.TextBox TxtQuorum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   " "
      Top             =   7860
      Width           =   1785
   End
   Begin VB.TextBox txtAbs 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   11250
      Locked          =   -1  'True
      TabIndex        =   43
      ToolTipText     =   "Cantidad de Abstenciones"
      Top             =   7515
      Width           =   450
   End
   Begin VB.TextBox txtSi 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8220
      Locked          =   -1  'True
      TabIndex        =   42
      ToolTipText     =   "Cantidad de Votos Afirmativos"
      Top             =   7515
      Width           =   450
   End
   Begin VB.TextBox txtTiempo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6210
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   " "
      ToolTipText     =   "Tiempo de Votación"
      Top             =   7515
      Width           =   690
   End
   Begin VB.TextBox txtPendientesEmitirVoto 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14190
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   " "
      ToolTipText     =   "Cantidad de Legisladores Pendientes de Emitir Voto"
      Top             =   7515
      Width           =   435
   End
   Begin VB.TextBox txtResultado 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   8205
      Locked          =   -1  'True
      TabIndex        =   39
      ToolTipText     =   "Resultado de la Votación"
      Top             =   7860
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancelarVotacion 
      Caption         =   "Cance&lar Votación"
      Height          =   330
      Left            =   11700
      TabIndex        =   38
      ToolTipText     =   "Cancelar Votación"
      Top             =   9540
      Width           =   3210
   End
   Begin VB.CommandButton cmdAbstenciones 
      Caption         =   "Selector de &Abstenciones"
      Height          =   330
      Left            =   11700
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Período legislativo y Tipo de Sesión"
      Top             =   9930
      Width           =   3210
   End
   Begin VB.CommandButton Salir 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   11700
      TabIndex        =   36
      Top             =   10920
      Width           =   3210
   End
   Begin VB.CommandButton cmdReconsiderar 
      Caption         =   "Reconsiderar"
      Height          =   330
      Left            =   11700
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Período legislativo y Tipo de Sesión"
      Top             =   10290
      Width           =   3210
   End
   Begin VB.CommandButton cmdVotacion 
      Caption         =   "Votac&ión"
      Height          =   330
      Left            =   11700
      TabIndex        =   34
      Top             =   9180
      Width           =   3210
   End
   Begin VB.CommandButton cmdMantenimiento 
      Caption         =   "Man&tenimiento"
      Height          =   330
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   10920
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Tomar Control"
      Height          =   330
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   10920
      Width           =   1590
   End
   Begin VB.CommandButton cmdCarteles 
      Caption         =   "E&ncender Carteles"
      Height          =   330
      Left            =   9780
      TabIndex        =   31
      ToolTipText     =   "Período legislativo y Tipo de Sesión"
      Top             =   10920
      Width           =   1620
   End
   Begin VB.CheckBox chkInformacion 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   600
      Picture         =   "frmConsolaOperacionAEB.frx":17DD3
      TabIndex        =   29
      Top             =   10950
      UseMaskColor    =   -1  'True
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CommandButton cmdTituloRapido 
      Caption         =   "&Título"
      Height          =   315
      Left            =   5940
      TabIndex        =   28
      Top             =   9570
      Width           =   1245
   End
   Begin VB.CommandButton cmdCambiarSesion 
      Caption         =   "&Cambiar Sesión"
      Height          =   345
      Left            =   4590
      TabIndex        =   25
      Top             =   8820
      Width           =   1275
   End
   Begin VB.CommandButton cmdNuevaSesion 
      Caption         =   "&Nueva Sesión"
      Height          =   345
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8820
      Width           =   1215
   End
   Begin VB.CommandButton cmdTituloBlanco 
      Caption         =   "&Limpiar"
      Height          =   315
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9930
      Width           =   1245
   End
   Begin VB.CommandButton cmdPresidente 
      Caption         =   "&Presidente"
      Height          =   315
      Left            =   5940
      TabIndex        =   22
      ToolTipText     =   "Seleccionar presidente del total de legisladores"
      Top             =   9210
      Width           =   1245
   End
   Begin VB.CommandButton cmdTitulo 
      Caption         =   "Orden del Día"
      Height          =   345
      Left            =   5940
      TabIndex        =   21
      Top             =   8820
      Width           =   1245
   End
   Begin VB.CommandButton cmdTiempo 
      Caption         =   "Tiemp&o de Votación"
      Height          =   330
      Left            =   2610
      TabIndex        =   20
      ToolTipText     =   "Período legislativo y Tipo de Sesión"
      Top             =   10920
      Width           =   1815
   End
   Begin VB.CommandButton cmdPeriodoLegislativo 
      Caption         =   "Período Legislativo y Tipo de Sesión"
      Height          =   345
      Left            =   510
      TabIndex        =   19
      ToolTipText     =   "Seleccionar Período legislativo y Tipo de Sesión"
      Top             =   8820
      Width           =   2760
   End
   Begin VB.CommandButton cmdControlSistema 
      Caption         =   "&Control de Sistema"
      Height          =   330
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Período legislativo y Tipo de Sesión"
      Top             =   10920
      Width           =   1755
   End
   Begin VB.TextBox txtTitulo 
      Height          =   675
      Left            =   1470
      MaxLength       =   80
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   9540
      Width           =   4395
   End
   Begin MSWinsockLib.Winsock Ws 
      Left            =   13920
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   14400
      Top             =   1800
   End
   Begin MSDataListLib.DataCombo dcTipoMayoria 
      Height          =   315
      Left            =   8550
      TabIndex        =   0
      Top             =   8820
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcBaseMayoria 
      Height          =   315
      Left            =   8550
      TabIndex        =   1
      Top             =   9180
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcTipoOperacion 
      Height          =   315
      Left            =   12630
      TabIndex        =   2
      Top             =   8790
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcAbstencion 
      Height          =   315
      Left            =   8550
      TabIndex        =   6
      Top             =   9900
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcTipoQuorum 
      Height          =   315
      Left            =   8550
      TabIndex        =   7
      Top             =   9540
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   11700
      TabIndex        =   133
      Top             =   8850
      Width           =   1665
   End
   Begin VB.Label lblModoPrueba 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Consola en modo PRUEBA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   6240
      TabIndex        =   132
      Top             =   5370
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   2
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   2
      Left            =   4980
      TabIndex        =   131
      Top             =   5640
      Width           =   930
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   3
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   3
      Left            =   5880
      TabIndex        =   130
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   4
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   4
      Left            =   5880
      TabIndex        =   129
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   5
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   5
      Left            =   5880
      TabIndex        =   128
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   6
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   6
      Left            =   5880
      TabIndex        =   127
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   7
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   7
      Left            =   5880
      TabIndex        =   126
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   8
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   8
      Left            =   5880
      TabIndex        =   125
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   9
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   9
      Left            =   5880
      TabIndex        =   124
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   10
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   10
      Left            =   5880
      TabIndex        =   123
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   11
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   11
      Left            =   5880
      TabIndex        =   122
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   12
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   12
      Left            =   5880
      TabIndex        =   121
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   13
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   13
      Left            =   5880
      TabIndex        =   120
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   14
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   14
      Left            =   5880
      TabIndex        =   119
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   15
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   15
      Left            =   5880
      TabIndex        =   118
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   16
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   16
      Left            =   5880
      TabIndex        =   117
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   17
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   17
      Left            =   5880
      TabIndex        =   116
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   18
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   18
      Left            =   5880
      TabIndex        =   115
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   19
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   19
      Left            =   5880
      TabIndex        =   114
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   20
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   20
      Left            =   5880
      TabIndex        =   113
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   21
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   21
      Left            =   5880
      TabIndex        =   112
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   22
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   22
      Left            =   5880
      TabIndex        =   111
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   23
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   23
      Left            =   5880
      TabIndex        =   110
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   24
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   24
      Left            =   5880
      TabIndex        =   109
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   25
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   25
      Left            =   5880
      TabIndex        =   108
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   26
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   26
      Left            =   5880
      TabIndex        =   107
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   27
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   27
      Left            =   5880
      TabIndex        =   106
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   28
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   28
      Left            =   5880
      TabIndex        =   105
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   29
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   29
      Left            =   5880
      TabIndex        =   104
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   30
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   30
      Left            =   5880
      TabIndex        =   103
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   31
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   31
      Left            =   5880
      TabIndex        =   102
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   32
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   32
      Left            =   5880
      TabIndex        =   101
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   33
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   33
      Left            =   5880
      TabIndex        =   100
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   34
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   34
      Left            =   5880
      TabIndex        =   99
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   35
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   35
      Left            =   5880
      TabIndex        =   98
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   36
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   36
      Left            =   5880
      TabIndex        =   97
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   37
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   37
      Left            =   5880
      TabIndex        =   96
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   38
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   38
      Left            =   5880
      TabIndex        =   95
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   39
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   39
      Left            =   5880
      TabIndex        =   94
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   40
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   40
      Left            =   5880
      TabIndex        =   93
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   41
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   41
      Left            =   5880
      TabIndex        =   92
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   42
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   42
      Left            =   5880
      TabIndex        =   91
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   43
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   43
      Left            =   5880
      TabIndex        =   90
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   44
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   44
      Left            =   5880
      TabIndex        =   89
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   45
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   45
      Left            =   5880
      TabIndex        =   88
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   46
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   46
      Left            =   5880
      TabIndex        =   87
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   47
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   47
      Left            =   5880
      TabIndex        =   86
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   48
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   48
      Left            =   5880
      TabIndex        =   85
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   49
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   49
      Left            =   5880
      TabIndex        =   84
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   50
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   50
      Left            =   5880
      TabIndex        =   83
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   51
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   51
      Left            =   5880
      TabIndex        =   82
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   52
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   52
      Left            =   5880
      TabIndex        =   81
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   53
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   53
      Left            =   5880
      TabIndex        =   80
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   54
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   54
      Left            =   5880
      TabIndex        =   79
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   55
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   55
      Left            =   5880
      TabIndex        =   78
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   56
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   56
      Left            =   5880
      TabIndex        =   77
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   57
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   57
      Left            =   5880
      TabIndex        =   76
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   58
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   58
      Left            =   5880
      TabIndex        =   75
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   59
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   59
      Left            =   5880
      TabIndex        =   74
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   60
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   60
      Left            =   5880
      TabIndex        =   73
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   61
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   61
      Left            =   5880
      TabIndex        =   72
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   62
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   62
      Left            =   5880
      TabIndex        =   71
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   63
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   63
      Left            =   5880
      TabIndex        =   70
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   64
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   64
      Left            =   5880
      TabIndex        =   69
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   65
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   65
      Left            =   5880
      TabIndex        =   68
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   66
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   66
      Left            =   5880
      TabIndex        =   67
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   67
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   67
      Left            =   5880
      TabIndex        =   66
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   68
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   68
      Left            =   5880
      TabIndex        =   65
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   69
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   69
      Left            =   5880
      TabIndex        =   64
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   70
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   70
      Left            =   5880
      TabIndex        =   63
      Top             =   5850
      Width           =   420
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   0
      Left            =   5880
      TabIndex        =   62
      Top             =   5850
      Width           =   480
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   765
      Index           =   1
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5850
      Width           =   795
   End
   Begin VB.Label ctrBanca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   1
      Left            =   5880
      TabIndex        =   61
      Top             =   5850
      Width           =   420
   End
   Begin VB.Shape shpBanca 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   765
      Index           =   0
      Left            =   8070
      Shape           =   3  'Circle
      Top             =   5730
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Identificados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   12735
      TabIndex        =   59
      Top             =   7905
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ausentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   3915
      TabIndex        =   58
      Top             =   7560
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Presentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   2475
      TabIndex        =   57
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quorum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   2655
      TabIndex        =   56
      Top             =   7905
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   7290
      TabIndex        =   55
      Top             =   7905
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Afirmativos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   7245
      TabIndex        =   54
      Top             =   7560
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Negativos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   8805
      TabIndex        =   53
      Top             =   7560
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Abstenidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   10260
      TabIndex        =   52
      Top             =   7560
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   615
      TabIndex        =   51
      Top             =   7905
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   510
      TabIndex        =   50
      Top             =   7560
      Width           =   540
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   5535
      TabIndex        =   49
      Top             =   7560
      Width           =   630
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pendientes de Emitir Voto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   11940
      TabIndex        =   48
      Top             =   7560
      Width           =   2205
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Modalidad Información"
      Height          =   255
      Left            =   870
      TabIndex        =   30
      Top             =   10950
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Presidente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   2
      Left            =   510
      TabIndex        =   27
      Top             =   9240
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Título"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   780
      TabIndex        =   26
      Top             =   9540
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mayoría Base"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   20
      Left            =   6720
      TabIndex        =   17
      Top             =   9255
      Width           =   1740
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mayoría Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   19
      Left            =   6720
      TabIndex        =   16
      Top             =   8910
      Width           =   1740
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Abstención"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   18
      Left            =   6750
      TabIndex        =   15
      Top             =   9960
      Width           =   1710
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Quorum Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   17
      Left            =   6750
      TabIndex        =   14
      Top             =   9615
      Width           =   1710
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   11730
      Picture         =   "frmConsolaOperacionAEB.frx":AB68F
      Top             =   8370
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   1
      Left            =   510
      Picture         =   "frmConsolaOperacionAEB.frx":ABD79
      Top             =   8370
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   2
      Left            =   6600
      Picture         =   "frmConsolaOperacionAEB.frx":AC463
      Top             =   10470
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   3
      Left            =   510
      Picture         =   "frmConsolaOperacionAEB.frx":ACB4D
      Top             =   7080
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   4
      Left            =   510
      Picture         =   "frmConsolaOperacionAEB.frx":AD237
      Top             =   10470
      Width           =   360
   End
   Begin VB.Label lblEtiqueta 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "102 Período Legislativo: Especial - 1ª Sesión Especial - Próximo Nº de Acta: 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   450
      TabIndex        =   13
      Top             =   240
      Width           =   14430
   End
   Begin VB.Label lblTituloActa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Título del acta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   480
      TabIndex        =   8
      Top             =   465
      Width           =   14460
   End
   Begin VB.Label lblSesion 
      Caption         =   "lblSesion"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   5565
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblNroSesion 
      Caption         =   "lnlNroSesion"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   5310
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblPeriodo 
      Caption         =   "lblPeriodo"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   5055
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblCodigoSesion 
      Caption         =   "lblCodigoSesion"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblPresidente 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Presidente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1470
      TabIndex        =   4
      Top             =   9210
      Width           =   4395
   End
   Begin VB.Label lblActa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   15840
      TabIndex        =   3
      Top             =   885
      Width           =   915
   End
   Begin VB.Menu mnuPopUP 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu IndexBanca 
         Caption         =   "Index"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuID 
         Caption         =   "Identificación"
      End
      Begin VB.Menu mnuBloque 
         Caption         =   "Bloque"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDepartamento 
         Caption         =   "Departamento"
         Visible         =   0   'False
      End
      Begin VB.Menu Sep00 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAsignarID 
         Caption         =   "Asignar Identificación"
      End
      Begin VB.Menu mnuAbstener 
         Caption         =   "Cancelar Abstención"
      End
      Begin VB.Menu sep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReinicio 
         Caption         =   "Reiniciar Banca"
      End
      Begin VB.Menu mnuHardReset 
         Caption         =   "Resetear Hardware Banca"
      End
   End
End
Attribute VB_Name = "frmConsolaOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Public IniciaActualizacion        As Boolean
Public MensajeEsperado            As Variant
Private xListar_automaticamente   As Long
Dim foco                          As Boolean
Private strconexion               As String
Dim strTipoP                      As String
'Variables para DataCombos
Private rstTipoOperacion          As New ADODB.Recordset
Private rstBaseMayoria            As New ADODB.Recordset
Private rstTipoMayoria            As New ADODB.Recordset
Private rstAbstencion             As New ADODB.Recordset
Private rstTipoQuorum             As New ADODB.Recordset
Private conFoco                   As Boolean
Private TiempoEsperaRespuesta     As Long 'Variables para gestión de esperas
Private pasoTiempo                As Long 'Variables para gestión de esperas
'Variables para actualización de consola
Private rstVector                 As New ADODB.Recordset
Private mVectorColores()          As String
Private mVectorPresencia()        As String
Private mVectorIdentificacion()   As String
Private mCantidadLegisladores     As Integer
Private mActualizarDatos          As Boolean
Private mMensajeOperador          As String
Private mIPActivo                 As String
Private mPreguntarIniciarServidor As Boolean
Private mRefrescarConsola         As Boolean
Private Const sqlVector           As String = "SELECT TOP 1 * FROM vector ORDER BY id DESC"
Private Const mDelimitadorVector  As String = ";"

'Variables para Votación
Private mEstadoVotacion           As String
'Variables para gestión de sesiones
Private mEstadoSesion             As String
'Variables para control de conexion al SQV
Private mCiclosControlConexion    As Integer
Private mHoraAnterior             As String
Private Const mTiempoVerificacionConectividad = 50
'Colores de bancas
Private mColores                  As New Scripting.Dictionary
Private mColoresFuentes       As New Scripting.Dictionary
'Cadena de legisladores identificados
Public mLegisladoresAsignados     As String
'Control del presidente
Private rstPresidente             As ADODB.Recordset
'Gestión de mantenimiento
Private mModoMantenimiento        As Boolean
'Variable para form de info de banca
Public info                       As frmInformacionBanca
'MAnejo de actas
Private mActaIniciada             As Long
Private mActaGrabada              As Long
Private mActaImpresa              As Long
Private strPath                   As String
Private blBanderaTimer            As Boolean
Private Tiempo1 As Date
Private Tiempo2 As Date
Private xTiempoDif As Double

'ANDRÉS
Private imagePath As String
Private Type tpBanca
    Legislador As String
    Presencia As String
End Type
Private datBanca(71) As tpBanca

Public Property Let TiempoEspera(ByVal pNewValue As Integer)
    determinarTiempoEspera pNewValue
End Property
Public Sub determinarTiempoEspera(Optional pTiempo As Integer = -1)
    Dim auxTiempo As Integer
    If pTiempo < 1 Then
        'Obtener la cantidad de multiplicacion (auxtiempo) desde un tabla de config
        auxTiempo = 10
    Else
        auxTiempo = pTiempo
    End If
    TiempoEsperaRespuesta = (1000 / Timer.Interval) * auxTiempo
End Sub


Private Sub cmdAbstenciones_Click()
    If LCase(Trim(dcTipoOperacion.Text)) = "votación numérica" Then
        frmAbstencionesNumericas.Show 1, Me
    ElseIf LCase(Trim(dcTipoOperacion.Text)) = "votación nominal" Then
        frmAbstenerLegisladores.Show 1, Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    MensajesSQV.cancelarConsola
End Sub
Private Sub cmdCancelarVotacion_Click()
    MensajesSQV.cancelaVotacion
End Sub
Private Sub cmdCarteles_Click()
    strCartelSQV = ""
    frmMensajesCartel.Show 1, Me
    Select Case strCartelSQV
        Case "0"
            MensajesSQV.cambiarCartelEncendido "0"
        Case "1"
            MensajesSQV.cambiarCartelEncendido "1"
        Case "2"
            MensajesSQV.cambiarCartelEncendido "2"
    End Select
End Sub
Private Sub EstadoCarteles(pEstado As Integer)
    If UCase(pEstado) = 0 Then
        cmdCarteles.Caption = "Encender Carteles"
    Else
        cmdCarteles.Caption = "Apagar Carteles"
    End If
End Sub

Private Sub cmdControlSistema_Click()
        frmControlSistema.Show vbModal
End Sub

Private Sub cmdMantenimiento_Click()
    If PermisosTotales.UsuarioMantenimiento = 1 Then
        FrmMantenimiento.Show vbModal
    Else
        MsgBox "El usuario no tiene permisos para realizar esta tarea", vbInformation + vbOKOnly
    End If
End Sub
Private Sub cmdPeriodoLegislativo_Click()
    Dim periodo As New frmElegirPeriodo
    periodo.ActualizarDatos = True
    periodo.Show vbModal
    Set periodo = Nothing
End Sub
Private Sub cmdpresidente_Click()
    frmElegirPresidente.Show 1, Me
    ' lblPresidente.Caption = strPresidente
End Sub
Private Sub ArmarBancas()
 
    Dim respuestaMensaje  As Variant
    Dim strEtiqueta       As String
    ' Valores para imprimir acta
    Dim strTipoOp         As String
    Dim strPeriodoLeg     As String
    Dim xSesion           As Long
    Dim xNroActaActual    As Long
 
    SetearRs sqlVector, rstVector
    ' Determinar tipo de operacion
    strTipoOp = rstVector!Identificador_tipo_de_operacion
    strPeriodoLeg = rstVector!Período_legislativo
    xSesion = rstVector!Sesión
    xNroActaActual = rstVector!Nro_de_Acta
    xListar_automaticamente = rstVector!Listar_automaticamente
    
    If rstVector.EOF = False Then
        txtFecha.Text = Format(rstVector!Fecha, "dd/mm/yyyy")
        txtHora.Text = rstVector!hora
        mIPActivo = Trim(rstVector!IP_Consola_Habilitada)
        ' Or True se agrego para forzar a la consola a tener todos los controles habilitados
        If mIPActivo = Datos.IPconsola Or True Then
            mRefrescarConsola = True
        Else
            mRefrescarConsola = False
        End If
        If txtHora.Text = mHoraAnterior Then
            If mCiclosControlConexion < mTiempoVerificacionConectividad Then
                mCiclosControlConexion = mCiclosControlConexion + 1
            Else
                If mPreguntarIniciarServidor = True Then
                    If mIPActivo = Datos.IPconsola Then
                        If MsgBox("Error de comunicación con el Servidor." & Chr(13) & "Desea intentar iniciar el servidor?.", vbQuestion + vbYesNo, "Error de conectividad") = vbYes Then
                            Call Levanta_Sqv
                            Call Levanta_Banca
                            mCiclosControlConexion = 0
                        Else
                            MsgBox "La consola le advertirá cuando el servidor esté disponible.", vbInformation + vbOKOnly
                        End If
                    Else
                        MsgBox "Error de comunicación con el Servidor." & Chr(13) & "La consola le advertirá cuando el servidor esté disponible.", vbInformation + vbOKOnly
                    End If
                    mPreguntarIniciarServidor = False
                End If
                mRefrescarConsola = False
            End If
        Else
            mHoraAnterior = txtHora.Text
            mCiclosControlConexion = 0
            If mPreguntarIniciarServidor = False Then
                MsgBox "Se ha restablecido la comunicación con el Servidor.", vbInformation + vbOKOnly
                mPreguntarIniciarServidor = True
            End If
            If mIPActivo = Datos.IPconsola Then
                mRefrescarConsola = True
            Else
                mRefrescarConsola = False
            End If
        End If
        If Trim(dcTipoOperacion.Tag) <> Trim(rstVector!Identificador_tipo_de_operacion) Then
            dcTipoOperacion.Tag = Trim(rstVector!Identificador_tipo_de_operacion)
            dcTipoOperacion.BoundText = Trim(rstVector!Identificador_tipo_de_operacion)
        End If
        If Trim(dcBaseMayoria.Tag) <> Trim(rstVector.Fields("Base_de_Mayoría")) Then
            dcBaseMayoria.Tag = Trim(rstVector.Fields("Base_de_Mayoría"))
            dcBaseMayoria.BoundText = Trim(rstVector.Fields("Base_de_Mayoría"))
        End If
        If Trim(dcTipoMayoria.Tag) <> Trim(rstVector.Fields("Tipo_de_Mayoría")) Then
            dcTipoMayoria.Tag = Trim(rstVector.Fields("Tipo_de_Mayoría"))
            dcTipoMayoria.BoundText = Trim(rstVector.Fields("Tipo_de_Mayoría"))
        End If
        If Trim(dcAbstencion.Tag) <> Trim(rstVector!Tipo_de_Abstención) Then
            dcAbstencion.Tag = Trim(rstVector!Tipo_de_Abstención)
            dcAbstencion.BoundText = Trim(rstVector!Tipo_de_Abstención)
        End If
        If Trim(dcTipoQuorum.Tag) <> Trim(rstVector!Tipo_Mayoria_Quorum) Then
            dcTipoQuorum.Tag = Trim(rstVector!Tipo_Mayoria_Quorum)
            dcTipoQuorum.BoundText = Trim(rstVector!Tipo_Mayoria_Quorum)
        End If
        ' ------------------------------------------------------------------------------------
        ' Mostrar valores del vector
        ' ------------------------------------------------------------------------------------
        If Not IsNull(rstVector!Titulo_Del_Acta) Then
            lblTituloActa.Caption = rstVector!Titulo_Del_Acta
        Else
            lblTituloActa.Caption = ""
        End If
        txtPresentes.Text = rstVector!Presentes
        txtAusentes.Text = rstVector!ausentes
        txtOcup.Text = rstVector!Ocupadas_no_identificadas
        TxtQuorum.Text = Trim(rstVector!Leyenda_Quorum)
        txtResultado.Text = rstVector!resultado
        txtSi.Text = rstVector!afirmativos
        txtNo.Text = rstVector!negativos
        txtAbs.Text = rstVector!Abstenciones
        'AP 041012 No se actualiza para permitir la edicion. txtTitulo.Text = rstVector!Titulo_Del_Acta
        'lblNroSesion.Caption = rstVector!Sesión
        lblActa.Caption = rstVector!Nro_de_Acta
        mActaGrabada = Val(rstVector!Acta_Grabada)
        
        If Not IsNull(rstVector!Período_legislativo) Then
                lblCodigoSesion.Caption = Left(rstVector!Período_legislativo, 3)
                strEtiqueta = Left(rstVector!Período_legislativo, 3) & " Período Legislativo: "
                
                lblCodigoSesion.Tag = rstVector!Período_legislativo
                Select Case UCase(Mid(rstVector!Período_legislativo, 4, 1))
                    Case "O"
                        strEtiqueta = strEtiqueta & "Ordinario "
                        lblPeriodo.Caption = "Ordinario"
                        strTipoP = "Ordinario"
                    Case "E"
                        strEtiqueta = strEtiqueta & "Extraordinario "
                        lblPeriodo.Caption = "Extraordinario"
                        strTipoP = "Extraordinario"
                    Case "P"
                        strEtiqueta = strEtiqueta & "Preparatoria "
                        lblPeriodo.Caption = "Preparatoria"
                        strTipoP = "Preparatoria"
                End Select
                If IsNull(rstVector!Sesión) = False Then
                    lblNroSesion.Caption = rstVector!Sesión
                    Select Case UCase(Mid(rstVector!Período_legislativo, 5, 1))
                        Case "T"
                            strEtiqueta = strEtiqueta & " - " & rstVector!Sesión & " Sesión"
                            strEtiqueta = strEtiqueta & "Tablas"
                            lblSesion.Caption = "Tablas"
                        Case "E"
                            strEtiqueta = strEtiqueta & "Especial"
                            lblSesion.Caption = "Especial"
                        Case "P"
                            strEtiqueta = strEtiqueta & "Preparatoria"
                            lblSesion.Caption = "Preparatoria"
                    End Select
                End If
                If IsNull(rstVector!Nro_de_Acta) = False Then
                    strEtiqueta = strEtiqueta & " - Próximo Nº de Acta: " & rstVector!Nro_de_Acta
                    lblActa.Caption = rstVector!Nro_de_Acta
                End If
                lblEtiqueta.Caption = strEtiqueta
        Else
            lblCodigoSesion.Caption = ""
            lblPeriodo.Caption = ""
            lblNroSesion.Caption = ""
            lblSesion.Caption = ""
            lblActa.Caption = ""
            cmdCambiarSesion.Enabled = False
            cmdNuevaSesion.Enabled = False
        End If
        EstadoCarteles rstVector!encender_carteles
        hacerSplitVector Trim(rstVector!vector_colores), mVectorColores
        hacerSplitVector Trim(rstVector!vector_presencia), mVectorPresencia
        hacerSplitVector Trim(rstVector!vector_identificacion), mVectorIdentificacion
        ' Mostrar Presidente
        If mVectorIdentificacion(0) <> strIdPresidenteRecinto And mVectorIdentificacion(0) <> "0" Then
            strIdPresidenteRecinto = mVectorIdentificacion(0)
            Call BuscarPresidentePredeterminado
        ElseIf mVectorIdentificacion(0) = "0" Then
            'aca
            lblPresidente.Caption = "Seleccione el presidente antes de continuar"
        End If
        ' sbStatus.Panels(2).Text = "Tiempo para votación:" & Str(rstVector!Tiempo_de_votación)
        'Andres  Label2(2).Caption = "Tiempo de Votación: " & Str(rstVector!Tiempo_de_votación)
        respuestaMensaje = Trim(rstVector!strError)
        mEstadoVotacion = LCase(Trim(rstVector!Estado_de_votacion))
        mEstadoSesion = Trim(rstVector!Estado_sesion)
        mMensajeOperador = Trim(rstVector!Mensaje_al_operador)
        'Controles de permisos de usuario
        If (mRefrescarConsola = True) Then
             EstadoControles
        End If
        
        
        
        'botones de votacion
        If mEstadoSesion <> "abierta" And Trim(rstVector!Identificador_tipo_de_operacion) <> "paslis" Then
            cmdVotacion.Caption = "Sesion NO Abierta"
            cmdVotacion.Enabled = False
            cmdCancelarVotacion.Enabled = False
            cmdNuevaSesion.Caption = "Abrir s&esión"
            cmdNuevaSesion.BackColor = &H8080FF
        Else
            cmdNuevaSesion.Caption = "Nueva s&esión"
            cmdNuevaSesion.BackColor = &H8000000F
            Select Case Trim(rstVector!Identificador_tipo_de_operacion)
                Case Is = "votnom", "votnum"
                   Select Case mEstadoVotacion
                        Case Is = "cancelada", "finalizada"
                            cmdVotacion.Caption = "&Inicializar"
                            If mRefrescarConsola = True Then
                                cmdVotacion.Enabled = True
                            End If
                            'txtPendientesEmitirVoto.Visible = False
                            txtPendientesEmitirVoto.Enabled = False
                            cmdCancelarVotacion.Enabled = False
                        Case Is = "espera"
                            'txtPendientesEmitirVoto.Visible = False
                            txtPendientesEmitirVoto.Enabled = False
                            If TxtQuorum.Text = "QUORUM" Then
                                If mRefrescarConsola = True Then
                                    cmdVotacion.Enabled = True
                                End If
                                cmdVotacion.Caption = "Votac&ión"
                                cmdCancelarVotacion.Enabled = False
                            Else
                                cmdVotacion.Enabled = False
                                cmdVotacion.Caption = "Sin QUORUM"
                                cmdCancelarVotacion.Enabled = False
                            End If
                         Case Is = "votando", "larga"
                                txtPendientesEmitirVoto.Enabled = True
                                'txtPendientesEmitirVoto.Visible = True
                                cmdVotacion.Enabled = False
                                cmdVotacion.Caption = "Votac&ión"
                                ControlesHabilitados = False
                                txtPendientesEmitirVoto.Text = IIf(rstVector!Pendientes_Emitir_Voto >= 0, rstVector!Pendientes_Emitir_Voto, "-")
                                If mRefrescarConsola = True Then
                                    cmdCancelarVotacion.Enabled = True
                                End If
                          Case Is = "empate"
                                cmdVotacion.Enabled = False
                                cmdVotacion.Caption = "Votac&ión"
                                ControlesHabilitados = False
                                'txtPendientesEmitirVoto.Visible = False
                                txtPendientesEmitirVoto.Enabled = False
                                If mRefrescarConsola = True Then
                                    cmdCancelarVotacion.Enabled = True
                                End If
                    End Select
                Case Is = "quorum"
                   cmdVotacion.Enabled = False
                   cmdVotacion.Caption = "Votación"
                   'txtPendientesEmitirVoto.Visible = False
                   txtPendientesEmitirVoto.Enabled = False
                Case Is = "paslis"
                    Select Case mEstadoVotacion
                        Case Is = "cancelada"
                            cmdVotacion.Caption = "&Inicializar"
                            If mRefrescarConsola = True Then
                                cmdVotacion.Enabled = True
                            End If
                        Case Is = "votando"
                            cmdVotacion.Caption = "Pase de l&ista"
                            cmdVotacion.Enabled = False
                        Case Else
                            cmdVotacion.Caption = "Pase de l&ista"
                            If mRefrescarConsola = True Then
                                cmdVotacion.Enabled = True
                            End If
                    End Select
                    'txtPendientesEmitirVoto.Visible = False
                   txtPendientesEmitirVoto.Enabled = False
            End Select
        End If
    End If
    rstVector.Close
    AsignarDatosBancas
    If mRefrescarConsola = True Then
        'Control de respuesta de mensajes
        If MensajeEsperado <> MensajeVacio Then
            pasoTiempo = pasoTiempo + 1
            ' sbStatus.Panels(3).Text = "Esperando confirmación del servidor..."
            consolaHabilitada = False
            Screen.MousePointer = vbHourglass
            If respuestaMensaje <> MensajeEsperado Then
                If respuestaMensaje = "**error" Then
                    If strOldErrMensaje <> mMensajeOperador Then
                        MsgBox mMensajeOperador, vbExclamation + vbOKOnly
                        Datos.GrabarMensaje "mensajemostrado", " "
                        strOldErrMensaje = mMensajeOperador
                    End If
                End If
                If pasoTiempo >= TiempoEsperaRespuesta Then
                    MensajesSQV.mensajeNoConfirmado
                    pasoTiempo = 0
                    ' sbStatus.Panels(3).Text = "Último mensaje no confirmado"
                    Screen.MousePointer = 0
                    consolaHabilitada = True
                    determinarTiempoEspera
                    If MensajeEsperado = "pruebascan" Then
                        frmInformacionBanca.ControlesHabilitados = True
                    End If
                End If
            Else
                determinarTiempoEspera
                If MensajeEsperado = "pruebascan" Then
                    If LCase(mMensajeOperador) <> "negative" Then
                        MensajesSQV.PruebaScanConfirmada mMensajeOperador
                    Else
                        MensajesSQV.PruebaScanIncorrecta
                    End If
                End If
                MensajeEsperado = MensajeVacio
                ' sbStatus.Panels(3).Text = "Preparado"
                pasoTiempo = 0
                Screen.MousePointer = 0
                consolaHabilitada = True
            End If
        End If
        If respuestaMensaje = "**error" Then
            If strOldErrMensaje <> mMensajeOperador Then
                MsgBox mMensajeOperador, vbInformation + vbOKOnly, "SQV Informa"
                Datos.GrabarMensaje "mensajemostrado", " "
                strOldErrMensaje = mMensajeOperador
            End If
        End If
        ' ---------------------------------------------------------------------------------
        ' Control de actas
        ' ---------------------------------------------------------------------------------
        'If xListar_automaticamente > 0 Then
       '     'cmdListar.Caption = "No &Listar"
      '  Else
     '       'cmdListar.Caption = "&Listar"
     '   End If
        If ((mActaGrabada > 0) And (mActaImpresa <> mActaGrabada)) Then
            mActaImpresa = mActaGrabada
            If xListar_automaticamente > 0 Then
                Call ImprimirActa(strTipoOp, strPeriodoLeg, xSesion, xNroActaActual)
            Else
                ''cmdListar.Caption = "&Listar"
                If MsgBox("¿Desea imprimir el acta pendiente?", vbYesNo + vbQuestion) = vbYes Then
                    Call ImprimirActa(strTipoOp, strPeriodoLeg, xSesion, xNroActaActual)
                End If
            End If
        End If
        If mActaGrabada = 0 Then
            mActaImpresa = 0
        End If
        ' *********************************************************************************
    Else
        ControlesHabilitados = False
        Screen.MousePointer = 0
        consolaHabilitada = True
        determinarTiempoEspera
    End If
End Sub
Private Sub ImprimirActa(strTipoOperacionActual As String, strPeriodo As String, xSesionActual As Long, xActualActa As Long)
    
    'Dim m_Report As New rptActas
    'Dim rstActa  As New ADODB.Recordset
    'Dim Sql      As String
    
    'strTipoOperacionActual = Trim(LCase(strTipoOperacionActual))
    
    'If strTipoOperacionActual = "votnum" Then
    '    Sql = "SELECT * From actas " & _
            " WHERE Período_Legislativo = '" & strPeriodo & "' " & _
            " AND sesión = " & xSesionActual & _
            " AND Número_de_Acta = " & xActualActa & _
            " AND Versión_Acta = 0"
    'Else
    '    Sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Ultima_Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables, actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
        " actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
        " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período Legislativo: ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Ordinario' WHEN 'E' THEN 'Especial' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión: ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Tabla' WHEN 'E' THEN 'Especial' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - Próximo Nº de acta: ' + CAST(actas.Número_de_Acta AS Varchar(5)) END AS DescripcionPeriodoLegislativo, CASE TipoMayoriaQuorum.descripcion WHEN '120' THEN 'Más de la mitad' WHEN '121' THEN 'La mitad más uno' ELSE 'MAN' END AS DescripcionTipoQuorum " & _
        " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
        " AND actas.Sesión = " & _
        " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.identificador_en_mensajes LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
        " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (Actas.Período_Legislativo='100OT') AND (Actas.Sesión=1) AND (Actas.Número_de_Acta=" & xActualActa & ") AND (Actas.Versión_Acta=0) ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
    'End If
    
    'Dim fViewer As frmVisor
    'Set fViewer = New frmVisor
    'SetearRs Sql, rstActa
    'm_Report.Database.SetDataSource rstActa
    'fViewer.CRViewer1.ReportSource = m_Report
    'fViewer.CRViewer1.PrintReport
    'fViewer.CRViewer1.Zoom 100
    'fViewer.Show vbModal
    
    'Set rstActa = Nothing
    'Set fViewer = Nothing
    'Set m_Report = Nothing
    
End Sub
Private Sub EstadoControles()
    If (LCase(dcTipoOperacion.Tag) = "votnom") Or (LCase(dcTipoOperacion.Tag) = "votnum") Then
        If LCase(Trim(rstVector!Estado_de_votacion)) = "espera" Then
            'Antes de comenzar la votacion
            If mModoMantenimiento = False Then
                dcTipoOperacion.Enabled = True
            Else
                dcTipoOperacion.Enabled = False
            End If
            'POZ 'POZ cmdConsulta.Enabled = False
            ' habilitadoBloqueControles1 = False
            habilitadoBloqueControles1 = True
            habilitadoBloqueControlesResultado = False
        ElseIf (LCase(Trim(rstVector!Estado_de_votacion)) = "larga") And (Trim(LCase(rstVector!Estado_sesion)) = "abierta") Then
            'durante una votacion larga
            'POZ cmdConsulta.Enabled = True
            dcTipoOperacion.Enabled = False
            habilitadoBloqueControles1 = False
            habilitadoBloqueControlesResultado = False
        Else
            'votando
            'POZ cmdConsulta.Enabled = False
            dcTipoOperacion.Enabled = False
            habilitadoBloqueControles1 = False
            If (LCase(Trim(rstVector!Estado_de_votacion)) = "finalizada") Or (LCase(Trim(rstVector!Estado_de_votacion)) = "empate") Then
                habilitadoBloqueControlesResultado = True
            Else
                'Fin de votacion
                habilitadoBloqueControlesResultado = False
            End If
        End If
        If (LCase(dcTipoOperacion.Tag) = "votnom") And LCase(Trim(rstVector!Estado_de_votacion)) = "espera" Then
            cmdReconsiderar.Enabled = True
        Else
            cmdReconsiderar.Enabled = False
        End If
    ElseIf (LCase(dcTipoOperacion.Tag) = "quorum") Or (LCase(dcTipoOperacion.Tag) = "paslis") Or (LCase(Trim(rstVector!Estado_de_votacion)) = "espera") Then
        'Durante censado de quorum
        dcTipoOperacion.Enabled = True
        cmdReconsiderar.Enabled = False
        habilitadoBloqueControles1 = True
        habilitadoBloqueControlesResultado = False
    Else
        cmdReconsiderar.Enabled = False
        dcTipoOperacion.Enabled = False
        habilitadoBloqueControles1 = False
        habilitadoBloqueControlesResultado = False
    End If
        
    'Prevencion de cambio de operacion sin Presidente seleccionado
    If Len(Trim(strIdPresidenteRecinto)) = 0 And (LCase(dcTipoOperacion.Tag) = "quorum") Then
        dcTipoOperacion.Enabled = False
    End If
End Sub
Private Property Let habilitadoBloqueControles1(pModo As Boolean)

    If PermisosTotales.HabilitaBotonesConsola = 1 Then
        dcBaseMayoria.Enabled = pModo
        cmdPresidente.Enabled = pModo
        dcTipoMayoria.Enabled = pModo
        cmdTituloBlanco.Enabled = pModo
        cmdCambiarSesion.Enabled = pModo
        cmdNuevaSesion.Enabled = pModo
        cmdTitulo.Enabled = pModo
        cmdTituloRapido.Enabled = pModo
        cmdPeriodoLegislativo.Enabled = pModo
        dcTipoQuorum.Enabled = pModo
        'dcAbstencion.Enabled = pModo
        dcAbstencion.Enabled = False
        cmdTiempo.Enabled = pModo
        'cmdListar.Enabled = pModo
    ElseIf PermisosTotales.HabilitaBotonesConsola = 0 Then
        dcBaseMayoria.Enabled = False
        cmdPresidente.Enabled = False
        dcTipoMayoria.Enabled = False
        cmdTituloBlanco.Enabled = False
        cmdCambiarSesion.Enabled = False
        cmdNuevaSesion.Enabled = False
        cmdTitulo.Enabled = False
        cmdTituloRapido.Enabled = False
        cmdPeriodoLegislativo.Enabled = False
        dcTipoQuorum.Enabled = False
        dcAbstencion.Enabled = False
        cmdTiempo.Enabled = False
        'cmdListar.Enabled = False
    End If

End Property
Private Property Let habilitadoBloqueControlesResultado(pModo As Boolean)
    txtSi.Enabled = pModo
    txtNo.Enabled = pModo
    txtAbs.Enabled = pModo
    txtResultado.Enabled = pModo
End Property
Private Sub BuscarPresidente(pPresidente As String)
    
    If pPresidente <> "" Then
        Set rstPresidente = New ADODB.Recordset
        SetearRs "SELECT rtrim(SURNAME) + ', ' + rtrim(NAME) as nombre FROM VIEW_GETALLDATA WHERE id='" & pPresidente & "'", rstPresidente
        If rstPresidente.EOF = False Then
            lblPresidente.Caption = rstPresidente!Nombre
        Else
            lblPresidente.Caption = ""
        End If
    Else
        lblPresidente.Caption = ""
    End If
    'free mem
    If rstPresidente.State = adStateOpen Then
        rstPresidente.Close
    End If
    Set rstPresidente = Nothing
End Sub
Private Sub AsignarDatosBancas()
    On Error Resume Next
    Dim i     As Integer
    Dim clave As String
    'busco los datos del presi sólo si cambia
    'If mVectorIdentificacion(0) <> shpBanca(0).Legislador Then
        'POZ  BuscarPresidente mVectorIdentificacion(0)
    'End If
    For i = 0 To UBound(mVectorColores)
        clave = i
        shpBanca(i).FillColor = mColores(mVectorColores(clave))
        ctrBanca(i).ForeColor = mColoresFuentes(mVectorColores(clave))
        datBanca(i).Legislador = mVectorIdentificacion(i)
        datBanca(i).Presencia = mVectorPresencia(i)
        If mVectorIdentificacion(i) <> "" Then
            If mLegisladoresAsignados = "" Then
                mLegisladoresAsignados = "'" & mVectorIdentificacion(i) & "'"
            Else
                If mVectorIdentificacion(i) <> "0" Then
                    mLegisladoresAsignados = mLegisladoresAsignados & ",'" & mVectorIdentificacion(i) & "'"
                End If
            End If
        End If
    Next i
End Sub
Private Sub cmdReconsiderar_Click()
    Dim xSesionReconsiderar As Long
    Dim xActaReconsiderar As Long
    ' frmListarActas.Show 1
    MsgBox "xSesionReconsiderar y xActaReconsiderar deben ser seleccionados por usuario"
    xSesionReconsiderar = 1
    xActaReconsiderar = 2
    Call VotacionReconsideracion(xSesionReconsiderar, xActaReconsiderar)
End Sub
Private Sub cmdTiempo_Click()
    frmTiempoVotacion.Show vbModal
End Sub
Private Sub cmdTitulo_Click()
    Dim titulo As New frmTituloActa
    
    If (lblCodigoSesion.Caption <> "") And (lblSesion.Caption <> "") Then
        If (gTipoUsuario = 0) Or (gTipoUsuario = 2) Then
            titulo.MostrarDatos lblCodigoSesion.Tag, lblNroSesion.Caption
            titulo.Show vbModal
            Set titulo = Nothing
        Else
            MsgBox "Ud. no dispone de permisos para realizar esta acción.", vbInformation + vbOKOnly
        End If
    Else
        MsgBox "Debe haber una sesión abierta para visualizar su Órden del día.", vbInformation + vbOKOnly
    End If
End Sub
Private Sub cmdTituloRapido_Click()
    MensajesSQV.CambioTituloActa 0, txtTitulo
End Sub
Private Sub cmdTituloBlanco_Click()
    'txtTitulo.Text = ""
    'lblTituloActa.Caption = ""
    MensajesSQV.CambioTituloActa 0, ""
End Sub
Private Sub cmdVotacion_Click()
    Select Case cmdVotacion.Caption
        Case Is = "Votac&ión"
            MensajesSQV.inicioVotacion
        Case Is = "&Inicializar"
            MensajesSQV.inicializar
        Case Is = "Pase de l&ista"
            If mActaIniciada = 0 Then
                mActaIniciada = Val(lblActa.Caption)
                MensajesSQV.paseLista
                ' sbStatus.Panels(3).Text = "Pase de lista INICIADO..."
            Else
                ' sbStatus.Panels(3).Text = "ACTA YA INICIADO"
            End If
    End Select
    cmdVotacion.Enabled = False
End Sub
Private Sub cmdCambiarSesion_Click()
    If lblCodigoSesion.Caption <> "" Then
        Dim Sesion As New frmCambiarSesion
        If Sesion.MostrarDatos(lblCodigoSesion.Tag) = True Then
            Sesion.ActualizarDatos = True
            Sesion.Show vbModal
        End If
        Set Sesion = Nothing
    End If
End Sub
Private Sub cmdNuevaSesion_Click()
    If cmdNuevaSesion.Caption = "Nueva s&esión" Then
        If lblCodigoSesion.Caption <> "" Then
            Dim Sesion As New frmCrearSesion
            Sesion.AgregarDatos lblCodigoSesion.Tag
            Sesion.Show vbModal
            Set Sesion = Nothing
        End If
    Else
        If (gTipoUsuario = 0) Or (gTipoUsuario = 2) Then
            If MsgBox("Está Ud. seguro de abrir la sesión actual?", vbQuestion + vbYesNo, "Confirma la operación?") = vbYes Then
                MensajesSQV.abrirSesion
            End If
        Else
            MsgBox "Ud. no dispone de permisos para realizar esta acción.", vbInformation + vbOKOnly
        End If
    End If
End Sub


Private Sub ctrBanca_DblClick(Index As Integer)
    Dim PermitirPrueba      As Boolean
    Dim PermitirIdentificar As Boolean
    Dim PermitirCambioVoto  As Boolean
    Dim PermitirAbstener    As Boolean
    
    If chkInformacion.Value = vbChecked Then
        'modo info
        Set info = New frmInformacionBanca
        EvaluarPermisosOperacionesBanca Index, PermitirPrueba, PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
        If datBanca(Index).Legislador <> "" Then
            info.MostrarDatos Index, datBanca(Index).Legislador, PermitirPrueba, PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
        Else
            info.MostrarDatos Index, , , PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
        End If
        info.EstadoVotacion = mEstadoVotacion
        info.Show vbModal
        Set info = Nothing
    Else
        If (dcTipoOperacion.BoundText = "votnom") Or (dcTipoOperacion.BoundText = "paslis") Then
            If mRefrescarConsola = True Then
                'modo cambio id
                Dim frmId As New frmAsignarLegislador
                frmId.mostrarLegisladores Index
                frmId.Show vbModal
                Set frmId = Nothing
            End If
        End If
    End If
End Sub

Private Sub ctrBanca_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NumeroBanca As Long
    Dim PermitirPrueba      As Boolean
    Dim PermitirIdentificar As Boolean
    Dim PermitirCambioVoto  As Boolean
    Dim PermitirAbstener    As Boolean
    Dim Legislador As String
    Dim Bloque As String
    Dim Departamento As String
    
    If Button = 2 Then
        IndexBanca.Caption = Index
        If chkInformacion.Value = vbChecked Then
            NumeroBanca = Val(ctrBanca(Index).Caption)
            EvaluarPermisosOperacionesBanca Index, PermitirPrueba, PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
            mnuAsignarID.Enabled = PermitirIdentificar
            mnuAbstener.Enabled = PermitirAbstener
            DeterminarLeyendaMenuAbstencion NumeroBanca
            Sep00.Visible = True
            TraerInfoLegislador datBanca(Index).Legislador, Legislador, Bloque, Departamento
            mnuID.Caption = Legislador
            mnuBloque.Visible = (Bloque <> "")
            mnuBloque.Caption = Bloque
            mnuDepartamento.Visible = (Departamento <> "")
            mnuDepartamento.Caption = Departamento
            Me.PopupMenu mnuPopUP, , , , mnuID
        End If
    End If
End Sub

Private Sub dcAbstencion_Change()
    dcAbstencion.Enabled = False
    If conFoco = True Then
        MensajesSQV.ModoVotacion dcAbstencion.BoundText
        conFoco = False
        dcAbstencion.Tag = ""
        TxtQuorum.SetFocus
    End If
End Sub
Private Sub dcAbstencion_GotFocus()
    conFoco = True
End Sub
Private Sub dcAbstencion_LostFocus()
    conFoco = False
End Sub
Private Sub dcBaseMayoria_Change()
    If conFoco = True Then
        MensajesSQV.cambiarBaseMayoria dcBaseMayoria.BoundText
        conFoco = False
        dcBaseMayoria.Tag = ""
        TxtQuorum.SetFocus
    End If
End Sub
Private Sub dcBaseMayoria_GotFocus()
    conFoco = True
End Sub
Private Sub dcBaseMayoria_LostFocus()
    conFoco = False
End Sub
Private Sub dcTipoMayoria_Change()
    If conFoco = True Then
        MensajesSQV.cambiarTipoMayoriaVotacion dcTipoMayoria.BoundText
        conFoco = False
        dcTipoMayoria.Tag = ""
        TxtQuorum.SetFocus
    End If
End Sub
Private Sub dcTipoMayoria_GotFocus()
    conFoco = True
End Sub
Private Sub dcTipoMayoria_LostFocus()
    conFoco = False
End Sub
Private Sub dcTipoOperacion_Change()
    foco = True
    cmdAbstenciones.Enabled = False
    If conFoco = True Then
        MensajesSQV.cambiarTipoOperacion dcTipoOperacion.BoundText
        dcTipoOperacion.Tag = ""
        conFoco = False
        TxtQuorum.SetFocus
    End If
    cmdAbstenciones.Enabled = True
End Sub
Private Sub determinarCantidadLegisladores()
    Dim rstAux As New ADODB.Recordset
    SetearRs "SELECT Cantidad_de_Legisladores FROM Config", rstAux
    If rstAux.EOF = False Then
        mCantidadLegisladores = (rstAux!Cantidad_de_Legisladores) - 1
    Else
         mCantidadLegisladores = 0
    End If
    
    ReDim mVectorColores(0 To mCantidadLegisladores)
    ReDim mVectorPresencia(0 To mCantidadLegisladores)
    ReDim mVectorIdentificacion(0 To mCantidadLegisladores)
    
    'libero memoria
    If rstAux.State = adStateOpen Then
        rstAux.Close
    End If
    Set rstAux = Nothing
End Sub
Private Sub dcTipoOperacion_GotFocus()
foco = True
    conFoco = True
End Sub
Private Sub dcTipoOperacion_LostFocus()
foco = False
    conFoco = False
End Sub
Private Sub dcTipoQuorum_Change()
    If conFoco = True Then
        MensajesSQV.cambiarTipoQuorum dcTipoQuorum.BoundText
        TxtQuorum.SetFocus
        dcTipoQuorum.Tag = ""
        conFoco = False
    End If
End Sub
Private Sub dcTipoQuorum_GotFocus()
    conFoco = True
End Sub
Private Sub dcTipoQuorum_LostFocus()
    conFoco = False
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0

    If FlagBasePrueba Then
        lblModoPrueba.Visible = True
    Else
        lblModoPrueba.Visible = False
    End If
    Call Levanta_Sqv
    Call Levanta_Banca

End Sub
Private Sub Form_Load()
    'cargo el winsock p/levantar srvr
    'Ws.RemoteHost = "192.168.1.254"
    'Ws.RemotePort = 10023
    'Ws.Close
    'Ws.Connect
    'fin
    mActualizarDatos = True
    mActaIniciada = 0
    IniciaActualizacion = False
    mPreguntarIniciarServidor = True
    mRefrescarConsola = True
    MensajeEsperado = MensajesSQV.MensajeVacio
    determinarTiempoEspera
    determinarCantidadLegisladores
    Call DeterminarValoresInicioServer
    cargarColores
    Call ArmarPantalla
    Call ArmarBancas
    Call BuscarPresidentePredeterminado
    IniciaActualizacion = True
    MensajesSQV.habilitarConsola
    ' sbStatus.Panels(1).Text = "Versión: " & App.Major & "." & App.Minor & "." & App.Revision
    reiniciarBancas
    blBanderaTimer = True
    'Andres
    ArmarBancasCartel
    'Andres End
End Sub
Private Sub BuscarPresidentePredeterminado()
    Dim strSql          As String
    Dim strIdPresidente As String
    Dim RsPres          As ADODB.Recordset
    Set RsPres = New ADODB.Recordset
    ' Leantar vector de identificacion
    strSql = "SELECT Apellido + ', ' + Nombre AS Presidente, Es_Legislador FROM Legisladores WHERE Id = '" & strIdPresidenteRecinto & "'"
    SetearRs strSql, RsPres
    If Not RsPres.EOF Then
        lblPresidente.Caption = RsPres.Fields("Presidente").Value
        If RsPres.Fields("Es_Legislador").Value = 0 Then
            flPresidenteLegislador = False
        Else
            flPresidenteLegislador = True
        End If
    Else
        lblPresidente.Caption = "Seleccione presidente"
        flPresidenteLegislador = False
    End If
    RsPres.Close
End Sub
Private Sub DeterminarValoresInicioServer()
    Dim rstAux As New ADODB.Recordset
    Dim strSql As String
    ' ------------------------------------------------------------------------------
    ' Esta funcion levanta los valores de configuracion default
    ' para levantar los servicios de sqv server y sb
    ' ------------------------------------------------------------------------------
    
    strSql = "SELECT Ejecutable_sqv, Ejecutable_sb, IP_levanta_ap, " _
           & "puerto_levanta_ap From config"
    SetearRs strSql, rstAux
        
    If rstAux.EOF = False Then
        If Not IsNull(rstAux.Fields("puerto_levanta_ap").Value) Then
            strPuerto = rstAux.Fields("puerto_levanta_ap").Value
        End If
        If Not IsNull(rstAux.Fields("IP_levanta_ap").Value) Then
            strIpServer = rstAux.Fields("IP_levanta_ap").Value
        End If
        If Not IsNull(rstAux.Fields("Ejecutable_sqv").Value) Then
            strExeSqv = rstAux.Fields("Ejecutable_sqv").Value
        End If
        If Not IsNull(rstAux.Fields("Ejecutable_sb").Value) Then
            strExeSb = rstAux.Fields("Ejecutable_sb").Value
        End If
    Else
        strPuerto = ""
        strIpServer = ""
        strExeSqv = ""
        strExeSb = ""
    End If
    rstAux.Close: Set rstAux = Nothing
End Sub
Private Sub reiniciarBancas()
    If mIPActivo = Datos.IPconsola Then
        If (gTipoUsuario = 0) Or (gTipoUsuario = 2) Then
            If MsgBox("Desea reiniciar las bancas?", vbQuestion + vbYesNo) = vbYes Then
                reiniciarTodasBancas
            End If
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mCiclosControlConexion = 0
    MensajesSQV.liberarConsola
    'Libero rstTipo operación
    If rstTipoOperacion.State = adStateOpen Then
        rstTipoOperacion.Close
    End If
    Set rstTipoOperacion = Nothing
    'Libero rst base mayoria
    If rstBaseMayoria.State = adStateOpen Then
        rstBaseMayoria.Close
    End If
    Set rstBaseMayoria = Nothing
    'libero rst tipo mayoria
    If rstTipoMayoria.State = adStateOpen Then
        rstTipoMayoria.Close
    End If
    Set rstTipoMayoria = Nothing
    'libero rst VECTOR
    If rstVector.State = adStateOpen Then
        rstVector.Close
    End If
    Set rstVector = Nothing
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
    
    mColores.RemoveAll
    Set mColores = Nothing
End Sub


Private Sub Salir_Click()
    Unload Me
End Sub

Private Sub ArmarPantalla()
   Dim strSql As String
   Dim i As Integer
    'txtHora = Hour(Now) & ":" & Minute(Now) la levanto del vector para verificar conexión
    strSql = "SELECT Tipo_de_operación, rtrim(identificador_en_mensajes) as identificador , Resultados_posibles " _
           & "From tipoop ORDER BY Tipo_de_operación"
    SetearRs strSql, rstTipoOperacion
    With dcTipoOperacion
        Set .RowSource = rstTipoOperacion
        .ListField = "Tipo_de_operación"
        .BoundColumn = "identificador"
    End With
    
    strSql = "SELECT rtrim(identificador_en_mensajes) as identificador, Descripcion FROM basemay ORDER BY Descripcion"
    SetearRs strSql, rstBaseMayoria
    With dcBaseMayoria
        Set .RowSource = rstBaseMayoria
        .BoundColumn = "identificador"
        .ListField = "Descripcion"
    End With
    
    strSql = "SELECT rtrim(Tipo_de_Mayoria) as tipo, Descripcion From tipmay WHERE Habilitado = 1 ORDER BY Descripcion"
    SetearRs strSql, rstTipoMayoria
    With dcTipoMayoria
        Set .RowSource = rstTipoMayoria
        .ListField = "Descripcion"
        .BoundColumn = "tipo"
    End With

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
    ' PRoblema de banca de presidente alfanumerica...
    'shpBanca(0).Caption = "P"
    'For i = 1 To shpBanca.Count - 1
    For i = 0 To shpBanca.Count - 1
        ctrBanca(i).Caption = i
    Next i
    
    conFoco = False
End Sub
Private Sub Timer_Timer()
    If blBanderaTimer Then
        blBanderaTimer = False
        Tiempo1 = Now
        If mActualizarDatos = True Then
            Call ArmarBancas
            DoEvents
        End If
        blBanderaTimer = True
        Tiempo2 = Now
        xTiempoDif = DateDiff("s", Tiempo1, Tiempo2)
        'If xTiempoDif > 1 Then
        '    MsgBox xTiempoDif
        'End If
    End If
End Sub
Private Sub hacerSplitVector(ByVal pCadena As String, ByRef pVector() As String)
    pVector = Split(pCadena, mDelimitadorVector)
End Sub
Private Property Let ControlesHabilitados(ByVal pModo As Boolean)
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If (UCase(ctrl.Name) <> "TXTQUORUM") And (ctrl.Name <> "Salir") And (ctrl.Name <> "cmdCancelar") Then
            Select Case TypeName(ctrl)
                Case Is = "ComboBox", "CommandButton", "CheckBox", "DataCombo"
                        ctrl.Enabled = pModo
            End Select
        End If
    Next
    If PermisosTotales.HabilitaBotonesConsola = 1 Then
        chkInformacion.Enabled = True
        cmdMantenimiento.Enabled = True
        cmdCarteles.Enabled = True
        cmdAbstenciones.Enabled = True
        cmdReconsiderar.Enabled = True
        cmdControlSistema.Enabled = True
        dcAbstencion.Enabled = False
    ElseIf PermisosTotales.HabilitaBotonesConsola = 0 Then
        chkInformacion.Enabled = False
        cmdMantenimiento.Enabled = False
        cmdCarteles.Enabled = False
        cmdAbstenciones.Enabled = False
        cmdReconsiderar.Enabled = False
        cmdControlSistema.Enabled = False
        dcAbstencion.Enabled = False
        cmdCancelar.Enabled = False
    End If
End Property
Private Property Let consolaHabilitada(pModo As Boolean)
    Me.Enabled = pModo
End Property
Private Sub Levanta_Sqv()
    Ws.Close
    Ws.RemoteHost = Trim(strIpServer)
    Ws.RemotePort = Trim(strPuerto)
    Ws.Connect
    strPath = strExeSqv
    DoEvents
    'While Ws.State = 6
    '    Screen.MousePointer = 11
    '    DoEvents
   ' Wend
    Screen.MousePointer = 0
End Sub
Private Sub Levanta_Banca()
    Ws.Close
    Ws.RemoteHost = Trim(strIpServer)
    Ws.RemotePort = Trim(strPuerto)
    Ws.Connect
    strPath = strExeSb
    DoEvents
    'While Ws.State = 6
    '    Screen.MousePointer = 11
   '     DoEvents
   ' Wend
    Screen.MousePointer = 0
End Sub
Private Sub Ws_Connect()
    If Ws.State = sckConnected Then
       Ws.SendData strPath & vbCrLf
       DoEvents
    Else
        MsgBox "Error : No se puedo levantar al Aplicacion, Reintente"
    End If
    Ws.Close
End Sub

Private Sub ArmarBancasCartel()
   
    Dim xBanca As Integer
    Dim radio As Double
    Dim pi As Double
    Dim xCentroIzquierdo As Double
    Dim xCentroDerecho As Double
    Dim yCentro As Double
    Dim xCentro As Double
    Dim xObjeto As Double
    Dim yObjeto As Double
    Dim offset As Double
    Dim Step As Double
    Dim Escala As Double
    Dim xUltimaBanca As Integer
    Escala = 0.9
    xCentroIzquierdo = (Me.Width / 2) - 215
    xCentroDerecho = (Me.Width / 2) + 215
    yCentro = 6480
    pi = 4 * Atn(1)   ' Calculo el valor de pi
    Me.Cls
    Me.Refresh
    xUltimaBanca = 70
    For xBanca = 0 To xUltimaBanca
        shpBanca(xBanca).Visible = True
        ctrBanca(xBanca).Visible = True
        shpBanca(xBanca).Width = 630
        shpBanca(xBanca).Height = 630
        shpBanca(xBanca).FillStyle = 0
        shpBanca(xBanca).BorderWidth = 2
        If xBanca >= 1 And xBanca <= 22 Then
            radio = (Me.Width / 2) * 0.768 * Escala
            Step = pi / 21
            xCentro = IIf(xBanca < 12, xCentroIzquierdo, xCentroDerecho)
            offset = 1
        ElseIf xBanca >= 23 And xBanca <= 42 Then
            radio = (Me.Width / 2) * 0.638 * Escala
            Step = pi / 19
            xCentro = IIf(xBanca < 33, xCentroIzquierdo, xCentroDerecho)
            offset = 23
        ElseIf xBanca >= 43 And xBanca <= 58 Then
            radio = (Me.Width / 2) * 0.5 * Escala
            Step = pi / 15
            xCentro = IIf(xBanca < 51, xCentroIzquierdo, xCentroDerecho)
            offset = 43
        ElseIf xBanca >= 59 And xBanca <= 70 Then
            radio = (Me.Width / 2) * 0.365 * Escala
            Step = pi / 11
            xCentro = IIf(xBanca < 65, xCentroIzquierdo, xCentroDerecho)
            offset = 59
        End If
        
        If xBanca = 0 Then
            xObjeto = Me.Width / 2
            yObjeto = yCentro
        Else
            xObjeto = xCentro - Cos(Step * (xBanca - offset)) * radio
            yObjeto = yCentro - Sin(Step * (xBanca - offset)) * radio
        End If
        shpBanca(xBanca).Left = xObjeto - (shpBanca(xBanca).Width / 2)
        shpBanca(xBanca).Top = yObjeto - (shpBanca(xBanca).Height / 2)
        shpBanca(xBanca).FillColor = vbWhite
        ctrBanca(xBanca).Caption = xBanca
        ctrBanca(xBanca).AutoSize = True
        ctrBanca(xBanca).Alignment = 2
        ctrBanca(xBanca).FontSize = 19
        ctrBanca(xBanca).AutoSize = False
        ctrBanca(xBanca).Width = 630
        ctrBanca(xBanca).Left = xObjeto - (ctrBanca(xBanca).Width / 2)
        ctrBanca(xBanca).Top = yObjeto - (ctrBanca(xBanca).Height / 2)
        shpBanca(xBanca).ZOrder 1
    Next
    Me.ZOrder 1
    Me.Refresh
End Sub


Private Sub CargarColoresFuente()
    
    Dim Color As String
    Dim clave As String
    
    ' Limpio el diccionario
    mColoresFuentes.RemoveAll
    'cargo el diccionario de manera estática
    'GRIS     'CONTRASTE NEGRO
    Color = "&H0": clave = "0":    mColoresFuentes.Add clave, Color
    'BLANCO     'CONTRASTE NEGRO
    Color = "&H0":  clave = "1":     mColoresFuentes.Add clave, Color
    'AMARILLO     'CONTRASTE NEGRO
    Color = "&H0":      clave = "2":  mColoresFuentes.Add clave, Color
    'ROJO     'CONTRASTE BLANCO
    Color = "&HFFFFFF":      clave = "3":    mColoresFuentes.Add clave, Color
    'CELESTE     'CONTRASTE NEGRO
    Color = "&H0":      clave = "4":      mColoresFuentes.Add clave, Color
    'NARANJA     'CONTRASTE NEGRO
    Color = "&H0":      clave = "5":     mColoresFuentes.Add clave, Color
    'VERDE     'CONTRASTE NEGRO
    Color = "&H0":      clave = "6":     mColoresFuentes.Add clave, Color
    'NEGRO     'CONTRASTE BLANCO
    Color = "&HFFFFFF":      clave = "7":   mColoresFuentes.Add clave, Color
End Sub

Private Sub cargarColores()
    Dim Color As String
    Dim clave As String
    
    ' Limpio el diccionario
    mColores.RemoveAll
    'cargo el diccionario de manera estática
    'GRIS
    Color = "&HC0C0C0"
    clave = "0"
    mColores.Add clave, Color
    'BLANCO
    Color = "&HFFFFFF"
    clave = "1"
    mColores.Add clave, Color
    'AMARILLO
    Color = "&HFFFF"
    clave = "2"
    mColores.Add clave, Color
    'ROJO
    Color = "&HFF"
    clave = "3"
    mColores.Add clave, Color
    'CELESTE
    Color = "&HFFFF00"
    clave = "4"
    mColores.Add clave, Color
    'NARANJA
    Color = "&H80FF"
    clave = "5"
    mColores.Add clave, Color
    'VERDE
    Color = "&HFF00"
    clave = "6"
    mColores.Add clave, Color
    'NEGRO
    Color = "&H0"
    clave = "7"
    mColores.Add clave, Color
End Sub


'MENU CONTEXTUALES ************************************************************
Private Sub DeterminarLeyendaMenuAbstencion(mNumeroBanca As Long)
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
        mnuAbstener.Caption = "Cancelar A&bstención"
    Else
        mnuAbstener.Caption = "&Abstener"
    End If
End Sub

Private Sub EvaluarPermisosOperacionesBanca(Index As Integer, bPermitirPrueba As Boolean, bPermitirIdentificar As Boolean, bPermitirCambioVoto As Boolean, bPermitirAbstener As Boolean)
        bPermitirPrueba = LCase(dcTipoOperacion.BoundText) = "quorum" And datBanca(Index).Presencia = "1"
        bPermitirIdentificar = (LCase(dcTipoOperacion.BoundText) = "votnom") Or (LCase(dcTipoOperacion.BoundText) = "paslis")
        ' Chequear el vector identificacion
        If ((dcTipoOperacion.BoundText = "votnum") Or (dcTipoOperacion.BoundText = "votnom")) And _
            (mEstadoVotacion = "votando") Or (mEstadoVotacion = "larga") Or mEstadoVotacion = "empate" Then
            If Trim(mVectorPresencia(Index)) = "1" Then
                If dcTipoOperacion.BoundText = "votnom" And mVectorIdentificacion(Index) > 0 Then
                   bPermitirCambioVoto = True
                ElseIf dcTipoOperacion.BoundText = "votnum" Then
                    bPermitirCambioVoto = True
                Else
                    bPermitirCambioVoto = False
                End If
            Else
                bPermitirCambioVoto = False
            End If
        Else
            bPermitirCambioVoto = False
        End If
        If (dcTipoOperacion.BoundText = "votnum") Or (dcTipoOperacion.BoundText = "votnom") Then
            If (mEstadoVotacion = "votando") Or (mEstadoVotacion = "larga") Or (mEstadoVotacion = "espera") Then
                bPermitirAbstener = True
            Else
                bPermitirAbstener = False
            End If
        Else
            bPermitirAbstener = False
        End If
End Sub

Private Sub mnuAbstener_Click()
    MensajesSQV.abstener Val(ctrBanca(Val(IndexBanca.Caption)).Caption)
End Sub

Private Sub mnuAsignarId_Click()
    Dim asignar As New frmAsignarLegislador
    asignar.mostrarLegisladores Val(ctrBanca(Val(IndexBanca.Caption)).Caption)
    asignar.Show vbModal
    Set asignar = Nothing
End Sub

Private Sub mnuHardReset_Click()
    MensajesSQV.reiniciarBancaHard ctrBanca(Val(IndexBanca.Caption))
End Sub

Private Sub mnuReinicio_Click()
    MensajesSQV.reiniciarBanca ctrBanca(Val(IndexBanca.Caption))
End Sub


Private Sub TraerInfoLegislador(pIdLegislador As String, Legislador As String, Optional Bloque As String, Optional Departamento As String, Optional Agrupacion As String)
    If IsNull(pIdLegislador) = False Then
        Dim rstAux As New ADODB.Recordset
        SetearRs "SELECT * FROM Legisladores WHERE id='" & pIdLegislador & "'", rstAux
        If rstAux.EOF = False Then
            If (IsNull(rstAux!Apellido) = False) And (IsNull(rstAux!Nombre) = False) Then
                Legislador = Trim(rstAux!Apellido) & ", " & Trim(rstAux!Nombre)
            End If
            If IsNull(rstAux!grupo_politico) = False Then
                Agrupacion = Trim(rstAux!grupo_politico)
            End If
            If IsNull(rstAux!bloque_politico) = False Then
                Bloque = Trim(rstAux!bloque_politico)
            End If
            If IsNull(rstAux!Departamento) = False Then
                Departamento = Trim(rstAux!Departamento)
            End If
        Else
            Legislador = "No Identificado"
        End If
    End If
End Sub
