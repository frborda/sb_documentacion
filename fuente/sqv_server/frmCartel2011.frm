VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmCartel2011 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "G "
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   Begin VB.Frame grpPendientes 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "TRIACA, Alberto Jorge"
      Height          =   7455
      Left            =   420
      TabIndex        =   37
      Top             =   4020
      Visible         =   0   'False
      Width           =   14835
      Begin VB.Label lblPendientes10 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   525
         Left            =   360
         TabIndex        =   48
         Top             =   6840
         Width           =   10125
      End
      Begin VB.Label lblPendientes9 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   585
         Left            =   360
         TabIndex        =   47
         Top             =   6240
         Width           =   10125
      End
      Begin VB.Label lblPendientes8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   585
         Left            =   360
         TabIndex        =   46
         Top             =   5580
         Width           =   10125
      End
      Begin VB.Label lblPendientes7 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   585
         Left            =   360
         TabIndex        =   45
         Top             =   4980
         Width           =   10125
      End
      Begin VB.Label lblPendientes6 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   585
         Left            =   360
         TabIndex        =   44
         Top             =   4320
         Width           =   10125
      End
      Begin VB.Label lblPendientes5 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   585
         Left            =   360
         TabIndex        =   43
         Top             =   3660
         Width           =   10125
      End
      Begin VB.Label lblPendientes4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   585
         Left            =   360
         TabIndex        =   42
         Top             =   2880
         Width           =   10125
      End
      Begin VB.Label lblPendientes3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   585
         Left            =   360
         TabIndex        =   41
         Top             =   2220
         Width           =   10125
      End
      Begin VB.Label lblPendientes2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   585
         Left            =   360
         TabIndex        =   40
         Top             =   1560
         Width           =   10125
      End
      Begin VB.Label lblPendientes1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   585
         Left            =   360
         TabIndex        =   39
         Top             =   960
         Width           =   10125
      End
      Begin VB.Label lblPendientesCantidad 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Abstenciones: 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   60
         TabIndex        =   38
         Top             =   120
         Width           =   14805
      End
   End
   Begin VB.Timer tmActualizar 
      Interval        =   10
      Left            =   1680
      Top             =   6240
   End
   Begin VB.Timer tmUpdate 
      Interval        =   1
      Left            =   7920
      Top             =   4050
   End
   Begin VB.Frame frameTimer 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2595
      Left            =   300
      TabIndex        =   50
      Top             =   1080
      Visible         =   0   'False
      Width           =   7515
      Begin MSWinsockLib.Winsock svWinsock 
         Left            =   6180
         Top             =   2220
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer tmSqvData 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   6360
         Top             =   1560
      End
      Begin MSWinsockLib.Winsock WebSender 
         Left            =   6300
         Top             =   420
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   8891
      End
      Begin MSWinsockLib.Winsock ProtoServer 
         Left            =   4020
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   8881
         LocalPort       =   8881
      End
      Begin VB.Shape shpTimer 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         Height          =   1305
         Left            =   1800
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   3705
      End
      Begin VB.Label lblTimerSegundos 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   4740
         TabIndex        =   57
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label lblTimerMinutos 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   66
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1395
         Left            =   3120
         TabIndex        =   56
         Top             =   1020
         Width           =   1545
      End
      Begin VB.Image Image1 
         Height          =   1200
         Left            =   1800
         Picture         =   "frmCartel2011.frx":0000
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label lblTimerTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Período de Prueba"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   60
         TabIndex        =   55
         Top             =   300
         Width           =   7245
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Presentes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   795
         Left            =   7740
         TabIndex        =   54
         Top             =   1020
         Width           =   3675
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ausentes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   795
         Left            =   7740
         TabIndex        =   53
         Top             =   1770
         Width           =   3525
      End
      Begin VB.Label lblTimerPresentes 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   795
         Left            =   13560
         TabIndex        =   52
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label lblTimerAusentes 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   795
         Left            =   13560
         TabIndex        =   51
         Top             =   1770
         Width           =   1215
      End
   End
   Begin VB.Label lblRecintoTentativo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Distribución de Bancas por Bloques"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   120
      TabIndex        =   49
      Top             =   4080
      Visible         =   0   'False
      Width           =   15225
   End
   Begin VB.Label lblTextoExtra 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4155
      Left            =   720
      TabIndex        =   36
      Top             =   4260
      Width           =   14535
   End
   Begin VB.Shape shpRivas 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   4575
      Left            =   420
      Shape           =   4  'Rounded Rectangle
      Top             =   4020
      Visible         =   0   'False
      Width           =   14835
   End
   Begin VB.Label lblOrador04 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   35
      Top             =   10800
      Visible         =   0   'False
      Width           =   14265
   End
   Begin VB.Label lblOrador03 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   34
      Top             =   10260
      Visible         =   0   'False
      Width           =   14265
   End
   Begin VB.Label lblOrador02 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   33
      Top             =   9720
      Visible         =   0   'False
      Width           =   14265
   End
   Begin VB.Label lblOrador01 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   600
      TabIndex        =   32
      Top             =   9120
      Visible         =   0   'False
      Width           =   14265
   End
   Begin VB.Label lblPaseListaFinalizado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASE DE LISTA FINALIZADO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   33.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1515
      Left            =   9090
      TabIndex        =   31
      Top             =   4110
      Visible         =   0   'False
      Width           =   5325
   End
   Begin VB.Label lblEmpate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EMPATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   8790
      TabIndex        =   30
      Top             =   4050
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblResultado 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EMPATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   915
      Left            =   7170
      TabIndex        =   15
      Top             =   4020
      Visible         =   0   'False
      Width           =   7185
   End
   Begin VB.Label lblTiempo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "15s"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10590
      TabIndex        =   29
      Top             =   4950
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Label lblLeyendaTiempo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TIEMPO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   9600
      TabIndex        =   28
      Top             =   4050
      Visible         =   0   'False
      Width           =   4155
   End
   Begin VB.Label lblLeyendaAbstenciones 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Abs."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   13500
      TabIndex        =   27
      Top             =   5010
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblLeyendaNegativos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Neg."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   11670
      TabIndex        =   26
      Top             =   5010
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblLeyendaAfirmativos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Afir."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   9870
      TabIndex        =   25
      Top             =   4980
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblMayoria 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Más de 1/2 de los V. Emitidos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   795
      Left            =   390
      TabIndex        =   24
      Top             =   5490
      Visible         =   0   'False
      Width           =   7185
   End
   Begin VB.Label lblTipoVotacion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VOTACION NOMINAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   33.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1515
      Left            =   1350
      TabIndex        =   23
      Top             =   3990
      Visible         =   0   'False
      Width           =   5325
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   256
      Left            =   12840
      Shape           =   3  'Circle
      Top             =   10590
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   255
      Left            =   12840
      Shape           =   3  'Circle
      Top             =   10320
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   254
      Left            =   12780
      Shape           =   3  'Circle
      Top             =   10050
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   253
      Left            =   12750
      Shape           =   3  'Circle
      Top             =   9780
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   252
      Left            =   12660
      Shape           =   3  'Circle
      Top             =   9510
      Width           =   315
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   251
      Left            =   12600
      Shape           =   3  'Circle
      Top             =   9240
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   250
      Left            =   12450
      Shape           =   3  'Circle
      Top             =   8970
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   249
      Left            =   12330
      Shape           =   3  'Circle
      Top             =   8700
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   248
      Left            =   12150
      Shape           =   3  'Circle
      Top             =   8460
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   247
      Left            =   12000
      Shape           =   3  'Circle
      Top             =   8220
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   246
      Left            =   11790
      Shape           =   3  'Circle
      Top             =   7980
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   245
      Left            =   11580
      Shape           =   3  'Circle
      Top             =   7770
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   244
      Left            =   11370
      Shape           =   3  'Circle
      Top             =   7560
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   243
      Left            =   11130
      Shape           =   3  'Circle
      Top             =   7380
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   242
      Left            =   10650
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   241
      Left            =   10350
      Shape           =   3  'Circle
      Top             =   6900
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   240
      Left            =   10080
      Shape           =   3  'Circle
      Top             =   6780
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   239
      Left            =   9780
      Shape           =   3  'Circle
      Top             =   6630
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   238
      Left            =   9510
      Shape           =   3  'Circle
      Top             =   6540
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   237
      Left            =   9150
      Shape           =   3  'Circle
      Top             =   6420
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   236
      Left            =   8850
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   235
      Left            =   8460
      Shape           =   3  'Circle
      Top             =   6300
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   234
      Left            =   8190
      Shape           =   3  'Circle
      Top             =   6270
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   233
      Left            =   7830
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   232
      Left            =   7500
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   231
      Left            =   7170
      Shape           =   3  'Circle
      Top             =   6270
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   230
      Left            =   6870
      Shape           =   3  'Circle
      Top             =   6300
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   229
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   228
      Left            =   6180
      Shape           =   3  'Circle
      Top             =   6420
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   227
      Left            =   5850
      Shape           =   3  'Circle
      Top             =   6510
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   226
      Left            =   5580
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   225
      Left            =   5220
      Shape           =   3  'Circle
      Top             =   6750
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   224
      Left            =   4950
      Shape           =   3  'Circle
      Top             =   6870
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   223
      Left            =   4650
      Shape           =   3  'Circle
      Top             =   7050
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   222
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   7380
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   221
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   7560
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   220
      Left            =   3750
      Shape           =   3  'Circle
      Top             =   7740
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   219
      Left            =   3540
      Shape           =   3  'Circle
      Top             =   7950
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   218
      Left            =   3330
      Shape           =   3  'Circle
      Top             =   8220
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   217
      Left            =   3180
      Shape           =   3  'Circle
      Top             =   8430
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   216
      Left            =   3030
      Shape           =   3  'Circle
      Top             =   8670
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   215
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   8910
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   214
      Left            =   2730
      Shape           =   3  'Circle
      Top             =   9210
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   213
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   9480
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   212
      Left            =   2550
      Shape           =   3  'Circle
      Top             =   9750
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   211
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   9990
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   210
      Left            =   2460
      Shape           =   3  'Circle
      Top             =   10320
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   209
      Left            =   2460
      Shape           =   3  'Circle
      Top             =   10560
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   208
      Left            =   12180
      Shape           =   3  'Circle
      Top             =   11160
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   207
      Left            =   12180
      Shape           =   3  'Circle
      Top             =   10920
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   206
      Left            =   12180
      Shape           =   3  'Circle
      Top             =   10620
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   205
      Left            =   12180
      Shape           =   3  'Circle
      Top             =   10380
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   204
      Left            =   12150
      Shape           =   3  'Circle
      Top             =   10080
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   203
      Left            =   12090
      Shape           =   3  'Circle
      Top             =   9810
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   202
      Left            =   12000
      Shape           =   3  'Circle
      Top             =   9540
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   201
      Left            =   11880
      Shape           =   3  'Circle
      Top             =   9270
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   200
      Left            =   11730
      Shape           =   3  'Circle
      Top             =   8940
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   199
      Left            =   11580
      Shape           =   3  'Circle
      Top             =   8700
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   198
      Left            =   11400
      Shape           =   3  'Circle
      Top             =   8460
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   197
      Left            =   11220
      Shape           =   3  'Circle
      Top             =   8250
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   196
      Left            =   10950
      Shape           =   3  'Circle
      Top             =   8010
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   195
      Left            =   10710
      Shape           =   3  'Circle
      Top             =   7800
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   194
      Left            =   10260
      Shape           =   3  'Circle
      Top             =   7500
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   193
      Left            =   10020
      Shape           =   3  'Circle
      Top             =   7350
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   192
      Left            =   9690
      Shape           =   3  'Circle
      Top             =   7200
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   191
      Left            =   9420
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   190
      Left            =   9090
      Shape           =   3  'Circle
      Top             =   6990
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   189
      Left            =   8820
      Shape           =   3  'Circle
      Top             =   6930
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   188
      Left            =   8490
      Shape           =   3  'Circle
      Top             =   6870
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   187
      Left            =   8190
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape shpBanca 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   186
      Left            =   7830
      Shape           =   3  'Circle
      Top             =   6810
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   185
      Left            =   7530
      Shape           =   3  'Circle
      Top             =   6810
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   184
      Left            =   7170
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   183
      Left            =   6870
      Shape           =   3  'Circle
      Top             =   6870
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   182
      Left            =   6510
      Shape           =   3  'Circle
      Top             =   6930
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   181
      Left            =   6210
      Shape           =   3  'Circle
      Top             =   6990
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   180
      Left            =   5910
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   179
      Left            =   5610
      Shape           =   3  'Circle
      Top             =   7200
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   178
      Left            =   5310
      Shape           =   3  'Circle
      Top             =   7350
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   177
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   7500
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   176
      Left            =   4620
      Shape           =   3  'Circle
      Top             =   7830
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   175
      Left            =   4380
      Shape           =   3  'Circle
      Top             =   8010
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   174
      Left            =   4110
      Shape           =   3  'Circle
      Top             =   8220
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   173
      Left            =   3900
      Shape           =   3  'Circle
      Top             =   8460
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   172
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   8700
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   171
      Left            =   3570
      Shape           =   3  'Circle
      Top             =   8940
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   170
      Left            =   3420
      Shape           =   3  'Circle
      Top             =   9240
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   169
      Left            =   3330
      Shape           =   3  'Circle
      Top             =   9480
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   168
      Left            =   3210
      Shape           =   3  'Circle
      Top             =   9810
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   167
      Left            =   3150
      Shape           =   3  'Circle
      Top             =   10080
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   166
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   10380
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   165
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   10620
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   164
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   10920
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   163
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   11160
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   162
      Left            =   11700
      Shape           =   3  'Circle
      Top             =   11160
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   161
      Left            =   11700
      Shape           =   3  'Circle
      Top             =   10920
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   160
      Left            =   11700
      Shape           =   3  'Circle
      Top             =   10650
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   159
      Left            =   11700
      Shape           =   3  'Circle
      Top             =   10410
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   158
      Left            =   11670
      Shape           =   3  'Circle
      Top             =   10110
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   157
      Left            =   11610
      Shape           =   3  'Circle
      Top             =   9840
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   156
      Left            =   11490
      Shape           =   3  'Circle
      Top             =   9540
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   155
      Left            =   11370
      Shape           =   3  'Circle
      Top             =   9270
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   154
      Left            =   11220
      Shape           =   3  'Circle
      Top             =   9000
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   153
      Left            =   11040
      Shape           =   3  'Circle
      Top             =   8760
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   152
      Left            =   10830
      Shape           =   3  'Circle
      Top             =   8520
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   151
      Left            =   10620
      Shape           =   3  'Circle
      Top             =   8310
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   150
      Left            =   10410
      Shape           =   3  'Circle
      Top             =   8130
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   149
      Left            =   9960
      Shape           =   3  'Circle
      Top             =   7830
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   148
      Left            =   9660
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   147
      Left            =   9270
      Shape           =   3  'Circle
      Top             =   7500
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   146
      Left            =   8970
      Shape           =   3  'Circle
      Top             =   7410
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   145
      Left            =   8550
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   144
      Left            =   8220
      Shape           =   3  'Circle
      Top             =   7260
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   143
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   7230
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   142
      Left            =   7500
      Shape           =   3  'Circle
      Top             =   7230
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   141
      Left            =   7080
      Shape           =   3  'Circle
      Top             =   7260
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   140
      Left            =   6750
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   139
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   7410
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   138
      Left            =   6030
      Shape           =   3  'Circle
      Top             =   7530
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   137
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   136
      Left            =   5370
      Shape           =   3  'Circle
      Top             =   7830
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   135
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   8130
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   134
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   8310
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   133
      Left            =   4470
      Shape           =   3  'Circle
      Top             =   8550
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   132
      Left            =   4260
      Shape           =   3  'Circle
      Top             =   8790
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   131
      Left            =   4080
      Shape           =   3  'Circle
      Top             =   9000
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   130
      Left            =   3930
      Shape           =   3  'Circle
      Top             =   9300
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   129
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   9540
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   128
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   9810
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   127
      Left            =   3660
      Shape           =   3  'Circle
      Top             =   10080
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   126
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   10410
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   125
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   10650
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   124
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   10920
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   123
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   11160
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   122
      Left            =   11070
      Shape           =   3  'Circle
      Top             =   11250
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   121
      Left            =   11070
      Shape           =   3  'Circle
      Top             =   11010
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   120
      Left            =   11070
      Shape           =   3  'Circle
      Top             =   10740
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   119
      Left            =   11040
      Shape           =   3  'Circle
      Top             =   10440
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   118
      Left            =   10980
      Shape           =   3  'Circle
      Top             =   10140
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   117
      Left            =   10920
      Shape           =   3  'Circle
      Top             =   9870
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   116
      Left            =   10800
      Shape           =   3  'Circle
      Top             =   9570
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   115
      Left            =   10650
      Shape           =   3  'Circle
      Top             =   9300
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   114
      Left            =   10440
      Shape           =   3  'Circle
      Top             =   9030
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   113
      Left            =   10230
      Shape           =   3  'Circle
      Top             =   8790
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   112
      Left            =   10020
      Shape           =   3  'Circle
      Top             =   8610
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   111
      Left            =   9600
      Shape           =   3  'Circle
      Top             =   8280
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   110
      Left            =   9330
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   109
      Left            =   9060
      Shape           =   3  'Circle
      Top             =   8040
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   108
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   7920
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   107
      Left            =   8460
      Shape           =   3  'Circle
      Top             =   7860
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   106
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   7800
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   105
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   7770
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   104
      Left            =   7500
      Shape           =   3  'Circle
      Top             =   7770
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   103
      Left            =   7140
      Shape           =   3  'Circle
      Top             =   7800
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   102
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   7830
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   101
      Left            =   6540
      Shape           =   3  'Circle
      Top             =   7920
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   100
      Left            =   6270
      Shape           =   3  'Circle
      Top             =   8010
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   99
      Left            =   5970
      Shape           =   3  'Circle
      Top             =   8130
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   98
      Left            =   5700
      Shape           =   3  'Circle
      Top             =   8280
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   97
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   8580
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   96
      Left            =   5070
      Shape           =   3  'Circle
      Top             =   8790
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   95
      Left            =   4830
      Shape           =   3  'Circle
      Top             =   9060
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   94
      Left            =   4650
      Shape           =   3  'Circle
      Top             =   9300
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   93
      Left            =   4500
      Shape           =   3  'Circle
      Top             =   9570
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   92
      Left            =   4380
      Shape           =   3  'Circle
      Top             =   9870
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   91
      Left            =   4290
      Shape           =   3  'Circle
      Top             =   10140
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   90
      Left            =   4260
      Shape           =   3  'Circle
      Top             =   10440
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   89
      Left            =   4230
      Shape           =   3  'Circle
      Top             =   10740
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   88
      Left            =   4230
      Shape           =   3  'Circle
      Top             =   11010
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   87
      Left            =   4230
      Shape           =   3  'Circle
      Top             =   11250
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   86
      Left            =   10560
      Shape           =   3  'Circle
      Top             =   11220
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   85
      Left            =   10560
      Shape           =   3  'Circle
      Top             =   10980
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   84
      Left            =   10560
      Shape           =   3  'Circle
      Top             =   10740
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   83
      Left            =   10560
      Shape           =   3  'Circle
      Top             =   10470
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   82
      Left            =   10470
      Shape           =   3  'Circle
      Top             =   10110
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   81
      Left            =   10380
      Shape           =   3  'Circle
      Top             =   9840
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   80
      Left            =   10230
      Shape           =   3  'Circle
      Top             =   9570
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   79
      Left            =   10080
      Shape           =   3  'Circle
      Top             =   9330
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   78
      Left            =   9840
      Shape           =   3  'Circle
      Top             =   9060
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   77
      Left            =   9630
      Shape           =   3  'Circle
      Top             =   8880
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   76
      Left            =   9210
      Shape           =   3  'Circle
      Top             =   8610
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   75
      Left            =   8940
      Shape           =   3  'Circle
      Top             =   8460
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   74
      Left            =   8550
      Shape           =   3  'Circle
      Top             =   8310
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   73
      Left            =   8250
      Shape           =   3  'Circle
      Top             =   8250
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   72
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   8220
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   71
      Left            =   7500
      Shape           =   3  'Circle
      Top             =   8220
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   70
      Left            =   7080
      Shape           =   3  'Circle
      Top             =   8250
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   69
      Left            =   6780
      Shape           =   3  'Circle
      Top             =   8310
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   68
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   8490
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   67
      Left            =   6060
      Shape           =   3  'Circle
      Top             =   8610
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   66
      Left            =   5670
      Shape           =   3  'Circle
      Top             =   8880
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   65
      Left            =   5430
      Shape           =   3  'Circle
      Top             =   9060
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   64
      Left            =   5220
      Shape           =   3  'Circle
      Top             =   9330
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   63
      Left            =   5070
      Shape           =   3  'Circle
      Top             =   9570
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   62
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   9870
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   61
      Left            =   4830
      Shape           =   3  'Circle
      Top             =   10110
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   60
      Left            =   4770
      Shape           =   3  'Circle
      Top             =   10470
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   59
      Left            =   4740
      Shape           =   3  'Circle
      Top             =   10740
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   58
      Left            =   4740
      Shape           =   3  'Circle
      Top             =   10980
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   57
      Left            =   4740
      Shape           =   3  'Circle
      Top             =   11220
      Width           =   285
   End
   Begin VB.Label lblFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   480
      TabIndex        =   22
      Top             =   150
      Width           =   2865
   End
   Begin VB.Label lblLeyendaSinIdentificar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No Ident."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   12480
      TabIndex        =   21
      Top             =   7080
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Label lblSinIdentificar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   795
      Left            =   12480
      TabIndex        =   20
      Top             =   7650
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Título de prueba para Cartel - Expediente Nº 7876 - Año 2018 - Pruebas 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   600
      TabIndex        =   19
      Top             =   2970
      Visible         =   0   'False
      Width           =   14325
   End
   Begin VB.Label lblAbstenciones 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   675
      Left            =   13500
      TabIndex        =   18
      Top             =   5580
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblNegativos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   11670
      TabIndex        =   17
      Top             =   5580
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblAfirmativos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "201"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   675
      Left            =   9840
      TabIndex        =   16
      Top             =   5580
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   56
      Left            =   9900
      Shape           =   3  'Circle
      Top             =   11220
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   55
      Left            =   9900
      Shape           =   3  'Circle
      Top             =   10980
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   54
      Left            =   9900
      Shape           =   3  'Circle
      Top             =   10710
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   53
      Left            =   9870
      Shape           =   3  'Circle
      Top             =   10440
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   52
      Left            =   9810
      Shape           =   3  'Circle
      Top             =   10170
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   51
      Left            =   9690
      Shape           =   3  'Circle
      Top             =   9870
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   50
      Left            =   9540
      Shape           =   3  'Circle
      Top             =   9600
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   49
      Left            =   9300
      Shape           =   3  'Circle
      Top             =   9330
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   48
      Left            =   8850
      Shape           =   3  'Circle
      Top             =   9030
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   47
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   8880
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   46
      Left            =   8190
      Shape           =   3  'Circle
      Top             =   8790
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   45
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   8760
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   44
      Left            =   7500
      Shape           =   3  'Circle
      Top             =   8760
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   43
      Left            =   7110
      Shape           =   3  'Circle
      Top             =   8790
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   42
      Left            =   6780
      Shape           =   3  'Circle
      Top             =   8880
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   41
      Left            =   6420
      Shape           =   3  'Circle
      Top             =   9030
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   40
      Left            =   5970
      Shape           =   3  'Circle
      Top             =   9330
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   39
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   9570
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   38
      Left            =   5580
      Shape           =   3  'Circle
      Top             =   9870
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   37
      Left            =   5490
      Shape           =   3  'Circle
      Top             =   10140
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   36
      Left            =   5430
      Shape           =   3  'Circle
      Top             =   10440
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   35
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   10710
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   34
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   10980
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   33
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   11220
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   32
      Left            =   9390
      Shape           =   3  'Circle
      Top             =   11160
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   31
      Left            =   9390
      Shape           =   3  'Circle
      Top             =   10920
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   30
      Left            =   9390
      Shape           =   3  'Circle
      Top             =   10680
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   29
      Left            =   9360
      Shape           =   3  'Circle
      Top             =   10350
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   28
      Left            =   9210
      Shape           =   3  'Circle
      Top             =   9990
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   27
      Left            =   9030
      Shape           =   3  'Circle
      Top             =   9750
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   26
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   9360
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   25
      Left            =   8220
      Shape           =   3  'Circle
      Top             =   9240
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   24
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   9180
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   23
      Left            =   7500
      Shape           =   3  'Circle
      Top             =   9180
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   22
      Left            =   7080
      Shape           =   3  'Circle
      Top             =   9240
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   21
      Left            =   6780
      Shape           =   3  'Circle
      Top             =   9360
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   20
      Left            =   6270
      Shape           =   3  'Circle
      Top             =   9750
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   19
      Left            =   6090
      Shape           =   3  'Circle
      Top             =   9990
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   18
      Left            =   5940
      Shape           =   3  'Circle
      Top             =   10350
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   17
      Left            =   5910
      Shape           =   3  'Circle
      Top             =   10680
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   16
      Left            =   5910
      Shape           =   3  'Circle
      Top             =   10920
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   15
      Left            =   5910
      Shape           =   3  'Circle
      Top             =   11160
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   14
      Left            =   8730
      Shape           =   3  'Circle
      Top             =   11190
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   13
      Left            =   8730
      Shape           =   3  'Circle
      Top             =   10950
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   12
      Left            =   8730
      Shape           =   3  'Circle
      Top             =   10710
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   11
      Left            =   8640
      Shape           =   3  'Circle
      Top             =   10350
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   10
      Left            =   8460
      Shape           =   3  'Circle
      Top             =   10080
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   9
      Left            =   8190
      Shape           =   3  'Circle
      Top             =   9870
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   8
      Left            =   7830
      Shape           =   3  'Circle
      Top             =   9750
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   7
      Left            =   7470
      Shape           =   3  'Circle
      Top             =   9750
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   6
      Left            =   7110
      Shape           =   3  'Circle
      Top             =   9870
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   5
      Left            =   6810
      Shape           =   3  'Circle
      Top             =   10110
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   4
      Left            =   6630
      Shape           =   3  'Circle
      Top             =   10380
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   3
      Left            =   6540
      Shape           =   3  'Circle
      Top             =   10710
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   2
      Left            =   6540
      Shape           =   3  'Circle
      Top             =   10950
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   1
      Left            =   6540
      Shape           =   3  'Circle
      Top             =   11190
      Width           =   285
   End
   Begin VB.Shape shpBanca 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   0
      Left            =   7620
      Shape           =   3  'Circle
      Top             =   10740
      Width           =   285
   End
   Begin VB.Label lblAusentes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   795
      Left            =   12660
      TabIndex        =   14
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label lblPresentes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   795
      Left            =   9360
      TabIndex        =   13
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label lblLeyendaAusentes 
      BackStyle       =   0  'Transparent
      Caption         =   "Ausentes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   12180
      TabIndex        =   12
      Top             =   1320
      Width           =   2505
   End
   Begin VB.Label lblLeyendaPresentes 
      BackStyle       =   0  'Transparent
      Caption         =   "Presentes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   8640
      TabIndex        =   11
      Top             =   1320
      Width           =   2475
   End
   Begin VB.Label lblLeyendaReunion 
      BackStyle       =   0  'Transparent
      Caption         =   "Reunión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   2040
      TabIndex        =   10
      Top             =   2850
      Width           =   2025
   End
   Begin VB.Label lblTipoSesion 
      BackStyle       =   0  'Transparent
      Caption         =   "Sesión - Prueba"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   2040
      TabIndex        =   9
      Top             =   2070
      Width           =   5985
   End
   Begin VB.Label lblTipoPeriodo 
      BackStyle       =   0  'Transparent
      Caption         =   "Período de Prueba"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   2040
      TabIndex        =   8
      Top             =   1320
      Width           =   6075
   End
   Begin VB.Label lblSeparadorReunion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   1560
      TabIndex        =   7
      Top             =   2820
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblSeparadorSesion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   1560
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblSeparacionPeriodo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   1560
      TabIndex        =   5
      Top             =   1290
      Width           =   345
   End
   Begin VB.Label lblNumeroReunion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   450
      TabIndex        =   4
      Top             =   2850
      Width           =   1065
   End
   Begin VB.Label lblNumeroSesion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   450
      TabIndex        =   3
      Top             =   2070
      Width           =   1065
   End
   Begin VB.Label lblNumeroPeriodo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   450
      TabIndex        =   2
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Shape shpQuorum 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   1005
      Left            =   8040
      Shape           =   4  'Rounded Rectangle
      Top             =   90
      Width           =   7185
   End
   Begin VB.Label lblQuorum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QUORUM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   795
      Left            =   8220
      TabIndex        =   1
      Top             =   150
      Width           =   6855
   End
   Begin VB.Label lblHora 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   5520
      TabIndex        =   0
      Top             =   150
      Width           =   1905
   End
   Begin VB.Shape shpHora 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1005
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   90
      Width           =   7185
   End
   Begin VB.Image imgBancas 
      Height          =   11520
      Left            =   0
      Picture         =   "frmCartel2011.frx":0B58
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "frmCartel2011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Verde = &HFF00&
Const Rojo = &HFF&
Const Celeste = &HFFFF00
Const Amarillo = &HFFFF&
Const Blanco = &HFFFFFF
Const Gris = &H404040
Public renglonIndex As Integer
Public leftVot As Integer
Public lPage As Integer
Public WebSenderTick As Long
Private svCuilsArray(6000) As String
Private svBancasArray(6000) As String
Private svTotal As String
Private svLastSent As String

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If (KeyAscii = 76 Or KeyAscii = 108) Then
        Dim r As Integer
        r = MsgBox("¿Cambiar a modo Asamblea Legislativa?", vbYesNo)
        If r = vbYes Then
            frmAsamblea.Show
            Me.Hide
        End If
    End If
    If KeyAscii = 27 Then
        Dim rx As Integer
        rx = MsgBox("¿Desea cerrar el servidor?", vbYesNo)
        If (rx = vbYes) Then
            End
        End If
    End If
    If KeyAscii = 118 Then 'V
        If (lblTextoExtra.Caption = "") Then
            Call ReadFromText
        End If
        If (shpRivas.Visible = True) Then
            shpRivas.Visible = False
            lblTextoExtra.Visible = False
            lblTitulo.Visible = False
            lblTipoVotacion.Visible = False
            lblTipoVotacion.Left = leftVot
        Else
            leftVot = lblTipoVotacion.Left
            lblTipoVotacion.Left = -7000
            shpRivas.Visible = True
            lblTextoExtra.Visible = True
            lblTitulo.Visible = True
            lblTipoVotacion.Visible = True
        End If
    End If
    If KeyAscii = 113 Then 'q
        'If renglonIndex > 0 Then
        '    renglonIndex = renglonIndex - 1
        '    Call ActualizarRenglones
        'End If
        Call PaginaAnterior
    End If
    If KeyAscii = 119 Then 'w
        'If (renglonIndex < UBound(renglonesExtra)) Then
        '    renglonIndex = renglonIndex + 1
        '    Call ActualizarRenglones
        'End If
        Call SiguientePagina
    End If
    If KeyAscii = 114 Then
        RecintoTentativo.modoTentativo = Not RecintoTentativo.modoTentativo
    End If
    If KeyAscii = 122 Then 'Z - Empieza servicio de Node
        svLastSent = ""
        Call RefreshNodeData
        Me.tmSqvData.Enabled = True
    End If
    If KeyAscii = 120 Then ' Detiene servicio de Node
        svLastSent = ""
        Me.tmSqvData.Enabled = False
    End If
End Sub

Private Sub RefreshNodeData()
On Error Resume Next
Dim rs As New ADODB.Recordset
frmMain.SetearRsAux "SELECT legisladores_activos.id, legisladores_activos.apellido, legisladores_activos.nombre, DiputadosCuil.cuil AS cuil, ISNULL(BancasProbables.banca, 300) AS banca FROM legisladores_activos LEFT JOIN DiputadosCuil ON DiputadosCuil.id = legisladores_activos.id LEFT JOIN BancasProbables ON BancasProbables.id_legislador = legisladores_activos.id", rs
While Not rs.EOF
    Dim svId As Integer
    Dim svCuil As String
    Dim svBanca As String
    svId = rs.Fields(0)
    svCuil = rs.Fields(3)
    svBanca = rs.Fields(4)
    svCuilsArray(svId) = svCuil
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
'Obtengo Bancas + IDs probables
Set rs = New ADODB.Recordset
frmMain.SetearRsAux "SELECT BancasProbables.id_legislador as id_legislador, BancasProbables.banca AS banca FROM BancasProbables WHERE BancasProbables.id_legislador IN (SELECT id FROM legisladores_activos)", rs
While Not rs.EOF
    Dim tsvId As Integer
    Dim tsvBanca As Integer
    tsvId = rs.Fields(0)
    tsvBanca = rs.Fields(1)
    svBancasArray(tsvBanca) = tsvId
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
End Sub

Private Sub SiguientePagina()
Dim i As Integer
Dim offToEnd As Integer
Dim offset As Integer
Dim nIndex As Integer
If (renglonIndex <> UBound(renglonesExtra)) Then
    nIndex = renglonIndex
    If (nIndex <> 0) Then
        nIndex = nIndex + 1
    End If
    offToEnd = UBound(renglonesExtra) - renglonIndex
    If (offToEnd > 7) Then
        If (nIndex <> 0) Then
            renglonIndex = renglonIndex + 8
        Else
            renglonIndex = renglonIndex + 7
        End If
    Else
        renglonIndex = renglonIndex + offToEnd
    End If
    Dim total As String
    For i = (nIndex) To renglonIndex
        total = total & renglonesExtra(i) & vbCrLf
    Next i
    If (renglonIndex = UBound(renglonesExtra)) Then
        lPage = offToEnd
    End If
    lblTextoExtra.Caption = total
End If
End Sub

Private Sub PaginaAnterior()
Dim lineas As Integer
Dim linBegg As Integer
If (renglonIndex = UBound(renglonesExtra)) Then
    linBegg = renglonIndex
    renglonIndex = renglonIndex - lPage
Else
    If (renglonIndex > 7) Then
        renglonIndex = renglonIndex - 8 'La linea de fin
    End If
End If
linBegg = renglonIndex - 7
Dim total As String
Dim i As Integer
For i = linBegg To renglonIndex
    total = total & renglonesExtra(i) & vbCrLf
Next i
lblTextoExtra.Caption = total
End Sub

Private Sub ActualizarRenglones()
On Error Resume Next
Dim i As Integer
Dim offToEnd As Integer
Dim offset As Integer
offToEnd = UBound(renglonesExtra) - renglonIndex
If (offToEnd < 7) Then
    offset = offToEnd
Else
    offset = 7
End If
Dim total As String
For i = renglonIndex To (renglonIndex + offset)
    total = total & renglonesExtra(i) & vbCrLf
Next i
lblTextoExtra.Caption = total
End Sub

Private Sub ReadFromText()
On Error Resume Next
Dim maxChars As Integer
maxChars = 60
renglonIndex = 0
lblTextoExtra.Caption = ""
Dim total As String
Dim linea As String
Dim letra As String
Dim arr() As String
Open App.Path & "\nota.txt" For Input As #1
While Not EOF(1)
    Line Input #1, linea
    total = total & linea
Wend
Close #1
Dim i As Integer
Dim actualChars As Integer
Dim renglon As String
renglon = ""
i = 0
While total <> ""
    letra = Left(total, 1)
    renglon = renglon & letra
    total = Right(total, Len(total) - 1)
    If (letra = ":") Then
        i = i + 1
        ReDim Preserve arr(i)
        arr(i - 1) = renglon
        renglon = ""
    Else
        If Len(renglon) >= maxChars Or total = "" Then
            If (total <> "") Then
                If (letra = " " Or letra = "." Or letra = ":") Then
                    i = i + 1
                    ReDim Preserve arr(i)
                    arr(i - 1) = renglon
                    renglon = ""
                End If
            Else
                i = i + 1
                ReDim Preserve arr(i)
                arr(i - 1) = renglon
                renglon = ""
            End If
    
        End If
        letra = ""
    End If
Wend
renglonesExtra = arr
End Sub

Private Sub Form_Load()
RecintoTentativo.modoTentativo = False
Call cargarRecintoTentativo
VL.modoExtendido = False
Pendientes.paginaActualPendientes = 0
MostrarCartel = True
frmMain.EjecutaSQLCartel "UPDATE config SET Mostrar_Recinto = 1"
PrimerControlLarga = False
lblOrador02.ForeColor = &HFFFF&
lblOrador03.ForeColor = &HFFFF&
lblOrador04.ForeColor = &HFFFF&

Call StartProtoServer
Call StartWebSender
'Call LoadVotoRemoto
Dim n As Integer
n = VotoRemoto.getPresentes()
End Sub

Private Sub cargarRecintoTentativo()
Dim i As Integer
i = -1
Open App.Path & "\recintotentativo.csv" For Input As #1
While Not EOF(1)
    Dim linea As String
    linea = ""
    Line Input #1, linea
    i = i + 1
    linea = Replace(linea, "#", "")
    Dim p1 As Variant
    Dim p2 As Variant
    Dim p3 As Variant
    Dim l As Long
    p1 = CVar("&H" & Mid(linea, 5, 2))
    p2 = CVar("&H" & Mid(linea, 3, 2))
    p3 = CVar("&H" & Mid(linea, 1, 2))
    l = RGB(CInt(p3), CInt(p2), CInt(p1))
    RecintoTentativo.bancasTentativas(i) = l
Wend
Close #1
End Sub

Private Sub ProtoServer_ConnectionRequest(ByVal requestID As Long)
ProtoServer.Close
ProtoServer.Accept requestID
End Sub

Private Sub ProtoServer_DataArrival(ByVal bytesTotal As Long)
Dim buffer As String
Dim arr() As String
ProtoServer.GetData buffer
arr = Split(buffer, ";")
If (arr(0) = "t") Then
    If Not Trim(EstadoActual.TipoDeOperacion) = "quorum" Then
        Me.frameTimer.Visible = False
    Else
        Me.frameTimer.Visible = True
    End If
    Me.lblTimerMinutos.Caption = arr(2)
    Me.lblTimerSegundos.Caption = arr(3)
    If (arr(1) = "+") Then
        If (arr(2) = "0" And Val(arr(3)) <= 15) Then
            shpTimer.BorderColor = Amarillo
            lblTimerMinutos.ForeColor = Amarillo
            lblTimerSegundos.ForeColor = Amarillo
        Else
            shpTimer.BorderColor = Blanco
            lblTimerMinutos.ForeColor = Blanco
            lblTimerSegundos.ForeColor = Blanco
        End If
    Else
        shpTimer.BorderColor = Rojo
        lblTimerMinutos.ForeColor = Rojo
        lblTimerSegundos.ForeColor = Rojo
    End If
ElseIf (arr(0) = "i") Then
    Me.lblTimerMinutos.Caption = ""
    Me.lblTimerSegundos.Caption = ""
    Me.frameTimer.Visible = False
End If
Call StartProtoServer
End Sub

Private Sub svWinsock_Connect()
On Error Resume Next
Dim Pack As String
svLastSent = svTotal
Pack = "POST / HTTP/1.1" & vbCrLf
Pack = Pack & "Host: 10.1.1.8:3001" & vbCrLf
Pack = Pack & "Content-Type: text/plain" & vbCrLf
Pack = Pack & "Content-Length: " & Len(svTotal) & vbCrLf & vbCrLf & svTotal
svWinsock.SendData Pack
End Sub

Private Sub svWinsock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim data As String
svWinsock.GetData data
End Sub

Private Sub tmActualizar_Timer()
Dim i As Integer
Dim Mostrar_Progreso As Boolean
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
frmMain.SetearRsAux "SELECT Mostrar_Progreso_Votacion FROM config", rsTemp
If rsTemp.EOF Then
    MsgBox "Error de Integridad"
Else
    If Not IsNull(rsTemp.Fields(0)) Then
        If rsTemp.Fields(0) = 1 Then
            Mostrar_Progreso = True
        Else
            Mostrar_Progreso = False
        End If
    Else
        Mostrar_Progreso = False
    End If
End If
rsTemp.Close
Set rsTemp = Nothing
If EstadoActual.TipoDeOperacion = "quorum" Then
    If (ProtoServer.State = sckClosed Or ProtoServer.State = sckError) Then
        StartProtoServer
    End If
    If (RecintoTentativo.modoTentativo) Then
        If Not (imgBancas.Visible) Then
            lblRecintoTentativo.Visible = True
            For i = 0 To 256
                shpBanca(i).Visible = True
                shpBanca(i).FillColor = RecintoTentativo.bancasTentativas(i)
                shpBanca(i).BorderColor = Blanco
            Next i
        End If
        shpBanca(0).Visible = True
        imgBancas.Visible = True
    Else
        lblRecintoTentativo.Visible = False
        shpBanca(0).Visible = False
        imgBancas.Visible = False
    End If
Else
    Me.frameTimer.Visible = False
    If EstadoActual.EstadoVotacion_y_PasList = "esperando" Then
        If Mostrar_Progreso = True Then
            shpBanca(0).Visible = True
            imgBancas.Visible = True
        Else
            shpBanca(0).Visible = False
            imgBancas.Visible = False
        End If
    End If
End If
If EstadoActual.EstadoVotacion_y_PasList = "empate" Then
    lblResultado.Visible = True
End If
If Not RecintoTentativo.modoTentativo Then
    For i = 0 To 256
        If EstadoActual.TipoDeOperacion = "quorum" Then
            shpBanca(i).Visible = False
            imgBancas.Visible = False
        ElseIf EstadoActual.EstadoVotacion_y_PasList = "larga" And VL.modoExtendido = True Then
            shpBanca(i).Visible = False
            imgBancas.Visible = False
            grpPendientes.Visible = True
        ElseIf EstadoActual.EstadoVotacion_y_PasList = "cancelada" And EstadoActual.TipoDeOperacion = "votnom" Then
            shpBanca(i).Visible = True
            imgBancas.Visible = True
            grpPendientes.Visible = False
            lblTiempo.Caption = ""
        Else
            If EstadoActual.EstadoVotacion_y_PasList = "votando" And EstadoActual.TipoDeOperacion <> "votnum" Then
                If Mostrar_Progreso = True Then
                    shpBanca(i).Visible = True
                    shpBanca(0).Visible = True
                    imgBancas.Visible = True
                Else
                    shpBanca(i).Visible = False
                    shpBanca(0).Visible = False
                    imgBancas.Visible = False
                End If
            End If
        End If
        If EstadoActual.EstadoVotacion_y_PasList = "votando" Or EstadoActual.EstadoVotacion_y_PasList = "larga" Then
            If Trim(EstadoActual.VectorResultados(i)) <> "" Then
                If lblTiempo.Visible = True Or VL.modoExtendido = True Then
                    With shpBanca(i)
                        .FillColor = Gris
                        .BorderColor = Gris
                    End With
                Else
                    Dim x23 As String
                    x23 = ""
                End If
            Else
                If EstadoActual.VectorIdentificacion(i) <> NO_IDENTIFICADO Then
                    With shpBanca(i)
                        .FillColor = Celeste
                        .BorderColor = Celeste
                    End With
                ElseIf EstadoActual.VectorPresencia(i) = "X" Or EstadoActual.VectorPresencia(i) = "0" Then
                    With shpBanca(i)
                        .FillColor = Blanco
                        .BorderColor = Blanco
                    End With
                Else
                    With shpBanca(i)
                        .FillColor = Amarillo
                        .BorderColor = Amarillo
                    End With
                End If
            End If
        ElseIf EstadoActual.EstadoVotacion_y_PasList = "cierre" Or EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate" Then
            grpPendientes.Visible = False
            If EstadoActual.TipoDeOperacion = "votnom" Then
                If shpBanca(0).Visible <> True Then
                    shpBanca(0).Visible = True
                End If
                If shpBanca(i).Visible <> True Then
                    shpBanca(i).Visible = True
                End If
                If imgBancas.Visible <> True Then
                    imgBancas.Visible = True
                End If
                If imgBancas.Visible <> True Then
                    imgBancas.Visible = True
                End If
                ActualizarColoresVotos
                lblTiempo.Caption = ""
                If EstadoActual.EstadoVotacion_y_PasList <> "empate" Then
                    ActualizarUnColor (0)
                End If
            Else
                shpBanca(0).Visible = False
                shpBanca(i).Visible = False
                imgBancas.Visible = False
            End If
        Else
            If Mostrar_Progreso = True And (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis") Then
                shpBanca(0).Visible = True
                shpBanca(i).Visible = True
                shpBanca(0).FillColor = Celeste
                shpBanca(0).BorderColor = Celeste
                imgBancas.Visible = True
                If EstadoActual.VectorIdentificacion(i) <> "0" Then
                    With shpBanca(i)
                        .FillColor = Celeste
                        .BorderColor = Celeste
                    End With
                Else
                   If EstadoActual.VectorPresencia(i) = "1" Then
                        With shpBanca(i)
                            .FillColor = Amarillo
                            .BorderColor = Amarillo
                        End With
                    Else
                        With shpBanca(i)
                            .FillColor = Blanco
                            .BorderColor = Blanco
                        End With
                    End If
                End If
            Else
                shpBanca(0).Visible = False
                shpBanca(i).Visible = False
                imgBancas.Visible = False
            End If
        End If
    Next i
End If
If grpPendientes.Visible = True Then
    Dim diputados() As Pendientes.DiputadoPendiente
    Dim top As Integer
    Dim low As Integer
    Dim tot As Integer
    tot = 0
    diputados = getPaginasPendientes()
    Dim d As Integer
    For d = LBound(diputados) To UBound(diputados)
        Dim dd As Integer
        For dd = 0 To 27
            If (diputados(d).diputado(dd) <> "") Then
                tot = tot + 1
            End If
        Next dd
    Next d
    top = UBound(diputados)
    low = LBound(diputados)
    lblPendientesCantidad.Caption = "Abstenciones: " & tot
    If Pendientes.paginaActualPendientes > top Then
        Pendientes.paginaActualPendientes = top
    ElseIf Pendientes.paginaActualPendientes < low Then
        Pendientes.paginaActualPendientes = low
    End If
    lblPendientes1.Caption = diputados(Pendientes.paginaActualPendientes).diputado(0)
    lblPendientes2.Caption = diputados(Pendientes.paginaActualPendientes).diputado(1)
    lblPendientes3.Caption = diputados(Pendientes.paginaActualPendientes).diputado(2)
    lblPendientes4.Caption = diputados(Pendientes.paginaActualPendientes).diputado(3)
    lblPendientes5.Caption = diputados(Pendientes.paginaActualPendientes).diputado(4)
    lblPendientes6.Caption = diputados(Pendientes.paginaActualPendientes).diputado(5)
    lblPendientes7.Caption = diputados(Pendientes.paginaActualPendientes).diputado(6)
    lblPendientes8.Caption = diputados(Pendientes.paginaActualPendientes).diputado(7)
    lblPendientes9.Caption = diputados(Pendientes.paginaActualPendientes).diputado(8)
    lblPendientes10.Caption = diputados(Pendientes.paginaActualPendientes).diputado(9)
End If
If Trim(frmMain.lblOrador02.Caption) <> "" And (Mostrar_Progreso = False Or EstadoActual.TipoDeOperacion = "quorum") Then
    lblOrador01.Visible = True
    lblOrador02.Visible = True
    lblOrador03.Visible = True
    lblOrador04.Visible = True
    lblOrador01.Caption = frmMain.lblOrador01.Caption
    lblOrador02.Caption = frmMain.lblOrador02.Caption
    lblOrador03.Caption = frmMain.lblOrador03.Caption
    lblOrador04.Caption = frmMain.lblOrador04.Caption
Else
    lblOrador01.Visible = False
    lblOrador02.Visible = False
    lblOrador03.Visible = False
    lblOrador04.Visible = False
End If
End Sub

Public Function GetGlobalData() As String
Dim s As String
Dim i As Integer
If shpBanca(0).FillColor = Verde Then
    s = s & "v"
ElseIf shpBanca(0).FillColor = Rojo Then
    s = s & "r"
ElseIf shpBanca(0).FillColor = Blanco Then
    s = s & "b"
ElseIf shpBanca(0).FillColor = Gris Then
    s = s & "g"
ElseIf shpBanca(0).FillColor = Celeste Then
    s = s & "c"
ElseIf shpBanca(0).FillColor = Amarillo Then
    s = s & "a"
End If
For i = 1 To 256
        If shpBanca(i).FillColor = Verde Then
            s = s & ";v"
        ElseIf shpBanca(i).FillColor = Rojo Then
            s = s & ";r"
        ElseIf shpBanca(i).FillColor = Blanco Then
            s = s & ";b"
        ElseIf shpBanca(i).FillColor = Gris Then
            s = s & ";g"
        ElseIf shpBanca(i).FillColor = Celeste Then
            s = s & ";c"
        ElseIf shpBanca(i).FillColor = Amarillo Then
            s = s & ";a"
        End If
Next i
Dim dp() As DiputadoPendiente
dp = Pendientes.getPaginasPendientes()
Dim lastPendIndex As Integer
lastPendIndex = -1
For i = LBound(dp) To UBound(dp)
    For z = 0 To 27
        If (Trim(dp(i).diputado(z)) = "") Then
            s = s & ";n"
        Else
            s = s & ";" & dp(i).diputado(z)
        End If
        lastPendIndex = lastPendIndex + 1
    Next z
Next i
For i = (lastPendIndex + 1) To 256
    s = s & ";n"
Next i
Dim a As String
Dim mTextTimer As String
mTextTimer = Trim(lblTiempo.Caption)
If (mTextTimer = "15s") Or (Not IsNumeric(mTextTimer)) Then
    mTextTimer = ""
End If
s = s & ";" & EstadoActual.TipoDeOperacion & ";" & Me.lblPresentes.Caption & ";" & Me.lblAusentes.Caption & ";" & Me.lblSinIdentificar.Caption & _
";" & Me.lblResultado.Caption & ";" & Me.lblAfirmativos.Caption & ";" & Me.lblNegativos.Caption & _
";" & Me.lblAbstenciones.Caption & ";" & Me.lblMayoria.Caption & ";" & mTextTimer & ";" & paginaActualPendientes
If (VL.modoExtendido) Then
    s = s & ";1"
Else
    s = s & ";0"
End If
s = s & ";" & Trim(EstadoActual.EstadoVotacion_y_PasList)
Dim arr() As String
arr = Split(s, ";")
If (UBound(arr) <> 525) Then
    s = s
End If
GetGlobalData = s
End Function

Public Sub ActualizarColoresVotos()
Dim i As Integer
If EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum" Then
    For i = 1 To 256
        If EstadoActual.VectorResultados(i) = AFIRMATIVO Then
            With shpBanca(i)
                .FillColor = Verde
                .BorderColor = Verde
            End With
        ElseIf EstadoActual.VectorResultados(i) = NEGATIVO Then
            With shpBanca(i)
                .FillColor = Rojo
                .BorderColor = Rojo
            End With
        ElseIf EstadoActual.VectorPresencia(i) = AUSENTE Then
            With shpBanca(i)
                .FillColor = Blanco
                .FillColor = Blanco
            End With
        End If
    Next i
End If
End Sub

Public Sub ActualizarUnColor(i As Integer)
If (i > 0) Then
    If EstadoActual.VectorResultados(i) = AFIRMATIVO Then
        With shpBanca(i)
            .FillColor = Verde
            .BorderColor = Verde
        End With
    ElseIf EstadoActual.VectorResultados(i) = NEGATIVO Then
        With shpBanca(i)
            .FillColor = Rojo
            .BorderColor = Rojo
        End With
    End If
Else
    If (EstadoActual.VectorResultados(i) <> " ") Then
        If EstadoActual.VectorResultados(i) = AFIRMATIVO Then
            With shpBanca(i)
                .FillColor = Verde
                .BorderColor = Verde
            End With
        ElseIf EstadoActual.VectorResultados(i) = NEGATIVO Then
            With shpBanca(i)
                .FillColor = Rojo
                .BorderColor = Rojo
            End With
        End If
    Else
        If EstadoActual.ResultadoVotoPresidente = AFIRMATIVO Then
            With shpBanca(i)
            .FillColor = Verde
            .BorderColor = Verde
            End With
        ElseIf EstadoActual.ResultadoVotoPresidente = NEGATIVO Then
            With shpBanca(i)
            .FillColor = Rojo
            .BorderColor = Rojo
            End With
        End If
    End If
End If

End Sub

Private Sub tmSqvData_Timer()
On Error Resume Next
Dim svLine As String
Dim i As Integer
'Obtengo IDS + Cuils
svTotal = ""
'Genero informacion
For i = 0 To 256
    svLine = i & ";"
    If (EstadoActual.VectorIdentificacion(i) <> NO_IDENTIFICADO Or EstadoActual.VectorPresencia(i) = PRESENTE) Then
        svLine = svLine & "1;"
    Else
        svLine = svLine & "0;"
    End If
    If (EstadoActual.VectorIdentificacion(i) <> NO_IDENTIFICADO) Then
        svLine = svLine & svCuilsArray(EstadoActual.VectorIdentificacion(i))
    Else
        svLine = svLine & "-1"
    End If
    If (svBancasArray(i) <> "" And svBancasArray(i) <> "300") Then
        svLine = svLine & ";" & svCuilsArray(svBancasArray(i))
    Else
        svLine = svLine & ";" & "-1"
    End If
    If i = 56 Then
        svLine = svLine
    End If
    
    If (svTotal = "") Then
        svTotal = svLine
    Else
        svTotal = svTotal & "," & svLine
    End If
Next i
Dim s As String
s = ""
svTotal = Replace(svTotal, vbCrLf, "")
svTotal = Replace(svTotal, " ", "")
'Clipboard.Clear
'Clipboard.SetText svTotal
svWinsock.Close
If svTotal <> svLastSent Then
    svWinsock.Connect "10.1.1.8", 3001
End If
End Sub

Private Sub tmUpdate_Timer()
Update
If ((GetTickCount - WebSenderTick) > 15000) Then
    StartWebSender
Else
    If (WebSender.State = sckClosed Or WebSender.State = sckError Or WebSender.State = sckClosing) Then
        StartWebSender
    End If
End If
End Sub

Private Function getSafeResult(param As String) As String
Dim final As String
final = param
If (IsNumeric(param)) Then
    Dim c As Integer
    c = Val(param)
    If (c < 0) Then
        final = "0"
    End If
End If
getSafeResult = final
End Function

Public Sub Update()
Const ORDINAL_MASCULINO = "°"
Const ORDINAL_FEMENINO = "ª"
If EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis" Then
    lblLeyendaSinIdentificar.Visible = True
    lblSinIdentificar.Visible = True
Else
    lblLeyendaSinIdentificar.Visible = False
    lblSinIdentificar.Visible = False
End If
lblAfirmativos.Caption = getSafeResult(frmMain.lblGeneralAfirmativosDato.Caption)
lblNegativos.Caption = getSafeResult(frmMain.lblGeneralNegativosDato.Caption)
lblAbstenciones.Caption = getSafeResult(frmMain.lblGeneralAbstencionesDato.Caption)
lblResultado.Caption = Trim(frmMain.lblGeneralResultadoDato.Caption)
lblFecha.Caption = Format(Now(), "dd/mm/yy")
lblHora.Caption = Format(Now(), "hh:mm")
lblTitulo.Caption = Trim(frmMain.lblTituloActa.Caption)
If VL.modoExtendido And (EstadoActual.EstadoVotacion_y_PasList = "larga") Then
    Dim i As Integer
    Dim totalSinVotar As Integer
    totalSinVotar = 0
    For i = 1 To 256
        If EstadoActual.VectorResultados(i) = ABSTENCION And EstadoActual.VectorIdentificacion(i) <> NO_IDENTIFICADO Then
            totalSinVotar = totalSinVotar + 1
        End If
    Next i
    If EstadoActual.PresidenteHabilitadoParaVotar And EstadoActual.VectorResultados(0) = ABSTENCION And EstadoActual.EstadoVotacion_y_PasList <> "empate" Then
        totalSinVotar = totalSinVotar + 1
    End If
    lblLeyendaSinIdentificar.Caption = "Sin Votar"
    lblSinIdentificar.Caption = Trim(CStr(totalSinVotar))
    lblSinIdentificar.ForeColor = Celeste
Else
    lblLeyendaSinIdentificar.Caption = "No Ident."
    If Not EstadoActual.EstadoVotacion_y_PasList = "finalizada" Then
        lblSinIdentificar.Caption = GetNoIdentificadosSobrePresentes
    Else
        'Si esta finalizada
        If Not sinIdentificarCongelado Then
            lblSinIdentificar.Caption = GetNoIdentificadosSobrePresentes
            sinIdentificarCongelado = True
        End If
    End If
    lblSinIdentificar.ForeColor = &HFFFF&
End If
If EstadoActual.Expresiones_Minoria = True Then
    lblTipoVotacion.Caption = "" 'Expresiones en Minorí
Else
    If lblTipoVotacion.Caption = "" Then
        lblTipoVotacion.Caption = UCase(frmMain.lblGeneralTipoOperacionDato)
    End If
End If
If LCase(Trim(EstadoActual.TipoDeOperacion)) <> "quorum" Then
    If EstadoActual.EstadoVotacion_y_PasList = "esperafin" Then
        lblPaseListaFinalizado.Visible = True
        lblTipoVotacion.Visible = False
    Else
        lblPaseListaFinalizado.Visible = False
        lblTipoVotacion.Visible = True
    End If
Else
    lblPaseListaFinalizado.Visible = False
    lblTipoVotacion.Visible = False
End If
If EstadoActual.Expresiones_Minoria = False Then
    If EstadoActual.TipoDeOperacion = "votnom" And Not InStr(lblTipoVotacion.Caption, "NOMINAL") > 0 Then
        lblTipoVotacion.Caption = "VOTACION" & vbCrLf & "NOMINAL"
    Else
        If EstadoActual.EstadoVotacion_y_PasList <> "esperafin" Then
            If EstadoActual.TipoDeOperacion <> "votnom" Then
                lblTipoVotacion.Caption = Trim(frmMain.lblGeneralInformacion.Caption)
            End If
        End If
    End If
End If
lblMayoria.Caption = Trim(frmMain.lblGeneralMayoriaDato(2).Caption)
If EstadoActual.Expresiones_Minoria = False Then
    If shpQuorum.Visible = False Then
        shpQuorum.Visible = True
    End If
    If lblQuorum.Visible = False Then
        lblQuorum.Visible = True
    End If
    If Trim(frmMain.lblGeneralLeyendaQuorumDato.Caption) = "NO HAY QUORUM" Then
        lblQuorum.Caption = "NO HAY QUORUM"
        lblQuorum.ForeColor = &HFF&
        shpQuorum.BorderColor = &HC0&
    Else
        If EstadoActual.TipoMayoriaQuorum = "man" Then
            lblQuorum.Caption = "MANTENIMIENTO"
        Else
            lblQuorum.Caption = "QUORUM"
        End If
        lblQuorum.ForeColor = &HFFFF&
        shpQuorum.BorderColor = &HFFFF&
    End If
Else
    shpQuorum.Visible = True
    lblQuorum.Visible = True
    lblQuorum.Caption = "Expr. en Minoría"
    shpQuorum.BorderColor = &HFF&
    lblQuorum.ForeColor = &HFF&
End If
lblNumeroPeriodo.Caption = NumeroPeriodo & ORDINAL_MASCULINO
If EstadoActual.Expresiones_Minoria = False And EstadoActual.Sesion <> -1 Then
    If lblNumeroSesion.Visible = False Then
        lblNumeroSesion.Visible = True
    End If
    lblNumeroSesion.Caption = EstadoActual.Sesion & ORDINAL_FEMENINO
Else
    If lblNumeroSesion.Visible = True Then
        lblNumeroSesion.Visible = False
    End If
End If
lblNumeroReunion.Caption = EstadoActual.Reunion & ORDINAL_FEMENINO
If (Left(Right(EstadoActual.PeriodoLegislativo, 2), 1) = "p") Then
    lblTipoPeriodo.Caption = "Prórroga Ordinario"
Else
    lblTipoPeriodo.Caption = "Período " & ObtenerTipoPeriodo
End If
lblTipoSesion.Caption = "Sesión - " & ObtenerTipoSesion
lblPresentes.Caption = Trim(frmMain.lblGeneralPresentesDato.Caption)
lblAusentes.Caption = Trim(frmMain.lblGeneralAusentesDato.Caption)
lblTimerPresentes.Caption = lblPresentes.Caption
lblTimerAusentes.Caption = lblAusentes.Caption
'TITULO
If EstadoActual.TipoDeOperacion = "quorum" Or EstadoActual.TipoDeOperacion = "paslis" Then
    'shpTitulo.Visible = False
    lblTitulo.Visible = False
    lblLeyendaReunion.Visible = True
    lblNumeroReunion.Visible = True
Else
    If lblTitulo.Caption <> "" Then
        'shpTitulo.Visible = True
        lblTitulo.Visible = True
    Else
        'shpTitulo.Visible = False
        lblTitulo.Visible = False
    End If
    lblLeyendaReunion.Visible = False
    lblNumeroReunion.Visible = False
End If
If EstadoActual.EstadoVotacion_y_PasList <> "esperafin" And (EstadoActual.TipoDeOperacion <> "quorum" Or EstadoActual.Expresiones_Minoria = True) Then
    lblTipoVotacion.Visible = True
    If EstadoActual.TipoDeOperacion <> "paslis" And EstadoActual.Expresiones_Minoria = False Then
        lblMayoria.Visible = True
    Else
        lblMayoria.Visible = False
    End If
Else
    lblTipoVotacion.Visible = False
    lblMayoria.Visible = False
End If
If EstadoActual.EstadoVotacion_y_PasList = "votando" Then
    If Trim(frmMain.lblGeneralTiempoDato.Caption) <> "" Then
        lblTiempo.Caption = Trim(frmMain.lblGeneralTiempoDato.Caption)
        lblLeyendaTiempo.Visible = True
        lblTiempo.Visible = True
    Else
        lblTiempo.Visible = False
        lblLeyendaTiempo.Visible = False
    End If
Else
    lblLeyendaTiempo.Visible = False
    lblTiempo.Visible = False
End If
If EstadoActual.EstadoVotacion_y_PasList = "cierre" Or EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate" Then
    MostrarResultados
ElseIf EstadoActual.EstadoVotacion_y_PasList = "empate" Then
    If EstadoActual.EstadoVotacion_y_PasList = "cierre" Or EstadoActual.EstadoVotacion_y_PasList = "finalizada" Then
        ActualizarUnColor (0)
    End If
Else
    If EstadoActual.EstadoVotacion_y_PasList <> "empate" Then
        lblLeyendaAfirmativos.Visible = False
        lblLeyendaNegativos.Visible = False
        lblLeyendaAbstenciones.Visible = False
        lblAfirmativos.Visible = False
        lblNegativos.Visible = False
        lblAbstenciones.Visible = False
        lblResultado.Visible = False
    End If
End If

'Timer sets
Dim tituloSesionTimer As String
Dim tipoSesion As String
tipoSesion = ObtenerTipoSesion
tituloSesionTimer = tipoSesion
If (LCase(tipoSesion) = "tablas") Then
    tituloSesionTimer = "Sesión de Tablas"
Else
    tituloSesionTimer = "Sesión " & tipoSesion
End If
If EstadoActual.Expresiones_Minoria = True Then
    Me.lblTimerTitulo.Caption = Me.lblNumeroPeriodo.Caption & " Período - " & Me.lblNumeroReunion.Caption & " Reunión "
    If (Me.frameTimer.Width <> 7515) Then
        Me.frameTimer.Width = 7515
    End If
    If Me.lblTimerTitulo.Width <> 7245 Then
        Me.lblTimerTitulo.Width = 7245
    End If
Else
    If (LCase(tipoSesion) = "preparatoria") Then
        Me.lblTimerTitulo.Caption = Me.lblNumeroPeriodo.Caption & " " & Me.lblTipoPeriodo.Caption & " - " & _
            tituloSesionTimer & " - " & _
            Me.lblNumeroReunion.Caption & " Reunión"
    Else
        Me.lblTimerTitulo.Caption = Me.lblNumeroPeriodo.Caption & " " & Me.lblTipoPeriodo.Caption & " - " & _
            Me.lblNumeroSesion.Caption & " " & tituloSesionTimer & " - " & _
            Me.lblNumeroReunion.Caption & " Reunión"
    End If
    If (Me.frameTimer.Width <> 14955) Then
        Me.frameTimer.Width = 14955
    End If
    If (Me.lblTimerTitulo.Width <> 14838) Then
        Me.lblTimerTitulo.Width = 14838
    End If
End If
End Sub
Private Function GetNoIdentificadosSobrePresentes() As Integer
Dim i As Integer
Dim total As Integer
total = 0
For i = 1 To 256 'Se evita la 0 porque siempre esta identificada
    If (EstadoActual.VectorIdentificacion(i) = 0 And EstadoActual.VectorPresencia(i) = PRESENTE) Then
    total = total + 1
    End If
Next i
GetNoIdentificadosSobrePresentes = total
End Function
Public Sub MostrarResultados()
    lblLeyendaAfirmativos.Visible = True
    lblLeyendaNegativos.Visible = True
    lblLeyendaAbstenciones.Visible = True
    If EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis" Then
        lblLeyendaSinIdentificar.Visible = True
        lblSinIdentificar.Visible = True
    Else
        lblLeyendaSinIdentificar.Visible = False
        lblSinIdentificar.Visible = False
    End If
    lblAfirmativos.Caption = getSafeResult(frmMain.lblGeneralAfirmativosDato.Caption)
    lblNegativos.Caption = getSafeResult(frmMain.lblGeneralNegativosDato.Caption)
    lblAbstenciones.Caption = getSafeResult(frmMain.lblGeneralAbstencionesDato.Caption)
    If lblResultado.Caption <> Trim(frmMain.lblGeneralResultadoDato.Caption) Then
        lblResultado.Caption = Trim(frmMain.lblGeneralResultadoDato.Caption)
    End If
    If lblResultado.Caption = "AFIRMATIVO" Then
        If lblResultado.ForeColor <> Verde Then
            lblResultado.ForeColor = Verde
            lblResultado.Left = 7980
        End If
    ElseIf lblResultado.Caption = "NEGATIVO" Then
        If lblResultado.ForeColor <> Rojo Then
            lblResultado.ForeColor = Rojo
            lblResultado.Left = 7740
        End If
    ElseIf lblResultado.Caption = "EMPATE" Then
        If lblResultado.ForeColor <> Amarillo Then
            lblResultado.ForeColor = Amarillo
        End If
    Else
        If lblResultado.ForeColor <> Blanco Then
            lblResultado.ForeColor = Blanco
            lblResultado.Left = 7170
        End If
    End If
    ActualizarUnColor (0)
    lblAfirmativos.Visible = True
    lblNegativos.Visible = True
    lblAbstenciones.Visible = True
    lblResultado.Visible = True
End Sub
Private Function NumeroPeriodo() As String
Dim temp As String
temp = Mid(EstadoActual.PeriodoLegislativo, 1, 3)
NumeroPeriodo = temp
End Function
Private Function ObtenerTipoPeriodo() As String
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
frmMain.SetearRsAux "SELECT leyenda_para_cartel FROM tipo_periodo WHERE id = '" & Mid(EstadoActual.PeriodoLegislativo, 4, 1) & "'", rsTemp
If rsTemp.EOF Then
    ObtenerTipoPeriodo = "Invalido"
Else
    ObtenerTipoPeriodo = rsTemp.Fields(0)
End If
rsTemp.Close
Set rsTemp = Nothing
End Function
Private Function ObtenerTipoSesion() As String
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
frmMain.SetearRsAux "SELECT leyenda_para_cartel FROM tipo_sesion WHERE id = '" & Mid(EstadoActual.PeriodoLegislativo, 5, 1) & "'", rsTemp
If rsTemp.EOF Then
    ObtenerTipoSesion = "Invalido"
Else
    ObtenerTipoSesion = rsTemp.Fields(0)
End If
rsTemp.Close
Set rsTemp = Nothing
End Function
Private Sub StartProtoServer()
On Error Resume Next
ProtoServer.Close
ProtoServer.Listen
End Sub
Private Sub StartWebSender()
On Error Resume Next
WebSenderTick = GetTickCount
WebSender.Close
WebSender.Listen
End Sub

Private Sub WebSender_ConnectionRequest(ByVal requestID As Long)
WebSender.Close
WebSender.Accept requestID
End Sub

Private Sub WebSender_DataArrival(ByVal bytesTotal As Long)
On Error GoTo AltSub
Dim buffer As String
Dim arr() As String
Dim send As String
send = Base64.Base64EncodeString(GetGlobalData)
WebSender.GetData buffer
If (buffer = "update") Then
    WebSender.SendData "|START|" & send & "|END|"
End If
WebSenderTick = GetTickCount
Exit Sub
AltSub:
StartWebSender
End Sub

