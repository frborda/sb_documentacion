VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmConsolaOperacion 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   15330
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   19170
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15330
   ScaleWidth      =   19170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Proyecto1.ButtonOffice cmdModoComponente 
      Height          =   420
      Left            =   300
      TabIndex        =   354
      Top             =   3000
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   741
      BackColor       =   12230304
      Caption         =   "Video"
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
   Begin Proyecto1.ButtonOffice cmdTaparQuorum 
      Height          =   735
      Left            =   14400
      TabIndex        =   353
      Top             =   13380
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1296
      BackColor       =   12230304
      Caption         =   "Apagar 5/6"
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
   Begin Proyecto1.ButtonOffice cmdListadoInmediato 
      Height          =   465
      Left            =   7080
      TabIndex        =   348
      Top             =   14580
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   820
      BackColor       =   8454016
      Caption         =   "Listado de Identificados"
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
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   -20000
      Picture         =   "frmConsolaOperacionAEBFranes.frx":0000
      ScaleHeight     =   4530
      ScaleWidth      =   15360
      TabIndex        =   6
      Top             =   6990
      Width           =   15360
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7250
         Picture         =   "frmConsolaOperacionAEBFranes.frx":BAA0
         ScaleHeight     =   330
         ScaleWidth      =   855
         TabIndex        =   270
         Top             =   100
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   4
         Left            =   540
         Picture         =   "frmConsolaOperacionAEBFranes.frx":CA38
         Top             =   3450
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   3
         Left            =   540
         Picture         =   "frmConsolaOperacionAEBFranes.frx":D122
         Top             =   60
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   2
         Left            =   6630
         Picture         =   "frmConsolaOperacionAEBFranes.frx":D80C
         Top             =   3450
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   1
         Left            =   540
         Picture         =   "frmConsolaOperacionAEBFranes.frx":DEF6
         Top             =   1350
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   0
         Left            =   11760
         Picture         =   "frmConsolaOperacionAEBFranes.frx":E5E0
         Top             =   1350
         Width           =   360
      End
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
   Begin VB.PictureBox HEA 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   -20000
      Picture         =   "frmConsolaOperacionAEBFranes.frx":ECCA
      ScaleHeight     =   780
      ScaleWidth      =   15360
      TabIndex        =   5
      Top             =   4200
      Width           =   15360
   End
   Begin VB.PictureBox MR1 
      BorderStyle     =   0  'None
      Height          =   15360
      Left            =   0
      Picture         =   "frmConsolaOperacionAEBFranes.frx":11D5C
      ScaleHeight     =   15360
      ScaleMode       =   0  'User
      ScaleWidth      =   19200
      TabIndex        =   177
      Top             =   0
      Width           =   19200
      Begin VB.Timer tmPresidente 
         Enabled         =   0   'False
         Interval        =   6000
         Left            =   17880
         Top             =   3480
      End
      Begin VB.Timer tmAScreen 
         Interval        =   100
         Left            =   9600
         Top             =   12330
      End
      Begin VB.Timer tmAutoCaptura 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   8610
         Top             =   12540
      End
      Begin VB.TextBox txtFoco 
         Height          =   375
         Left            =   2400
         TabIndex        =   344
         Text            =   "Text1"
         Top             =   15000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTituloTemp 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   3720
         TabIndex        =   342
         Top             =   14940
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.PictureBox pctInfo 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   3915
         Left            =   60
         ScaleHeight     =   3915
         ScaleWidth      =   7695
         TabIndex        =   336
         Top             =   6720
         Width           =   7695
         Begin VB.Label lblPICBanca 
            BackStyle       =   0  'Transparent
            Caption         =   "-1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000F&
            Height          =   795
            Left            =   120
            TabIndex        =   341
            Top             =   180
            Width           =   4455
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00808080&
            X1              =   60
            X2              =   4620
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00808080&
            X1              =   60
            X2              =   4620
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Label lblPICNombre 
            BackStyle       =   0  'Transparent
            Caption         =   "APELLIDO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   60
            TabIndex        =   340
            Top             =   1800
            Width           =   4515
         End
         Begin VB.Image picFlotante 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   3795
            Left            =   4680
            Stretch         =   -1  'True
            Top             =   60
            Width           =   2955
         End
         Begin VB.Label lblPICProvincia 
            BackStyle       =   0  'Transparent
            Caption         =   "PROVINCIA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   60
            TabIndex        =   339
            Top             =   3480
            Width           =   4575
         End
         Begin VB.Label lblPICBloque 
            BackStyle       =   0  'Transparent
            Caption         =   "Frente para la Victoria - Partido Bloqui"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   675
            Left            =   60
            TabIndex        =   338
            Top             =   2520
            Width           =   4575
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            X1              =   0
            X2              =   4560
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label lblPICApellido 
            BackStyle       =   0  'Transparent
            Caption         =   "APELLIDO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   60
            TabIndex        =   337
            Top             =   1380
            Width           =   4515
         End
      End
      Begin Proyecto1.ButtonOffice cmdSimular 
         Height          =   285
         Left            =   60
         TabIndex        =   316
         Top             =   14970
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         BackColor       =   16744576
         Caption         =   "Simular Voto"
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
      Begin Proyecto1.ButtonOffice cmdCierreVotacion 
         Height          =   630
         Left            =   15480
         TabIndex        =   315
         Top             =   14550
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   1111
         BackColor       =   12230304
         Caption         =   "Cier&re Votación"
         Enabled         =   0   'False
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
         State           =   3
      End
      Begin Proyecto1.ButtonOffice cmdModoVotaPresidente 
         Height          =   630
         Left            =   15480
         TabIndex        =   314
         Top             =   13800
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   1111
         BackColor       =   12230304
         Caption         =   "Habilitar voto del Presidente"
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
      Begin Proyecto1.ButtonOffice cmdVotacion 
         Height          =   630
         Left            =   15450
         TabIndex        =   313
         Top             =   13020
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   1111
         BackColor       =   12230304
         Caption         =   "Votac&ión"
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
      Begin Proyecto1.ButtonOffice cmdModoNominal 
         Height          =   585
         Left            =   13350
         TabIndex        =   312
         Top             =   14160
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1032
         BackColor       =   12230304
         Caption         =   "Habilitar identificación"
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
      Begin Proyecto1.ButtonOffice cmdCarteles 
         Height          =   735
         Left            =   13350
         TabIndex        =   311
         Top             =   13380
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1296
         BackColor       =   12230304
         Caption         =   "E&ncender Carteles"
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
      Begin Proyecto1.ButtonOffice cmdMantenimiento 
         Height          =   510
         Left            =   120
         TabIndex        =   310
         Top             =   13080
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   900
         BackColor       =   12230304
         Caption         =   "Man&tenimiento"
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
      Begin Proyecto1.ButtonOffice cmdLimpiarOrador 
         Height          =   465
         Left            =   11160
         TabIndex        =   309
         Top             =   13620
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   820
         BackColor       =   12230304
         Caption         =   "Limpiar Orador"
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
      Begin Proyecto1.ButtonOffice cmdSeleccionarOrador 
         Height          =   465
         Left            =   8940
         TabIndex        =   308
         Top             =   13620
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   820
         BackColor       =   12230304
         Caption         =   "Cambiar &Orador"
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
      Begin Proyecto1.ButtonOffice cmdCambiarNumeroReunion 
         Height          =   465
         Left            =   8940
         TabIndex        =   307
         Top             =   12690
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   820
         BackColor       =   12230304
         Caption         =   "Cambiar &Reunion"
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
      Begin Proyecto1.ButtonOffice cmdTituloBlanco 
         Height          =   585
         Left            =   7020
         TabIndex        =   306
         Top             =   13890
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1032
         BackColor       =   12230304
         Caption         =   "&Limpiar Título"
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
      Begin Proyecto1.ButtonOffice cmdTituloRapido 
         Height          =   465
         Left            =   2760
         TabIndex        =   305
         Top             =   14580
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   820
         BackColor       =   12230304
         Caption         =   "&Título Rápido"
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
      Begin Proyecto1.ButtonOffice cmdPresidente 
         Height          =   345
         Left            =   6990
         TabIndex        =   304
         Top             =   13290
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         BackColor       =   12230304
         Caption         =   "&Presidente"
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
      Begin Proyecto1.ButtonOffice cmdTitulo 
         Height          =   465
         Left            =   4950
         TabIndex        =   303
         Top             =   14580
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   820
         BackColor       =   12230304
         Caption         =   "Títulos Predefinidos"
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
      Begin Proyecto1.ButtonOffice cmdCambiarSesion 
         Height          =   465
         Left            =   6990
         TabIndex        =   302
         Top             =   12690
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   820
         BackColor       =   12230304
         Caption         =   "&Cambiar Sesión"
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
      Begin Proyecto1.ButtonOffice cmdNuevaSesion 
         Height          =   465
         Left            =   5610
         TabIndex        =   301
         Top             =   12690
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   820
         BackColor       =   12230304
         Caption         =   "Nueva s&esión"
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
      Begin Proyecto1.ButtonOffice cmdPeriodoLegislativo 
         Height          =   465
         Left            =   2730
         TabIndex        =   300
         Top             =   12690
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   820
         BackColor       =   12230304
         Caption         =   "P. Legislativo y Tipo de Sesión"
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
      Begin Proyecto1.ButtonOffice cmdCancelar 
         Height          =   420
         Left            =   120
         TabIndex        =   299
         Top             =   12630
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   741
         BackColor       =   12230304
         Caption         =   "Tomar Control"
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
      Begin Proyecto1.ButtonOffice cmdCancelarVotacion 
         Height          =   630
         Left            =   13350
         TabIndex        =   298
         Top             =   12720
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1111
         BackColor       =   16761024
         Caption         =   "Cance&lar Votación"
         Enabled         =   0   'False
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
         State           =   3
      End
      Begin Proyecto1.ButtonOffice Salir 
         Height          =   420
         Left            =   120
         TabIndex        =   297
         Top             =   14160
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   741
         BackColor       =   16761024
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
         PicOpacity      =   0
      End
      Begin VB.TextBox txtNumeroReunion 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   8940
         TabIndex        =   283
         Text            =   "0"
         Top             =   12720
         Width           =   1275
      End
      Begin VB.TextBox txtTitulo 
         Height          =   675
         Left            =   2760
         MaxLength       =   80
         MultiLine       =   -1  'True
         TabIndex        =   282
         Top             =   13830
         Width           =   4185
      End
      Begin VB.Frame frmControl 
         BackColor       =   &H00000000&
         Caption         =   "Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   4065
         Left            =   17880
         TabIndex        =   281
         Top             =   4650
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Opciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   2805
         Left            =   17880
         TabIndex        =   280
         Top             =   7020
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame frmActividades 
         BackColor       =   &H00000000&
         Caption         =   "Acciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   2805
         Left            =   16770
         TabIndex        =   279
         Top             =   5850
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Frame frmSesion 
         BackColor       =   &H00000000&
         Caption         =   "Sesión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   2805
         Left            =   17880
         TabIndex        =   278
         Top             =   7260
         Visible         =   0   'False
         Width           =   11235
      End
      Begin VB.Frame frmInformacion 
         BackColor       =   &H00000000&
         Caption         =   "Información"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1215
         Left            =   18390
         TabIndex        =   277
         Top             =   6900
         Visible         =   0   'False
         Width           =   17115
      End
      Begin VB.CommandButton cmdReconsiderar 
         Caption         =   "Reconsiderar"
         Height          =   330
         Left            =   7380
         Style           =   1  'Graphical
         TabIndex        =   274
         ToolTipText     =   "Período legislativo y Tipo de Sesión"
         Top             =   15285
         Visible         =   0   'False
         Width           =   3210
      End
      Begin VB.CommandButton cmdAbstenciones 
         Caption         =   "Selector de &Abstenciones"
         Height          =   330
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   273
         ToolTipText     =   "Período legislativo y Tipo de Sesión"
         Top             =   15285
         Visible         =   0   'False
         Width           =   3210
      End
      Begin VB.Timer tmCheckTemp 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   13680
         Top             =   120
      End
      Begin VB.Timer tmScreens 
         Enabled         =   0   'False
         Left            =   12840
         Top             =   -120
      End
      Begin MSDataListLib.DataCombo dcTipoMayoria 
         Height          =   345
         Left            =   10200
         TabIndex        =   284
         Top             =   14310
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   609
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483643
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
      Begin MSDataListLib.DataCombo dcBaseMayoria 
         Height          =   345
         Left            =   10200
         TabIndex        =   285
         Top             =   14880
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   609
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483643
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
      Begin MSDataListLib.DataCombo dcAbstencion 
         Height          =   315
         Left            =   9840
         TabIndex        =   286
         Top             =   15330
         Visible         =   0   'False
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
         Left            =   9690
         TabIndex        =   287
         Top             =   15120
         Visible         =   0   'False
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcTipoOperacion 
         Height          =   315
         Left            =   16410
         TabIndex        =   295
         Top             =   12630
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin Proyecto1.ButtonOffice cmdExpresionesMinoria 
         Height          =   375
         Left            =   13350
         TabIndex        =   343
         Top             =   14820
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         BackColor       =   8421631
         Caption         =   "Expr. en Minoría"
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
      Begin Proyecto1.ButtonOffice cmdImpresionEnConsola 
         Height          =   510
         Left            =   120
         TabIndex        =   347
         Top             =   13620
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   900
         BackColor       =   12230304
         Caption         =   "Impresión"
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
      Begin Proyecto1.ButtonOffice cmdListaSiguiente 
         Height          =   525
         Left            =   17160
         TabIndex        =   350
         Top             =   11340
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   926
         BackColor       =   12230304
         Caption         =   ">"
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
      Begin Proyecto1.ButtonOffice cmdListaAnterior 
         Height          =   525
         Left            =   16740
         TabIndex        =   351
         Top             =   11340
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   926
         BackColor       =   12230304
         Caption         =   "<"
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
      Begin Proyecto1.ButtonOffice cmdModoDatos 
         Height          =   420
         Left            =   1680
         TabIndex        =   352
         Top             =   3000
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   741
         BackColor       =   12230304
         Caption         =   "Datos"
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
      Begin Proyecto1.ButtonOffice cmdReiniciarIzquierdo 
         Height          =   660
         Left            =   240
         TabIndex        =   355
         Top             =   3480
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   1164
         BackColor       =   12230304
         Caption         =   "Reiniciar cartel izquierdo"
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
      Begin Proyecto1.ButtonOffice cmdReiniciarDerecho 
         Height          =   660
         Left            =   240
         TabIndex        =   356
         Top             =   4200
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   1164
         BackColor       =   12230304
         Caption         =   "Reiniciar cartel derecho"
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
      Begin VB.Label lblVotacionLarga 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Votación Larga"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   15960
         TabIndex        =   349
         Top             =   3000
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label lblEModoPrueba 
         BackStyle       =   0  'Transparent
         Caption         =   "MODO PRUEBA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   360
         TabIndex        =   346
         Top             =   2520
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Label lblUltimaAccion 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
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
         Height          =   345
         Left            =   8250
         TabIndex        =   345
         Top             =   10470
         Width           =   3075
      End
      Begin VB.Label lblAusentes 
         BackStyle       =   0  'Transparent
         Caption         =   "Ausentes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   15720
         TabIndex        =   335
         Top             =   1980
         Width           =   1695
      End
      Begin VB.Label lblPresentes 
         BackStyle       =   0  'Transparent
         Caption         =   "Presentes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   15720
         TabIndex        =   334
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label txtPresentes 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "257"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   17760
         TabIndex        =   333
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label txtAusentes 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "257"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   17760
         TabIndex        =   332
         Top             =   1980
         Width           =   795
      End
      Begin VB.Label TxtQuorum 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NO HAY QUORUM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   15540
         TabIndex        =   331
         Top             =   2520
         Width           =   3195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No Identificados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   405
         Index           =   16
         Left            =   225
         TabIndex        =   330
         Top             =   11880
         Width           =   2520
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pendientes de Votar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   405
         Index           =   3
         Left            =   4050
         TabIndex        =   329
         Top             =   12240
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.Label lblSi 
         BackStyle       =   0  'Transparent
         Caption         =   "AFIRMATIVOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   240
         TabIndex        =   328
         Top             =   11370
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label lblAbs 
         BackStyle       =   0  'Transparent
         Caption         =   "ABSTENCIONES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   7965
         TabIndex        =   327
         Top             =   11370
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.Label lblNo 
         BackStyle       =   0  'Transparent
         Caption         =   "NEGATIVOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   4365
         TabIndex        =   326
         Top             =   11370
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label txtNo 
         BackStyle       =   0  'Transparent
         Caption         =   "257"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   1005
         Left            =   6525
         TabIndex        =   325
         Top             =   11010
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label txtAbs 
         BackStyle       =   0  'Transparent
         Caption         =   "257"
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
         Height          =   1005
         Left            =   10965
         TabIndex        =   324
         Top             =   11010
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label txtSi 
         BackStyle       =   0  'Transparent
         Caption         =   "257"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   735
         Left            =   2685
         TabIndex        =   323
         Top             =   11010
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label txtResultado 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(RESULTADO)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   975
         Left            =   12765
         TabIndex        =   322
         Top             =   10920
         Visible         =   0   'False
         Width           =   5940
      End
      Begin VB.Label txtPendientesEmitirVoto 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "257"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   6300
         TabIndex        =   321
         Top             =   12150
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label txtOcup 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "257"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   2505
         TabIndex        =   320
         Top             =   11790
         Width           =   1020
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   480
         X2              =   2820
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Label txtHora 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "03:12:32"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   690
         TabIndex        =   319
         Top             =   1980
         Width           =   1995
      End
      Begin VB.Label txtFecha 
         BackStyle       =   0  'Transparent
         Caption         =   "12/12/2012"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   900
         TabIndex        =   318
         Top             =   1440
         Width           =   1875
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   615
         Left            =   720
         TabIndex        =   317
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   15480
         TabIndex        =   296
         Top             =   12660
         Width           =   1665
      End
      Begin VB.Shape shpVotaPresidente 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   735
         Left            =   15480
         Top             =   13740
         Visible         =   0   'False
         Width           =   3465
      End
      Begin VB.Label lblOrador 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orador"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8940
         TabIndex        =   294
         Top             =   13260
         Width           =   3975
      End
      Begin VB.Label lblPresidente 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Presidente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2730
         TabIndex        =   293
         Top             =   13320
         Width           =   4185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Quorum Tipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   17
         Left            =   7890
         TabIndex        =   292
         Top             =   15180
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Abstención"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   18
         Left            =   8040
         TabIndex        =   291
         Top             =   15390
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mayoría Tipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   19
         Left            =   8220
         TabIndex        =   290
         Top             =   14340
         Width           =   1740
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mayoría Base"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   20
         Left            =   8280
         TabIndex        =   289
         Top             =   14910
         Width           =   1740
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Título"
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
         Height          =   210
         Index           =   0
         Left            =   2070
         TabIndex        =   288
         Top             =   13830
         Width           =   645
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2370
         TabIndex        =   275
         Top             =   420
         Width           =   14460
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2280
         TabIndex        =   276
         Top             =   120
         Width           =   14430
      End
      Begin VB.Label lblNombreRapido 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Info: Carlos 01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -15000
         TabIndex        =   271
         Top             =   0
         Width           =   1815
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
         Height          =   600
         Index           =   0
         Left            =   0
         TabIndex        =   272
         Top             =   0
         Width           =   570
      End
      Begin VB.Image pctFotoRapida 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   13800
         Stretch         =   -1  'True
         Top             =   720
         Width           =   615
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
         Left            =   -120
         TabIndex        =   248
         Top             =   60
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   2
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   0
         TabIndex        =   247
         Top             =   270
         Width           =   930
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   3
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   246
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   4
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   245
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   5
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   244
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   6
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   243
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   7
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   242
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   8
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   241
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   9
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   240
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   10
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   239
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   11
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   238
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   12
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   237
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   13
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   236
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   14
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   235
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   15
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   234
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   16
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   233
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   17
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   232
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   18
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   231
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   19
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   230
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   20
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   229
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   21
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   228
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   22
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   227
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   23
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   226
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   24
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   225
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   25
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   224
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   26
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   223
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   27
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   222
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   28
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   221
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   29
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   220
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   30
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   219
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   31
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   218
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   32
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   217
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   33
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   216
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   34
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   215
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   35
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   214
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   36
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   213
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   37
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   212
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   38
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   211
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   39
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   210
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   40
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   209
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   41
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   208
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   42
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   207
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   43
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   206
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   44
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   205
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   45
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   204
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   46
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   203
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   47
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   202
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   48
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   201
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   49
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   200
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   50
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   199
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   51
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   198
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   52
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   197
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   53
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   196
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   54
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   195
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   55
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   194
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   56
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   193
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   57
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   192
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   58
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   191
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   59
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   190
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   60
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   189
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   61
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   188
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   62
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   187
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   63
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   186
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   64
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   185
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   65
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   184
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   66
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   183
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   67
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   182
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   68
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   181
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   69
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   180
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   70
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   179
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Height          =   765
         Index           =   1
         Left            =   900
         Shape           =   3  'Circle
         Top             =   480
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
         Left            =   900
         TabIndex        =   178
         Top             =   480
         Width           =   420
      End
      Begin VB.Shape shpBanca 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   765
         Index           =   0
         Left            =   3120
         Shape           =   3  'Circle
         Top             =   120
         Width           =   795
      End
   End
   Begin VB.PictureBox MR2 
      BorderStyle     =   0  'None
      Height          =   6210
      Left            =   -840
      ScaleHeight     =   6210
      ScaleMode       =   0  'User
      ScaleWidth      =   16005
      TabIndex        =   7
      Top             =   780
      Visible         =   0   'False
      Width           =   16000
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   20
         Left            =   0
         TabIndex        =   269
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   480
         TabIndex        =   268
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   210
         TabIndex        =   267
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   270
         TabIndex        =   266
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   0
         TabIndex        =   265
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   480
         TabIndex        =   264
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   210
         TabIndex        =   263
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   270
         TabIndex        =   262
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   0
         TabIndex        =   261
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   480
         TabIndex        =   260
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   210
         TabIndex        =   259
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   270
         TabIndex        =   258
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   0
         TabIndex        =   257
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   480
         TabIndex        =   256
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   210
         TabIndex        =   255
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   270
         TabIndex        =   254
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   0
         TabIndex        =   253
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   3750
         TabIndex        =   252
         Top             =   1020
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3480
         TabIndex        =   251
         Top             =   900
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3540
         TabIndex        =   250
         Top             =   540
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3270
         TabIndex        =   249
         Top             =   420
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   72
         Left            =   405
         TabIndex        =   174
         Top             =   1950
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   71
         Left            =   315
         TabIndex        =   173
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   70
         Left            =   315
         TabIndex        =   172
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   69
         Left            =   315
         TabIndex        =   171
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   68
         Left            =   315
         TabIndex        =   170
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   67
         Left            =   315
         TabIndex        =   169
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   66
         Left            =   315
         TabIndex        =   168
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   65
         Left            =   315
         TabIndex        =   167
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   64
         Left            =   315
         TabIndex        =   166
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   63
         Left            =   315
         TabIndex        =   165
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   62
         Left            =   315
         TabIndex        =   164
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   61
         Left            =   315
         TabIndex        =   163
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   60
         Left            =   315
         TabIndex        =   162
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   59
         Left            =   315
         TabIndex        =   161
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   58
         Left            =   315
         TabIndex        =   160
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   57
         Left            =   315
         TabIndex        =   159
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   56
         Left            =   315
         TabIndex        =   158
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   55
         Left            =   315
         TabIndex        =   157
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   54
         Left            =   315
         TabIndex        =   156
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   53
         Left            =   315
         TabIndex        =   155
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   52
         Left            =   315
         TabIndex        =   154
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   51
         Left            =   315
         TabIndex        =   153
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   50
         Left            =   315
         TabIndex        =   152
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   49
         Left            =   315
         TabIndex        =   151
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   48
         Left            =   315
         TabIndex        =   150
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   47
         Left            =   315
         TabIndex        =   149
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   46
         Left            =   315
         TabIndex        =   148
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   45
         Left            =   315
         TabIndex        =   147
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   44
         Left            =   315
         TabIndex        =   146
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   43
         Left            =   315
         TabIndex        =   145
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   42
         Left            =   315
         TabIndex        =   144
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   41
         Left            =   315
         TabIndex        =   143
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   40
         Left            =   315
         TabIndex        =   142
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   39
         Left            =   315
         TabIndex        =   141
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   38
         Left            =   315
         TabIndex        =   140
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   37
         Left            =   315
         TabIndex        =   139
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   36
         Left            =   315
         TabIndex        =   138
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   35
         Left            =   315
         TabIndex        =   137
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   34
         Left            =   315
         TabIndex        =   136
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   33
         Left            =   315
         TabIndex        =   135
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   32
         Left            =   315
         TabIndex        =   134
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   31
         Left            =   315
         TabIndex        =   133
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   30
         Left            =   315
         TabIndex        =   132
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   29
         Left            =   315
         TabIndex        =   131
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   28
         Left            =   315
         TabIndex        =   130
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   27
         Left            =   315
         TabIndex        =   129
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   26
         Left            =   315
         TabIndex        =   128
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   25
         Left            =   315
         TabIndex        =   127
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   24
         Left            =   315
         TabIndex        =   126
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   23
         Left            =   315
         TabIndex        =   125
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   22
         Left            =   315
         TabIndex        =   124
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   21
         Left            =   315
         TabIndex        =   123
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   20
         Left            =   315
         TabIndex        =   122
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   19
         Left            =   315
         TabIndex        =   121
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   18
         Left            =   315
         TabIndex        =   120
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   17
         Left            =   315
         TabIndex        =   119
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   16
         Left            =   315
         TabIndex        =   118
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   15
         Left            =   315
         TabIndex        =   117
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   14
         Left            =   315
         TabIndex        =   116
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   13
         Left            =   315
         TabIndex        =   115
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   12
         Left            =   315
         TabIndex        =   114
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   315
         TabIndex        =   113
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   315
         TabIndex        =   112
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   315
         TabIndex        =   111
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   315
         TabIndex        =   110
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   315
         TabIndex        =   109
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   315
         TabIndex        =   108
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   315
         TabIndex        =   107
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   315
         TabIndex        =   106
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   315
         TabIndex        =   105
         Top             =   1890
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   315
         TabIndex        =   104
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   315
         TabIndex        =   103
         Top             =   1020
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblBanca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   420
         TabIndex        =   102
         Top             =   2580
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   72
         Left            =   660
         TabIndex        =   101
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   71
         Left            =   660
         TabIndex        =   100
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   70
         Left            =   660
         TabIndex        =   99
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   69
         Left            =   660
         TabIndex        =   98
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   68
         Left            =   660
         TabIndex        =   97
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   67
         Left            =   660
         TabIndex        =   96
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   66
         Left            =   660
         TabIndex        =   95
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   65
         Left            =   660
         TabIndex        =   94
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   64
         Left            =   660
         TabIndex        =   93
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   63
         Left            =   660
         TabIndex        =   92
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   62
         Left            =   660
         TabIndex        =   91
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   61
         Left            =   660
         TabIndex        =   90
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   60
         Left            =   660
         TabIndex        =   89
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   59
         Left            =   660
         TabIndex        =   88
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   58
         Left            =   660
         TabIndex        =   87
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   57
         Left            =   660
         TabIndex        =   86
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   56
         Left            =   660
         TabIndex        =   85
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   55
         Left            =   660
         TabIndex        =   84
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   54
         Left            =   660
         TabIndex        =   83
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   53
         Left            =   660
         TabIndex        =   82
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   52
         Left            =   660
         TabIndex        =   81
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   51
         Left            =   660
         TabIndex        =   80
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   50
         Left            =   660
         TabIndex        =   79
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   49
         Left            =   660
         TabIndex        =   78
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   48
         Left            =   660
         TabIndex        =   77
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   47
         Left            =   660
         TabIndex        =   76
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   46
         Left            =   660
         TabIndex        =   75
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   45
         Left            =   660
         TabIndex        =   74
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   44
         Left            =   660
         TabIndex        =   73
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   43
         Left            =   660
         TabIndex        =   72
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   42
         Left            =   660
         TabIndex        =   71
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   41
         Left            =   660
         TabIndex        =   70
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   40
         Left            =   660
         TabIndex        =   69
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   39
         Left            =   660
         TabIndex        =   68
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   38
         Left            =   660
         TabIndex        =   67
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   37
         Left            =   660
         TabIndex        =   66
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   36
         Left            =   660
         TabIndex        =   65
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   35
         Left            =   660
         TabIndex        =   64
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   34
         Left            =   660
         TabIndex        =   63
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   33
         Left            =   660
         TabIndex        =   62
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   32
         Left            =   660
         TabIndex        =   61
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   31
         Left            =   660
         TabIndex        =   60
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   30
         Left            =   660
         TabIndex        =   59
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   29
         Left            =   660
         TabIndex        =   58
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   28
         Left            =   660
         TabIndex        =   57
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   27
         Left            =   660
         TabIndex        =   56
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   26
         Left            =   660
         TabIndex        =   55
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   25
         Left            =   660
         TabIndex        =   54
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   24
         Left            =   660
         TabIndex        =   53
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   23
         Left            =   660
         TabIndex        =   52
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   660
         TabIndex        =   51
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   660
         TabIndex        =   50
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   660
         TabIndex        =   49
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   660
         TabIndex        =   48
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   660
         TabIndex        =   47
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   660
         TabIndex        =   46
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   660
         TabIndex        =   45
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   660
         TabIndex        =   44
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   660
         TabIndex        =   43
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   660
         TabIndex        =   42
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   660
         TabIndex        =   41
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   660
         TabIndex        =   40
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   660
         TabIndex        =   39
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   660
         TabIndex        =   38
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   660
         TabIndex        =   37
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   660
         TabIndex        =   36
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   660
         TabIndex        =   35
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   660
         TabIndex        =   34
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   660
         TabIndex        =   33
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   660
         TabIndex        =   32
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   660
         TabIndex        =   31
         Top             =   1500
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   660
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   60
         TabIndex        =   29
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   210
         TabIndex        =   28
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   210
         TabIndex        =   27
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   210
         TabIndex        =   26
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   210
         TabIndex        =   25
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   210
         TabIndex        =   24
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   210
         TabIndex        =   23
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   210
         TabIndex        =   22
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   210
         TabIndex        =   21
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   210
         TabIndex        =   20
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   210
         TabIndex        =   19
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   210
         TabIndex        =   18
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   210
         TabIndex        =   17
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   210
         TabIndex        =   16
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   210
         TabIndex        =   15
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   210
         TabIndex        =   14
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   210
         TabIndex        =   13
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   210
         TabIndex        =   12
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   210
         TabIndex        =   11
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   210
         TabIndex        =   10
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   660
         TabIndex        =   9
         Top             =   570
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   210
         TabIndex        =   8
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   71
         Left            =   2820
         Shape           =   4  'Rounded Rectangle
         Top             =   1860
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   70
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   69
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   68
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   67
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   66
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   65
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   64
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   63
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   62
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   61
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   60
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   59
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   58
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   57
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   56
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   55
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   54
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   53
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   52
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   51
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   50
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   49
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   48
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   47
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   46
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   45
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   44
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   43
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   42
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   41
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   40
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   39
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   38
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   37
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   36
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   35
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   34
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   33
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   32
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   31
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   30
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   29
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   28
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   27
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   26
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   25
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   24
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   23
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   22
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   21
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   20
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   19
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   18
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   17
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   16
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   15
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   14
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   13
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   12
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   11
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   10
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   9
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   8
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   7
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   6
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   5
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   4
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   3
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   2
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   1
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape shpBanka 
         Height          =   315
         Index           =   0
         Left            =   180
         Shape           =   4  'Rounded Rectangle
         Top             =   540
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.PictureBox ML 
      BorderStyle     =   0  'None
      Height          =   6210
      Index           =   1
      Left            =   4440
      Picture         =   "frmConsolaOperacionAEBFranes.frx":3D1D9E
      ScaleHeight     =   6210
      ScaleWidth      =   990
      TabIndex        =   175
      Top             =   780
      Visible         =   0   'False
      Width           =   990
      Begin VB.Image imgA 
         Height          =   360
         Index           =   1
         Left            =   450
         Picture         =   "frmConsolaOperacionAEBFranes.frx":3D3205
         ToolTipText     =   "Ver mapa de Bancas"
         Top             =   150
         Width           =   360
      End
      Begin VB.Image imgB 
         Height          =   360
         Index           =   1
         Left            =   450
         Picture         =   "frmConsolaOperacionAEBFranes.frx":3D38EF
         ToolTipText     =   "Ver Legisladores Presentes"
         Top             =   750
         Width           =   360
      End
   End
   Begin VB.PictureBox ML 
      BorderStyle     =   0  'None
      Height          =   6210
      Index           =   0
      Left            =   3240
      Picture         =   "frmConsolaOperacionAEBFranes.frx":3D3FD9
      ScaleHeight     =   6210
      ScaleWidth      =   990
      TabIndex        =   176
      Top             =   840
      Width           =   990
      Begin VB.Image imgA 
         Height          =   360
         Index           =   0
         Left            =   450
         Picture         =   "frmConsolaOperacionAEBFranes.frx":3D53F6
         ToolTipText     =   "Ver mapa de Bancas"
         Top             =   150
         Width           =   360
      End
      Begin VB.Image imgB 
         Height          =   360
         Index           =   0
         Left            =   450
         Picture         =   "frmConsolaOperacionAEBFranes.frx":3D5AE0
         ToolTipText     =   "Ver Legisladores Presentes"
         Top             =   750
         Width           =   360
      End
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   3720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblSesion 
      Caption         =   "lblSesion"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   5565
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblNroSesion 
      Caption         =   "lnlNroSesion"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   5310
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblPeriodo 
      Caption         =   "lblPeriodo"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   5055
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblCodigoSesion 
      Caption         =   "lblCodigoSesion"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblActa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   15840
      TabIndex        =   0
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
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Variables para el armado del recinto
Dim conta As Integer
Dim currDiputados As New Dictionary
Dim AumentoTop As Integer
Dim OffLeft As Integer
Dim Left_Inicial As Integer
Dim Top_Inicial As Integer
Dim Exp As Integer
Dim EsInicializacion            As Boolean
Dim nImpresionSesionActual As String
Dim nImpresionPeriodoActual As String
'Otras
Public IniciaActualizacion        As Boolean
Public TipMayCargo                As Boolean
Private Impresa As Boolean
Public MensajeEsperado            As Variant
Private xListar_automaticamente   As Long
Dim foco                          As Boolean
Private strconexion               As String
Dim strTipoP                      As String
Dim ModificarTiempo               As Boolean
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
Private mCantidadLegisladores     As Integer
Private mActualizarDatos          As Boolean
Private mMensajeOperador          As String
Private mIPActivo                 As String
Private mPreguntarIniciarServidor As Boolean
Private mRefrescarConsola         As Boolean
Private Const sqlVector           As String = "SELECT TOP 1 * FROM vector ORDER BY id DESC"
Private Const mDelimitadorVector  As String = ";"
Private Const MAX_BANCA = 256
Private CambiandoReunion As Boolean
Private Fue9999 As Boolean


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
'Control del presidente
Private rstPresidente             As ADODB.Recordset
Private rstOrador                 As ADODB.Recordset

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
    legislador As String
    Presencia As String
    LegisladorDefecto As String
    LegisladorDefectoNombre As String
    BloqueNombre As String
    BloqueIndex As Integer
    BancaDefecto As String
End Type

Private Type tpBloque
    Nombre As String
    TotalLegisladores As Integer
    Presentes As Integer
    Ausentes As Integer
    Identificados As Integer
    NoIdentificados As Integer
End Type

Private datBanca(MAX_BANCA + 1) As tpBanca
Private datLista(MAX_BANCA + 1) As tpBanca
Private datBloque() As tpBloque
Private flgRepintarLista As Boolean
Private cntRepintarLista As Long
'091026
Private mOrador       As String
Private mHabilitarIdentificacion As Integer
Private mPresenciaConIdentificacion As Integer
Dim MouseEnBanca As Boolean
'Variables para las Screens
Dim Segundos_Screen As Integer
Dim Carpeta_Screen As String
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Public Property Let TiempoEspera(ByVal pNewValue As Integer)
    determinarTiempoEspera pNewValue
End Property
Public Function IDRepetida(id As String) As Boolean
Dim i As Integer
For i = 0 To MAX_BANCA
    If Trim(datBanca(i).legislador) = Trim(id) Then
        IDRepetida = True
        i = MAX_BANCA
    Else
        IDRepetida = False
    End If
Next i
End Function
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

Private Sub cmdCambiarNumeroReunion_Click()
    Dim xForm As frmCambiarReunion
    Set xForm = New frmCambiarReunion
    xForm.txtReunion.Text = txtNumeroReunion.Text
    xForm.Show vbModal
    Set xForm = Nothing
    'Call cambiarReunion(txtNumeroReunion.Text)
    'CambiandoReunion = False
End Sub

Private Sub cmdCancelar_Click()
    MensajesSQV.cancelarConsola
End Sub
Private Sub cmdCancelarVotacion_Click()
    MensajesSQV.cancelaVotacion
End Sub
Private Sub cmdCarteles_Click()
    strCartelSQV = ""
If cmdCarteles.Caption = "Apagar Carteles" Then
    MensajesSQV.cambiarCartelEncendido "0"
Else
    MensajesSQV.cambiarCartelEncendido "2"
End If
End Sub
Private Sub EstadoCarteles(pEstado As Integer)
    If UCase(pEstado) = 0 Or UCase(pEstado) = 1 Then
        cmdCarteles.Caption = "Encender Carteles"
        If (cmdTaparQuorum.Enabled = True) Then
            cmdTaparQuorum.Enabled = False
        End If
    Else
        If (cmdTaparQuorum.Enabled = False) Then
            cmdTaparQuorum.Enabled = True
        End If
        cmdCarteles.Caption = "Apagar Carteles"
    End If
End Sub

Private Sub cmdCierreVotacion_Click()
If (Screens_Habilitadas = True) Then
    tmScreens.Enabled = False
    tmCheckTemp.Enabled = False
    MsgBox ("Se detuvo la captura de pantalla automática")
End If
    MensajesSQV.cierreVotacion
End Sub
Private Sub cmdHabilitaVotoPresidente_Click()
Datos.GrabarMensaje "cambio?modovotapresidente", " ", , True
End Sub
Private Sub cmdExpresionesMinoria_Click()
Datos.GrabarMensaje "cambio?expresiones_minoria", "1", , True
End Sub

Private Sub cmdImpresionEnConsola_Click()
Dim fActas As New frmMostrarActas
ImpresionDeConsola = True
fActas.sesion = nImpresionSesionActual
fActas.periodo = nImpresionPeriodoActual
fActas.Show vbModal, Me
End Sub

Private Sub cmdLimpiarOrador_Click()
    cambiarOrador ("0")
End Sub

Private Sub cmdListaAnterior_Click()
Datos.GrabarMensaje "listapendientes?anterior", " ", , True
End Sub

Private Sub cmdListadoInmediato_Click()
Dim qIn As String
Dim i As Integer
qIn = ""
For i = LBound(mVectorIdentificacion) To UBound(mVectorIdentificacion)
    If (mVectorIdentificacion(i) <> NO_IDENTIFICADO) Then
        If i = UBound(mVectorIdentificacion) Then
            qIn = qIn & mVectorIdentificacion(i) & ","
        Else
            qIn = qIn & mVectorIdentificacion(i)
        End If
    End If
Next i
If (qIn <> "") Then
    qIn = "(" & qIn & ")"
    Dim sql As String
    sql = "SELECT legisladores_activos.apellido + ', ' + legisladores_activos.nombre AS diputado, " & _
    "bloque_politico AS bloque FROM legisladores_activos WHERE legisladores_activos.id IN " & qIn
    Dim rs As New ADODB.Recordset
    SetearRs sql, rs
    If (rs.EOF) Then
        MsgBox "No hay diputados identificados!"
    Else
        Dim rpt As New rptListadoInmediato
        rpt.DataControl1.Recordset = rs
        rpt.PrintReport False
    End If
End If
End Sub

Private Sub cmdListaSiguiente_Click()
Datos.GrabarMensaje "listapendientes?siguiente", " ", , True
End Sub

Private Sub cmdMantenimiento_Click()
    If PermisosTotales.UsuarioMantenimiento = 1 Then
        frmMantenimiento.cmdTiempo.Enabled = ModificarTiempo
        frmMantenimiento.Show vbModal
    Else
        MsgBox "El usuario no tiene permisos para realizar esta tarea", vbInformation + vbOKOnly
    End If
End Sub

Private Sub cmdModoAv_Click()
Call Shell("BarcoTool.exe 2")
End Sub

Private Sub cmdModoComponente_Click()
Call Shell(App.Path & "\video.bat")
End Sub

Private Sub cmdModoDatos_Click()
'EjecutarSQL ("UPDATE barco SET modo_video = 0, cambiar_modo = 1")
Call Shell(App.Path & "\datos.bat")
End Sub

Private Sub cmdModoVideo_Click()
EjecutarSQL ("UPDATE barco SET modo_video = 1, cambiar_modo = 1")
End Sub

Private Sub cmdModoVotaPresidente_Click()
If (gTipoUsuario = 0) Or (gTipoUsuario = 2) Then
    'Datos.GrabarMensaje "cambio?modovotapresidente", " ", , True
    If tmPresidente.Enabled = True Then
        MsgBox "Ya se ha solicitado un cambio de estado para el presidente"
    Else
        'cmdModoVotaPresidente.Enabled = False
        cmdModoVotaPresidente.Caption = "Cargando, espere..."
        tmPresidente.Enabled = True
    End If
Else
    MsgBox "Ud. no dispone de permisos para realizar esta acción.", vbInformation + vbOKOnly
End If
End Sub

Private Sub cmdPeriodoLegislativo_Click()
    Dim periodo As New frmElegirPeriodo
    periodo.ActualizarDatos = True
    periodo.Show vbModal
    Set periodo = Nothing
End Sub
Private Sub cmdpresidente_Click()
On Error GoTo TrapError
Presidente_Label = lblPresidente.Caption
    Dim xF As frmElegirPresidente
    Set xF = New frmElegirPresidente
    If Trim(lblPresidente.Caption) <> "" Then
        xF.cmdCancelar.Enabled = True
    End If
    xF.Show vbModal, Me
    Dim nT As Long
    nT = GetTickCount
    While GetTickCount - nT < 2000
    Wend
    cntRepintarLista = 8
    Exit Sub
TrapError:
    Exit Sub
End Sub
Private Sub ArmarBancas()
    Dim cTemp As Boolean
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
    strPeriodoLeg = rstVector!Período_Legislativo
    xSesion = rstVector!Sesión
    nImpresionSesionActual = xSesion
    nImpresionPeriodoActual = strPeriodoLeg
    xNroActaActual = rstVector!Nro_de_Acta
    xListar_automaticamente = rstVector!Listar_automaticamente
    If (xListar_automaticamente = 1) Then
        ImpresionAutomaticaActivada = True
    Else
        ImpresionAutomaticaActivada = False
    End If
    If rstVector.EOF = False Then
        txtFecha.Caption = Format(rstVector!fecha, "dd/mm/yyyy")
        txtHora.Caption = rstVector!hora
'        txtHora.Caption = Format(Now(), "H:mm:ss")
        mIPActivo = Trim(rstVector!IP_Consola_Habilitada)
        ' Or True se agrego para forzar a la consola a tener todos los controles habilitados
        If rstVector.Fields("expresiones_minoria") = 1 Then
            cmdExpresionesMinoria.Caption = "Expr. Minoria OFF"
        Else
            cmdExpresionesMinoria.Caption = "Expr. Minoria ON"
        End If
        If mIPActivo = Datos.IPconsola Or mIPActivo = "0" Then
            mRefrescarConsola = True
        Else
            mRefrescarConsola = False
        End If
        If txtHora.Caption = mHoraAnterior Then
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
            mHoraAnterior = txtHora.Caption
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
        If Not Trim(dcAbstencion.Tag) = "absaut" Then
            lblVotacionLarga.Visible = True
        Else
            lblVotacionLarga.Visible = False
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
            txtTitulo.Text = lblTituloActa.Caption
        Else
            lblTituloActa.Caption = ""
            txtTitulo.Text = lblTituloActa.Caption
        End If
        'Controlar el caso de haber mas bancas dando presencia que miembros del cuerpo
        ' Se cambia el color de fondo
        If rstVector!Ausentes < 0 Then
            txtPresentes.BackColor = &H80&
            txtAusentes.BackColor = &H80&
            txtPresentes.Caption = rstVector!Presentes + rstVector!Ausentes
            txtAusentes.Caption = 0
        Else
            txtPresentes.Caption = max(Min(rstVector!Presentes, 257 - 1), 1) ' rstVector!Presentes
            txtAusentes.Caption = Min(max(rstVector!Ausentes, 1), 257 - 1) 'rstVector!Ausentes
            txtPresentes.BackColor = &H8000&
            txtAusentes.BackColor = &H8000&
        End If
        'txtOcup.Caption = rstVector!Ocupadas_no_identificadas
        Dim X As Integer
        Dim cOcup As Integer
        cOcup = 0
        Dim mContent As String
        Dim v As Integer
        v = -1
        If (IsNumeric(Trim(mVectorIdentificacion(0)))) Then
            v = Val(Trim(mVectorIdentificacion(0)))
        End If
        mContent = Trim(mVectorIdentificacion(0)) & ";" & currDiputados(v) & ";" & X
        For X = 1 To 256
            If mVectorIdentificacion(X) = "0" And mVectorPresencia(X) = "1" Then
                cOcup = cOcup + 1
            End If
            If mVectorIdentificacion(X) <> "0" Or mVectorPresencia(X) = "1" Then
                If (mVectorPresencia(X) = "1" And mVectorIdentificacion(X) = "0") Then
                    'Presente sin identificar
                    'Escribirtxt
                    mContent = mContent & vbCrLf & "0;;;" & X
                Else
                    'Presente identificado
                    v = -1
                    If (IsNumeric(Trim(mVectorIdentificacion(X)))) Then
                        v = Val(Trim(mVectorIdentificacion(X)))
                    End If
                    mContent = mContent & vbCrLf & Trim(mVectorIdentificacion(X)) & ";" & currDiputados(v) & ";" & X
                End If
            End If
        Next X
        If False Then
            Dim sPres As String
            Dim mDate As String
            Dim ms As String
            Dim mFolderPrefix As String
            ms = CStr(GetTickCount())
            
            sPres = Trim(txtPresentes.Caption)
            mDate = Format(Now, "dd_MM_yyyy_H_m_s")
            mFolderPrefix = "logs\sin_quorum\"
            If (IsNumeric(sPres)) Then
                If (Val(sPres) > 128) Then
                    mFolderPrefix = "logs\con_quorum\"
                Else
                    mFolderPrefix = "logs\sin_quorum\"
                End If
            End If
                    
            Open "C:\" & mFolderPrefix & sPres & "_" & mDate & "_" & ms & ".txt" For Output As #1
            Print #1, mContent
            Close #1
        End If
        
        txtOcup.Caption = Trim(Str(cOcup))
        TxtQuorum.Caption = Trim(rstVector!Leyenda_Quorum)
        If TxtQuorum.Caption = "QUORUM" Then
            If TxtQuorum.ForeColor <> vbYellow Then
                TxtQuorum.ForeColor = vbYellow
            End If
        Else
            If TxtQuorum.ForeColor <> vbRed Then
                TxtQuorum.ForeColor = vbRed
            End If
        End If
        txtResultado.Caption = rstVector!Resultado
        txtSi.Caption = rstVector!Afirmativos
        txtNo.Caption = rstVector!Negativos
        txtAbs.Caption = rstVector!Abstenciones
        'AP 041012 No se actualiza para permitir la edicion. txtTitulo.Text = rstVector!Titulo_Del_Acta
        'lblNroSesion.Caption = rstVector!Sesión
        lblActa.Caption = rstVector!Nro_de_Acta
        mActaGrabada = Val(rstVector!Acta_Grabada) 'Se podria actualizar mActaImpresa tambien con un valor del vector para que no pida dos veces el mismo acta.
        
        If Not IsNull(rstVector!Período_Legislativo) Then
                lblCodigoSesion.Caption = Left(rstVector!Período_Legislativo, 3)
                strEtiqueta = Left(rstVector!Período_Legislativo, 3)
                
                lblCodigoSesion.Tag = rstVector!Período_Legislativo
                Select Case UCase(mId(rstVector!Período_Legislativo, 4, 1))
                    Case "O"
                        'strEtiqueta = strEtiqueta & "Ordinario "
                        strEtiqueta = strEtiqueta & " Período Ordinario "
                        'lblPeriodo.Caption = "Ordinario"
                        lblPeriodo.Caption = "Ordinario"
                        strTipoP = "Ordinario"
                    Case "E"
                        strEtiqueta = strEtiqueta & " Período Extraordinario "
                        lblPeriodo.Caption = "Extraordinario"
                        strTipoP = "Extraordinario"
                    Case "P"
                        strEtiqueta = strEtiqueta & " Prórroga Ordinario "
                        lblPeriodo.Caption = "Prórroga Ordinario"
                        strTipoP = "Prórroga Ordinario"
                End Select
                If IsNull(rstVector!Sesión) = False Then
                    If Trim(rstVector!Sesión) = "0" Then
                        Call CreaSesion(lblCodigoSesion.Tag, True)
                        Exit Sub
                    End If
                    lblNroSesion.Caption = rstVector!Sesión
                    Select Case UCase(mId(rstVector!Período_Legislativo, 5, 1))
                        Case "T"
                            strEtiqueta = strEtiqueta & " - " & rstVector!Sesión & "ª Sesión "
                            strEtiqueta = strEtiqueta & "- Tablas"
                            lblSesion.Caption = "Tablas"
                        Case "E"
                            strEtiqueta = strEtiqueta & " - " & rstVector!Sesión & "ª Sesión "
                            strEtiqueta = strEtiqueta & "- Especial"
                            lblSesion.Caption = "Especial"
                        Case "O"
                            strEtiqueta = strEtiqueta & " - " & rstVector!Sesión & "ª Sesión "
                            strEtiqueta = strEtiqueta & "- Ordinaria"
                            lblSesion.Caption = "Ordinaria"
                        Case "X"
                            strEtiqueta = strEtiqueta & " - " & rstVector!Sesión & "ª Sesión "
                            strEtiqueta = strEtiqueta & "- Extraordinaria"
                            lblSesion.Caption = "Extraordinaria"
                        Case "A"
                            strEtiqueta = strEtiqueta & " - " & rstVector!Sesión & "ª "
                            strEtiqueta = strEtiqueta & "- Asamblea Legislativa"
                            lblSesion.Caption = "Asamblea Legislativa"
                        Case "P"
                            strEtiqueta = strEtiqueta & " - Sesión Preparatoria"
                            lblSesion.Caption = "Preparatoria"
                        Case "I"
                            strEtiqueta = strEtiqueta & " - " & rstVector!Sesión & "ª Sesión "
                            strEtiqueta = strEtiqueta & "- Informativa"
                            lblSesion.Caption = "Informativa"
                        Case "H"
                            strEtiqueta = strEtiqueta & " - " & rstVector!Sesión & "ª Sesión "
                            strEtiqueta = strEtiqueta & "- Homenajes"
                    End Select
                End If
                If IsNull(rstVector!Nro_de_Acta) = False Then
                    strEtiqueta = strEtiqueta & " - Próximo Nº de Acta: " & rstVector!Nro_de_Acta
                    lblActa.Caption = rstVector!Nro_de_Acta
                End If
                If IsNull(rstVector!Reunion) = False Then
                    If CambiandoReunion = False Then
                        txtNumeroReunion.Text = rstVector!Reunion
                        strEtiqueta = strEtiqueta & " - Reunión: " & rstVector!Reunion
                    End If
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
        If mVectorIdentificacion(0) <> strIdPresidenteRecinto Then ' And mVectorIdentificacion(0) <> "0" Then
            strIdPresidenteRecinto = mVectorIdentificacion(0)
            Call BuscarPresidentePredeterminado
        ElseIf mVectorIdentificacion(0) = "0" Then
            'aca
            lblPresidente.Caption = "Seleccione el presidente antes de continuar"
            If frmCargando.Visible = False Then
                cmdpresidente_Click
                If Error_Carga = True Then
                    Unload Me
                    frmMenu.Show
                    frmMenu.cmdConsola.Enabled = True
                    Exit Sub
                End If
            End If
        End If
        ' sbStatus.Panels(2).Text = "Tiempo para votación:" & Str(rstVector!Tiempo_de_votación)
        'Andres  Label2(2).Caption = "Tiempo de Votación: " & Str(rstVector!Tiempo_de_votación)
        respuestaMensaje = Trim(rstVector!strError)
        mEstadoVotacion = LCase(Trim(rstVector!Estado_de_votacion))
        mEstadoSesion = Trim(rstVector!Estado_sesion)
        mMensajeOperador = Trim(rstVector!Mensaje_al_operador)
        'mPresenciaConIdentificacion = Trim(rstVector!????)
        'mHabilitarIdentificacion = Trim(rstVector!????)
        'Controles de permisos de usuario
        If (mRefrescarConsola = True) Then
             EstadoControles
        End If
        
        'Boton de Vota Presidente
        If Modo_Vota_Presidente = True Then
            If shpVotaPresidente.Visible = False Then
                shpVotaPresidente.Visible = True
                cmdModoVotaPresidente.BackColor = vbRed
                cmdModoVotaPresidente.Caption = "Deshabilitar voto del Presidente"
            End If
        Else
            If shpVotaPresidente.Visible = True Then
                shpVotaPresidente.Visible = False
                cmdModoVotaPresidente.BackColor = &H8000000F
                cmdModoVotaPresidente.Caption = "Habilitar voto del Presidente"
            End If
        End If
        
        'botones de votacion
        If mEstadoSesion <> "abierta" Then 'Aca se revisaba si era paslis, pero no creo sea necesario.
            cmdModoVotaPresidente.Enabled = False
            cmdCancelarVotacion.Enabled = False
            cmdCierreVotacion.Enabled = False
            cmdNuevaSesion.Caption = "Abrir s&esión"
            If cmdExpresionesMinoria.Enabled <> True And (LCase(dcTipoOperacion.Tag) = "quorum") Then
                cmdExpresionesMinoria.Enabled = True
            ElseIf cmdExpresionesMinoria.Enabled = True And (LCase(dcTipoOperacion.Tag) <> "quorum") Then
                cmdExpresionesMinoria.Enabled = False
            End If
            cmdNuevaSesion.BackColor = &H8080FF
            cmdModoVotaPresidente.Enabled = False
            If (LCase(dcTipoOperacion.Tag) = "paslis") Then
                If cmdVotacion.Caption <> "Pase de l&ista" Then
                    cmdVotacion.Caption = "Pase de l&ista"
                End If
                If cmdVotacion.Enabled <> True Then
                    cmdVotacion.Enabled = True
                End If
                Select Case mEstadoVotacion
                    Case Is = "cancelada", "canpas"
                        cmdVotacion.Caption = "&Inicializar"
                        If mRefrescarConsola = True Then
                            cmdVotacion.Enabled = True
                        End If
                    Case Is = "esperafin"
                        cmdVotacion.Caption = "&Inicializar"
                        dcTipoOperacion.Enabled = False
                        If mRefrescarConsola = True Then
                            cmdVotacion.Enabled = True
                        End If
                    Case Is = "inipas"
                        cmdCancelarVotacion.Enabled = True
                        cmdVotacion.Enabled = False
                        dcTipoOperacion.Enabled = False
                    Case Is = "votando"
                        cmdVotacion.Caption = "Pase de l&ista"
                        cmdVotacion.Enabled = False
                    Case Else
                        cmdVotacion.Caption = "Pase de l&ista"
                        If mRefrescarConsola = True Then
                            cmdVotacion.Enabled = True
                        End If
                End Select
            Else
                If cmdVotacion.Caption <> "Sesion NO Abierta" Then
                    cmdVotacion.Caption = "Sesion NO Abierta"
                End If
                If cmdVotacion.Enabled <> False Then
                    cmdVotacion.Enabled = False
                End If
            End If
        Else
            cmdNuevaSesion.Caption = "Nueva s&esión"
            cmdNuevaSesion.BackColor = &H8000000F
            If InStr(cmdExpresionesMinoria.Caption, "OFF") > 0 Then
                cmdExpresionesMinoria_Click
            End If
            If cmdExpresionesMinoria.Enabled <> False Then
                cmdExpresionesMinoria.Enabled = False
            End If
            Select Case Trim(rstVector!Identificador_tipo_de_operacion)
                Case Is = "votnom", "votnum"
                   Select Case mEstadoVotacion
                        Case Is = "cancelada", "finalizada"
                            If mEstadoVotacion = "finalizada" And mRefrescarConsola = True Then
                                If nImpresion = False Then
                                    tmAutoCaptura.Enabled = False
                                    frmImpresion.Show vbModal, Me
                                    nImpresion = True
                                End If
                            End If
                            cmdVotacion.Caption = "&Inicializar"
                            cmdListaSiguiente.Visible = False
                            cmdListaAnterior.Visible = False
                            If mRefrescarConsola = True Then
                                cmdVotacion.Enabled = True
                            End If
                            'txtPendientesEmitirVoto.Visible = False
                            txtPendientesEmitirVoto.Enabled = False
                            cmdCancelarVotacion.Enabled = False
                            cmdCierreVotacion.Enabled = False
                            cmdModoVotaPresidente.Enabled = False
                        Case Is = "espera"
                            'txtPendientesEmitirVoto.Visible = False
                            cmdModoVotaPresidente.Enabled = True
                            txtPendientesEmitirVoto.Enabled = False
                            If TxtQuorum.Caption = "QUORUM" Then
                                If mRefrescarConsola = True Then
                                    cmdVotacion.Enabled = True
                                End If
                                cmdVotacion.Caption = "Votac&ión"
                                cmdCancelarVotacion.Enabled = False
                                cmdCierreVotacion.Enabled = False
                                cmdModoVotaPresidente.Enabled = True
                            Else
                                cmdVotacion.Enabled = False
                                cmdVotacion.Caption = "Sin QUORUM"
                                cmdCancelarVotacion.Enabled = False
                                cmdCierreVotacion.Enabled = False
                                cmdModoVotaPresidente.Enabled = False
                            End If
                         Case Is = "votando", "larga"
                                cmdModoVotaPresidente.Enabled = False
                                txtPendientesEmitirVoto.Enabled = True
                                'txtPendientesEmitirVoto.Visible = True
                                cmdVotacion.Enabled = False
                                cmdModoVotaPresidente.Enabled = False
                                cmdVotacion.Caption = "Votac&ión"
                                ControlesHabilitados = False
                                txtPendientesEmitirVoto.Caption = IIf(rstVector!Pendientes_Emitir_Voto >= 0, rstVector!Pendientes_Emitir_Voto, "-")
                                If mRefrescarConsola = True Then
                                    If txtTitulo.Text <> "MANTENIMIENTO DEL SISTEMA SQV" Then
                                        cmdCancelarVotacion.Enabled = True
                                    Else
                                        cmdCancelarVotacion.Enabled = False
                                    End If
                                    If Trim(mEstadoVotacion) = "larga" Then
                                        cmdCierreVotacion.Enabled = True
                                        cmdListaSiguiente.Visible = True
                                        cmdListaAnterior.Visible = True
                                    Else
                                        cmdCierreVotacion.Enabled = False
                                    End If
                                End If
                          Case Is = "empate"
                                cmdModoVotaPresidente.Enabled = False
                                cmdVotacion.Enabled = False
                                cmdCancelarVotacion.Enabled = False
                                cmdModoVotaPresidente.Enabled = False
                                cmdVotacion.Caption = "Votac&ión"
                                ControlesHabilitados = False
                                'txtPendientesEmitirVoto.Visible = False
                                txtPendientesEmitirVoto.Enabled = False
                                If mRefrescarConsola = True Then
                                    'cmdCancelarVotacion.Enabled = True
                                    cmdCierreVotacion.Enabled = True
                                End If
                    End Select
                Case Is = "quorum"
                   cmdModoVotaPresidente.Enabled = False
                   cmdVotacion.Enabled = False
                   cmdModoVotaPresidente.Enabled = False
                   cmdListaSiguiente.Visible = False
                   cmdListaAnterior.Visible = False
                   cmdVotacion.Caption = "Votación"
                   'txtPendientesEmitirVoto.Visible = False
                   txtPendientesEmitirVoto.Enabled = False
                Case Is = "paslis"
                    cmdModoVotaPresidente.Enabled = False
                    Select Case mEstadoVotacion
                        Case Is = "cancelada", "canpas"
                            cmdVotacion.Caption = "&Inicializar"
                            If mRefrescarConsola = True Then
                                cmdVotacion.Enabled = True
                            End If
                        Case Is = "esperafin"
                            cmdVotacion.Caption = "&Inicializar"
                            If dcTipoOperacion.Enabled = True Then
                                dcTipoOperacion.Enabled = False
                            End If
                            If mRefrescarConsola = True Then
                                cmdVotacion.Enabled = True
                            End If
                            If mEstadoVotacion = "esperafin" Then
                                If nImpresion = False Then
                                    frmImpresion.Show vbModal, Me
                                    nImpresion = True
                                End If
                            End If
                        Case Is = "inipas"
                            cmdCancelarVotacion.Enabled = True
                            cmdVotacion.Enabled = False
                            If dcTipoOperacion.Enabled = True Then
                                dcTipoOperacion.Enabled = False
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
                    cmdModoVotaPresidente.Enabled = False
                    'txtPendientesEmitirVoto.Visible = False
                   txtPendientesEmitirVoto.Enabled = False
            End Select
        End If
        'orador
        If Not (Trim(mOrador) = (Trim(rstVector!Orador))) Then
            mOrador = (Trim(rstVector!Orador))
            Call BuscarOrador(mOrador)
        End If
        'Modos de identificacion
        mModo_Ident_Nom = mId(Trim(rstVector!Identificador_de_Formulario), 1, 1) = "1"
        Modo_Vota_Presidente = IIf(rstVector("modo_vota_presidente") = 1, True, False)
        mModo_Presencia_Nom = mId(Trim(rstVector!Identificador_de_Formulario), 2, 1) = "1"
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
                    MensajesSQV.PruebaScanConfirmada mMensajeOperador
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

        ' *********************************************************************************
        If Trim(txtTitulo.Text) = "MANTENIMIENTO DEL SISTEMA SQV" Then
            cTemp = False
        Else
            cTemp = True
        End If
        'cmdCancelarVotacion.Enabled = cTemp
        cmdPeriodoLegislativo.Enabled = cTemp
        cmdNuevaSesion.Enabled = cTemp
        cmdCambiarSesion.Enabled = cTemp
        cmdTitulo.Enabled = cTemp
        cmdPresidente.Enabled = cTemp
        cmdTituloRapido.Enabled = cTemp
        cmdTituloBlanco.Enabled = cTemp
        cmdCambiarNumeroReunion.Enabled = cTemp
        cmdSeleccionarOrador.Enabled = cTemp
        cmdLimpiarOrador.Enabled = cTemp
        If Trim(txtTitulo.Text) = "MANTENIMIENTO DEL SISTEMA SQV" Then
            cmdMantenimiento.Enabled = True
            txtTitulo.Enabled = False
        Else
            txtTitulo.Enabled = cTemp
            cmdMantenimiento.Enabled = cTemp
        End If
        If txtTitulo.Text <> "MANTENIMIENTO DEL SISTEMA SQV" Then
            cmdCarteles.Enabled = cTemp
        Else
            cmdCarteles.Enabled = True
        End If
        If (LCase(dcTipoOperacion.Tag) = "quorum") Then
            cmdModoNominal.Enabled = True
        Else
            cmdModoNominal.Enabled = False
        End If
    Else
        cmdModoVotaPresidente.Enabled = False
        cTemp = False
        cmdVotacion.Enabled = False
        cmdCancelarVotacion.Enabled = cTemp
        cmdExpresionesMinoria.Enabled = cTemp
        cmdPeriodoLegislativo.Enabled = cTemp
        cmdNuevaSesion.Enabled = cTemp
        cmdCambiarSesion.Enabled = cTemp
        cmdTitulo.Enabled = cTemp
        cmdPresidente.Enabled = cTemp
        cmdTituloRapido.Enabled = cTemp
        cmdTituloBlanco.Enabled = cTemp
        cmdCambiarNumeroReunion.Enabled = cTemp
        cmdSeleccionarOrador.Enabled = cTemp
        cmdLimpiarOrador.Enabled = cTemp
        cmdMantenimiento.Enabled = cTemp
        cmdCarteles.Enabled = cTemp
        cmdModoNominal.Enabled = False
        ControlesHabilitados = False
        cmdMantenimiento.Enabled = False
        cmdCarteles.Enabled = False
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
        " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Ordinario' WHEN 'E' THEN 'Especial' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión: ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Tabla' WHEN 'E' THEN 'Especial' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - Próximo Nº de acta: ' + CAST(actas.Número_de_Acta AS Varchar(5)) END AS DescripcionPeriodoLegislativo, CASE TipoMayoriaQuorum.descripcion WHEN '120' THEN 'Más de la mitad' WHEN '121' THEN 'La mitad más uno' ELSE 'MAN' END AS DescripcionTipoQuorum " & _
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
        If LCase(Trim(rstVector!Estado_de_votacion)) = "inipas" Or LCase(Trim(rstVector!Estado_de_votacion)) = "esperafin" Then
            dcTipoOperacion.Enabled = False
        Else
            dcTipoOperacion.Enabled = True
        End If
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
        ModificarTiempo = pModo
        'cmdListar.Enabled = pModo
        txtNumeroReunion.Enabled = pModo
        cmdCambiarNumeroReunion.Enabled = pModo
        cmdSeleccionarOrador.Enabled = pModo
        cmdLimpiarOrador.Enabled = pModo
        'cmdSubirBajarReunion.Enabled = pModo
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
        ModificarTiempo = False
        'cmdListar.Enabled = False
        txtNumeroReunion.Enabled = False
        cmdCambiarNumeroReunion.Enabled = False
        cmdSeleccionarOrador.Enabled = False
        cmdLimpiarOrador.Enabled = False
        'cmdSubirBajarReunion.Enabled = False
    End If

End Property
Private Property Let habilitadoBloqueControlesResultado(pModo As Boolean)
    If txtSi.Visible <> pModo Then
        txtSi.Visible = pModo
        lblSi.Visible = pModo
        txtNo.Visible = pModo
        lblNo.Visible = pModo
        txtAbs.Visible = pModo
        lblAbs.Visible = pModo
        txtResultado.Visible = pModo
        If pModo = True Then
            cmdListaSiguiente.Visible = False
            cmdListaAnterior.Visible = False
        End If
    End If
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
Private Sub BuscarOrador(pOrador As String)
    
    If pOrador <> "" Then
        Set rstOrador = New ADODB.Recordset
        SetearRs "SELECT Apellido + ', ' + Nombre AS Orador, Es_Legislador FROM Legisladores WHERE Es_Legislador " & IIf(mModoMantenimiento, ">=0", "=1") & " and Id = '" & Trim(pOrador) & "'", rstOrador
        If rstOrador.EOF = False Then
            lblOrador.Caption = rstOrador!Orador
        Else
            lblOrador.Caption = ""
        End If
    Else
        lblOrador.Caption = ""
    End If
    'free mem
    If rstOrador.State = adStateOpen Then
        rstOrador.Close
    End If
    Set rstOrador = Nothing
End Sub
Private Sub AsignarDatosBancas()
    On Error Resume Next
    Dim i     As Integer
    Dim clave As String
    Dim bIdx As Long
    Dim strBloque   As String
    Dim c As Integer
    Dim Col As Single
    Dim Row As Single
    Dim Xo As Single
    Dim Yo As Single
    Dim Xp As Single
    Dim Yp As Single
    Dim BloqueContador As Integer
    Dim ColWidth As Single
    Dim RowHeight As Single
    Dim Rows As Single
    Dim ShpHeight As Single
    Dim shpWidth As Single
    Dim FontHeight    As Single
    Dim bcIdx As Integer
    ColWidth = 2600
    RowHeight = 330
    FontHeight = 14
    Rows = 18
    Xo = 240
    Yo = 60
    ShpHeight = 315
    shpWidth = 315
    cntRepintarLista = cntRepintarLista + 1
    If MR2.Visible Then
        For i = 0 To UBound(datBloque)
            datBloque(i).Ausentes = 0
            datBloque(i).Presentes = 0
            datBloque(i).Identificados = 0
            datBloque(i).NoIdentificados = 0
        Next
    End If
    strBloque = ""
    BloqueContador = 0
    c = 0
    For i = 0 To UBound(mVectorColores)
        Dim X As String
        clave = i
        If i = 1 Then
            i = i
        End If
        'Los colores se procesan de lado SQV (definidos en CONSTANTES)
        'Sin embargo se asignan en Consola
        If i = 0 Or (cmdVotacion.Caption = "&Inicializar" Or cmdVotacion.Enabled = False) And dcTipoOperacion.Tag <> "quorum" Then
            shpBanca(i).FillColor = mColores(mVectorColores(clave))
        Else
            If mVectorIdentificacion(i) <> "0" Then
                shpBanca(i).FillColor = &HC0C000
            Else
                shpBanca(i).FillColor = mColores(mVectorColores(clave))
            End If
        End If
        X = mVectorColores(clave)
        If shpBanca(i).FillColor = 12632064 Then
            ctrBanca(i).ForeColor = vbBlack
        Else
            ctrBanca(i).ForeColor = mColoresFuentes(X)
        End If
        datBanca(i).legislador = mVectorIdentificacion(i)
        datBanca(i).Presencia = mVectorPresencia(i)
        If MR2.Visible Then
            bcIdx = datLista(i).BancaDefecto
            bIdx = datLista(i).BloqueIndex
            If mVectorPresencia(bcIdx) <> "1" Then
                datBloque(bIdx).Ausentes = datBloque(bIdx).Ausentes + 1
            Else
                datBloque(bIdx).Presentes = datBloque(bIdx).Presentes + 1
            End If
            If mVectorIdentificacion(bcIdx) <> "0" Then
                datBloque(bIdx).Identificados = datBloque(bIdx).Identificados + 1
            Else
                datBloque(bIdx).NoIdentificados = datBloque(bIdx).NoIdentificados + 1
            End If
            If mVectorIdentificacion(bcIdx) <> "0" Then
                lblLegis(i).FontBold = True
                If datLista(i).LegisladorDefecto <> mVectorIdentificacion(bcIdx) And bcIdx > 0 Then
                    MsgBox "La consola detecta un legislador identificado en la banca " & bcIdx & " y esta no es la banca que normalmente tiene asignada." & vbCrLf & "La modalidad de información que usted está utilizando puede presentar información erronea ya que trabaja en base a la asignación habitual de legisladores a bancas y dicho patrón no se está cumpliendo." & vbCrLf & "Se sugiere volver al modo mapa del recinto.", vbExclamation
                    flgRepintarLista = True
                End If
            Else
                lblLegis(i).FontBold = False
            End If
            If mVectorPresencia(bcIdx) <> "1" Then
                lblLegis(i).ForeColor = vbRed
            Else
                lblLegis(i).ForeColor = vbBlack
            End If
            shpBanka(i).FillColor = mColores(mVectorColores(bcIdx))
            lblBanca(i).ForeColor = mColoresFuentes(mVectorColores(bcIdx))
            
            If cntRepintarLista Mod 10 = 0 Then
                cntRepintarLista = 0
                'Determino posición de prox elemento
                If strBloque <> datBloque(bIdx).Nombre Then
                    Col = c \ Rows
                    Row = c Mod Rows
                    Xp = Xo + (Col * ColWidth)
                    Yp = Yo + (Row * RowHeight)
                    strBloque = datBloque(bIdx).Nombre
                    lblBloque(bIdx).Caption = strBloque
                    lblBloque(bIdx).Width = ColWidth
                    lblBloque(bIdx).AutoSize = True
                    lblBloque(bIdx).Left = Xp
                    lblBloque(bIdx).Top = Yp
                    lblBloque(bIdx).Visible = True
                    c = c + 1
                    Col = c \ Rows
                    Row = c Mod Rows
                    Xp = Xo + (Col * ColWidth)
                    Yp = Yo + (Row * RowHeight)
                    lblBloqueInfo(bIdx).Width = ColWidth
                    lblBloqueInfo(bIdx).AutoSize = True
                    lblBloqueInfo(bIdx).Left = Xp
                    lblBloqueInfo(bIdx).Top = Yp
                    lblBloqueInfo(bIdx).Visible = True
                    lblBloqueInfo(bIdx).FontBold = False
                    lblBloqueInfo(bIdx).FontSize = 10
                    c = c + 1
                End If
                If (lblLegis(i).Caption <> "") And Not (flPresidenteLegislador And bcIdx = 0) Then
                    lblBanca(i).Tag = strBloque
                    lblBanca(i).Caption = datLista(i).BancaDefecto
                    If c Mod Rows = 0 Then
                        c = c + 1
                    End If
                    Col = c \ Rows
                    Row = c Mod Rows
                    Xp = Xo + (Col * ColWidth)
                    Yp = Yo + (Row * RowHeight)
                    shpBanka(i).Height = ShpHeight
                    shpBanka(i).Width = shpWidth
                    shpBanka(i).FillStyle = vbFSSolid
                    lblBanca(i).FontSize = FontHeight
                    lblBanca(i).Visible = True
                    lblLegis(i).Visible = True
                    lblLegis(i).Width = ColWidth
                    lblBanca(i).AutoSize = False
                    lblBanca(i).Width = shpWidth
                    shpBanka(i).Left = Xp
                    shpBanka(i).Top = Yp
                    shpBanka(i).Visible = True
                    shpBanka(i).BorderColor = vbBlack
                    lblBanca(i).Top = shpBanka(i).Top - ((lblBanca(i).Height - shpBanka(i).Height) / 2)
                    lblBanca(i).Left = shpBanka(i).Left - ((lblBanca(i).Width - shpBanka(i).Width) / 2)
                    lblLegis(i).Top = shpBanka(i).Top - ((lblLegis(i).Height - shpBanka(i).Height) / 2)
                    lblLegis(i).Left = shpBanka(i).Left + shpBanka(i).Width + 45
                    c = c + 1
                Else
                    lblBanca(i).Visible = False
                    lblLegis(i).Visible = False
                    shpBanka(i).Visible = False
                End If
            End If
            If Trim(strIdPresidenteRecinto) = CStr(datLista(i).LegisladorDefecto) Then
                If flPresidenteLegislador Then
                    lblLegis(i).Caption = lblPresidente.Caption
                    lblLegis(i).Tag = strIdPresidenteRecinto
                    lblBanca(i).Caption = "0"
                    shpBanka(i).BorderColor = vbWhite
                    shpBanka(i).FillColor = vbCyan
                    lblLegis(i).FontBold = True
                    lblLegis(i).ForeColor = vbBlack
                Else
                    lblLegis(i).Caption = datLista(i).LegisladorDefectoNombre
                    lblLegis(i).Tag = datLista(i).LegisladorDefecto
                End If
            Else
                lblLegis(i).Caption = datLista(i).LegisladorDefectoNombre
                lblLegis(i).Tag = datLista(i).LegisladorDefecto
            End If
        End If
    Next i
    flgRepintarLista = False
    If MR2.Visible Then
        For i = 0 To UBound(datBloque)
            lblBloque(i).Caption = datBloque(i).Nombre
            lblBloqueInfo(i).Caption = "Presentes " & datBloque(i).Presentes & "/" & datBloque(i).TotalLegisladores & ". Identificados: " & datBloque(i).Identificados
        Next
    End If
End Sub

Private Sub cmdPresidente2_Click()

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

Private Sub cmdSubirBajarReunion_Change()
   'MsgBox cmdSubirBajarReunion.Value & "-" & CInt(txtNumeroReunion.Text)
   'If cmdSubirBajarReunion.Value > 0 And cmdSubirBajarReunion.Value >= CInt(txtNumeroReunion.Text) Then
   '     txtNumeroReunion.Text = CInt(txtNumeroReunion.Text) + 1
   ' Else
   '     If CInt(txtNumeroReunion.Text) > 0 Then 'CInt(txtNumeroReunion.Text) > 0 Then
   '         txtNumeroReunion.Text = CInt(txtNumeroReunion.Text) - 1
   '     End If
   ' End If
    
End Sub

Private Sub cmdReiniciarDerecho_Click()
If MsgBox("¿Desea reiniciar el cartel derecho?", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
    Call Shell(App.Path & "\reset2.bat")
End If
End Sub

Private Sub cmdReiniciarIzquierdo_Click()
If MsgBox("¿Desea reiniciar el cartel izquierdo?", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
    Call Shell(App.Path & "\reset1.bat")
End If
End Sub

Private Sub cmdSimular_Click()
    Datos.GrabarMensaje "simulacion?votonegativo", " ", , True
End Sub

Private Sub cmdTaparQuorum_Click()
MensajesSQV.cambiarCartelEncendido "1"
End Sub

Private Sub cmdTitulo_Click()
    Dim Titulo As New frmTituloActa
    
    If (lblCodigoSesion.Caption <> "") And (lblSesion.Caption <> "") Then
        If (gTipoUsuario = 0) Or (gTipoUsuario = 2) Then
            Titulo.MostrarDatos lblCodigoSesion.Tag, lblNroSesion.Caption
            Titulo.Show vbModal
            Set Titulo = Nothing
        Else
            MsgBox "Ud. no dispone de permisos para realizar esta acción.", vbInformation + vbOKOnly
        End If
    Else
        MsgBox "Debe haber una sesión abierta para visualizar su Órden del día.", vbInformation + vbOKOnly
    End If
End Sub
Private Sub SeleccionarOrador()
    Dim strId       As String
    Dim blCondicion As Boolean
    If PermisosTotales.ConsultaABMLegislador = 0 Then
        MsgBox "El usuario no tiene permisos para esta tarea", vbInformation + vbOKOnly, "Usuario sin permisos!!"
        Exit Sub
    End If
    
    EsSeleccionDeOrador = True
    Dim xFrm As New frmAsignarLegislador
    xFrm.Caption = "SELECCIÓN DE ORADOR"
    xFrm.cmdAsignar.Caption = "&Asignar Orador"
    xFrm.mostrarLegisladores (0)
    xFrm.Show vbModal, Me
    Set xFrm = Nothing
'    EsSeleccionDeOrador = False
'    strId = Trim(frmABMLegisladores.lblid.Caption)
'    blCondicion = True
'    If LCase(strId) <> "nothing" Then
'        cambiarOrador (strId)
'    Else
'        MsgBox strId
'    End If
    cmdSeleccionarOrador.Enabled = True
End Sub

Private Sub cmdTituloRapido_Click()
    MensajesSQV.CambioTituloActa 0, txtTituloTemp.Text
End Sub
Private Sub cmdTituloBlanco_Click()
    'txtTitulo.Text = ""
    'lblTituloActa.Caption = ""
    MensajesSQV.CambioTituloActa 0, ""
    txtTitulo.Text = ""
    txtTituloTemp.Text = ""
End Sub
Private Sub cmdVotacion_Click()
    Fue9999 = False
    Impresa = False
    Ultimo_Periodo = Trim(lblCodigoSesion.Caption) & mId(lblPeriodo.Caption, 1, 1) & mId(lblSesion.Caption, 1, 1)
    Ultima_Sesion = lblNroSesion.Caption
    Ultimo_Acta = lblActa.Caption
    Select Case cmdVotacion.Caption
        Case Is = "Votac&ión"
            MensajesSQV.inicioVotacion
            cmdSimular.Enabled = True
            nImpresion = False
        Case Is = "&Inicializar"
            MensajesSQV.inicializar
            cmdSimular.Enabled = False
        Case Is = "Pase de l&ista"
            nImpresion = False
            If cmdNuevaSesion.Caption = "Abrir s&esión" Then
                Fue9999 = True
            End If
            'If mActaIniciada = 0 Then
                mActaIniciada = Val(lblActa.Caption)
                MensajesSQV.paseLista
                ' sbStatus.Panels(3).Text = "Pase de lista INICIADO..."
            'Else
                ' sbStatus.Panels(3).Text = "ACTA YA INICIADO"
            'End If
    End Select
    cmdVotacion.Enabled = False
End Sub
Private Sub cmdCambiarSesion_Click()
    If lblCodigoSesion.Caption <> "" Then
        Dim sesion As New frmCambiarSesion
        If sesion.MostrarDatos(lblCodigoSesion.Tag) = True Then
            sesion.ActualizarDatos = True
            sesion.Show vbModal
        End If
        Set sesion = Nothing
    End If
End Sub
Private Sub cmdNuevaSesion_Click()
Dim LPeriodo As String
LPeriodo = ""
    If cmdNuevaSesion.Caption = "Nueva s&esión" Then
        If lblCodigoSesion.Caption <> "" Then
            Dim xF As frmElegirPresidente
            Set xF = New frmElegirPresidente
            xF.cmdMantener.Enabled = True
            xF.cmdCancelar.Enabled = False
            xF.Show vbModal, Me
            Set xF = Nothing
'            Dim Sesion As New frmCrearSesion
'            Sesion.AgregarDatos lblCodigoSesion.Tag
'            Sesion.Show vbModal
'            Set Sesion = Nothing
            LPeriodo = lblCodigoSesion.Tag
            Call CreaSesion(LPeriodo)
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
Private Function nuevoNumeroSesion(cPeriodo As String) As Long
    Dim rstNumero As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT max(Sesión) as maximo FROM sesion WHERE (Sesión <> 9999) AND (Sesión <> -1) AND Período_Legislativo LIKE '" & cPeriodo & "%'"
    SetearRs consulta, rstNumero
    If rstNumero.EOF = False Then
        If IsNull(rstNumero!maximo) = False Then
            nuevoNumeroSesion = rstNumero!maximo + 1
        Else
            nuevoNumeroSesion = 1
        End If
    Else
        nuevoNumeroSesion = 1
    End If
    If rstNumero.State = adStateOpen Then
        rstNumero.Close
    End If
    Set rstNumero = Nothing
End Function
Private Sub CreaSesion(cPerLeg As String, Optional Alerta As Boolean)
Dim consulta As String
Dim nSesion As Long
consulta = "DELETE FROM sesion WHERE RTrim(LTrim(Estado_sesión)) = 'nueva' AND Período_Legislativo LIKE '" & mId(cPerLeg, 1, 4) & "%'"
Call EjecutarSQL(consulta)
nSesion = nuevoNumeroSesion(mId(cPerLeg, 1, 4))
If (Right(cPerLeg, 1) = "p") Then
    nSesion = -1
End If
consulta = "INSERT INTO sesion (Período_Legislativo, Sesión,Fecha_de_inicio, Próximo_Acta, Estado_sesión, Prorroga) " _
                   & " VALUES ('" & cPerLeg & "','" & nSesion & "','" & Format(Now(), "yyyymmdd HH:MM:SS") & "','" & 1 & "','nueva',0)"
If (nSesion = -1) Then
    Dim s As String
    Dim rs As New ADODB.Recordset
    SetearRs "SELECT * FROM sesion WHERE Período_Legislativo = '" & cPerLeg & "' AND Sesión = -1", rs
    If (rs.EOF) Then
        EjecutarSQL (consulta)
    End If
Else
    EjecutarSQL (consulta)
End If
MensajesSQV.cambiosesion Trim(Str(nSesion))
If Alerta = False Then
    Call MsgBox("Sesión Nº" & nSesion & " creada satisfactoriamente", vbInformation)
End If
End Sub



Private Sub ctrBanca_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim t_id As Long
Dim Rinfo As ADODB.Recordset
Dim pic As ADODB.Stream
t_id = datBanca(Index).legislador
If Index = 3 Then
    t_id = t_id
End If
If t_id <> 0 Then
    Set Rinfo = New ADODB.Recordset
    SetearRs "SELECT nombre,apellido,bloque_politico,distritos.distrito AS Provincia FROM Legisladores LEFT OUTER JOIN distritos ON Legisladores.distrito = distritos.id_distrito WHERE id='" & t_id & "'", Rinfo
    If Rinfo.EOF Then
        lblPICNombre.Caption = "NO ENCONTRADO"
    Else
        picFlotante.Picture = LoadPicture(GetFoto(Trim(Str(t_id))))
        lblPICBanca.Caption = Trim(Str(Index))
        lblPICNombre.Caption = Rinfo.Fields("nombre")
        lblPICApellido.Caption = Rinfo.Fields("apellido")
        lblPICBloque.Caption = IIf(IsNull(Rinfo.Fields("bloque_politico")), "Bloque Vacio", Rinfo.Fields("bloque_politico"))
        lblPICProvincia.Caption = IIf(IsNull(Rinfo.Fields("Provincia")), "Provincia Vacio", Rinfo.Fields("Provincia"))
    End If
    Rinfo.Close
    Set Rinfo = Nothing
    If ctrBanca(Index).Top > 4820 Then 'Limite banca 174
        If pctInfo.Top <> 900 Then
            pctInfo.Top = 900
        End If
    Else
        If pctInfo.Top <> 6720 Then
            pctInfo.Top = 6720
        End If
    End If
    If Y > 15 Then
        If pctInfo.Visible <> True Then
            pctInfo.Visible = True
        End If
    Else
        If pctInfo.Visible <> False Then
            pctInfo.Visible = False
        End If
    End If
Else
    If pctInfo.Visible <> False Then
        pctInfo.Visible = False
    End If
End If
'On Error Resume Next
'Dim pic As ADODB.Stream
'If MouseEnBanca = False Then
'    Dim t_id As Long
'    t_id = datBanca(Index).legislador
'    MouseEnBanca = True
'    If (t_id <> 0) Then
'        Dim Rinfo As New ADODB.Recordset
'        Dim Nombre_Legislador As String
'        Dim Apellido_Legislador As String
'        Dim Grupo As String
'        Dim Bloque As String
'        If IMAGENES_RAPIDAS_HABILITADAS = False Then
'            SetearRs "SELECT nombre,apellido,bloque_plitico,grupo_politico FROM Legisladores WHERE id='" & t_id & "'", Rinfo
'            If (Rinfo.RecordCount > 0) Then
'                Nombre_Legislador = Rinfo.Fields("nombre")
'                Apellido_Legislador = Rinfo.Fields("apellido")
'                Bloque = IIf(IsNull(Rinfo.Fields("bloque_politico")), "Nulo", Rinfo.Fields("bloque_politico"))
'                Grupo = IIf(IsNull(Rinfo.Fields("grupo_politico")), "Nulo", Rinfo.Fields("grupo_politico"))
'                SetLabelInfo Nombre_Legislador & " " & Apellido_Legislador, Index
'            Else
'                SetLabelInfo "ID Inexistente", Index
'            End If
'            Rinfo.Close
'        Else
'            SetearRs "SELECT nombre,apellido,PICTURE,bloque_politico,grupo_politico FROM Legisladores WHERE id='" & t_id & "'", Rinfo
'            If (Rinfo.RecordCount > 0) Then
'                Call BorrarImgTemp
'                If (IsNull(Rinfo.Fields("nombre")) = False) Then
'                    Nombre_Legislador = Rinfo.Fields("nombre")
'                End If
'                If (IsNull(Rinfo.Fields("apellido")) = False) Then
'                    Apellido_Legislador = Rinfo.Fields("apellido")
'                End If
'                Bloque = IIf(IsNull(Rinfo.Fields("bloque_politico")), "Nulo", Rinfo.Fields("bloque_politico"))
'                Grupo = IIf(IsNull(Rinfo.Fields("grupo_politico")), "Nulo", Rinfo.Fields("grupo_politico"))
'                If (IsNull(Rinfo.Fields("PICTURE")) = False) Then
'                    Set pic = New ADODB.Stream
'                    pic.Type = adTypeBinary
'                    pic.Open
'                    pic.Write Rinfo.Fields("PICTURE")
'                    pic.SaveToFile App.Path & "\temp.jpg", adSaveCreateOverWrite
'                    pctFotoRapida.Picture = LoadPicture(App.Path & "\temp.jpg")
'                Else
'                    Set pctFotoRapida.Picture = Nothing
'                End If
'                SetLabelInfo Nombre_Legislador & " " & Apellido_Legislador, Index
'            Else
'                SetLabelInfo "ID Inexistente", Index
'            End If
'            Rinfo.Close
'        End If
'    End If
'End If
End Sub
Private Sub BorrarImgTemp()
On Error Resume Next
Kill (App.Path & "\temp.jpg")
End Sub
Private Sub ctrBanca_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NumeroBanca As Long
    Dim PermitirPrueba      As Boolean
    Dim PermitirIdentificar As Boolean
    Dim PermitirCambioVoto  As Boolean
    Dim PermitirAbstener    As Boolean
    Dim legislador As String
    Dim bloque As String
    Dim Departamento As String
    
    If Button = 2 Then
        EsSeleccionDeOrador = False
        EvaluarPermisosOperacionesBanca Index, PermitirPrueba, PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
        If datBanca(Index).legislador = "0" Then
            'info.MostrarDatos Index, datBanca(Index).legislador, PermitirPrueba, PermitirIdentificar, PermitirCambioVoto, PermitirAbstener}
            frmNuevoInfo.BancaID = Index
            frmNuevoInfo.Show vbModal, Me
        Else
            'info.MostrarDatos Index, , , PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
            frmNuevoInfo.BancaID = Index
            frmNuevoInfo.Show vbModal, Me
        End If
        IndexBanca.Caption = Index
        If False Then
            NumeroBanca = Val(ctrBanca(Index).Caption)
            EvaluarPermisosOperacionesBanca Index, PermitirPrueba, PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
            mnuAsignarID.Enabled = IIf(NumeroBanca = 0, False, PermitirIdentificar)
            mnuAbstener.Enabled = PermitirAbstener
            DeterminarLeyendaMenuAbstencion NumeroBanca
            Sep00.Visible = True
            TraerInfoLegislador datBanca(Index).legislador, legislador, bloque, Departamento
            mnuID.Caption = legislador
            mnuBloque.Visible = (bloque <> "")
            mnuBloque.Caption = bloque
            mnuDepartamento.Visible = (Departamento <> "")
            mnuDepartamento.Caption = Departamento
            Me.PopupMenu mnuPopUP, , , , mnuID
        End If
    Else
        If mEstadoVotacion = "espera" Then
            ctrBanca(Index).Enabled = False
            If pctInfo.Visible = True Then
                pctInfo.Visible = False
            End If
            MouseEnBanca = False
            lblNombreRapido.Left = -5000
            If False Then
                'modo info
                Set info = New frmInformacionBanca
                EvaluarPermisosOperacionesBanca Index, PermitirPrueba, PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
                If datBanca(Index).legislador <> "" Then
                    info.MostrarDatos Index, datBanca(Index).legislador, PermitirPrueba, PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
                Else
                    info.MostrarDatos Index, , , PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
                End If
                info.EstadoVotacion = mEstadoVotacion
                info.Show vbModal
                Set info = Nothing
            Else
                If Val(ctrBanca(Index).Caption) = "0" Then
                    'Nada por ahora
                ElseIf (dcTipoOperacion.BoundText = "votnom") Or (dcTipoOperacion.BoundText = "paslis") Or (dcTipoOperacion.BoundText = "quorum" And cmdModoNominal.Enabled = True And cmdModoNominal.Caption = "Deshabilitar identificación") Then
                    If IsNumeric(mVectorPresencia(Index)) = True Then
                        If mVectorPresencia(Index) = 1 And mVectorIdentificacion(Index) = "0" Then
                            If mRefrescarConsola = True Then
                                'modo cambio id5
                                Dim frmId As New frmAsignarLegislador
                                frmId.mostrarLegisladores Index
                                frmId.Show vbModal
                                Set frmId = Nothing
                            End If
                        End If
                    End If
                End If
            End If
            ctrBanca(Index).Enabled = True
        Else
            MsgBox "No puede asignar una identificación en este momento"
        End If
    End If
End Sub

Private Sub dcAbstencion_Change()
    dcAbstencion.Enabled = False
    If conFoco = True Then
        MensajesSQV.ModoVotacion dcAbstencion.BoundText
        conFoco = False
        dcAbstencion.Tag = ""
        'TxtQuorum.SetFocus
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
        dcBaseMayoria.Tag = ""
        txtTitulo.SetFocus
    End If
End Sub
Private Sub dcBaseMayoria_GotFocus()
    conFoco = True
End Sub
Private Sub dcBaseMayoria_LostFocus()
    conFoco = False
End Sub
Private Sub dcTipoMayoria_Change()
If TipMayCargo = True Then
    If conFoco = True Then
        MensajesSQV.cambiarTipoMayoriaVotacion dcTipoMayoria.BoundText
        dcTipoMayoria.Tag = ""
        txtTitulo.SetFocus
    End If
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
        'conFoco = False
        txtTitulo.SetFocus
    End If
    cmdAbstenciones.Enabled = True
    If LCase(dcTipoOperacion.BoundText) = "quorum" Then
        cmdModoNominal.Enabled = True
    Else
        cmdModoNominal.Enabled = False
    End If
End Sub
Private Sub determinarCantidadLegisladores()
    Dim rstAux As New ADODB.Recordset 'banana
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
        'TxtQuorum.SetFocus
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
    'MR2.Height = 8000
'    Screen.MousePointer = 0
'
'    If FlagBasePrueba Then
'        lblModoPrueba.Visible = True
'    Else
'        lblModoPrueba.Visible = False
'    End If
'    Call Levanta_Sqv
'    Call Levanta_Banca
If Modo_Prueba_Seleccionado = True Then
    lblEModoPrueba.Visible = True
    cmdSimular.Visible = True
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
    EntroAConsola = True
    If Error_Carga = True Then
        Unload Me
        Exit Sub
    End If
    TipMayCargo = False
    txtTituloTemp.Width = txtTitulo.Width
    txtTituloTemp.Height = txtTitulo.Height
    txtTituloTemp.Left = txtTitulo.Left
    txtTituloTemp.Top = txtTitulo.Top
    pctInfo.Visible = False
    frmFlotante.Show
    frmFlotante.Visible = False
    PrimeraVezMantenimiento = True
    MouseEnBanca = False
    lblNombreRapido.Left = -5000
    pctFotoRapida.Visible = False
    If (Screens_Habilitadas = True) Then
        Segundos_Screen = Obtener_Tiempo_Screens()
        Carpeta_Screen = Obtener_Directorio_Screens()
        tmScreens.Interval = Segundos_Screen
        tmScreens.Enabled = True
    End If
    'cargo el winsock p/levantar srvr
    'Ws.RemoteHost = "192.168.1.254"
    'Ws.RemotePort = 10023
    'Ws.Close
    'Ws.Connect
    'fin
    xInicializando = True
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
    CargarColoresFuente
    Call ArmarPantalla
    'Antes de ArmarBancas cacheo los diputados
    Dim rsCur As New ADODB.Recordset
    Dim currSql As String
    currSql = "SELECT id, apellido, nombre, bloque_politico FROM legisladores_activos ORDER BY bloque_politico, apellido, nombre"
    Datos.SetearRs currSql, rsCur
    While Not rsCur.EOF
        currDiputados.Add Val(rsCur.Fields("id")), rsCur.Fields("apellido") & ", " & rsCur.Fields("nombre") & ";" & rsCur.Fields("bloque_politico")
        rsCur.MoveNext
    Wend
    Call ArmarBancas
    If Error_Carga = True Then
        Exit Sub
        Unload Me
    End If
    IniciaActualizacion = True
    Call BuscarPresidentePredeterminado
    Call BuscarOrador(mOrador)
    MensajesSQV.habilitarConsola
    ' sbStatus.Panels(1).Text = "Versión: " & App.Major & "." & App.Minor & "." & App.Revision
    reiniciarBancas
    blBanderaTimer = True
    'Andres
    ArmarBancasCartel
    CargarVectorLegisladoresPorBloque
    flgRepintarLista = True
    TipMayCargo = True
    xInicializando = False
'Andres End
nImpresion = False
End Sub
Private Sub BuscarPresidentePredeterminado()
    Dim strSql          As String
    Dim strIdPresidente As String
    Dim RsPres          As ADODB.Recordset
    Set RsPres = New ADODB.Recordset
    ' Levantar vector de identificacion
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
If MsgBox("¿Desea reiniciar las bancas?", vbQuestion + vbYesNo) = vbYes Then
    reiniciarTodasBancas
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    EntroAConsola = False
    If Error_Carga = False Then
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
    End If
End Sub
Private Sub frmInfoRapida_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseEnBanca = True Then
    MouseEnBanca = False
    'frmInfoRapida.Left = -5000
End If
End Sub
Private Sub imgA_Click(Index As Integer)
    ML(0).Visible = True
    ML(1).Visible = False
    MR1.Visible = True
    MR2.Visible = False
End Sub

Private Sub imgB_Click(Index As Integer)
    'Mostrar legisladores
    If False Then
        ML(0).Visible = False
        ML(1).Visible = True
        MR1.Visible = False
        MR2.Visible = True
        cntRepintarLista = 9
    Else
        'MsgBox "Función no habilitada"
    End If
End Sub
Private Sub lblBanca_DblClick(Index As Integer)
    Dim PermitirPrueba      As Boolean
    Dim PermitirIdentificar As Boolean
    Dim PermitirCambioVoto  As Boolean
    Dim PermitirAbstener    As Boolean
    Index = Val(lblBanca(Index).Caption)
    If False Then
        'modo info
        Set info = New frmInformacionBanca
        EvaluarPermisosOperacionesBanca Index, PermitirPrueba, PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
        If datLista(Index).legislador <> "" Then
            info.MostrarDatos Index, datBanca(Index).legislador, PermitirPrueba, PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
        Else
            info.MostrarDatos Index, , , PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
        End If
        info.EstadoVotacion = mEstadoVotacion
        info.Show vbModal
        Set info = Nothing
    Else
        If Val(lblBanca(Index).Caption) = "0" Then
            cmdpresidente_Click
        ElseIf (dcTipoOperacion.BoundText = "votnom") Or (dcTipoOperacion.BoundText = "paslis") Or (dcTipoOperacion.BoundText = "quorum" And cmdModoNominal.Enabled = True And cmdModoNominal.Caption = "Deshabilitar identificación") Then
            If mRefrescarConsola = True Then
                'modo cambio id5
                Dim frmId As New frmAsignarLegislador
                frmId.mostrarLegisladores Index
                frmId.Show vbModal
                Set frmId = Nothing
            End If
        End If
    End If
End Sub
Private Sub lblBanca_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NumeroBanca As Long
    Dim PermitirPrueba      As Boolean
    Dim PermitirIdentificar As Boolean
    Dim PermitirCambioVoto  As Boolean
    Dim PermitirAbstener    As Boolean
    Dim legislador As String
    Dim bloque As String
    Dim Departamento As String
    
    If Button = 2 Then
        IndexBanca.Caption = lblBanca(Index).Caption
        If False Then
            NumeroBanca = Val(lblBanca(Index).Caption)
            EvaluarPermisosOperacionesBanca Val(lblBanca(Index).Caption), PermitirPrueba, PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
            mnuAsignarID.Enabled = PermitirIdentificar
            mnuAbstener.Enabled = PermitirAbstener
            DeterminarLeyendaMenuAbstencion NumeroBanca
            Sep00.Visible = True
            TraerInfoLegislador datBanca(Index).legislador, legislador, bloque, Departamento
            mnuID.Caption = legislador
            mnuBloque.Visible = (bloque <> "")
            mnuBloque.Caption = bloque
            mnuDepartamento.Visible = (Departamento <> "")
            mnuDepartamento.Caption = Departamento
            Me.PopupMenu mnuPopUP, , , , mnuID
        End If
    End If
End Sub

Private Sub lblNombreRapido_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseEnBanca = True Then
    MouseEnBanca = False
    lblNombreRapido.Left = -5000
    If IMAGENES_RAPIDAS_HABILITADAS = True Then
        pctFotoRapida.Visible = False
    End If
End If
End Sub

Private Sub lblPresidente_Change()
'If Trim(lblPresidente.Caption) = "Seleccione el presidente antes de continuar" Then
'    cmdpresidente_Click
'End If
End Sub

Private Sub lblTipoAbstencion_Click()

End Sub

Private Sub MR1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pctInfo.Visible = True Then
    pctInfo.Visible = False
End If
End Sub
Private Sub pctFotoRapida_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseEnBanca = True Then
    MouseEnBanca = False
    lblNombreRapido.Left = -5000
    If IMAGENES_RAPIDAS_HABILITADAS = True Then
        pctFotoRapida.Visible = False
    End If
End If
End Sub
Private Sub pctInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctInfo.Visible = False
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub PicFlotante_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctInfo.Visible = False
End Sub

Private Sub Salir_Click()
    EntroAConsola = False
    Unload Me
    frmMenu.cmdConsola.Enabled = True
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

Private Sub cmdSeleccionarOrador_Click()
cmdSeleccionarOrador.Enabled = False
SeleccionarOrador
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
'    If dcTipoOperacion.SelectedItem = 2 Then
'        cmdModoNominal.Enabled = True
'    Else
'        cmdModoNominal.Enabled = False
'    End If
    If Error_Carga = False Then
        If mModo_Ident_Nom Then
            cmdModoNominal.Caption = "Deshabilitar identificación"
        Else
            cmdModoNominal.Caption = "Habilitar identificación"
        End If
    Else
        Unload Me
    End If
End Sub
Public Sub hacerSplitVector(ByVal pCadena As String, ByRef pVector() As String)
    pVector = Split(pCadena, mDelimitadorVector)
End Sub
Private Sub cmdModoNominal_Click()
    If (gTipoUsuario = 0) Or (gTipoUsuario = 2) Then
        If cmdModoNominal.Caption = "Habilitar identificación" Then
            Datos.GrabarMensaje "cambio?forzarids", " ", , True
        Else
            Datos.GrabarMensaje "cambio?forzaroffids", " ", , True
        End If
    Else
        MsgBox "Ud. no dispone de permisos para realizar esta acción.", vbInformation + vbOKOnly
    End If
End Sub
Private Property Let ControlesHabilitados(ByVal pModo As Boolean)
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If (UCase(ctrl.Name) <> "TXTQUORUM") And (ctrl.Name <> "Salir") And (ctrl.Name <> "cmdCancelar") And (ctrl.Name <> "cmdSimular") And (ctrl.Name <> "cmdModoNominal") Then
            Select Case TypeName(ctrl)
                Case Is = "ComboBox", "CommandButton", "CheckBox", "DataCombo"
                        ctrl.Enabled = pModo
            End Select
        End If
    Next
    If PermisosTotales.HabilitaBotonesConsola = 1 Then
        'chkInformacion.Enabled = True
        cmdMantenimiento.Enabled = True
        cmdCarteles.Enabled = True
        cmdAbstenciones.Enabled = True
        cmdReconsiderar.Enabled = True
        'cmdControlSistema.Enabled = True
        dcAbstencion.Enabled = False
    ElseIf PermisosTotales.HabilitaBotonesConsola = 0 Then
        'chkInformacion.Enabled = False
        cmdMantenimiento.Enabled = False
        cmdCarteles.Enabled = False
        cmdAbstenciones.Enabled = False
        cmdReconsiderar.Enabled = False
        'cmdControlSistema.Enabled = False
        dcAbstencion.Enabled = False
        cmdCancelar.Enabled = False
    End If
End Property
Private Property Let consolaHabilitada(pModo As Boolean)
    Me.Enabled = pModo
End Property
Private Sub Levanta_Sqv()
    If True Then
        Call IniciarSQVporBandera
    Else
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
    End If
    Screen.MousePointer = 0
End Sub
Private Sub Levanta_Banca()
    If True Then
    Else
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
    End If
    Screen.MousePointer = 0
End Sub

Private Sub tmAScreen_Timer()
If AutoCaptura = True Then
    'tmAutoCaptura.Enabled = True
    If frmConsolaOperacion.dcTipoOperacion.Tag <> "quorum" Then
        If cmdVotacion.Caption = "Votac&ión" And cmdVotacion.Enabled = False And nImpresion = False Then
            If tmAutoCaptura.Enabled = False Then
                tmAutoCaptura.Enabled = True
                PreHTML = "<html>" & vbCrLf
                PreHTML = PreHTML & "<title>" & UCase(frmConsolaOperacion.lblCodigoSesion.Tag) & "_" & Format(frmConsolaOperacion.lblNroSesion.Caption, "00") & _
"_" & Format(frmConsolaOperacion.lblActa.Caption, "00") & "_R" & Format(frmConsolaOperacion.txtNumeroReunion.Text, "00") & "</title>" & vbCrLf
                PreHTML = PreHTML & "<body>" & vbCrLf
                PostHTML = "</body>" & vbCrLf & "</html>"
            End If
        ElseIf (cmdVotacion.Caption = "&Inicializar" And cmdVotacion.Enabled = True) Or nImpresion = True Then
            If tmAutoCaptura.Enabled = True Then
                tmAutoCaptura.Enabled = False
                html = ""
            End If
        End If
    End If
End If
End Sub
Private Sub tmAutoCaptura_Timer()
Dim CarpetaBuff As String
CarpetaBuff = UCase(frmConsolaOperacion.lblCodigoSesion.Tag) & "_" & Format(frmConsolaOperacion.lblNroSesion.Caption, "00") & _
"_" & Format(frmConsolaOperacion.lblActa.Caption, "00") & "_R" & Format(frmConsolaOperacion.txtNumeroReunion.Text, "00")
If cmdVotacion.Caption = "Votac&ión" And cmdVotacion.Enabled = False Then
    CreaSeparador (CarpetaBuff)
    'Si esta en proceso de votacion
    GuardarScreen (CarpetaBuff)
End If
End Sub
Private Sub GuardarScreen(dir As String)
Dim Prefijo As String
Dim Archivo As String
On Error GoTo cErr
Prefijo = Format(Now(), "hhmmss")
Archivo = App.Path & "\capturas\" & dir & "\" & Prefijo & ".bmp"
Call SacarScreen
SavePicture Clipboard.GetData(vbCFBitmap), Archivo
html = html & "<img src=" & """" & Prefijo & ".bmp" & """" & "></img><br>"
Open App.Path & "\capturas\" & dir & "\index.html" For Output As #1
Print #1, PreHTML & html & PostHTML
Close #1
Exit Sub
cErr:
Exit Sub
End Sub
Private Sub CreaSeparador(bc As String)
On Error Resume Next
MkDir App.Path & "\capturas"
MkDir App.Path & "\capturas\" & bc
End Sub
Private Sub tmCheckTemp_Timer()
Dim Nombre_Archivo As String
Nombre_Archivo = Now
Nombre_Archivo = Replace(Nombre_Archivo, " ", "_")
Nombre_Archivo = Replace(Nombre_Archivo, "/", ".")
Nombre_Archivo = Replace(Nombre_Archivo, ":", "_")
End Sub

Private Sub tmPresidente_Timer()
Datos.GrabarMensaje "cambio?modovotapresidente", " ", , True
tmPresidente.Enabled = False
End Sub

Private Sub tmScreens_Timer()
On Error GoTo pTE
If (GetForegroundWindow = Me.hwnd) Then
    Call SacarScreen
    SavePicture Clipboard.GetData(vbCFBitmap), Carpeta_Screen & "temp.bmp"
    tmCheckTemp.Enabled = True
End If
Exit Sub
pTE:
End Sub

Private Sub txtNumeroReunion_GotFocus()
CambiandoReunion = True
End Sub
Private Sub txtNumeroReunion_LostFocus()
CambiandoReunion = False
End Sub

Private Sub txtTitulo_GotFocus()
txtTituloTemp.Visible = True
txtTituloTemp.Text = txtTitulo.Text
txtTituloTemp.SetFocus
End Sub
Private Sub txtTituloTemp_LostFocus()
txtTituloTemp.Visible = False
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
On Error Resume Next
    'Dim xBanca As Integer
    'Dim xBancaAux As Integer
    'Dim xPosicionBanca As Integer
    'Dim xCorrimientoVertical As Integer
    'Dim radio As Double
    'Dim pi As Double
    'Dim xCentroIzquierdo As Double
    'Dim xCentroDerecho As Double
    'Dim yCentro As Double
    'Dim xCentro As Double
    'Dim xObjeto As Double
    'Dim yObjeto As Double
    'Dim offset As Double
    'Dim Step As Double
    'Dim Escala As Double
    'Dim EscalaX As Double
    'Dim xUltimaBanca As Integer
    'Escala = 0.9
    'EscalaX = 1.35
    'xCentroIzquierdo = ((Me.Width - ML(0).Width - MR1.Left) / 2) - 215
    'xCentroDerecho = ((Me.Width - ML(0).Width - MR1.Left) / 2) + 215
    'yCentro = ML(0).Height - 2600     '6480
    'pi = 4 * Atn(1)   ' Calculo el valor de pi
    'Me.Cls
    'Me.Refresh
    'xUltimaBanca = MAX_BANCA
    'For xBanca = 0 To xUltimaBanca
    '    shpBanca(xBanca).Visible = True
    '    ctrBanca(xBanca).Visible = True
    '    shpBanca(xBanca).Width = 630
    '    shpBanca(xBanca).Height = 630
    '    shpBanca(xBanca).FillStyle = 0
    '    shpBanca(xBanca).BorderWidth = 2
    '    xPosicionBanca = xBanca
    '    xCorrimientoVertical = 0
    '
    '    If xBanca >= 25 And xBanca <= 46 Then 'arco externo
    '        radio = (Me.Width / 2) * 0.46 * Escala
    '        Step = pi / (43 - 28)
    '        xCentro = IIf(xBanca < 36, xCentroIzquierdo, xCentroDerecho)
    '        offset = 28
    '    ElseIf xBanca >= 7 And xBanca <= 24 Then
    '        radio = (Me.Width / 2) * 0.33 * Escala
    '        Step = pi / (21 - 10)
    '        xCentro = IIf(xBanca < 16, xCentroIzquierdo, xCentroDerecho)
    '        offset = 10
    '    ElseIf xBanca >= 1 And xBanca <= 6 Then 'arco interno
    '        radio = (Me.Width / 2) * 0.17 * Escala
    '        Step = pi / 5
    '        xCentro = IIf(xBanca < 4, xCentroIzquierdo, xCentroDerecho)
    '        offset = 1
    '    End If
    '
    '    If xBanca >= 7 And xBanca <= 9 Then
    '        xPosicionBanca = 10
    '        xCorrimientoVertical = 10 - xBanca
    '    ElseIf xBanca >= 22 And xBanca <= 24 Then
    '        xPosicionBanca = 21
    '        xCorrimientoVertical = -21 + xBanca
    '    ElseIf xBanca >= 25 And xBanca <= 27 Then
   ''         xPosicionBanca = 28
   ''         xCorrimientoVertical = 28 - xBanca
   ''     ElseIf xBanca >= 44 And xBanca <= 46 Then
   ''         xPosicionBanca = 43
   ''         xCorrimientoVertical = -43 + xBanca
   ''     End If
   ''
   ''
   ''
   ''     If xBanca = 0 Then
   '         xObjeto = (Me.Width - ML(0).Width - MR1.Left) / 2
   '         yObjeto = yCentro + 3 * 700
   '     Else
   '         xObjeto = xCentro - Cos(Step * (xPosicionBanca - offset)) * radio * EscalaX
   '         yObjeto = yCentro - Sin(Step * (xPosicionBanca - offset)) * radio + xCorrimientoVertical * 700
   '     End If
   '     shpBanca(xBanca).Left = xObjeto - (shpBanca(xBanca).Width / 2)
   '     shpBanca(xBanca).Top = yObjeto - (shpBanca(xBanca).Height / 2)
   '     shpBanca(xBanca).FillColor = vbWhite
   '     ctrBanca(xBanca).Caption = xBanca
   '     ctrBanca(xBanca).AutoSize = True
   '     ctrBanca(xBanca).Alignment = 2
   '     ctrBanca(xBanca).FontSize = 19
   '     ctrBanca(xBanca).AutoSize = False
   '     ctrBanca(xBanca).Width = 630
   '     ctrBanca(xBanca).Left = xObjeto - (ctrBanca(xBanca).Width / 2)
   '     ctrBanca(xBanca).Top = yObjeto - (ctrBanca(xBanca).Height / 2)
   '     shpBanca(xBanca).ZOrder 1
   '
    'Next
    ''cambio de lugar
    'For xBanca = 1 To 22
    '    Call IntercambiarBancas(xBanca, xBanca + 25 - 1)
    'Next xBanca
    'For xBanca = 41 To 46
    '    Call IntercambiarBancas(xBanca, xBanca - (41 - 25))
    'Next xBanca
    'xBancaAux = 40
    'For xBanca = 31 To 35
    '    Call IntercambiarBancas(xBanca, xBancaAux)
    '    xBancaAux = xBancaAux - 1
    'Next xBanca
    'xBancaAux = 30
    'For xBanca = 25 To 27
    '    Call IntercambiarBancas(xBanca, xBancaAux)
    '    xBancaAux = xBancaAux - 1
    'Next xBanca
    'Call IntercambiarBancas(23, 24)
    'For xBanca = MAX_BANCA + 1 To 70
    '    shpBanca(xBanca).Visible = False
    '    ctrBanca(xBanca).Visible = False
    'Next
    
    
    'Me.ZOrder 1
    
    ' ap 080905 para demo: ajusta posicion arco de bancas
    'NUEVA VERSION
    Dim i As Integer
    Dim B_Top As Integer
    Dim B_Left As Integer
    Dim Offset As Integer
    Dim TLeft As Integer
    Dim TTop As Integer
    Dim r As Integer
    Dim OffTop As Integer
    Dim TopOffset As Integer
    Dim LefTOffset As Integer
    TopOffset = 550
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    SetearRs "SELECT b_top,b_left FROM posiciones WHERE destino = 'C' AND banca = 0", RsTemp
    With shpBanca(0)
        .Visible = True
        .Left = RsTemp.Fields("b_left")
        .Top = RsTemp.Fields("b_top") + TopOffset
        .BorderStyle = vbSolid
        .BorderColor = &HFFFFFF
        .FillColor = &HFFFFFF
        .FillStyle = vbSolid
        .Height = 375
        .Width = 375
    End With
    RsTemp.Close
    Set RsTemp = Nothing
    Set RsTemp = New ADODB.Recordset
    SetearRs "SELECT banca,b_top,b_left FROM posiciones WHERE banca > 0 AND destino = 'C' ORDER BY banca", RsTemp
    While Not RsTemp.EOF
        Load shpBanca(RsTemp.Fields("banca"))
        With shpBanca(RsTemp.Fields("banca"))
            .Visible = True
            .Left = RsTemp.Fields("b_left")
            .Top = RsTemp.Fields("b_top") + TopOffset
            .BorderStyle = vbSolid
            .BorderColor = &HFFFFFF
            .FillColor = &HFFFFFF
            .FillStyle = vbSolid
            .Height = 375
            .Width = 375
        End With
        RsTemp.MoveNext
    Wend
    RsTemp.Close
    Set RsTemp = Nothing
    ctrBanca(0).Left = shpBanca(0).Left + 250
    ctrBanca(0).Top = shpBanca(0).Top + 50
    ctrBanca(0).Visible = True
    With ctrBanca(0)
        .FontName = "Arial"
        .FontSize = 10
        .FontBold = True
        .Alignment = vbRightJustify
        .Caption = "  " & .Caption & "  "
    End With
    For i = 1 To shpBanca.UBound
    Load ctrBanca(i)
    shpBanca(i).Height = 380
    shpBanca(i).Width = 430
    With ctrBanca(i)
        .FontName = "Arial"
        .FontSize = 10
        .FontBold = True
        .Alignment = vbRightJustify
        .Refresh
        .Caption = i
        .ZOrder 0
        If i <= 9 Then
            .Left = shpBanca(i).Left + 270
            .Top = shpBanca(i).Top + 60
            .Caption = "  " & .Caption & "  "
        ElseIf i > 9 And i <= 99 Then
            .Left = shpBanca(i).Left + 140
            .Top = shpBanca(i).Top + 40
            .Caption = " " & .Caption & " "
        Else
            .Left = shpBanca(i).Left + 50
            .Top = shpBanca(i).Top + 40
        End If
        .Visible = True
    End With
Next i
    MR1.Align = 0
    MR1.Left = 80
    MR1.Height = 15360
    MR1.Width = 19200
    MR1.Picture = LoadPicture(App.Path & "\IMAGENES\black.bmp")
    MR2.Width = 17000
'    Set MR2.Picture = Nothing 'manzana
'    Set MR1.Picture = Nothing 'manzana
    ' fin ap 080905
    Me.Refresh
End Sub
Private Sub IntercambiarBancas(nBanca1 As Integer, nBanca2 As Integer)
    Dim nshpLeft As Integer
    Dim nctrLeft As Integer
    Dim nshpTop As Integer
    Dim nctrTop As Integer
    
    nshpLeft = shpBanca(nBanca1).Left
    nctrLeft = ctrBanca(nBanca1).Left
    nshpTop = shpBanca(nBanca1).Top
    nctrTop = ctrBanca(nBanca1).Top

    shpBanca(nBanca1).Left = shpBanca(nBanca2).Left
    shpBanca(nBanca1).Top = shpBanca(nBanca2).Top
    ctrBanca(nBanca1).Left = ctrBanca(nBanca2).Left
    ctrBanca(nBanca1).Top = ctrBanca(nBanca2).Top

    shpBanca(nBanca2).Left = nshpLeft
    shpBanca(nBanca2).Top = nshpTop
    ctrBanca(nBanca2).Left = nctrLeft
    ctrBanca(nBanca2).Top = nctrTop
    
End Sub
Private Sub ArmarBancasCartelCBA2003() 'Obsoleto
   
    Dim xBanca As Integer
    Dim radio As Double
    Dim pi As Double
    Dim xCentroIzquierdo As Double
    Dim xCentroDerecho As Double
    Dim yCentro As Double
    Dim xCentro As Double
    Dim xObjeto As Double
    Dim yObjeto As Double
    Dim Offset As Double
    Dim Step As Double
    Dim Escala As Double
    Dim xUltimaBanca As Integer
    Escala = 0.9
    xCentroIzquierdo = ((Me.Width - ML(0).Width - MR1.Left) / 2) - 215
    xCentroDerecho = ((Me.Width - ML(0).Width - MR1.Left) / 2) + 215
    yCentro = ML(0).Height - 470     '6480
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
            Offset = 1
        ElseIf xBanca >= 23 And xBanca <= 42 Then
            radio = (Me.Width / 2) * 0.638 * Escala
            Step = pi / 19
            xCentro = IIf(xBanca < 33, xCentroIzquierdo, xCentroDerecho)
            Offset = 23
        ElseIf xBanca >= 43 And xBanca <= 58 Then
            radio = (Me.Width / 2) * 0.5 * Escala
            Step = pi / 15
            xCentro = IIf(xBanca < 51, xCentroIzquierdo, xCentroDerecho)
            Offset = 43
        ElseIf xBanca >= 59 And xBanca <= 70 Then
            radio = (Me.Width / 2) * 0.365 * Escala
            Step = pi / 11
            xCentro = IIf(xBanca < 65, xCentroIzquierdo, xCentroDerecho)
            Offset = 59
        End If
        
        If xBanca = 0 Then
            xObjeto = (Me.Width - ML(0).Width - MR1.Left) / 2
            yObjeto = yCentro
        Else
            xObjeto = xCentro - Cos(Step * (xBanca - Offset)) * radio
            yObjeto = yCentro - Sin(Step * (xBanca - Offset)) * radio
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
    
    ' ap 080905 para demo: ajusta posicion arco de bancas
    MR1.Align = 0
    MR1.Left = 990
    ' fin ap 080905
    
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
    'OLIVA     'CONTRASTE BLANCO
    Color = "&HFFFFFF":      clave = "8":   mColoresFuentes.Add clave, Color
    'NEGRO     'CONTRASTE BLANCO
    Color = "&HFFFFFF":      clave = "9":   mColoresFuentes.Add clave, Color
    'MARRON     CONTASTE BLANCO
    Color = "&HFFFFFF":      clave = "10":   mColoresFuentes.Add clave, Color
    Color = "&HFFFFFF":        clave = "11":   mColoresFuentes.Add clave, Color
End Sub

Private Sub cargarColores()
    Dim Color As String
    Dim clave As String
    
    ' Limpio el diccionario
    mColores.RemoveAll
    'cargo el diccionario de manera estática
    'GRIS
    Color = "&H808080"
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
    Color = "&H00C0C000"
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
    'OLIVA para INDICAR QUE HA REALIZADO EL VOTO
    Color = "&H00004040"
    clave = "8"
    mColores.Add clave, Color
    'AZUL: banca con error de switch (IOC)
    Color = "&HC00000"
    clave = "9"
    mColores.Add clave, Color
    'MARRON: banca con problema técnico
    Color = "&H404080"
    clave = "10"
    mColores.Add clave, Color
    'Marron Claro: Banca con TIDINV
    Color = "&H80FF"
    clave = "11"
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
        bPermitirIdentificar = (LCase(dcTipoOperacion.BoundText) = "votnom") Or _
                               (LCase(dcTipoOperacion.BoundText) = "paslis") Or _
                               mModo_Ident_Nom
        ' Chequear el vector identificacion
        If ((dcTipoOperacion.BoundText = "votnum") Or (dcTipoOperacion.BoundText = "votnom")) And _
            (mEstadoVotacion = "votando") Or (mEstadoVotacion = "larga") Or mEstadoVotacion = "empate" Then
            If Trim(mVectorPresencia(Index)) = "1" Or (Index = 0) Then
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


Private Sub TraerInfoLegislador(pIdLegislador As String, legislador As String, Optional bloque As String, Optional Departamento As String, Optional Agrupacion As String)
    If IsNull(pIdLegislador) = False Then
        Dim rstAux As New ADODB.Recordset
        SetearRs "SELECT * FROM Legisladores WHERE id='" & pIdLegislador & "'", rstAux
        If rstAux.EOF = False Then
            If (IsNull(rstAux!Apellido) = False) And (IsNull(rstAux!Nombre) = False) Then
                legislador = Trim(rstAux!Apellido) & ", " & Trim(rstAux!Nombre)
            End If
            If IsNull(rstAux!grupo_politico) = False Then
                Agrupacion = Trim(rstAux!grupo_politico)
            End If
            If IsNull(rstAux!bloque_politico) = False Then
                bloque = Trim(rstAux!bloque_politico)
            End If
            If IsNull(rstAux!Departamento) = False Then
                Departamento = Trim(rstAux!Departamento)
            End If
        Else
            legislador = "No Identificado"
        End If
    End If
End Sub

Private Sub CargarVectorLegisladoresPorBloque()
    Dim sql As String
    Dim rstB As New ADODB.Recordset
    Dim rstAux As New ADODB.Recordset
    Dim BCont As Integer
    Dim b As Integer
    Dim c As Integer
    Dim idxB As Integer
    Dim bolPresidente As Boolean
    sql = "SELECT Count(legisladores_activos.ID) AS TotalLegisladores, Legisladores.bloque_politico AS BLOQUE "
    sql = sql & "FROM legisladores_activos "
    sql = sql & "INNER JOIN Legisladores ON legisladores_activos.ID = Legisladores.id "
    sql = sql & "GROUP BY Legisladores.bloque_politico "
    SetearRs sql, rstB
    If Error_Carga = True Then
        Unload Me
        Exit Sub
    End If
    If Not rstB.EOF Then
        rstB.MoveLast
        BCont = rstB.RecordCount
        ReDim datBloque(BCont)
        BCont = 0
        rstB.MoveFirst
        c = 0
        bolPresidente = False
        Do Until rstB.EOF
            
            sql = "SELECT legisladores_activos.ID AS ID, legisladores_activos.DESKID AS BANCA, Legisladores.apellido AS APELLIDO, Legisladores.nombre AS NOMBRE, Legisladores.bloque_politico AS BLOQUE "
            sql = sql & "FROM legisladores_activos "
            sql = sql & "INNER JOIN Legisladores ON legisladores_activos.ID = Legisladores.id "
            sql = sql & "WHERE Legisladores.bloque_politico = '" & rstB("Bloque").Value & "' "
            sql = sql & "ORDER BY Legisladores.bloque_politico , Legisladores.apellido, Legisladores.nombre  "
            SetearRs sql, rstAux
            Do Until rstAux.EOF
                If Not bolPresidente Then
                    bolPresidente = (rstAux!Banca = "0")
                End If
                datLista(c).BloqueIndex = BCont
                datLista(c).BloqueNombre = rstB("Bloque").Value & ""
                datLista(c).LegisladorDefecto = rstAux("ID").Value & ""
                datLista(c).LegisladorDefectoNombre = UCase(rstAux("APELLIDO").Value) & ", " & rstAux("NOMBRE").Value
                datLista(c).BancaDefecto = rstAux!Banca
                rstAux.MoveNext
                c = c + 1
            Loop
            datBloque(BCont).Nombre = rstB("Bloque").Value
            datBloque(BCont).TotalLegisladores = rstB("TotalLegisladores").Value + IIf(bolPresidente, -1, 0)
            bolPresidente = False
            BCont = BCont + 1
            rstB.MoveNext
        Loop
    Else
        MsgBox "No hay legisladores activos", vbCritical + vbOKOnly, "Consola de Legisladores"
        
    End If
End Sub
Private Function SetLabelInfo(m As String, BancaInd As Integer)
'Para actualizar ventana flotante
Dim AMostrar As String
Dim Prox_Left As Integer
Dim i As Integer
If (IMAGENES_RAPIDAS_HABILITADAS = True) Then
    pctFotoRapida.Height = 1800
End If
lblNombreRapido.Alignment = vbCenter
If (ctrBanca(BancaInd).Top < 4115) Then
    lblNombreRapido.Top = ctrBanca(BancaInd).Top + 580
Else
    lblNombreRapido.Top = ctrBanca(BancaInd).Top - 600
End If
Prox_Left = ctrBanca(BancaInd).Left
lblNombreRapido.Width = 15
AMostrar = "Info: " + m
lblNombreRapido.Caption = AMostrar
For i = 1 To Len(AMostrar) 'Acomodo el ancho del label
    lblNombreRapido.Width = lblNombreRapido.Width + 120
Next i
pctFotoRapida.Width = lblNombreRapido.Width
i = (lblNombreRapido.Width / 2) - 400 '400 es el offset para que quede en el medio
If (IMAGENES_RAPIDAS_HABILITADAS = True) Then
    If (ctrBanca(BancaInd).Top < 4115) Then
        pctFotoRapida.Top = lblNombreRapido.Top + lblNombreRapido.Height - 20
    Else
        pctFotoRapida.Top = ctrBanca(BancaInd).Top - (pctFotoRapida.Height + lblNombreRapido.Height + 210)
    End If
End If
lblNombreRapido.Left = Prox_Left - i
If (IMAGENES_RAPIDAS_HABILITADAS = True) Then
    pctFotoRapida.Left = lblNombreRapido.Left
    pctFotoRapida.Visible = True
End If
End Function
Private Sub Crea1(img As Object)
Dim i As Integer
On Error Resume Next
ctrBanca(0).Left = Left_Inicial + 1300
ctrBanca(0).Top = Top_Inicial + 300
ctrBanca(0).Caption = "0"
ctrBanca(0).Font = "Arial"
ctrBanca(0).FontSize = 10
ctrBanca(0).FontBold = True
ctrBanca(0).Visible = True
For i = 1 To 5
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial + 100
            .Top = Top_Inicial + 250
            OffLeft = 0
        ElseIf i = 3 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + AumentoTop / 2 + OffLeft - 100
            .Top = img(conta - 1).Top - AumentoTop
        ElseIf i = 4 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + AumentoTop / 2 + OffLeft
            .Top = img(conta - 1).Top - AumentoTop
        ElseIf i = 5 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + AumentoTop / 1.5 + OffLeft
            .Top = img(conta - 1).Top - AumentoTop + 30
        Else
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + OffLeft
            .Top = img(conta - 1).Top - AumentoTop
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea2(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 4
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial + 925
            .Top = Top_Inicial - 1350
            OffLeft = 0
        ElseIf i = 2 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 460 + Exp
            .Top = img(conta - 1).Top - 100
        ElseIf i = 3 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 500 + Exp
            .Top = img(conta - 1).Top
        Else
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 470 + Exp
            .Top = img(conta - 3).Top
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea3(img As Object)
Dim i As Integer
On Error Resume Next
Dim xT As Integer
xT = conta + 5
conta = conta + 6
For i = 1 To 5
    conta = conta - 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial + 2900
            .Top = Top_Inicial + 250
            OffLeft = 0
        ElseIf i = 3 Then
            OffLeft = OffLeft + 0
            .Left = img(conta + 1).Left - AumentoTop / 2 + OffLeft + 100
            .Top = img(conta + 1).Top - AumentoTop
        ElseIf i = 4 Then
            OffLeft = OffLeft + 0
            .Left = img(conta + 1).Left - AumentoTop / 2 - OffLeft
            .Top = img(conta + 1).Top - AumentoTop
        ElseIf i = 5 Then
            OffLeft = OffLeft + 0
            .Left = img(conta + 1).Left - AumentoTop / 1.5 - OffLeft
            .Top = img(conta + 1).Top - AumentoTop + 30
        Else
            OffLeft = OffLeft + 0
            .Left = img(conta + 1).Left - OffLeft
            .Top = img(conta + 1).Top - AumentoTop
        End If
        .Visible = True
    End With
Next i
conta = xT
End Sub
Private Sub Crea4(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 6
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial - 550
            .Top = Top_Inicial + 300
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 30
            .Left = img(conta - 1).Left + OffLeft
            .Top = img(conta - 1).Top - AumentoTop - 20
        ElseIf i = 4 Then
            OffLeft = OffLeft + 30
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop - 10
        ElseIf i = 6 Then
            OffLeft = OffLeft + 30
            .Left = img(conta - 1).Left + OffLeft + 100
            .Top = img(conta - 1).Top - AumentoTop
        Else
            OffLeft = OffLeft + 30
            .Left = img(conta - 1).Left + OffLeft + 70
            .Top = img(conta - 1).Top - AumentoTop + 10
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea5(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 6
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial + 550
            .Top = Top_Inicial - 1850
            OffLeft = 0
        ElseIf i = 2 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 460 + Exp
            .Top = img(conta - 1).Top - 200
        ElseIf i = 3 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 500 + Exp
            .Top = img(conta - 1).Top - 70
        ElseIf i = 4 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 500 + Exp
            .Top = img(conta - 1).Top
        ElseIf i = 5 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 500 + Exp
            .Top = img(conta - 3).Top
        Else
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 470 + Exp
            .Top = img(conta - 5).Top
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea6(img As Object)
Dim i As Integer
On Error Resume Next
Dim xT As Integer
xT = conta + 6
conta = conta + 7
For i = 1 To 6
    conta = conta - 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial + 3550
            .Top = Top_Inicial + 300
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 30
            .Left = img(conta + 1).Left - OffLeft
            .Top = img(conta + 1).Top - AumentoTop - 20
        ElseIf i = 4 Then
            OffLeft = OffLeft + 30
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop - 10
        ElseIf i = 6 Then
            OffLeft = OffLeft + 30
            .Left = img(conta + 1).Left - OffLeft - 100
            .Top = img(conta + 1).Top - AumentoTop
        Else
            OffLeft = OffLeft + 30
            .Left = img(conta + 1).Left - OffLeft - 70
            .Top = img(conta + 1).Top - AumentoTop + 10
        End If
        .Visible = True
    End With
Next i
conta = xT
End Sub
Private Sub Crea7(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 8
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial - 1250
            .Top = Top_Inicial + 350
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 30
            .Left = img(conta - 1).Left + OffLeft
            .Top = img(conta - 1).Top - AumentoTop - 30
        ElseIf i = 4 Then
            OffLeft = OffLeft + 10
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop - 10
        ElseIf i = 5 Then
            OffLeft = OffLeft + 10
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop
        ElseIf i = 6 Then
            OffLeft = OffLeft + 10
            .Left = img(conta - 1).Left + OffLeft + 100
            .Top = img(conta - 1).Top - AumentoTop + 20
        ElseIf i = 7 Then
            OffLeft = OffLeft + 30
            .Left = img(conta - 1).Left + OffLeft + 70
            .Top = img(conta - 1).Top - AumentoTop
        Else
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft + 120
            .Top = img(conta - 1).Top - AumentoTop + 20
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea8(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 8
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial + 150
            .Top = Top_Inicial - 2300
            OffLeft = 0
        ElseIf i = 2 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 360 + Exp
            .Top = img(conta - 1).Top - 250
        ElseIf i = 3 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 430 + Exp
            .Top = img(conta - 1).Top - 180
        ElseIf i = 4 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 500 + Exp
            .Top = img(conta - 1).Top - 70
        ElseIf i = 5 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 500 + Exp
            .Top = img(conta - 1).Top
        ElseIf i = 6 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 470 + Exp
            .Top = img(conta - 3).Top
        ElseIf i = 7 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 420 + Exp
            .Top = img(conta - 5).Top
        Else
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 380 + Exp
            .Top = img(conta - 7).Top
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea9(img As Object)
Dim i As Integer
On Error Resume Next
Dim xT As Integer
xT = conta + 8
conta = conta + 9
For i = 1 To 8
    conta = conta - 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial + 4250
            .Top = Top_Inicial + 350
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 30
            .Left = img(conta + 1).Left - OffLeft
            .Top = img(conta + 1).Top - AumentoTop - 30
        ElseIf i = 4 Then
            OffLeft = OffLeft + 10
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop - 10
        ElseIf i = 5 Then
            OffLeft = OffLeft + 10
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop
        ElseIf i = 6 Then
            OffLeft = OffLeft + 10
            .Left = img(conta + 1).Left - OffLeft - 100
            .Top = img(conta + 1).Top - AumentoTop + 20
        ElseIf i = 7 Then
            OffLeft = OffLeft + 30
            .Left = img(conta + 1).Left - OffLeft - 70
            .Top = img(conta + 1).Top - AumentoTop
        Else
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft - 120
            .Top = img(conta + 1).Top - AumentoTop + 20
        End If
        .Visible = True
    End With
Next i
conta = xT
End Sub
Private Sub Crea10(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 10
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial - 1950
            .Top = Top_Inicial + 400
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 30
            .Left = img(conta - 1).Left + OffLeft
            .Top = img(conta - 1).Top - AumentoTop - 30
        ElseIf i = 4 Then
            OffLeft = OffLeft - 10
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop - 10
        ElseIf i = 5 Then
            OffLeft = OffLeft - 10
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop
        ElseIf i = 6 Then
            OffLeft = OffLeft + 10
            .Left = img(conta - 1).Left + OffLeft + 100
            .Top = img(conta - 1).Top - AumentoTop + 20
        ElseIf i = 7 Then
            OffLeft = OffLeft + 30
            .Left = img(conta - 1).Left + OffLeft + 70
            .Top = img(conta - 1).Top - AumentoTop
        ElseIf i = 10 Then
            OffLeft = OffLeft + 150
            .Left = img(conta - 1).Left + OffLeft + 20
            .Top = img(conta - 1).Top - AumentoTop + 50
        Else
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft + 120
            .Top = img(conta - 1).Top - AumentoTop + 20
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea11(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 10
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial - 300
            .Top = Top_Inicial - 2700
            OffLeft = 0
        ElseIf i = 2 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 460 + Exp
            .Top = img(conta - 1).Top - 250
        ElseIf i = 3 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 500 + Exp
            .Top = img(conta - 1).Top - 200
        ElseIf i = 4 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 500 + Exp
            .Top = img(conta - 1).Top - 140
        ElseIf i = 5 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 500 + Exp
            .Top = img(conta - 1).Top - 50
        ElseIf i = 6 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 550 + Exp
            .Top = img(conta - 1).Top
        ElseIf i = 7 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 500 + Exp
            .Top = img(conta - 3).Top
        ElseIf i = 8 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 500 + Exp
            .Top = img(conta - 5).Top
        ElseIf i = 9 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 470 + Exp
            .Top = img(conta - 7).Top
        Else
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 470 + Exp
            .Top = img(conta - 9).Top
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea12(img As Object)
Dim i As Integer
On Error Resume Next
Dim TF As Integer
TF = conta + 10
conta = conta + 11
For i = 1 To 10
    conta = conta - 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial + 4950
            .Top = Top_Inicial + 400
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 30
            .Left = img(conta + 1).Left - OffLeft
            .Top = img(conta + 1).Top - AumentoTop - 30
        ElseIf i = 4 Then
            OffLeft = OffLeft - 10
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop - 10
        ElseIf i = 5 Then
            OffLeft = OffLeft - 10
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop
        ElseIf i = 6 Then
            OffLeft = OffLeft + 10
            .Left = img(conta + 1).Left - OffLeft - 100
            .Top = img(conta + 1).Top - AumentoTop + 20
        ElseIf i = 7 Then
            OffLeft = OffLeft + 30
            .Left = img(conta + 1).Left - OffLeft - 70
            .Top = img(conta + 1).Top - AumentoTop
        ElseIf i = 10 Then
            OffLeft = OffLeft + 150
            .Left = img(conta + 1).Left - OffLeft - 20
            .Top = img(conta + 1).Top - AumentoTop + 50
        Else
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft - 120
            .Top = img(conta + 1).Top - AumentoTop + 20
        End If
        .Visible = True
    End With
Next i
conta = TF
End Sub
Private Sub Crea13(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 11
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial - 2600
            .Top = Top_Inicial + 200
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 20
            .Left = img(conta - 1).Left + OffLeft
            .Top = img(conta - 1).Top - AumentoTop - 30
        ElseIf i = 4 Then
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop - 10
        ElseIf i = 5 Then
            OffLeft = OffLeft - 10
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop
        ElseIf i = 6 Then
            OffLeft = OffLeft + 10
            .Left = img(conta - 1).Left + OffLeft + 100
            .Top = img(conta - 1).Top - AumentoTop + 20
        ElseIf i = 7 Then
            OffLeft = OffLeft + 50
            .Left = img(conta - 1).Left + OffLeft + 70
            .Top = img(conta - 1).Top - AumentoTop
        ElseIf i = 8 Or i = 9 Then
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft + 120
            .Top = img(conta - 1).Top - AumentoTop + 20
        ElseIf i = 10 Then
            OffLeft = OffLeft + 150
            .Left = img(conta - 1).Left + OffLeft + 20
            .Top = img(conta - 1).Top - AumentoTop + 50
        Else
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft - 20
            .Top = img(conta - 1).Top - AumentoTop + 40
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea14(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 14
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial - 770
            .Top = Top_Inicial - 3250
            OffLeft = 0
        ElseIf i = 2 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 340 + Exp
            .Top = img(conta - 1).Top - 250
        ElseIf i = 3 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 400 + Exp
            .Top = img(conta - 1).Top - 200
        ElseIf i = 4 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 400 + Exp
            .Top = img(conta - 1).Top - 180
        ElseIf i = 5 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 150
        ElseIf i = 6 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 80
        ElseIf i = 7 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 40
        ElseIf i = 8 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 470 + Exp
            .Top = img(conta - 1).Top
        ElseIf i = 9 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 3).Top
        ElseIf i = 10 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 5).Top
        ElseIf i = 11 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 7).Top
        ElseIf i = 12 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 400 + Exp
            .Top = img(conta - 9).Top
        ElseIf i = 13 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 380 + Exp
            .Top = img(conta - 11).Top
        Else
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 350 + Exp
            .Top = img(conta - 13).Top
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea15(img As Object)
Dim i As Integer
On Error Resume Next
Dim Tx As Integer
Tx = conta + 11
conta = conta + 12
For i = 1 To 11
    conta = conta - 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial + 5600
            .Top = Top_Inicial + 200
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 20
            .Left = img(conta + 1).Left - OffLeft
            .Top = img(conta + 1).Top - AumentoTop - 30
        ElseIf i = 4 Then
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop - 10
        ElseIf i = 5 Then
            OffLeft = OffLeft - 10
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop
        ElseIf i = 6 Then
            OffLeft = OffLeft + 10
            .Left = img(conta + 1).Left - OffLeft - 100
            .Top = img(conta + 1).Top - AumentoTop + 20
        ElseIf i = 7 Then
            OffLeft = OffLeft + 50
            .Left = img(conta + 1).Left - OffLeft - 70
            .Top = img(conta + 1).Top - AumentoTop
        ElseIf i = 8 Or i = 9 Then
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft - 120
            .Top = img(conta + 1).Top - AumentoTop + 20
        ElseIf i = 10 Then
            OffLeft = OffLeft + 150
            .Left = img(conta + 1).Left - OffLeft - 20
            .Top = img(conta + 1).Top - AumentoTop + 50
        Else
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft + 20
            .Top = img(conta + 1).Top - AumentoTop + 40
        End If
        .Visible = True
    End With
Next i
conta = Tx
End Sub
Private Sub Crea16(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 13
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial - 3600
            .Top = Top_Inicial + 380
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 20
            .Left = img(conta - 1).Left + OffLeft
            .Top = img(conta - 1).Top - AumentoTop - 30
        ElseIf i = 4 Then
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop - 10
        ElseIf i = 5 Then
            OffLeft = OffLeft - 10
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop
        ElseIf i = 6 Then
            OffLeft = OffLeft + 10
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop + 20
        ElseIf i = 7 Then
            OffLeft = OffLeft + 40
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop
        ElseIf i = 8 Then
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft + 80
            .Top = img(conta - 1).Top - AumentoTop + 20
        ElseIf i = 9 Then
            .Left = img(conta - 1).Left + OffLeft + 120
            .Top = img(conta - 1).Top - AumentoTop + 20
        ElseIf i = 10 Then
            OffLeft = OffLeft + 150
            .Left = img(conta - 1).Left + OffLeft - 50
            .Top = img(conta - 1).Top - AumentoTop + 40
        ElseIf i = 11 Then '133
            OffLeft = OffLeft + 150
            .Left = img(conta - 1).Left + OffLeft - 120
            .Top = img(conta - 1).Top - AumentoTop + 60
        ElseIf i = 12 Then '134
            OffLeft = OffLeft + 150
            .Left = img(conta - 1).Left + OffLeft - 250
            .Top = img(conta - 1).Top - AumentoTop + 60
        Else '135
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft - 200
            .Top = img(conta - 1).Top - AumentoTop + 100
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea17(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 14
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial - 1300
            .Top = Top_Inicial - 3900
            OffLeft = 0
        ElseIf i = 2 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 340 + Exp
            .Top = img(conta - 1).Top - 300
        ElseIf i = 3 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 400 + Exp
            .Top = img(conta - 1).Top - 200
        ElseIf i = 4 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 400 + Exp
            .Top = img(conta - 1).Top - 200
        ElseIf i = 5 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 150
        ElseIf i = 6 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 80
        ElseIf i = 7 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 40
        ElseIf i = 8 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 470 + Exp
            .Top = img(conta - 1).Top
        ElseIf i = 9 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 3).Top
        ElseIf i = 10 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 5).Top
        ElseIf i = 11 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 7).Top
        ElseIf i = 12 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 400 + Exp
            .Top = img(conta - 9).Top
        ElseIf i = 13 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 380 + Exp
            .Top = img(conta - 11).Top
        Else
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 300 + Exp
            .Top = img(conta - 13).Top
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea18(img As Object)
Dim i As Integer
On Error Resume Next
Dim Tx As Integer
Tx = conta + 13
conta = conta + 14
For i = 1 To 13
    conta = conta - 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial + 6600
            .Top = Top_Inicial + 380
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 20
            .Left = img(conta + 1).Left - OffLeft
            .Top = img(conta + 1).Top - AumentoTop - 30
        ElseIf i = 4 Then
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop - 10
        ElseIf i = 5 Then
            OffLeft = OffLeft - 10
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop
        ElseIf i = 6 Then
            OffLeft = OffLeft + 10
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop + 20
        ElseIf i = 7 Then
            OffLeft = OffLeft + 40
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop
        ElseIf i = 8 Then
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft - 80
            .Top = img(conta + 1).Top - AumentoTop + 20
        ElseIf i = 9 Then
            .Left = img(conta + 1).Left - OffLeft - 120
            .Top = img(conta + 1).Top - AumentoTop + 20
        ElseIf i = 10 Then
            OffLeft = OffLeft + 150
            .Left = img(conta + 1).Left - OffLeft + 50
            .Top = img(conta + 1).Top - AumentoTop + 40
        ElseIf i = 11 Then '133
            OffLeft = OffLeft + 150
            .Left = img(conta + 1).Left - OffLeft + 120
            .Top = img(conta + 1).Top - AumentoTop + 60
        ElseIf i = 12 Then '134
            OffLeft = OffLeft + 150
            .Left = img(conta + 1).Left - OffLeft + 250
            .Top = img(conta + 1).Top - AumentoTop + 60
        Else '135
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft + 200
            .Top = img(conta + 1).Top - AumentoTop + 100
        End If
        .Visible = True
    End With
Next i
conta = Tx
End Sub
Private Sub Crea19(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 14
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial - 4600
            .Top = Top_Inicial + 300
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 20
            .Left = img(conta - 1).Left + OffLeft
            .Top = img(conta - 1).Top - AumentoTop - 30
        ElseIf i = 4 Then
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop - 10
        ElseIf i = 5 Then
            OffLeft = OffLeft - 10
            .Left = img(conta - 1).Left + OffLeft + 10
            .Top = img(conta - 1).Top - AumentoTop
        ElseIf i = 6 Then
            OffLeft = OffLeft + 10
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop + 20
        ElseIf i = 7 Then
            OffLeft = OffLeft + 40
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop
        ElseIf i = 8 Then
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop + 20
        ElseIf i = 9 Then
            .Left = img(conta - 1).Left + OffLeft + 150
            .Top = img(conta - 1).Top - AumentoTop + 20
        ElseIf i = 10 Then
            OffLeft = OffLeft + 150
            .Left = img(conta - 1).Left + OffLeft + 40
            .Top = img(conta - 1).Top - AumentoTop + 40
        ElseIf i = 11 Then '133
            OffLeft = OffLeft + 150
            .Left = img(conta - 1).Left + OffLeft - 120
            .Top = img(conta - 1).Top - AumentoTop + 60
        ElseIf i = 12 Then '134
            OffLeft = OffLeft + 150
            .Left = img(conta - 1).Left + OffLeft - 200
            .Top = img(conta - 1).Top - AumentoTop + 60
        ElseIf i = 13 Then
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft - 250
            .Top = img(conta - 1).Top - AumentoTop + 100
        Else
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft - 200
            .Top = img(conta - 1).Top - AumentoTop + 100
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea20(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 18
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial - 1850
            .Top = Top_Inicial - 4500
            OffLeft = 0
        ElseIf i = 2 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 340 + Exp
            .Top = img(conta - 1).Top - 300
        ElseIf i = 3 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 400 + Exp
            .Top = img(conta - 1).Top - 200
        ElseIf i = 4 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 400 + Exp
            .Top = img(conta - 1).Top - 200
        ElseIf i = 5 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 150
        ElseIf i = 6 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 120
        ElseIf i = 7 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 80
        ElseIf i = 8 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 470 + Exp
            .Top = img(conta - 1).Top - 40
        ElseIf i = 9 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 20
        ElseIf i = 10 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top
        ElseIf i = 11 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top
        ElseIf i = 12 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 400 + Exp
            .Top = img(conta - 5).Top
        ElseIf i = 13 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 380 + Exp
            .Top = img(conta - 7).Top
        ElseIf i = 14 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 380 + Exp
            .Top = img(conta - 9).Top
        ElseIf i = 15 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 380 + Exp
            .Top = img(conta - 11).Top
        ElseIf i = 16 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 380 + Exp
            .Top = img(conta - 13).Top
        ElseIf i = 17 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 380 + Exp
            .Top = img(conta - 15).Top
        ElseIf i = 18 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 340 + Exp
            .Top = img(conta - 17).Top - 50
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea21(img As Object)
Dim i As Integer
On Error Resume Next
Dim Tx As Integer
Tx = conta + 14
conta = conta + 15
For i = 1 To 14
    conta = conta - 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial + 7600
            .Top = Top_Inicial + 300
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 20
            .Left = img(conta + 1).Left - OffLeft
            .Top = img(conta + 1).Top - AumentoTop - 30
        ElseIf i = 4 Then
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop - 10
        ElseIf i = 5 Then
            OffLeft = OffLeft - 10
            .Left = img(conta + 1).Left - OffLeft - 10
            .Top = img(conta + 1).Top - AumentoTop
        ElseIf i = 6 Then
            OffLeft = OffLeft + 10
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop + 20
        ElseIf i = 7 Then
            OffLeft = OffLeft + 40
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop
        ElseIf i = 8 Then
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop + 20
        ElseIf i = 9 Then
            .Left = img(conta + 1).Left - OffLeft - 150
            .Top = img(conta + 1).Top - AumentoTop + 20
        ElseIf i = 10 Then
            OffLeft = OffLeft + 150
            .Left = img(conta + 1).Left - OffLeft - 40
            .Top = img(conta + 1).Top - AumentoTop + 40
        ElseIf i = 11 Then '133
            OffLeft = OffLeft + 150
            .Left = img(conta + 1).Left - OffLeft + 120
            .Top = img(conta + 1).Top - AumentoTop + 60
        ElseIf i = 12 Then '134
            OffLeft = OffLeft + 150
            .Left = img(conta + 1).Left - OffLeft + 200
            .Top = img(conta + 1).Top - AumentoTop + 60
        ElseIf i = 13 Then
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft + 250
            .Top = img(conta + 1).Top - AumentoTop + 100
        Else
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft + 200
            .Top = img(conta + 1).Top - AumentoTop + 100
        End If
        .Visible = True
    End With
Next i
conta = Tx
End Sub
Private Sub Crea22(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 14
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial - 5400
            .Top = Top_Inicial + 250
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 20
            .Left = img(conta - 1).Left + OffLeft
            .Top = img(conta - 1).Top - AumentoTop - 30
        ElseIf i = 4 Then
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop - 10
        ElseIf i = 5 Then
            OffLeft = OffLeft - 10
            .Left = img(conta - 1).Left + OffLeft + 10
            .Top = img(conta - 1).Top - AumentoTop
        ElseIf i = 6 Then
            OffLeft = OffLeft + 10
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop + 20
        ElseIf i = 7 Then
            OffLeft = OffLeft + 40
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop
        ElseIf i = 8 Then
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft + 50
            .Top = img(conta - 1).Top - AumentoTop + 20
        ElseIf i = 9 Then
            .Left = img(conta - 1).Left + OffLeft + 150
            .Top = img(conta - 1).Top - AumentoTop + 20
        ElseIf i = 10 Then
            OffLeft = OffLeft + 150
            .Left = img(conta - 1).Left + OffLeft + 40
            .Top = img(conta - 1).Top - AumentoTop + 40
        ElseIf i = 11 Then '133
            OffLeft = OffLeft + 150
            .Left = img(conta - 1).Left + OffLeft - 60
            .Top = img(conta - 1).Top - AumentoTop + 60
        ElseIf i = 12 Then '134
            OffLeft = OffLeft + 150
            .Left = img(conta - 1).Left + OffLeft - 200
            .Top = img(conta - 1).Top - AumentoTop + 60
        ElseIf i = 13 Then
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft - 150
            .Top = img(conta - 1).Top - AumentoTop + 100
        Else
            OffLeft = OffLeft
            .Left = img(conta - 1).Left + OffLeft - 100
            .Top = img(conta - 1).Top - AumentoTop + 100
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea23(img As Object)
Dim i As Integer
On Error Resume Next
For i = 1 To 20
    conta = conta + 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial - 2400
            .Top = Top_Inicial - 5200
            OffLeft = 0
        ElseIf i = 2 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 340 + Exp
            .Top = img(conta - 1).Top - 300
        ElseIf i = 3 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 400 + Exp
            .Top = img(conta - 1).Top - 200
        ElseIf i = 4 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 400 + Exp
            .Top = img(conta - 1).Top - 200
        ElseIf i = 5 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 150
        ElseIf i = 6 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 120
        ElseIf i = 7 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 80
        ElseIf i = 8 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 470 + Exp
            .Top = img(conta - 1).Top - 40
        ElseIf i = 9 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top - 20
        ElseIf i = 10 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top
        ElseIf i = 11 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 440 + Exp
            .Top = img(conta - 1).Top
        ElseIf i = 12 Then '234
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 400 + Exp
            .Top = img(conta - 3).Top
        ElseIf i = 13 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 380 + Exp
            .Top = img(conta - 5).Top
        ElseIf i = 14 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 380 + Exp
            .Top = img(conta - 7).Top
        ElseIf i = 15 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 380 + Exp
            .Top = img(conta - 9).Top
        ElseIf i = 16 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 380 + Exp
            .Top = img(conta - 11).Top
        ElseIf i = 17 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 380 + Exp
            .Top = img(conta - 13).Top
        ElseIf i = 18 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 340 + Exp
            .Top = img(conta - 15).Top - 50
        ElseIf i = 19 Then
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 340 + Exp
            .Top = img(conta - 17).Top - 50
        Else
            OffLeft = OffLeft + 0
            .Left = img(conta - 1).Left + 340 + Exp
            .Top = img(conta - 19).Top - 50
        End If
        .Visible = True
    End With
Next i
End Sub
Private Sub Crea24(img As Object)
Dim i As Integer
On Error Resume Next
Dim Tx As Integer
Tx = conta + 14
conta = conta + 15
For i = 1 To 14
    conta = conta - 1
    Load img(conta)
    With img(conta)
        .Visible = True
        .BorderStyle = vbSolid
        .BorderColor = &H4080&
        .FillStyle = vbSolid
        .Height = 350
        .Width = 400
        If i = 1 Then
            .Left = Left_Inicial + 8500
            .Top = Top_Inicial + 250
            OffLeft = 0
        ElseIf i <= 3 Then
            OffLeft = OffLeft + 20
            .Left = img(conta + 1).Left - OffLeft
            .Top = img(conta + 1).Top - AumentoTop - 30
        ElseIf i = 4 Then
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop - 10
        ElseIf i = 5 Then
            OffLeft = OffLeft - 10
            .Left = img(conta + 1).Left - OffLeft - 10
            .Top = img(conta + 1).Top - AumentoTop
        ElseIf i = 6 Then
            OffLeft = OffLeft + 10
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop + 20
        ElseIf i = 7 Then
            OffLeft = OffLeft + 40
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop
        ElseIf i = 8 Then
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft - 50
            .Top = img(conta + 1).Top - AumentoTop + 20
        ElseIf i = 9 Then
            .Left = img(conta + 1).Left - OffLeft - 150
            .Top = img(conta + 1).Top - AumentoTop + 20
        ElseIf i = 10 Then
            OffLeft = OffLeft + 150
            .Left = img(conta + 1).Left - OffLeft - 40
            .Top = img(conta + 1).Top - AumentoTop + 40
        ElseIf i = 11 Then '133
            OffLeft = OffLeft + 150
            .Left = img(conta + 1).Left - OffLeft + 60
            .Top = img(conta + 1).Top - AumentoTop + 60
        ElseIf i = 12 Then '134
            OffLeft = OffLeft + 150
            .Left = img(conta + 1).Left - OffLeft + 200
            .Top = img(conta + 1).Top - AumentoTop + 60
        ElseIf i = 13 Then
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft + 150
            .Top = img(conta + 1).Top - AumentoTop + 100
        Else
            OffLeft = OffLeft
            .Left = img(conta + 1).Left - OffLeft + 100
            .Top = img(conta + 1).Top - AumentoTop + 100
        End If
        .Visible = True
    End With
Next i
End Sub
Private Function max(a As Long, b As Long) As Long
    max = IIf(a > b, a, b)
End Function
Private Function Min(a As Long, b As Long) As Long
    Min = IIf(a < b, a, b)
End Function
Private Function HayQueImprimir() As Boolean
Dim RsTemp As ADODB.Recordset
Set RsTemp = New ADODB.Recordset
SetearRs "SELECT ImprimirActa FROM ComunicacionRapida", RsTemp
If RsTemp.Fields(0) = 1 Then
    HayQueImprimir = True
Else
    HayQueImprimir = False
End If
End Function
Private Sub BorrarImpresion()
EjecutarSQL ("UPDATE ComunicacionRapida SET ImprimirActa = 0")
End Sub
Public Sub Impresion()
lblUltimaAccion.Caption = "Enviando a imprimir..."
Dim strTipoOp         As String
Dim strPeriodoLeg     As String
Dim xSesion           As Long
Dim xNroActaActual    As Long
Dim xRs As Recordset
Dim CantAfirmativos As Integer
Dim CantNegativos As Integer
Dim CantAbstenciones As Integer
Dim CantAusentes As Integer
Dim Bkp As Long
Dim bkpActa As Long
BorrarImpresion
'**********************
xSesion = Ultima_Sesion
strTipoOp = frmConsolaOperacion.dcTipoOperacion.BoundText
strPeriodoLeg = Ultimo_Periodo
mActaGrabada = Ultimo_Acta
'**********************
Bkp = xSesion
bkpActa = mActaGrabada
If Fue9999 = True Then
    xSesion = 9999
End If
If Fue9999 = True Then
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    SetearRs "SELECT max(Número_de_Acta) as maximo FROM actas WHERE Sesión = 9999 AND Período_Legislativo = '" & lblCodigoSesion.Tag & "'", RsTemp
    If RsTemp.EOF Then
        mActaGrabada = 1
    Else
        mActaGrabada = RsTemp.Fields(0)
    End If
    RsTemp.Close
    Set RsTemp = Nothing
End If
If strTipoOp <> "paslis" Then
    TotalPaginas = -1
    'Tipo_PreActa = "abs"
    'Solo abstencion
    'Call frmConsultarActa.imprimirActaFiltrada(strTipoOp, strPeriodoLeg, Val(Str(xSesion)), Val(Str(mActaGrabada)), 0, 1, 0, 0, 0, True)
    'Imprime todo
    Call frmConsultarActa.imprimirUnActa(strTipoOp, strPeriodoLeg, Val(Str(xSesion)), Val(Str(mActaGrabada)), 0)
    Dim rsCant As New ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT COUNT(Resultado) AS Abstenciones FROM detalleactas WHERE Período_Legislativo = '" & strPeriodoLeg & "' AND Sesión = " & xSesion & " AND Nro_de_Acta = " & mActaGrabada & " AND Versión_Acta = 0 AND LTrim(RTrim(Resultado)) = 'ABSTENCION'"
    SetearRs consulta, rsCant
    If (Not rsCant.EOF) Then
        If (rsCant.Fields(0) > 0) Then
            'Call frmConsultarActa.imprimirActaFiltrada(strTipoOp, strPeriodoLeg, Val(Str(xSesion)), Val(Str(mActaGrabada)), 0, 1, 0, 0, 0, True)
        Else
            'No hay abstenidos!
            'Imprimir solo el encabezado
            
        End If
    End If
End If
Tipo_PreActa = ""
If strTipoOp <> "votnum" Then
    TotalPaginas = -1
Else
    TotalPaginas = 1
End If
If strTipoOp <> "votnom" Then
    Call frmConsultarActa.imprimirUnActa(strTipoOp, strPeriodoLeg, Val(Str(xSesion)), Val(Str(mActaGrabada)), 0) 'Val(Str(xNroActaActual)), 0)
End If
xSesion = Bkp
mActaGrabada = bkpActa
mActaIniciada = 0
If mActaGrabada = 0 Then
    mActaImpresa = 0
    mActaIniciada = 0
End If
End Sub
