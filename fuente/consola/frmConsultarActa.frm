VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmConsultarActa 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de actas grabadas"
   ClientHeight    =   12015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12015
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNegativoDesempate 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   11010
      Locked          =   -1  'True
      TabIndex        =   75
      Text            =   "Vacio"
      Top             =   6900
      Width           =   1005
   End
   Begin VB.TextBox txtAfirmativosDesempate 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   73
      Text            =   "Vacio"
      Top             =   6900
      Width           =   855
   End
   Begin Proyecto1.ButtonOffice cmdReporte 
      Height          =   555
      Left            =   11160
      TabIndex        =   68
      Top             =   8040
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   979
      BackColor       =   33023
      Caption         =   "&Reporte"
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
   Begin VB.TextBox txtAbstencionesTotales 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   11010
      Locked          =   -1  'True
      TabIndex        =   67
      Text            =   "Vacio"
      Top             =   7380
      Width           =   1005
   End
   Begin VB.TextBox txtNegativoTotales 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   64
      Text            =   "Vacio"
      Top             =   7380
      Width           =   825
   End
   Begin VB.TextBox txtAfirmativosTotal 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   62
      Text            =   "Vacio"
      Top             =   7350
      Width           =   825
   End
   Begin VB.TextBox txtAusentesTotal 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   60
      Text            =   "Vacio"
      Top             =   6900
      Width           =   825
   End
   Begin VB.TextBox txtPresentesTotal 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   11010
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "Vacio"
      Top             =   6450
      Width           =   1005
   End
   Begin VB.TextBox txtPresentesNoId 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   56
      Text            =   "Vacio"
      Top             =   6450
      Width           =   855
   End
   Begin VB.TextBox txtPresentesId 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   3930
      Locked          =   -1  'True
      TabIndex        =   54
      Text            =   "Vacio"
      Top             =   6450
      Width           =   825
   End
   Begin VB.TextBox txtObservaciones 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1965
      Left            =   2550
      MaxLength       =   1024
      MultiLine       =   -1  'True
      TabIndex        =   53
      Top             =   4050
      Width           =   10005
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   3645
      Left            =   180
      TabIndex        =   37
      Top             =   2430
      Width           =   12495
      Begin VB.TextBox txtVotoPresidente 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   10770
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "Vacio"
         Top             =   1020
         Width           =   2445
      End
      Begin VB.TextBox txtTipoQuorum 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Vacio"
         Top             =   0
         Width           =   2445
      End
      Begin VB.TextBox txtMiembros 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   7980
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "Vacio"
         Top             =   0
         Width           =   1035
      End
      Begin VB.TextBox txtDesempate 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   10770
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "Vacio"
         Top             =   0
         Width           =   1035
      End
      Begin VB.TextBox txtTipoMayoria 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Vacio"
         Top             =   480
         Width           =   2445
      End
      Begin VB.TextBox txtBase 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   5700
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "Vacio"
         Top             =   480
         Width           =   2445
      End
      Begin VB.TextBox txtVotacion 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   10770
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Vacio"
         Top             =   420
         Width           =   2445
      End
      Begin VB.TextBox txtNombrePresidente 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Vacio"
         Top             =   1020
         Width           =   5925
      End
      Begin VB.Label lblVotoPresidente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voto Pres."
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
         Height          =   330
         Left            =   9120
         TabIndex        =   78
         Top             =   1020
         Width           =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de quorum"
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
         Height          =   330
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   2220
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Miembros del Cuerpo"
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
         Height          =   330
         Left            =   4890
         TabIndex        =   51
         Top             =   0
         Width           =   3000
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desempate"
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
         Height          =   330
         Left            =   9120
         TabIndex        =   50
         Top             =   0
         Width           =   1560
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de mayoría"
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
         Height          =   330
         Left            =   30
         TabIndex        =   49
         Top             =   480
         Width           =   2250
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base"
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
         Height          =   330
         Left            =   4890
         TabIndex        =   48
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Votación"
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
         Height          =   330
         Left            =   9120
         TabIndex        =   47
         Top             =   420
         Width           =   1230
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Presidente"
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
         Height          =   330
         Left            =   60
         TabIndex        =   46
         Top             =   1020
         Width           =   1500
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         Height          =   330
         Left            =   60
         TabIndex        =   45
         Top             =   1530
         Width           =   2115
      End
   End
   Begin VB.TextBox txtReunion 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   7410
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "Vacio"
      Top             =   600
      Width           =   1905
   End
   Begin VB.TextBox txtHora 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   3510
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "Vacio"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtFecha 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "Vacio"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtVersion 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Left            =   10620
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "Vacio"
      Top             =   660
      Width           =   3105
   End
   Begin VB.TextBox txtNroActa 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   10590
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "Vacio"
      Top             =   120
      Width           =   3105
   End
   Begin VB.TextBox txtSesion 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   7380
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Vacio"
      Top             =   90
      Width           =   1965
   End
   Begin VB.CommandButton cmdPresidente 
      Caption         =   "Cam&biar presidente"
      Height          =   375
      Left            =   12690
      TabIndex        =   13
      Top             =   12150
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txtBuscar 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2190
      TabIndex        =   12
      Top             =   11400
      Width           =   8805
   End
   Begin VB.ComboBox cmbResultados 
      Height          =   315
      Left            =   1080
      TabIndex        =   11
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAbstencionesNoId 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   19650
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   12510
      Width           =   615
   End
   Begin VB.TextBox txtAbstencionesId 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   19650
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   12120
      Width           =   615
   End
   Begin VB.TextBox txtNegativoNoId 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   16650
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   12510
      Width           =   615
   End
   Begin VB.TextBox txtNegativoID 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   16650
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   12120
      Width           =   615
   End
   Begin VB.TextBox txtAfirmativosNoId 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   13860
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   12510
      Width           =   615
   End
   Begin VB.TextBox txtAfirmativosId 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   13860
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   12120
      Width           =   615
   End
   Begin VB.TextBox txtCodigoPresidente 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   12060
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   12090
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   840
      Left            =   2730
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmConsultarActa.frx":0000
      Top             =   1110
      Width           =   9195
   End
   Begin VB.TextBox txtTipoOperacion 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   2820
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Vacio"
      Top             =   120
      Width           =   3105
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exportar Ascii"
      Height          =   375
      Left            =   12570
      TabIndex        =   1
      Top             =   12180
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.ComboBox cmbImpresoras 
      Height          =   315
      Left            =   12240
      TabIndex        =   0
      Text            =   "Impresora Default"
      Top             =   12120
      Visible         =   0   'False
      Width           =   2595
   End
   Begin MSFlexGridLib.MSFlexGrid vsGrilla 
      Height          =   3315
      Left            =   360
      TabIndex        =   14
      Top             =   7980
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   5847
      _Version        =   393216
      Cols            =   6
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Proyecto1.ButtonOffice cmdAnularActa 
      Height          =   555
      Left            =   11160
      TabIndex        =   69
      Top             =   8820
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   979
      BackColor       =   33023
      Caption         =   "ANULAR ACTA"
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
   Begin Proyecto1.ButtonOffice cmdMOdificar 
      Height          =   555
      Left            =   11160
      TabIndex        =   70
      Top             =   10020
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   979
      BackColor       =   33023
      Caption         =   "&Modificar"
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
   Begin Proyecto1.ButtonOffice cmdAceptar 
      Height          =   555
      Left            =   11160
      TabIndex        =   71
      Top             =   10620
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   979
      BackColor       =   33023
      Caption         =   "&Aceptar"
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
   Begin Proyecto1.ButtonOffice cmdCerrar 
      Height          =   555
      Left            =   11160
      TabIndex        =   72
      Top             =   11220
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   979
      BackColor       =   33023
      Caption         =   "&Cerrar"
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
   Begin Proyecto1.ButtonOffice cmdVivavoz 
      Height          =   555
      Left            =   11160
      TabIndex        =   15
      Top             =   9420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   979
      BackColor       =   33023
      Caption         =   "Viva voz"
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
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desemp. Neg"
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
      Height          =   330
      Left            =   9060
      TabIndex        =   76
      Top             =   6900
      Width           =   1830
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desemp. Afirm."
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
      Height          =   330
      Left            =   5280
      TabIndex        =   74
      Top             =   6900
      Width           =   2100
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   11460
      X2              =   12060
      Y1              =   8730
      Y2              =   8730
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   210
      X2              =   12570
      Y1              =   7860
      Y2              =   7860
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Abstenciones"
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
      Height          =   330
      Left            =   9060
      TabIndex        =   66
      Top             =   7380
      Width           =   1905
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Votos Negativos"
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
      Height          =   330
      Left            =   5280
      TabIndex        =   65
      Top             =   7350
      Width           =   2265
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Votos Afirmativos"
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
      Height          =   330
      Left            =   540
      TabIndex        =   63
      Top             =   7350
      Width           =   2460
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ausentes Totales"
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
      Height          =   330
      Left            =   540
      TabIndex        =   61
      Top             =   6900
      Width           =   2415
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Height          =   330
      Left            =   9060
      TabIndex        =   59
      Top             =   6450
      Width           =   705
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Identificados"
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
      Height          =   330
      Left            =   5280
      TabIndex        =   57
      Top             =   6450
      Width           =   2235
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Presentes Identificados"
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
      Height          =   330
      Left            =   510
      TabIndex        =   55
      Top             =   6450
      Width           =   3270
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   210
      X2              =   12570
      Y1              =   6300
      Y2              =   6300
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reunión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   6120
      TabIndex        =   36
      Top             =   600
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   210
      X2              =   12570
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Acta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   150
      TabIndex        =   34
      Top             =   1110
      Width           =   2280
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   2730
      TabIndex        =   33
      Top             =   600
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   150
      TabIndex        =   31
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   9420
      TabIndex        =   29
      Top             =   660
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº acta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   9510
      TabIndex        =   27
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sesión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   6300
      TabIndex        =   24
      Top             =   90
      Width           =   960
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   120
      TabIndex        =   23
      Top             =   90
      Width           =   2535
   End
   Begin VB.Label lblBuscarLegislador 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar Diputado"
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
      Height          =   225
      Left            =   420
      TabIndex        =   22
      Top             =   11490
      Width           =   1410
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Abstenciones No Identificables"
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
      Height          =   240
      Left            =   16920
      TabIndex        =   21
      Top             =   12570
      Width           =   2910
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Abstenciones identificables"
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
      Height          =   240
      Left            =   16920
      TabIndex        =   20
      Top             =   12180
      Width           =   2595
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Votos Neg. No Identificables"
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
      Height          =   240
      Left            =   14070
      TabIndex        =   19
      Top             =   12570
      Width           =   2670
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Votos Neg. Identificables"
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
      Height          =   240
      Left            =   14070
      TabIndex        =   18
      Top             =   12180
      Width           =   2355
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Votos Afirm. No Identificables"
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
      Height          =   240
      Left            =   11220
      TabIndex        =   17
      Top             =   12570
      Width           =   2805
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Votos Afirm. Identificables"
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
      Height          =   240
      Left            =   11220
      TabIndex        =   16
      Top             =   12180
      Width           =   2490
   End
End
Attribute VB_Name = "frmConsultarActa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mVivavozChanged                 As Boolean
Private mVivavozEntered                 As Boolean
Private rstActa                         As New ADODB.Recordset
Private rstActa2                        As New ADODB.Recordset
Private WithEvents rs                   As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private mActa                           As Integer
Private mPeriodo                        As String
Private mSesion                         As Integer
Private mVersion                        As Integer
Private mValorAnteriorVoto              As String
Private mPaseLista                      As Boolean
Private mCambios                        As Boolean
Private strVotoPresidente               As String
' ------------------------------------------------------------------------------------------------
' Variables usadas para Calculo de Resultado y validación previa a grabar acta editada
' ------------------------------------------------------------------------------------------------
Private strTipoOperacion                  As String
Private strPeriodo_Legislativo            As String
Private strSesion                         As String
Private xNumeroActa                       As Long
Private xVersionActa                      As Long
Private xUltimaVersionActa                As Long
Private strNombreActa                     As String
Private strTipoQuorum                     As String
Private strTipoMayoria                    As String
Private strBaseMayoria                    As String
Private xMiembrosDelCuerpo                As Long
Private strResultadoEsperado              As String  ' Para validar que no modifiquen el resultado de la votacion en edicion de acta
Private xNroOrdenDia                      As Long
Private blEsLegislador                    As Boolean
Private xMinimoParaQuorum                 As Long

Private xPresentesPaseDeLista             As Long
'-------------------------------------------------------------------------------------------------
Private xResultadoVotaPresidente          As String
'-------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------
' Variables usadas para control
' ------------------------------------------------------------------------------------------------
Private xVotosAfirmativosIdentificables   As Long
Private xVotosAfirmativosTotal            As Long
Private xVotosAfirmativosNoIdentificables As Long
Private xVotosAfirmativosDesempate        As Long
Private xVotosNegativosIdentificables     As Long
Private xVotosNegativosNoIdentificables   As Long
Private xVotosNegativosDesempate          As Long
Private xVotosNegativosTotal              As Long
Private xAbstencionesIdentificables       As Long
Private xAbstencionesNoIdentificables     As Long
Private xAbstencionesTotal                As Long
Private xPresentesIdentificables          As Long
Private xPresentesNoIdentificables        As Long
Private xPresentesTotal                   As Long
' ------------------------------------------------------------------------------------------------
' Variables usadas para parser RTF
' ------------------------------------------------------------------------------------------------
Private strCadenaRtfLogo                      As String

Private Sub cmdConexion_Click()
    frmConfig.Show vbModal
End Sub

Private Sub cmdUsuarios_Click()
    FrmGestionarUsuarios.Show vbModal
End Sub

Private Sub cmdVolver_Click()
    Unload Me
End Sub

Private Sub cmbResultados_Click()
      cmbResultados.Visible = False
End Sub

Private Sub cmbResultados_LostFocus()
      cmbResultados.Visible = False
End Sub
Private Sub cmbResultados_Validate(Cancel As Boolean)
    
    Dim strCambioNuevo As String
    
    strCambioNuevo = Trim(UCase(cmbResultados.Text))
    Cancel = False
    
    If (mValorAnteriorVoto <> cmbResultados.Text) Then
        If mValorAnteriorVoto = "AUSENTE" Then
            If Val(txtPresentesNoId.Text) <= 0 Then
                MsgBox "No se encuentran legisladores presentes no identificados para asignar.", vbInformation + vbOKOnly
                Cancel = True
            End If
        End If
    End If
    
    If Cancel = False Then
        xVotosAfirmativosIdentificables = xVotosAfirmativosIdentificables + IIf(mValorAnteriorVoto = "AFIRMATIVO", -1, 0) + IIf(strCambioNuevo = "AFIRMATIVO", 1, 0)
        xVotosNegativosIdentificables = xVotosNegativosIdentificables + IIf(mValorAnteriorVoto = "NEGATIVO", -1, 0) + IIf(strCambioNuevo = "NEGATIVO", 1, 0)
        xAbstencionesIdentificables = xAbstencionesIdentificables + IIf(mValorAnteriorVoto = "ABSTENCION", -1, 0) + IIf(strCambioNuevo = "ABSTENCION", 1, 0)
        xPresentesIdentificables = xPresentesIdentificables + IIf(InStr("AFIRMATIVO NEGATIVO ABSTENCION PRESENTE", mValorAnteriorVoto) > 0, -1, 0) + IIf(InStr("AFIRMATIVO NEGATIVO ABSTENCION PRESENTE", strCambioNuevo) > 0, 1, 0)
        xPresentesNoIdentificables = xPresentesNoIdentificables + IIf(mValorAnteriorVoto = "AUSENTE", -1, 0) + IIf(strCambioNuevo = "AUSENTE", 1, 0)
        xPresentesTotal = xPresentesIdentificables + xPresentesNoIdentificables
        xVotosAfirmativosTotal = xVotosAfirmativosIdentificables + xVotosAfirmativosNoIdentificables + xVotosAfirmativosDesempate
        xVotosNegativosTotal = xVotosNegativosIdentificables + xVotosNegativosNoIdentificables + xVotosNegativosDesempate
        xAbstencionesTotal = xAbstencionesIdentificables + xAbstencionesNoIdentificables
        vsGrilla.Text = cmbResultados.Text
        mCambios = True
    Else
        vsGrilla.Text = mValorAnteriorVoto
    End If
    Call MostrarCuentas
End Sub
Private Sub cmdAceptar_Click()
    ControlesHabilitados = False
    cmdMOdificar.Enabled = True
    Call GuardarCambios
End Sub

Private Sub cmdAnularActa_Click()
Dim r As Integer
Dim consulta As String
Dim Resultado_Original As String
If InStr(LCase(cmdAnularActa.Caption), "anular") > 0 Then
    r = MsgBox("¿Está seguro que desea anular este acta?", vbYesNo, "Consola")
    If r = vbYes Then
        consulta = "UPDATE actas SET Votacion = 'ANULADA' WHERE " & "(Actas.Período_Legislativo='" & mPeriodo & "') AND (Actas.Sesión=" & mSesion & ") AND (Actas.Número_de_Acta=" & mActa & ") AND (Actas.Versión_Acta=" & mVersion & ")"
        Call EjecutarSQL(consulta)
        Call MsgBox("El resultado ha sido modificado con éxito", vbInformation, "Consola")
        cmdAnularActa.Caption = "RESTAURAR RESULTADO"
        txtVotacion.Text = "ANULADA"
    End If
Else
    Resultado_Original = LCase(Trim(CalculoResultado(strBaseMayoria, strTipoMayoria, xMiembrosDelCuerpo, xPresentesTotal, xVotosAfirmativosTotal, xVotosNegativosTotal, "", 0, 0, 0, strVotoPresidente, IIf(blEsLegislador, 1, 0))))
    r = MsgBox("¿Desea restaurar el resultado del acta?", vbYesNo, "Consola")
    If r = vbYes Then
        consulta = "UPDATE actas SET Votacion = '" & UCase(Resultado_Original) & "' WHERE " & "(Actas.Período_Legislativo='" & mPeriodo & "') AND (Actas.Sesión=" & mSesion & ") AND (Actas.Número_de_Acta=" & mActa & ") AND (Actas.Versión_Acta=" & mVersion & ")"
        Call EjecutarSQL(consulta)
        Call MsgBox("El resultado ha sido restaurado con éxito: " & UCase(Resultado_Original), vbInformation, "Consola")
        cmdAnularActa.Caption = "ANULAR ACTA"
        txtVotacion.Text = UCase(Resultado_Original)
    End If
End If
End Sub

Private Sub cmdCerrar_Click()
Dim i As Integer
If mCambios = True Or mVivavozChanged = True Then
    i = MsgBox("Ha realizado cambios en el acta. ¿Desea salir sin guardar?", vbYesNo)
    If i = vbYes Then
        Unload Me
    End If
Else
    'mCambios = True
    'GuardarCambios
    Unload Me
End If
End Sub
Private Sub GuardarCambios()
    Dim strNuevoResultado As String
    Dim xRespuesta        As Integer
    
    If mCambios = False And mVivavozChanged = False Then
        Exit Sub
    End If
    strTipoOperacion = Trim(LCase(strTipoOperacion))
    strResultadoEsperado = Trim(LCase(strResultadoEsperado))
    ' Confirmar Edición de actas
    xRespuesta = MsgBox("Está Ud. seguro de guardar las modificaciones realizadas?", vbQuestion + vbYesNo, "Confirmar Operación")
    If xRespuesta = vbNo Then
        Exit Sub
    End If
    
    ' ------------------------------------------------------------------------------------
    ' Votacion Nominal
    ' ------------------------------------------------------------------------------------
    If LCase(strTipoOperacion) = "votnom" Then
        ' Verificar que no se haya modificado el resultado
        Call BuscarVotoPresidente
        strNuevoResultado = LCase(Trim(CalculoResultado(strBaseMayoria, strTipoMayoria, xMiembrosDelCuerpo, xPresentesTotal, xVotosAfirmativosTotal, xVotosNegativosTotal, "", 0, 0, 0, strVotoPresidente, IIf(blEsLegislador, 1, 0))))
        If strNuevoResultado <> strResultadoEsperado Then
            ' El usuario no puede grabar acta con cambios en los resultados. Opciones posibles:
            xRespuesta = MsgBox("El resultado " & UCase(strNuevoResultado) & " difiere del original. Elija Si para cancelar y No para reingresar", vbYesNo + vbCritical, "Validación de ACtualización de Actas")
            If xRespuesta = vbYes Then
                Unload Me
            Else
                Call cmdModificar_Click
            End If
        Else
            Call AlmacenarCambios
        End If
    End If
    ' ------------------------------------------------------------------------------------
    ' PASE DE LISTA
    ' ------------------------------------------------------------------------------------
    If strTipoOperacion = "paslis" Then
         ' If xPresentesTotal >= CalcularMinimoParaQuorum(strTipoQuorum, xMiembrosDelCuerpo) Then
         Call ContarPresentes
         If xPresentesPaseDeLista >= CalcularMinimoParaQuorum(strTipoQuorum, xMiembrosDelCuerpo) Then
            strNuevoResultado = "quorum"
         Else
            strNuevoResultado = "no hay quorum"
         End If
        If strNuevoResultado <> strResultadoEsperado Then
            ' El usuario no puede grabar acta con cambios en los resultados. Opciones posibles:
            xRespuesta = MsgBox("El resultado " & UCase(strNuevoResultado) & " difiere del original. Elija Si para cancelar y No para reingresar", vbYesNo + vbCritical, "Validación de ACtualización de Actas")
            If xRespuesta = vbYes Then
                Unload Me
            Else
                Call cmdModificar_Click
            End If
        Else
            Call AlmacenarCambios
        End If
    End If
    ' ------------------------------------------------------------------------------------
    ' VOTACION NUMERICA
    ' ------------------------------------------------------------------------------------
    If strTipoOperacion = "votnum" Then
        Call AlmacenarCambios
    End If
End Sub
Private Sub ContarPresentes()
    ' recorrer la grilla y contar presentes totales
    Dim X          As Long
    Dim strColumna As String
    xPresentesPaseDeLista = 0
    With vsGrilla
        For X = 1 To .Rows - 1
            strColumna = Trim(LCase(.TextMatrix(X, 4)))
            If strColumna = "presente" Then
                xPresentesPaseDeLista = xPresentesPaseDeLista + 1
            End If
        Next X
    End With
End Sub
Private Sub BuscarVotoPresidente()
    Dim xBanca       As Long
    Dim strResultado As String
    Dim X            As Long
    
    With vsGrilla
        For X = 1 To .Rows - 1
            xBanca = Int(.TextMatrix(X, 1))
            strResultado = UCase(Trim(.TextMatrix(X, 4)))
            If xBanca = 0 Then
                strVotoPresidente = Trim(LCase(strResultado))
                Select Case strVotoPresidente
                    Case "afirmativo"
                        strVotoPresidente = "s"
                        Exit For
                    Case "negativo"
                        strVotoPresidente = "n"
                        Exit For
                    Case Else
                        strVotoPresidente = " " ' Hubo algun error de logica interna
                End Select
            End If
        Next X
    End With

End Sub

Private Sub AlmacenarCambios()
    Dim strSql       As String
    Dim RsGrabar     As ADODB.Recordset
    Dim xPresentes   As Long
    Dim xAfirmativos As Long
    Dim xNegativos   As Long
    Dim X            As Long
    Dim xBanca       As Long
    Dim strLeg       As String
    Dim strResultado As String
    Dim strNombre    As String
    
    Dim strOperacion          As String
    Dim strLegisladorAsignado As String
    Dim strBloquePolitico     As String
    Dim strDepartamento       As String
    Dim strGrupoPolitico      As String
    
    'Dim strLegPres As String
    'Guardo Manifestaciones
    If (mVivavozChanged) Then
        Call Me.GuardarManifestaciones
    Else
        Call Me.CopiarManifestaciones
    End If
    Set RsGrabar = New ADODB.Recordset
    Call AbrirDB
    xUltimaVersionActa = xUltimaVersionActa + 1
    Dim PHabilitado As String
    Dim TempP As String
    Dim rsAC As New ADODB.Recordset
    TempP = "SELECT presidente_habilitado_votar FROM actas " _
           & "WHERE período_legislativo = '" & mPeriodo & "' AND sesión = " & mSesion _
           & " AND Número_de_acta = " & mActa & " AND versión_acta = " & mVersion
    SetearRs TempP, rsAC
    If rsAC.EOF Then
        PHabilitado = "0"
    Else
        PHabilitado = rsAC.Fields(0)
    End If
    If strTipoOperacion <> "votnum" Then
        ' Grabar Detalles Actas
        Screen.MousePointer = 11
            For X = 1 To vsGrilla.Rows - 1
                ' leer datos de la grilla
                xBanca = Int(vsGrilla.TextMatrix(X, 1))
                strLeg = Trim(vsGrilla.TextMatrix(X, 2))
                strNombre = vsGrilla.TextMatrix(X, 3)
                strResultado = UCase(Trim(vsGrilla.TextMatrix(X, 4)))
                ' Buscar los datos que faltan de la base de datos
                strSql = "SELECT * FROM DetalleActas " _
                       & "WHERE período_legislativo = '" & mPeriodo & "' AND sesión = " & mSesion _
                       & " AND nro_de_acta = " & mActa & " AND versión_acta = " & mVersion & " " _
                       & "AND Numero_de_banca = " & Str(xBanca)
                SetearRs strSql, rstActa2
                strOperacion = Trim(rstActa2.Fields("Operación").Value)
                strLegisladorAsignado = Trim(rstActa2.Fields("Legislador_asignado").Value)
                strBloquePolitico = IIf(IsNull((Trim(rstActa2.Fields("Bloque_político").Value))), "", Trim(rstActa2.Fields("Bloque_político").Value))
                strDepartamento = IIf(IsNull(Trim(rstActa2.Fields("Departamento").Value)), "", Trim(rstActa2.Fields("Departamento").Value))
                If Not IsNull(rstActa2.Fields("Grupo_politico").Value) Then
                    strGrupoPolitico = Trim(rstActa2.Fields("Grupo_politico").Value)
                Else
                    strGrupoPolitico = ""
                End If
                rstActa2.Close
                
                ' Grabar nueva version del acta
                strSql = "UPDATE DetalleActas SET versión_acta = " & xUltimaVersionActa & " WHERE período_legislativo = '" & mPeriodo & "' AND sesión = " & mSesion _
                       & " AND nro_de_acta = " & mActa & " AND versión_acta = " & Str(mVersion) _
                       & " AND Numero_de_banca = " & Str(xBanca)
                ' Debug.Print strSql
                Call SenteciaSQl(strSql)
                strSql = "INSERT INTO DetalleActas (Período_Legislativo, Sesión, Nro_de_Acta, Versión_Acta, " _
                       & "Operación, Numero_de_banca, Resultado, Legislador_asignado, " _
                       & "Bloque_político, Departamento, Grupo_politico) " _
                       & "VALUES ('" & mPeriodo & "', " & mSesion & ", " & mActa & ", 0, " _
                       & "'" & strOperacion & "', " & Str(xBanca) & ", '" & strResultado & "', '" & strLegisladorAsignado & "', " _
                       & "'" & strBloquePolitico & "', '" & strDepartamento & "', '" & strGrupoPolitico & "')"
                Call SenteciaSQl(strSql)
            Next X
        Screen.MousePointer = 0
    End If
    
    ' Grabar Actas
    strSql = "UPDATE Actas SET Versión_Acta = " & xUltimaVersionActa & " " _
           & "WHERE período_legislativo = '" & mPeriodo & "' AND sesión = " & mSesion _
           & " AND Número_de_Acta = " & mActa & " AND versión_acta = 0"
    Call SenteciaSQl(strSql)
    strSql = "SELECT * FROM Actas WHERE 1 = 0"
    SetearRsW strSql, rstActa2
    With rstActa2
        .AddNew
        .Fields("Tipo_de_operación").Value = Trim(strTipoOperacion)
        .Fields("Período_Legislativo").Value = mPeriodo
        .Fields("Sesión").Value = mSesion
        .Fields("Número_de_Acta").Value = mActa
        .Fields("Versión_Acta").Value = 0
        .Fields("Ultima_Versión_Acta").Value = xUltimaVersionActa
        .Fields("Nombre_del_Acta").Value = strNombreActa
        .Fields("Tipo_de_Quorum").Value = UCase(strTipoQuorum)
        .Fields("Base_de_Mayoria").Value = strBaseMayoria
        .Fields("Tipo_de_Mayoria").Value = strTipoMayoria
        .Fields("Miembros_del_cuerpo").Value = txtMiembros.Text
        .Fields("Desempate").Value = Trim(txtDesempate.Text)
        .Fields("Votacion").Value = Trim(txtVotacion.Text)
        .Fields("Presidente").Value = Trim(txtCodigoPresidente.Text)
        .Fields("Presentes_Identificables").Value = xPresentesIdentificables
        .Fields("Presentes_No_Identificables").Value = xPresentesNoIdentificables
        .Fields("Presentes_Total").Value = xPresentesTotal
        .Fields("Ausentes_Total").Value = txtAusentesTotal.Text
        .Fields("Votos_Afirm_Identificables").Value = xVotosAfirmativosIdentificables
        .Fields("Votos_Afirm_No_Identificables").Value = xVotosAfirmativosNoIdentificables
        .Fields("Votos_Afirm_Desempate").Value = xVotosAfirmativosDesempate
        .Fields("Votos_Afirm_Total").Value = xVotosAfirmativosTotal
        .Fields("Votos_Neg_Identificables").Value = xVotosNegativosIdentificables
        .Fields("Votos_Neg_No_Identificables").Value = xVotosNegativosNoIdentificables
        .Fields("Votos_Neg_Desempate").Value = xVotosNegativosDesempate
        .Fields("Votos_Neg_Total").Value = xVotosNegativosTotal
        .Fields("Abstenciones_Identificables").Value = xAbstencionesIdentificables
        .Fields("Abstenciones_No_Identificables").Value = xAbstencionesNoIdentificables
        .Fields("Abstenciones_Total").Value = xAbstencionesTotal
        .Fields("Fecha_Modificacion").Value = Format(Date, "DD/MM/YYYY")
        .Fields("Hora_Modificacion").Value = Format(Time, "HH:MM")
        .Fields("Usuario_Modificacion").Value = "Consola SQV"
        .Fields("IP_Modificacion").Value = ""
        .Fields("Observaciones").Value = Trim(txtObservaciones)
        .Fields("vota_presidente").Value = IIf(blEsLegislador = True, 1, 0)
        .Fields("NroOrdenDia").Value = xNroOrdenDia
        .Fields("Fecha").Value = txtFecha.Text 'Format(txtFecha, "DD/MM/YYYY")
        .Fields("Hora").Value = "" 'Format(txtHora, "HH:MM")
        .Fields("Reunion").Value = txtReunion.Text
        .Fields("presidente_habilitado_votar") = PHabilitado
        .Fields("resultado_voto_presidente") = xResultadoVotaPresidente
        .Update
        .Close
    End With
    Screen.MousePointer = 0
    Set rstActa2 = Nothing
    mCambios = False
End Sub
Private Sub cmdModificar_Click()
    If PermisosTotales.ModificaActas = 1 Then
        ControlesHabilitados = True
        cmdMOdificar.Enabled = False
    Else
        MsgBox "El usuario no tiene permisos para realizar esta tarea", vbInformation + vbOKOnly
    End If
End Sub
Private Sub cmdpresidente_Click()
    Call cambiarPresidente
End Sub
Private Sub cambiarPresidente()
    Dim cambio As New frmElegirPresidenteActas
    If cambio.MostrarDatos(vsGrilla, 2, 3, 4, Me) = True Then
        cambio.Show vbModal
    End If
    Set cambio = Nothing
End Sub
Public Sub RealizarCambioPresidente(pCodigoLEgislador As String, pNombre As String, pFilaNuevo As Integer)
    Dim pExPresidente       As String
    Dim pExPresidenteCodigo As String
    
    If pFilaNuevo > 0 Then
        pExPresidente = Trim(txtNombrePresidente.Text)
        pExPresidenteCodigo = Trim(txtCodigoPresidente.Text)
        txtNombrePresidente.Text = Trim(pNombre)
        txtCodigoPresidente.Text = Trim(pCodigoLEgislador)
        pExPresidenteCodigo = Trim(pCodigoLEgislador)
        vsGrilla.TextMatrix(pFilaNuevo, 2) = pExPresidenteCodigo
        vsGrilla.TextMatrix(pFilaNuevo, 3) = pExPresidente
        mCambios = True
    End If
End Sub

Public Sub cmdReporte_Click()
    Dim xRs As Recordset
    Dim CantAfirmativos As Integer
    Dim CantNegativos As Integer
    Dim CantAbstenciones As Integer
    Dim CantAusentes As Integer
    Dim CantidadPaginas As Integer
    Dim RsTemp As ADODB.Recordset
    Dim RsTemp2 As ADODB.Recordset
    cmdReporte.Caption = "Cargando..."
    cmdReporte.Enabled = False
    DoEvents
    CantidadPaginasAuditoria = 0
    If strTipoOperacion = "votnum" Then
        EsNum = True
        Tipo_PreActa = "consulta"
        If (mVersion > 0) Then
            Call imprimirUnActa(strTipoOperacion, mPeriodo, mSesion, mActa, mVersion)
        Else
            EsNum = False
            Tipo_PreActa = ""
            Call imprimirUnActa(strTipoOperacion, mPeriodo, mSesion, mActa, mVersion)
        End If
        Tipo_PreActa = ""
        EsNum = False
    Else
        Dim consulta As String
        Set RsTemp = New ADODB.Recordset
        consulta = "SELECT * FROM actas WHERE Período_Legislativo = '" & mPeriodo & "' AND Sesión = " & mSesion & " AND Número_de_Acta = " & mActa & " AND Ultima_Versión_Acta > 0"
        SetearRs consulta, RsTemp
        If RsTemp.EOF Then
            TotalPaginas = -1
            Tipo_PreActa = "consulta"
            imprimirActaFiltrada strTipoOperacion, mPeriodo, mSesion, mActa, mVersion, 1, 0, 0, 0, False
            Tipo_PreActa = ""
        Else
            Acta_Titulo = ""
            Acta_SubTitulo = ""
            Acta_Numero = ""
            Acta_Version = ""
            Acta_Miembros = ""
            Acta_Presidente = ""
            Acta_TipoMayoria = ""
            Acta_BaseMayoria = ""
            Acta_Fecha = ""
            Acta_Hora = ""
            Acta_TipoQuorum = ""
            Acta_ResultadoVotacion = ""
            Set RsTemp2 = New ADODB.Recordset
            SetearRs "SELECT Observaciones,Fecha_Modificacion,Hora_Modificacion FROM actas WHERE " & "(Actas.Período_Legislativo='" & mPeriodo & "') AND (Actas.Sesión=" & mSesion & ") AND (Actas.Número_de_Acta=" & mActa & ") AND (Actas.Versión_Acta=" & mVersion & ")", RsTemp2
            If RsTemp2.EOF Then
                Acta_Observaciones = "Ninguna."
                Acta_FechaModificacion = ""
                Acta_HoraModificacion = ""
            Else
                Acta_Observaciones = Trim(RsTemp2.Fields("Observaciones"))
                Acta_FechaModificacion = Trim(RsTemp2.Fields("Fecha_Modificacion"))
                Acta_HoraModificacion = Trim(RsTemp2.Fields("Hora_Modificacion"))
            End If
            RsTemp2.Close
            Set RsTemp2 = Nothing
            Set TodoElReporte = New ActiveReport
            Auditoria_Index = 0
            CantidadPaginasAuditoria = ObtenerCantidadAuditoria(strTipoOperacion, mPeriodo, mSesion, mActa, mVersion, 1, 0, 0, 0, False)
            CantidadPaginas = ObtenerCantidadDePaginas(strTipoOperacion, mPeriodo, mSesion, mActa, mVersion)
            TotalPaginas = CantidadPaginas + CantidadPaginasAuditoria
            Paginas_Fijas_Auditoria = TotalPaginas
            Paginas_Auditoria = Paginas_Fijas_Auditoria
            Auditoria_Index = TotalPaginas - CantidadPaginasAuditoria
            Call frmConsultarActa.imprimirUnActaCompleta(strTipoOperacion, mPeriodo, Val(Str(mSesion)), Val(Str(mActa)), mVersion) 'Val(Str(xNroActaActual)), 0)
            Call TodoElReporte.Pages.Remove(TodoElReporte.Pages.Count - 1)
            TodoElReporte.Pages.Commit
            Dim nTick As Long
            nTick = GetTickCount
'            While GetTickCount - nTick < 1500
'                DoEvents
'            Wend
            imprimirActaAuditoria strTipoOperacion, mPeriodo, mSesion, mActa, mVersion, 1, 0, 0, 0, False
            nTick = GetTickCount
'            While GetTickCount - nTick < 1500
'                DoEvents
'            Wend
            TodoElReporte.Pages(TodoElReporte.Pages.Count - 1).Width = TodoElReporte.Pages(TodoElReporte.Pages.Count - 1).Width - 2000
            TodoElReporte.Pages.Commit
            nTick = GetTickCount
'            While GetTickCount - nTick < 1500
'                DoEvents
'            Wend
            Dim X As Integer
            For X = 0 To TodoElReporte.Pages.Count - 1
                TodoElReporte.Pages(X).Width = 300
                TodoElReporte.Pages.Commit
            Next X
            TodoElReporte.Pages.Commit
            TodoElReporte.Run False
            nTick = GetTickCount
'            While GetTickCount - nTick < 1500
'                DoEvents
'            Wend
            TotalPaginas = TodoElReporte.Pages.Count
            CantidadPaginas = TotalPaginas - CantidadPaginasAuditoria
            Paginas_Fijas_Auditoria = TotalPaginas
            TodoElReporte.Pages.RemoveAll
            TodoElReporte.Pages.Commit
            nTick = GetTickCount
'            While GetTickCount - nTick < 1500
'                DoEvents
'            Wend
            'Ahora se hace todo de nuevo sabiendo la cantidad de paginas
            Call frmConsultarActa.imprimirUnActaCompleta(strTipoOperacion, mPeriodo, Val(Str(mSesion)), Val(Str(mActa)), mVersion) 'Val(Str(xNroActaActual)), 0)
            TodoElReporte.Pages.Commit
            nTick = GetTickCount
'            While GetTickCount - nTick < 1500
'                DoEvents
'            Wend
            imprimirActaAuditoria strTipoOperacion, mPeriodo, mSesion, mActa, mVersion, 1, 0, 0, 0, False
            nTick = GetTickCount
'            While GetTickCount - nTick < 1500
'                DoEvents
'            Wend
            'MANIFESTACIONES
            'imprimirReporteManifestaciones
            TodoElReporte.Pages(TodoElReporte.Pages.Count - 1).Width = TodoElReporte.Pages(TodoElReporte.Pages.Count - 1).Width - 2000
            TodoElReporte.Pages.Commit
            nTick = GetTickCount
'            While GetTickCount - nTick < 1500
'                DoEvents
'            Wend
            For X = 0 To TodoElReporte.Pages.Count - 1
                TodoElReporte.Pages(X).Width = 300
                TodoElReporte.Pages.Commit
            Next X
            TodoElReporte.Pages.Commit
            TodoElReporte.Run False
            nTick = GetTickCount
            While GetTickCount - nTick < 1500
                DoEvents
            Wend
            If ImpresionDeConsola = True Then
                TodoElReporte.PrintReport False
            Else
                TodoElReporte.PrintReport True
            End If
        End If
        RsTemp.Close
        Set RsTemp = Nothing
    End If
    CantidadPaginasAuditoria = 0
    cmdReporte.Caption = "&Reporte"
    cmdReporte.Enabled = True
End Sub
Public Sub imprimirReporteManifestaciones(rAudit As rptAuditoria)
Dim rpt As New rptManifestaciones
Dim rs As New ADODB.Recordset
Dim s As String
s = "SELECT id_diputado, Legisladores.apellido + ', ' + Legisladores.nombre as diputado, comentario " & _
" FROM manifestaciones_vivavoz INNER JOIN Legisladores ON Legisladores.id = manifestaciones_vivavoz.id_diputado " & _
" WHERE manifestaciones_vivavoz.periodo = '" & mPeriodo & "'" & _
" AND manifestaciones_vivavoz.sesion = " & mSesion & _
" AND manifestaciones_vivavoz.nro_acta = " & mActa & _
" AND manifestaciones_vivavoz.version_acta = " & Str(xUltimaVersionActa) + 1 & _
" ORDER BY Legisladores.apellido, Legisladores.nombre"
SetearRs s, rs
If rs.EOF Then
    Exit Sub
End If
rpt.DataControl1.Recordset = rs
rAudit.Sections("GroupFooter3").Controls("SubReportVivavoz").Object = rpt
End Sub
Public Function getUltimaEdicion() As String
Dim ret As String
Dim rs As New ADODB.Recordset
Dim s As String
ret = ""
s = ""
s = "SELECT TOP 1 CONVERT(varchar, ultima_edicion, 103) AS fecha, CONVERT(varchar(5), ultima_edicion, 108) AS hora " & _
" FROM manifestaciones_vivavoz WHERE " & _
" manifestaciones_vivavoz.periodo = '" & mPeriodo & "'" & _
" AND manifestaciones_vivavoz.sesion = " & mSesion & _
" AND manifestaciones_vivavoz.nro_acta = " & mActa & _
" AND manifestaciones_vivavoz.version_acta = " & mVersion & _
" ORDER BY manifestaciones_vivavoz.ultima_edicion DESC"
SetearRs s, rs
If rs.EOF Then
    getUltimaEdicion = ""
    Exit Function
End If
ret = "Manifestaciones actualizadas al día " & rs.Fields("fecha") & " a las " & rs.Fields("hora") & "HS."
getUltimaEdicion = ret
End Function
Public Sub Reportear()
    Dim xRs As Recordset
    Dim CantAfirmativos As Integer
    Dim CantNegativos As Integer
    Dim CantAbstenciones As Integer
    Dim CantAusentes As Integer
    Dim CantidadPaginas As Integer
    Dim CantidadPaginasAuditoria As Integer
    Dim RsTemp As ADODB.Recordset
    Dim RsTemp2 As ADODB.Recordset
    If strTipoOperacion = "votnum" Then
        EsNum = True
        Call imprimirUnActa(strTipoOperacion, mPeriodo, mSesion, mActa, mVersion)
        EsNum = False
    Else
        Dim consulta As String
        Set RsTemp = New ADODB.Recordset
        consulta = "SELECT * FROM actas WHERE Período_Legislativo = '" & mPeriodo & "' AND Sesión = " & mSesion & " AND Número_de_Acta = " & mActa & " AND Ultima_Versión_Acta > 0"
        SetearRs consulta, RsTemp
        If RsTemp.EOF Then
            TotalPaginas = -1
            Tipo_PreActa = "consulta"
            imprimirActaFiltrada strTipoOperacion, mPeriodo, mSesion, mActa, mVersion, 1, 0, 0, 0, False
            Tipo_PreActa = ""
        Else
            Acta_Titulo = ""
            Acta_SubTitulo = ""
            Acta_Numero = ""
            Acta_Version = ""
            Acta_Miembros = ""
            Acta_Presidente = ""
            Acta_TipoMayoria = ""
            Acta_BaseMayoria = ""
            Acta_Fecha = ""
            Acta_Hora = ""
            Acta_TipoQuorum = ""
            Acta_ResultadoVotacion = ""
            Set RsTemp2 = New ADODB.Recordset
            SetearRs "SELECT Observaciones,Fecha_Modificacion,Hora_Modificacion FROM actas WHERE " & "(Actas.Período_Legislativo='" & mPeriodo & "') AND (Actas.Sesión=" & mSesion & ") AND (Actas.Número_de_Acta=" & mActa & ") AND (Actas.Versión_Acta=" & mVersion & ")", RsTemp2
            If RsTemp2.EOF Then
                Acta_Observaciones = "Ninguna."
                Acta_FechaModificacion = ""
                Acta_HoraModificacion = ""
            Else
                Acta_Observaciones = Trim(RsTemp2.Fields("Observaciones"))
                Acta_FechaModificacion = Trim(RsTemp2.Fields("Fecha_Modificacion"))
                Acta_HoraModificacion = Trim(RsTemp2.Fields("Hora_Modificacion"))
            End If
            RsTemp2.Close
            Set RsTemp2 = Nothing
            Set TodoElReporte = New ActiveReport
            CantidadPaginasAuditoria = ObtenerCantidadAuditoria(strTipoOperacion, mPeriodo, mSesion, mActa, mVersion, 1, 0, 0, 0, False)
            CantidadPaginas = ObtenerCantidadDePaginas(strTipoOperacion, mPeriodo, mSesion, mActa, mVersion)
            TotalPaginas = CantidadPaginas + CantidadPaginasAuditoria
            Paginas_Fijas_Auditoria = TotalPaginas
            Paginas_Auditoria = Paginas_Fijas_Auditoria
            Call frmConsultarActa.imprimirUnActaCompleta(strTipoOperacion, mPeriodo, Val(Str(mSesion)), Val(Str(mActa)), mVersion) 'Val(Str(xNroActaActual)), 0)
            Call TodoElReporte.Pages.Remove(TodoElReporte.Pages.Count - 1)
            TodoElReporte.Pages.Commit
            Dim nTick As Long
            nTick = GetTickCount
            While GetTickCount - nTick < 1500
                DoEvents
            Wend
            imprimirActaAuditoria strTipoOperacion, mPeriodo, mSesion, mActa, mVersion, 1, 0, 0, 0, False
            nTick = GetTickCount
            While GetTickCount - nTick < 1500
                DoEvents
            Wend
            TodoElReporte.Pages(TodoElReporte.Pages.Count - 1).Width = TodoElReporte.Pages(TodoElReporte.Pages.Count - 1).Width - 2000
            TodoElReporte.Pages.Commit
            nTick = GetTickCount
            While GetTickCount - nTick < 1500
                DoEvents
            Wend
            Dim X As Integer
            For X = 0 To TodoElReporte.Pages.Count - 1
                TodoElReporte.Pages(X).Width = 300
                TodoElReporte.Pages.Commit
            Next X
            TodoElReporte.Pages.Commit
            TodoElReporte.Run False
            TodoElReporte.PrintReport False
        End If
        RsTemp.Close
        Set RsTemp = Nothing
    End If
End Sub
Public Sub imprimirUnActa(strTipoOperacion As String, pPeriodo As String, pSesion As Integer, pActa As Integer, pVersion As Integer)
   ' 'On Error GoTo TrapError
    'IMPRESION ACTA 2009

    If PermisosTotales.ConsultaActas = 0 Then
        MsgBox "No posee permisos para consulta de actas", vbInformation + vbOKOnly, "Usuario sin permisos!!"
        Exit Sub
    End If

    Dim m_Report As New rptActas
    Dim rstActa  As New ADODB.Recordset
    'Dim fViewer  As frmVisor
    Dim sql      As String
    Dim sql_voto_afirmativo_presidente As String
    Dim sql_voto_negativo_presidente As String
    Dim sql_voto_abstencion_presidente As String
    Dim rsPresi As ADODB.Recordset
    Dim PresidenteVotosAfirmativos As Integer
    Dim PresidenteVotosNegativos As Integer
    Dim PresidenteAbstenciones As Integer
    Dim IDPresidente As Integer
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    SetearRs "SELECT Presidente FROM actas WHERE " & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ")", RsTemp
    If RsTemp.EOF Then
        IDPresidente = 0
    Else
        If Trim(RsTemp.Fields(0)) = "" Then
            IDPresidente = 0
        Else
            IDPresidente = Val(Trim(RsTemp.Fields(0)))
        End If
    End If
    RsTemp.Close
    Set RsTemp = Nothing
    PresidenteVotosAfirmativos = 0
    PresidenteVotosNegativos = 0
    PresidenteAbstenciones = 0
    Set rsPresi = New ADODB.Recordset
    SetearRs "SELECT resultado_voto_presidente FROM Actas WHERE presidente_habilitado_votar = 1 AND " & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ")", rsPresi
    If rsPresi.EOF Then 'Reviso si votó el presidente, y qué votó
        PresidenteVotosAfirmativos = 0
        PresidenteVotosNegativos = 0
        PresidenteAbstenciones = 0
    Else
        Select Case Trim(LCase(rsPresi.Fields(0)))
        Case "s"
            PresidenteVotosAfirmativos = 1
        Case "n"
            PresidenteVotosNegativos = 1
        Case ""
            PresidenteAbstenciones = 1
        End Select
    End If
    rsPresi.Close
    Set rsPresi = Nothing
    sql_voto_afirmativo_presidente = Trim(Str(PresidenteVotosAfirmativos))
    sql_voto_abstencion_presidente = Trim(Str(PresidenteAbstenciones))
    sql_voto_negativo_presidente = Trim(Str(PresidenteVotosNegativos))
    'Set fViewer = New frmVisor
    'leyenda tipo periodo legislativo. Cambiar los CASE WHEN
    If strTipoOperacion = "votnum" Then
'        sql = "SELECT * From actas " & _
'              " WHERE Período_Legislativo = '" & pPeriodo & "' " & _
'              " AND sesión = " & pSesion & _
'              " AND Número_de_Acta = " & pActa & _
'              " AND Versión_Acta =" & pVersion
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables, actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa,actas.Reunion, actas.Tipo_de_operación , Legisladores.distrito,tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
              " WHERE (Actas.Período_Legislativo='" & pPeriodo & "')" & _
              " AND (Actas.Sesión=" & pSesion & ")" & _
              " AND (Actas.Número_de_Acta=" & pActa & ")" & _
              " AND (Actas.Versión_Acta=" & pVersion & ")" & _
              " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
        
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, 0 AS Votos_Afirm_Identificables, (actas.Votos_Afirm_Identificables + actas.Votos_Afirm_No_Identificables) AS Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, O AS Votos_Neg_Identificables , (actas.Votos_Neg_Identificables + actas.Votos_Neg_No_Identificables) AS Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, " & _
              " actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa,actas.Reunion, actas.Tipo_de_operación , Legisladores.distrito,tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
              " WHERE (Actas.Período_Legislativo='" & pPeriodo & "')" & _
              " AND (Actas.Sesión=" & pSesion & ")" & _
              " AND (Actas.Número_de_Acta=" & pActa & ")" & _
              " AND (Actas.Versión_Acta=" & pVersion & ")" & _
              " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
        
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, 0 AS Votos_Afirm_Identificables, (actas.Votos_Afirm_Identificables + actas.Votos_Afirm_No_Identificables + " & _
        sql_voto_afirmativo_presidente & ") AS Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, 0 AS Votos_Neg_Identificables , (actas.Votos_Neg_Identificables + actas.Votos_Neg_No_Identificables + " & sql_voto_negativo_presidente & ") AS Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total , actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion" & _
              ", actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa,actas.Reunion, actas.Tipo_de_operación , Legisladores.distrito,tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Ordinario' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Prórroga' WHEN 'L' THEN 'Legislativo' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'P' THEN 'Preparatoria 'WHEN 'I' THEN 'Informativa' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
              " WHERE (Actas.Período_Legislativo='" & pPeriodo & "')" & _
              " AND (Actas.Sesión=" & pSesion & ")" & _
              " AND (Actas.Número_de_Acta=" & pActa & ")" & _
              " AND (Actas.Versión_Acta=" & pVersion & ")" & _
              " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
              
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, 0 AS Votos_Afirm_Identificables, (actas.Votos_Afirm_Identificables + actas.Votos_Afirm_No_Identificables + " & _
        sql_voto_afirmativo_presidente & ") AS Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, 0 AS Votos_Neg_Identificables , (actas.Votos_Neg_Identificables + actas.Votos_Neg_No_Identificables + " & sql_voto_negativo_presidente & ") AS Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total , actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion" & _
              ", actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa,actas.Reunion, actas.Tipo_de_operación , Legisladores.distrito,tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa' WHEN 'H' THEN ' - Homenajes' WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
              " WHERE (Actas.Período_Legislativo='" & pPeriodo & "')" & _
              " AND (Actas.Sesión=" & pSesion & ")" & _
              " AND (Actas.Número_de_Acta=" & pActa & ")" & _
              " AND (Actas.Versión_Acta=" & pVersion & ")" & _
              " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
              
        
        m_Report.Campo31.SummaryType = ddSMNone
        m_Report.Campo31.Text = 1
        'm_Report.Section5.Suppress = True
        m_Report.Detail.Visible = False
        'm_Report.Cuadro1.Suppress = True
        m_Report.delimitador1.Visible = False
        m_Report.delimitador2.Visible = False
        m_Report.delimitador3.Visible = False
        m_Report.delimitador4.Visible = False
        m_Report.Label2.Visible = False
        m_Report.Label3.Visible = False
        m_Report.Campo25.Visible = False
        m_Report.Campo26.Visible = False
        m_Report.Campo23.DataField = "Abstenciones_Total"
        m_Report.lblVotacion.Caption = "Votación Numérica"
        'm_Report.Texto25.Suppress = True
        m_Report.Texto25.Visible = False
        'm_Report.Línea5.Suppress = True
        m_Report.Línea5.Visible = False
        'm_Report.Texto23.Suppress = True
        m_Report.Texto23.Visible = False
        'm_Report.Texto1.Suppress = True
        m_Report.Texto1.Visible = False
        'm_Report.Línea6.Suppress = True
        m_Report.Línea6.Visible = False
        'm_Report.Línea7.Suppress = True
        m_Report.Línea7.Visible = False
        'm_Report.Línea8.Suppress = True
        'm_Report.Línea8.Visible = False
        'm_Report.Línea9.Suppress = True
        'm_Report.Línea9.Visible = False
        'm_Report.Texto22.Suppress = True
        m_Report.Texto22.Visible = False
                        
        With m_Report
            '.Campo14.Visible = False
            '.Campo18.Visible = False
            '.Campo22.Visible = False
            '.Texto18.Visible = False
            .Texto19.Caption = "Diputados"
            .lblPresidente.Visible = False
            .vap.Visible = False
            .vnp.Visible = False
            .vabsp.Visible = False
        End With
    Else
        'version sin voto presidente
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables, actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
              
        'version con voto presidente, pero sin restar de la columna de identificados
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables, actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                " AND actas.Sesión = " & _
                " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
    
        'Version restando el voto de la columna identificados
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - " & sql_voto_afirmativo_presidente & " as Votos_Afirm_Identificables " & _
                ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables - " & sql_voto_negativo_presidente & " as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables - " & sql_voto_abstencion_presidente & " as Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                " AND actas.Sesión = " & _
                " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
    
'        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - " & sql_voto_afirmativo_presidente & " as Votos_Afirm_Identificables " & _
                ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables - " & sql_voto_negativo_presidente & " as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables - " & sql_voto_abstencion_presidente & " as Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                " AND actas.Sesión = " & _
                " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (RTrim(DetalleActas.Resultado) <> 'AUSENTE') AND (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
                
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - 0 as Votos_Afirm_Identificables " & _
                ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, actas.Votos_Neg_Identificables - 0 as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total, actas.Abstenciones_Identificables - " & sql_voto_abstencion_presidente & " as Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total + " & sql_voto_abstencion_presidente & " AS Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + " & _
                "(CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa' WHEN 'H' THEN ' - Homenajes' WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                " AND actas.Sesión = " & _
                " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (RTrim(DetalleActas.Resultado) <> 'AUSENTE') AND (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "

        m_Report.Campo23.DataField = "Abstenciones_Identificables"
        m_Report.Campo15.DataField = "Votos_Afirm_Identificables"
        m_Report.Campo19.DataField = "Votos_Neg_Identificables"
        m_Report.Campo24.DataField = "Abstenciones_Identificables"
        
'NUEVO

'        sql = " SELECT " & TotalPaginas & " AS Ultima_Pagina, DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - 0 as Votos_Afirm_Identificables " & _
'                ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, actas.Votos_Neg_Identificables - 0 as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total, actas.Abstenciones_Identificables + '" & sql_voto_abstencion_presidente & "' as Abstenciones_Identificables, actas.Abstenciones_Identificables AS Absten_Sin_Presi, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total + " & sql_voto_abstencion_presidente & " AS Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
'                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
'                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
'                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
'                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
'                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
'                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
'                " AND actas.Sesión = " & _
'                " detalleactas.Sesión AND DetalleActas.Legislador_asignado <> '" & IDPresidente & "' AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
'                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (RTrim(DetalleActas.Resultado) <> 'AUSENTE') AND (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
        sql = " SELECT " & TotalPaginas & " AS Ultima_Pagina, DetalleActas.Grupo_Politico, DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - 0 as Votos_Afirm_Identificables " & _
                ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, actas.Votos_Neg_Identificables - 0 as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total, actas.Abstenciones_Identificables + '" & sql_voto_abstencion_presidente & "' as Abstenciones_Identificables, actas.Abstenciones_Identificables AS Absten_Sin_Presi, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total + " & sql_voto_abstencion_presidente & " AS Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + " & _
                "(CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'H' THEN 'Homenajes' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                " AND actas.Sesión = " & _
                " detalleactas.Sesión AND DetalleActas.Legislador_asignado <> '" & IDPresidente & "' AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (RTrim(DetalleActas.Resultado) <> 'AUSENTE') AND (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
        m_Report.Campo23.DataField = "Absten_Sin_Presi"
        m_Report.Campo15.DataField = "Votos_Afirm_Identificables"
        m_Report.Campo19.DataField = "Votos_Neg_Identificables"
        m_Report.Campo24.DataField = "Abstenciones_Identificables"
        
        
    End If
    If Trim(txtTipoOperacion.Text) = "Pase de Lista" Or strTipoOperacion = "paslis" Then
        Const Corrimiento As Integer = 3250
        With m_Report
        .lblVotacion.Caption = "Pase de Lista"
        .Texto12.Caption = "Resultado :"
        'Se eliminan todos los textos
        '.Texto18.Visible = False
        .Texto19.Visible = False
        .Texto20.Visible = False
        .Texto21.Visible = False
        .Texto13.Visible = False
        .Texto14.Visible = False
        .Texto15.Visible = False
        .Texto2.Visible = False
        .Texto8.Visible = False
        .Texto7.Visible = False
        '.Label1.Visible = False
        'Se eliminan todos los campos con Datafields
        '.Campo14.Visible = False
        .Campo15.Visible = False
        .Campo16.Visible = False
        .Campo17.Visible = False
        '.Campo18.Visible = False
        .Campo19.Visible = False
        .Campo20.Visible = False
        .Campo21.Visible = False
        '.Campo22.Visible = False
        .Campo23.Visible = False
        .Campo24.Visible = False
        .Campo33.Visible = False
        .Campo24.Visible = False
        .Campo33.Visible = False
        .Campo30.Visible = False
        .Campo35.Visible = False
        .Campo6.Visible = False
        'Se elimina el delimitador
        .Line7.Visible = False
        'Se acomodan los controles
        .Label2.Left = .Label2.Left + Corrimiento
        .Label3.Left = .Label3.Left + Corrimiento
        .Label4.Left = .Label4.Left + Corrimiento
        .Texto16.Left = .Texto16.Left + Corrimiento
        .Texto17.Left = .Texto17.Left + Corrimiento
        .Campo25.Left = .Campo25.Left + Corrimiento
        .Campo26.Left = .Campo26.Left + Corrimiento
        .Campo27.Left = .Campo27.Left + Corrimiento
        .Campo28.Left = .Campo28.Left + Corrimiento
        .lblPresidente.Visible = False
        .vap.Visible = False
        .vnp.Visible = False
        .vabsp.Visible = False
        End With
    End If
    SetearRs sql, rstActa
    'm_Report.Database.SetDataSource rstActa
    Set m_Report.DataControl1.Recordset = rstActa
    If True Then
        If strTipoOperacion = "votnum" Then
            m_Report.Campo31.Text = "1"
            m_Report.Run False
            'm_Report.Pages(0).Width = m_Report.Pages(0).Width - 1000
            m_Report.Pages.Remove (1)
            m_Report.Pages.Commit
        Else
            m_Report.Run False
        End If
        Dim X As Integer
        For X = 0 To m_Report.Pages.Count - 1
            'm_Report.Pages(X).Width = 300
            m_Report.Pages(X).Width = m_Report.Pages(X).Width - 1000
            m_Report.Pages.Commit
        Next X
        If VistaPrevia = True Then
            m_Report.Show vbModal
        Else
            If strTipoOperacion <> "votnum" Then
                m_Report.Run False
            End If
            Dim tempTick As Long
            tempTick = GetTickCount
            While GetTickCount - tempTick < 1000
                DoEvents
            Wend
            If Not m_Report.Printer.Status = ddJSPrinting Then
                If Tipo_PreActa = "consulta" Then
                    m_Report.PrintReport True
                Else
                    m_Report.PrintReport False
                End If
            End If
        End If
    Else
        'm_Report.Printer.Copies = 1
        'm_Report.Printer.StartJob "acta"
        m_Report.PrintReport False
    End If
'    fViewer.CRViewer1.ReportSource = m_Report
'    If PermisosTotales.ImprimeActas = 1 Then
'        fViewer.CRViewer1.EnablePrintButton = True
'    Else
'        fViewer.CRViewer1.EnablePrintButton = False
'    End If
'
'    fViewer.CRViewer1.ViewReport
'    fViewer.CRViewer1.Zoom 100
'    fViewer.Show vbModal
    
    Set rstActa = Nothing
    'Set fViewer = Nothing
    Set m_Report = Nothing
Exit Sub
TrapError:
    Select Case err.Number
        Case Else
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            End
    End Select
Return
End Sub

Private Sub ExportarAcaVotacionNumerica()
    'On Error GoTo TrapError
    Dim strTipoOperacion As String
    Dim strSesion        As String
    Dim strActaNro       As String
    Dim strFecha         As String
    Dim strHora          As Variant
    Dim strBaseMayoria   As String
    Dim strTipoMayoria   As String
    Dim strTipoQuorum    As String
    Dim strMiembros      As String
    Dim strDesempate     As String
    Dim strPresidente    As String
    Dim strResultado     As String
    Dim x1 As Long: Dim X2 As Long: Dim x3 As Long: Dim x4 As Long
    Dim x5 As Long: Dim x6 As Long: Dim x7 As Long: Dim x8 As Long
    Dim x9 As Long: Dim x10 As Long: Dim x11 As Long: Dim x12 As Long
    Dim x13 As Long: Dim x14 As Long: Dim x15 As Long: Dim x16 As Long
    Dim x17 As Long: Dim x18 As Long: Dim x19 As Long
    Dim strArchivo As String
    Dim xFile As Long
    

    strTipoOperacion = txtTipoOperacion.Text
    strSesion = IIf(txtSesion.Text = "0", "-1", txtSesion.Text)
    strActaNro = txtNroActa.Text
    strFecha = mId(txtFecha.Text, 1, 10)
    strHora = mId(txtFecha.Text, 11, 13)
    strBaseMayoria = txtBase.Text
    strTipoMayoria = txtTipoMayoria.Text
    strTipoQuorum = txtTipoQuorum.Text
    strMiembros = txtMiembros.Text
    strDesempate = txtDesempate.Text
    strPresidente = txtNombrePresidente.Text
    strResultado = txtVotacion.Text
    
    
    x1 = Trim(Str(xVotosAfirmativosIdentificables))
    X2 = Trim(Str(xVotosAfirmativosNoIdentificables))
    x3 = Trim(Str(xVotosAfirmativosDesempate))
    x4 = Trim(Str(xVotosAfirmativosTotal))
    x5 = Trim(Str(xVotosNegativosIdentificables))
    x6 = Trim(Str(xVotosNegativosNoIdentificables))
    x7 = Trim(Str(xVotosNegativosDesempate))
    x8 = Trim(Str(xVotosNegativosTotal))
    x9 = Trim(Str(xAbstencionesIdentificables))
    x12 = Trim(Str(xAbstencionesTotal))
    
    x13 = Trim(Str(xPresentesIdentificables))
    x14 = Trim(Str(xPresentesNoIdentificables))
    x16 = Trim(Str(xPresentesTotal))
    Dim txtExportarAscii
    Call ObtenerLogoPL ' Obtener cadena RTF con logo de la legislatura
    With txtExportarAscii
        .Text = ""
        .Text = .Text & "Tipo de Operacion: " & strTipoOperacion & vbCrLf
        .Text = .Text & "Acta Nro: " & strActaNro & vbCrLf
        .Text = .Text & "Fecha: " & strFecha & vbCrLf
        .Text = .Text & "Hora : " & strHora & vbCrLf
        .Text = .Text & "Base Mayoría : " & strBaseMayoria & vbCrLf
        .Text = .Text & "Tipo Mayoría: " & strTipoMayoria & vbCrLf
        .Text = .Text & "Tipo Quorum: " & strTipoQuorum & vbCrLf
        .Text = .Text & "Miembros del Cuerpo: " & strMiembros & vbCrLf
        .Text = .Text & "Desempate: " & strDesempate & vbCrLf
        .Text = .Text & "Presidente: " & strPresidente & vbCrLf
        .Text = .Text & "Resultado: " & strResultado & vbCrLf

        .Text = .Text & "                                        Identificados    S/Identif    Desempate      Total " & vbCrLf
        .Text = .Text & "Votos Afirmativos:                    " & Format(x1, "00") & "                  " & Format(X2, "00") & "                 " & Format(x3, "00") & "                 " & Format(x4, "00") & vbCrLf
        .Text = .Text & "Votoso Negativos:                   " & Format(x5, "00") & "                  " & Format(x6, "00") & "                 " & Format(x7, "00") & "                 " & Format(x8, "00") & vbCrLf
        .Text = .Text & "Abstenciones:                          " & Format(x9, "00") & "                  " & Format(x10, "00") & "                " & x11 & "                " & Format(x12, "00") & vbCrLf
        .Text = .Text & "Presentes:                               " & Format(x13, "00") & "                 " & Format(x14, "00") & "                " & Format(x15, "00") & "                " & Format(x16, "00") & vbCrLf
        .Text = .Text & "Ausentes:                                 " & Format(x17, "00") & "                 " & Format(x18, "00") & "                " & Format(x19, "00") & "                " & vbCrLf
    End With
    
    ' txtExportarAscii.TextRTF = strCadenaRtfLogo & txtExportarAscii.TextRTF
    ' terminar de dar formato RTF a las actas!
    
    xFile = FreeFile
    strArchivo = strActaNro & "_" & strFecha
    strArchivo = Replace(strArchivo, "/", "_")
    strArchivo = Replace(strArchivo, " ", "_")
    strArchivo = Replace(strArchivo, ":", "-")
    strArchivo = Replace(strArchivo, ".", "-")
    strArchivo = App.Path & "\actas\acta_nro" & strArchivo
    strArchivo = strArchivo & ".rtf"
    
    Open strArchivo For Binary As #xFile
        Put #xFile, , txtExportarAscii.TextRTF
    Close #xFile
    MsgBox "Acta generada!", vbInformation + vbOKOnly

Exit Sub
TrapError:
    Select Case err.Number
        Case 76
            If MsgBox("Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source & vbCrLf & "Verifique que existe la carpeta necesaria para almacenar " & strArchivo, vbQuestion + vbYesNo, "Reintentar?") = vbYes Then
                Resume
            Else
                Exit Sub
            End If
        Case Else
            If MsgBox("Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source, vbQuestion + vbYesNo, "Reintentar?") = vbYes Then
                Resume
            Else
                Exit Sub
            End If
    End Select
End Sub

Private Sub ObtenerLogoPL()
    ' --------------------------------------------------------------------------------
    ' Obtener string RTF con imagen de logo de la legislatura
    ' --------------------------------------------------------------------------------
    Dim strArchivoLogo As String
    Dim xFile          As Long
    strArchivoLogo = App.Path & "\logo_pl_rtf.dat"
    xFile = FreeFile
    Open strArchivoLogo For Binary As #xFile
        strCadenaRtfLogo = Space(LOF(xFile))
        Get #xFile, , strCadenaRtfLogo
    Close #xFile
End Sub

Private Sub cmdVivavoz_Click()
Dim s As String
s = "SELECT * FROM manifestaciones_vivavoz WHERE periodo = '" & mPeriodo & "' " & _
    " AND sesion = " & mSesion & " AND nro_acta = " & mActa & _
    " AND version_acta = " & mVersion
Dim rs As New Recordset
Call SetearRs(s, rs)
While Not rs.EOF
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
Dim f As New frmManifestacionesVivavoz
f.mPeriodo = mPeriodo
f.mSesion = mSesion
f.mActa = mActa
f.mVersion = Str(xUltimaVersionActa + 1)
If (mVivavozEntered = False) Then
    f.mCargarOriginales = True
Else
    f.mCargarOriginales = False
End If
f.Show vbModal, Me
mVivavozChanged = f.mChanged
If (f.mChanged = True) Then
    mVivavozEntered = True
    cmdAceptar.Enabled = True
Else
    If (mCambios = False) Then
        cmdAceptar.Enabled = False
    End If
End If
End Sub

Private Sub Command1_Click()

    If strTipoOperacion = "votnum" Then
        Call ExportarAcaVotacionNumerica
    End If


End Sub

Private Sub Command2_Click()
MsgBox txtFecha.Text
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
mVivavozChanged = False
mVivavozEntered = False
establecerPermisos
ponerTitulosGrilla
' Centrar frame
'With frameActas
'    .Left = (Screen.Width - .Width) / 2
'    .Top = (Screen.Height - .Height) / 2
'End With
End Sub
Private Sub MostrarCuentas()
    
    
    txtAfirmativosId.Text = Trim(Str(xVotosAfirmativosIdentificables))
    txtAfirmativosTotal.Text = Trim(Str(xVotosAfirmativosTotal))
    txtAfirmativosNoId.Text = Trim(Str(xVotosAfirmativosNoIdentificables))
    txtAfirmativosDesempate.Text = Trim(Str(xVotosAfirmativosDesempate))
    
    
    txtNegativoID.Text = Trim(Str(xVotosNegativosIdentificables))
    txtNegativoTotales.Text = Trim(Str(xVotosNegativosTotal))
    txtNegativoNoId.Text = Trim(Str(xVotosNegativosNoIdentificables))
    txtNegativoDesempate.Text = Trim(Str(xVotosNegativosDesempate))
    
    txtAbstencionesId.Text = Trim(Str(xAbstencionesIdentificables))
    txtAbstencionesTotales.Text = Trim(Str(xAbstencionesTotal))
    
    txtPresentesId.Text = Trim(Str(xPresentesIdentificables))
    txtPresentesNoId.Text = Trim(Str(xPresentesNoIdentificables))
    txtPresentesTotal.Text = Trim(Str(xPresentesTotal))
    
    
End Sub

Private Function CalcularMinimoParaQuorum(strTipoMayoria As String, xMiembrosDelCuerpo As Long)
    strTipoMayoria = Trim(LCase(strTipoMayoria))
    ' variable que permite calcular el minimo necesario para obtener quorum
    CalcularMinimoParaQuorum = IIf(strTipoMayoria = "man", 1, Fix(xMiembrosDelCuerpo / 2) + IIf(strTipoMayoria = "121", 1, 0))
    If xMiembrosDelCuerpo Mod 2 = 1 Then
        CalcularMinimoParaQuorum = CalcularMinimoParaQuorum + 1
    Else
        CalcularMinimoParaQuorum = CalcularMinimoParaQuorum
    End If
End Function


Private Sub MostrarDatosSesion()
    'On Error GoTo TrapError
    
    Dim strSql As String
    strSql = "SELECT actas.*, " _
        & " tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo,Actas.Reunion AS Reunion_Acta, basemay.Descripcion AS descBaseMay, " _
        & " tipmay.Descripcion AS descTipoMay, rtrim(Legisladores.apellido) + ', ' + rtrim(legisladores.nombre) AS Legislador, Actas.Tipo_de_Quorum " _
        & " FROM Legisladores RIGHT OUTER JOIN actas ON Legisladores.id = actas.Presidente LEFT OUTER JOIN " _
        & " tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN " _
        & " basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON " _
        & " actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT Outer Join tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes " _
        & " WHERE (Período_Legislativo='" & mPeriodo & "') AND (Sesión=" & mSesion & ") AND (Número_de_Acta=" & mActa & ") AND (Versión_Acta=" & mVersion & ") "
    SetearRs strSql, rstActa
    If rstActa.EOF = False Then
        With rstActa
            xUltimaVersionActa = !Ultima_Versión_Acta
            strResultadoEsperado = Trim(UCase(Trim(!Votacion)))
            strSesion = Trim(!Sesión)
            xNumeroActa = !Número_de_Acta
            xResultadoVotaPresidente = IIf(IsNull(!resultado_voto_presidente), "", !resultado_voto_presidente)
            xVersionActa = !Versión_Acta
            strTipoQuorum = Trim(!Tipo_de_Quorum)
            strTipoMayoria = Trim(!Tipo_de_Mayoria)
            strBaseMayoria = Trim(!Base_de_Mayoria)
            If IsNull(!NroOrdenDia) Then
                xNroOrdenDia = 0
            Else
                xNroOrdenDia = !NroOrdenDia
            End If
            strTipoOperacion = Trim(!Tipo_de_operación)
            strPeriodo_Legislativo = Trim(!Período_Legislativo)
            strNombreActa = Trim(!Nombre_del_Acta)
            xMiembrosDelCuerpo = !Miembros_del_cuerpo
            xPresentesTotal = !Presentes_Total
            xVotosAfirmativosTotal = !Votos_Afirm_Total
            xVotosNegativosTotal = !Votos_Neg_Total
            xVotosAfirmativosIdentificables = !Votos_Afirm_Identificables
            xVotosAfirmativosTotal = !Votos_Afirm_Total
            xVotosAfirmativosNoIdentificables = !Votos_Afirm_No_Identificables
            xVotosAfirmativosDesempate = !Votos_Afirm_Desempate
            xVotosNegativosIdentificables = !Votos_Neg_Identificables
            xVotosNegativosTotal = !Votos_Neg_Total
            xVotosNegativosNoIdentificables = !Votos_Neg_No_Identificables
            xVotosNegativosDesempate = !Votos_Neg_Desempate
            xAbstencionesIdentificables = !Abstenciones_Identificables
            xAbstencionesNoIdentificables = !Abstenciones_No_Identificables
            xAbstencionesTotal = !Abstenciones_Total
            xPresentesIdentificables = !Presentes_Identificables
            xPresentesNoIdentificables = !Presentes_No_Identificables
            xPresentesTotal = !Presentes_Total
            If (!vota_presidente = 1) Then
                blEsLegislador = True
            Else
                blEsLegislador = False
            End If
            If IsNull(!descTipoOp) = False Then
                txtTipoOperacion.Text = !descTipoOp
            End If
            If IsNull(!Sesión) = False Then
                txtSesion.Text = IIf(!Sesión = "-1", "0", !Sesión)
            End If
            If IsNull(!Número_de_Acta) = False Then
                txtNroActa.Text = !Número_de_Acta
            End If
            txtReunion.Text = !Reunion_Acta
            If IsNull(!Versión_Acta) = False Then
                txtVersion.Text = !Versión_Acta
                If !Ultima_Versión_Acta = 0 Then
                    txtVersion.Text = "Original"
                Else
                   If !Versión_Acta = 0 Then
                        txtVersion.Text = "Ult.Mod.Ver. " & Val(!Ultima_Versión_Acta) + 1
                   Else
                        txtVersion.Text = "Ver. " & Val(!Ultima_Versión_Acta) + 1
                   End If
                End If
                txtVersion.Tag = !Versión_Acta
            End If
            If IsNull(!Nombre_del_Acta) = False Then
                txtNombre.Text = Trim(!Nombre_del_Acta)
            End If
            If IsNull(!fecha) = False Then
                txtFecha.Text = !fecha
                If IsNull(!hora) = False Then
                    If Trim(!hora) = "" Then
                        txtHora.Text = Format(!fecha, "HH:MM")
                    Else
                        txtHora.Text = !hora
                    End If
                Else
                    txtHora.Text = Format(!fecha, "HH:MM")
                End If
            End If
            If IsNull(!descTipoMayQuo) = False Then
                txtTipoQuorum.Text = !descTipoMayQuo
            End If
            If IsNull(!Miembros_del_cuerpo) = False Then
                txtMiembros.Text = !Miembros_del_cuerpo
            End If
            If IsNull(!Desempate) = False Then
                txtDesempate.Text = !Desempate
            End If
            If IsNull(!descTipoMay) = False Then
                txtTipoMayoria.Text = !descTipoMay
            End If
            If IsNull(!descBaseMay) = False Then
                txtBase.Text = !descBaseMay
            End If
            If IsNull(!Votacion) = False Then
                txtVotacion.Text = !Votacion
                If Trim(txtVotacion.Text) = "ANULADA" Then
                    cmdAnularActa.Caption = "RESTAURAR RESULTADO"
                End If
            End If
            If IsNull(!Presidente) = False Then
                txtCodigoPresidente.Text = !Presidente
            End If
            If IsNull(!legislador) = False Then
                txtNombrePresidente.Text = !legislador
            End If
            If IsNull(!Observaciones) = False Then
                txtObservaciones.Text = Trim(!Observaciones)
            End If
            If IsNull(!Presentes_Identificables) = False Then
                txtPresentesId.Text = !Presentes_Identificables
            Else
                txtPresentesId.Text = "0"
            End If
            If IsNull(!Presentes_No_Identificables) = False Then
                txtPresentesNoId.Text = !Presentes_No_Identificables
            Else
                txtPresentesNoId.Text = "0"
            End If
            If IsNull(!Presentes_Total) = False Then
                txtPresentesTotal.Text = !Presentes_Total
            Else
                txtPresentesTotal.Text = "0"
            End If
            If IsNull(!Ausentes_Total) = False Then
                txtAusentesTotal.Text = !Ausentes_Total
            Else
                txtAusentesTotal.Text = "0"
            End If
            If IsNull(!Votos_Afirm_Identificables) = False Then
                txtAfirmativosId.Text = !Votos_Afirm_Identificables
            Else
                txtAfirmativosId.Text = "0"
            End If
            If IsNull(!Votos_Afirm_No_Identificables) = False Then
                txtAfirmativosNoId.Text = !Votos_Afirm_No_Identificables
            Else
                txtAfirmativosNoId.Text = "0"
            End If
            If IsNull(!Votos_Afirm_Desempate) = False Then
                txtAfirmativosDesempate.Text = !Votos_Afirm_Desempate
            Else
                txtAfirmativosDesempate.Text = "0"
            End If
            If IsNull(!Votos_Afirm_Total) = False Then
                txtAfirmativosTotal.Text = !Votos_Afirm_Total
            Else
                txtAfirmativosTotal.Text = "0"
            End If
            If IsNull(!Votos_Neg_Identificables) = False Then
                txtNegativoID.Text = !Votos_Neg_Identificables
            Else
                txtNegativoID.Text = "0"
            End If
            If IsNull(!Votos_Neg_No_Identificables) = False Then
                txtNegativoNoId.Text = !Votos_Neg_No_Identificables
            Else
                txtNegativoNoId.Text = "0"
            End If
            If IsNull(!Votos_Neg_Desempate) = False Then
                txtNegativoDesempate.Text = !Votos_Neg_Desempate
            Else
                txtNegativoDesempate.Text = "0"
            End If
            If IsNull(!Votos_Neg_Total) = False Then
                txtNegativoTotales.Text = !Votos_Neg_Total
            Else
                txtNegativoTotales.Text = "0"
            End If
            If IsNull(!Abstenciones_Identificables) = False Then
                txtAbstencionesId.Text = !Abstenciones_Identificables
            Else
                txtAbstencionesId.Text = "0"
            End If
            If IsNull(!Abstenciones_No_Identificables) = False Then
                txtAbstencionesNoId.Text = !Abstenciones_No_Identificables
            Else
                txtAbstencionesNoId.Text = "0"
            End If
            If IsNull(!Abstenciones_Total) = False Then
                If Trim(!Tipo_de_operación) = "votnom" Then
                    txtAbstencionesTotales.Text = !Abstenciones_Identificables
                Else
                    txtAbstencionesTotales.Text = !Abstenciones_Total
                End If
            Else
                txtAbstencionesTotales.Text = "0"
            End If
            If IsNull(!resultado_voto_presidente) = False Then 'Si voto
                If Trim(!resultado_voto_presidente) = "s" Then
                    txtVotoPresidente.Text = "Afirmativo"
                Else
                    txtVotoPresidente.Text = "Negativo"
                End If
            Else
                lblVotoPresidente.Visible = False
                txtVotoPresidente.Visible = False
            End If
        End With
    End If
    If strTipoOperacion <> "votnum" Then
        Call MostrarDetalleActa(rstActa!Período_Legislativo, IIf(txtSesion.Text = "0", "-1", txtSesion.Text), txtNroActa.Text, txtVersion.Tag)
        Call CargarComboResultados
    Else
        vsGrilla.Visible = False
        lblBuscarLegislador.Visible = False
        txtBuscar.Visible = False
    End If
    ControlesHabilitados = False
Exit Sub
TrapError:
    Select Case err.Number
        Case Else
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            Resume
    End Select
End Sub
Private Property Let ControlesHabilitados(vNew As Boolean)
    txtNombre.Enabled = vNew
    txtObservaciones.Enabled = vNew
    cmdAceptar.Enabled = vNew
    cmdPresidente.Enabled = vNew
    If (Trim(UCase(txtTipoOperacion)) = "VOTACIÓN NUMÉRICA") Or (Trim(UCase(txtTipoOperacion)) = "VOTACIÓN NOMINAL") Then
        vsGrilla.Enabled = vNew
    End If
End Property
Private Sub MostrarDetalleActa(pPeriodo As String, pSesion As Long, pActa As Long, pVersion As Long)
    'On Error GoTo TrapError
    Dim rstAux As New ADODB.Recordset
    Dim strSql As String
    strSql = "SELECT *, rtrim(Legisladores.Apellido) + ', '+ rtrim(Legisladores.Nombre) as legislador " _
            & " FROM detalleactas LEFT OUTER JOIN Legisladores ON detalleactas.Legislador_asignado = Legisladores.id " _
            & " WHERE (Período_Legislativo='" & pPeriodo & "') AND (Sesión=" & pSesion & ") AND (Nro_de_Acta=" & pActa & ") AND (Versión_Acta= " & pVersion & " ) " _
            & " ORDER BY Legisladores.Apellido,Legisladores.Nombre" ' Numero_de_banca"
    SetearRs strSql, rstAux
    If rstAux.EOF = False Then
        Do While Not (rstAux.EOF)
            If rstAux.Fields("estado") = 3 Then
                vsGrilla.AddItem vbTab & rstAux!Numero_de_banca & vbTab & rstAux!Legislador_asignado & vbTab & "Legislador no incorporado" & vbTab & rstAux!Resultado
            Else
                vsGrilla.AddItem vbTab & rstAux!Numero_de_banca & vbTab & rstAux!Legislador_asignado & vbTab & rstAux!legislador & vbTab & rstAux!Resultado
            End If
            
            'vsgrilla.AddItem vbTab & rstAux!Legislador_asignado & vbTab & rstAux!legislador & vbTab & rstAux!Resultado
            rstAux.MoveNext
        Loop
    Else
       MsgBox "Ha ocurrido un error al recuperar el detalle del acta.", vbInformation + vbOKOnly
    End If
    vsGrilla.RemoveItem 1
Exit Sub
TrapError:
    Select Case err.Number
        Case Else
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Origen: " & err.Source
            Resume Next
    End Select
Return
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim s As String
s = "DELETE FROM manifestaciones_vivavoz WHERE " & _
" manifestaciones_vivavoz.periodo = '" & mPeriodo & "'" & _
" AND manifestaciones_vivavoz.sesion = " & mSesion & _
" AND manifestaciones_vivavoz.nro_acta = " & mActa & _
" AND manifestaciones_vivavoz.version_acta = -1"
Call InsertSQL(s)
End Sub

Private Sub txtBuscar_GotFocus()
    Funciones.seleccionadoTxt txtBuscar
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim Col As Integer
        Dim Row As Integer
        Funciones.BuscarEnGrilla vsGrilla, 4, txtBuscar.Text, Col, Row
        If (Col <> -1) And (Row <> -1) Then
            If vsGrilla.Enabled = True Then
                vsGrilla.SetFocus
            End If
            vsGrilla.Row = Row
            vsGrilla.RowSel = Row
            vsGrilla.ColSel = 5
            If vsGrilla.RowSel <> 0 Then
                vsGrilla.TopRow = vsGrilla.RowSel
            End If
        Else
            MsgBox "No se ha encontrado el texto deseado." & Chr(13) & "Intente con otra búsqueda.", vbInformation + vbOKOnly
        End If
    End If
End Sub

Private Sub txtNombre_Change()
    mCambios = True
End Sub

Private Sub txtNombre_LostFocus()
    strNombreActa = txtNombre.Text
End Sub

Private Sub txtObservaciones_Change()
    mCambios = True
End Sub

Private Sub vsGrilla_Click()
    If cmbResultados.Visible = True Then
      cmbResultados.Visible = False
    End If
    If vsGrilla.Col = 4 Then      ' Position and size the ComboBox, then show it.
        cmbResultados.Width = vsGrilla.CellWidth
        cmbResultados.Left = vsGrilla.CellLeft + vsGrilla.Left
        cmbResultados.Top = vsGrilla.CellTop + vsGrilla.Top
        cmbResultados.Text = vsGrilla.Text
        mValorAnteriorVoto = Trim(UCase(vsGrilla.Text))
        cmbResultados.Visible = True
    End If
    Clipboard.Clear
    Clipboard.SetText ""
    Clipboard.SetText vsGrilla.TextMatrix(vsGrilla.RowSel, 3)
End Sub
Private Function CalculoResultado(pBase_de_Mayoria As String, _
                                     pTipo_Mayoria As String, _
                              pMiembros_del_cuerpo As Long, _
                                        pPresentes As Long, _
                                      pAfirmativos As Long, _
                                        pNegativos As Long, _
                               Optional pResultado As String, _
                          Optional pMin_Afirmativa As Long, _
                                   Optional pResto As Long, _
                Optional pMin_p_afirmativa_Calculo As Long, _
                          Optional pVotoPresidente As String, _
                        Optional pVotaElPresidente As Long) As String
    'On Error GoTo TrapError
    
    Dim rs                    As New ADODB.Recordset
    Dim strSql                As String
    Dim pAuxMinParaAfirmativa As Long
    Dim xNumerador            As Long
    Dim xDenominador          As Long
    Dim xVotosEmitidos        As Long
    Dim xResto                As Long
    Dim xBase_para_Mayoria    As Long

    pTipo_Mayoria = Trim(LCase(pTipo_Mayoria))

    strSql = "SELECT * From tipmay WHERE Tipo_de_Mayoria = '" & pTipo_Mayoria & "'"
    Call SetearRs(strSql, rs)

    xVotosEmitidos = pAfirmativos + pNegativos
    xNumerador = rs.Fields("Numerador").Value ' NUMERADOR DE LA TABLA
    xDenominador = rs.Fields("Denominador").Value  ' MODIFICAR POR DENOMINADOR DE LA TABLA
    xBase_para_Mayoria = IIf(pBase_de_Mayoria = "legpre", pPresentes, IIf(pBase_de_Mayoria = "miecue", pMiembros_del_cuerpo, IIf(pBase_de_Mayoria = "votemi", xVotosEmitidos, 0)))
    xResto = xBase_para_Mayoria * xNumerador Mod xDenominador
    pAuxMinParaAfirmativa = Fix(xBase_para_Mayoria * xNumerador / xDenominador)
    pMin_p_afirmativa_Calculo = IIf(xResto > 0, pAuxMinParaAfirmativa + 1, pAuxMinParaAfirmativa)

    If pAfirmativos > pMin_p_afirmativa_Calculo Then
         CalculoResultado = "AFIRMATIVO"
    Else
        If pAfirmativos < pMin_p_afirmativa_Calculo Then
            If (pAfirmativos = pNegativos And pAfirmativos + 1 = pMin_p_afirmativa_Calculo) Then ' Caso 2
                CalculoResultado = "EMPATE"
            Else
                CalculoResultado = "NEGATIVO"
            End If
        Else ' (equivale a la condicion pAfirmativos == pMin_p_afirmativa_Calculo,)
            If xResto > 0 Then
                If rs.Fields("Rdo_si_Af_igual_Min_y_Resto_mayor_0").Value = "A" Then
                    CalculoResultado = "AFIRMATIVO"
                Else
                    If rs.Fields("Rdo_si_Af_igual_Min_y_Resto_mayor_0").Value = "E" Then ' Caso 1
                        If pAfirmativos = pNegativos Then
                            CalculoResultado = "EMPATE"
                        Else
                            CalculoResultado = "AFIRMATIVO"
                        End If
                    Else
                        CalculoResultado = "NEGATIVO"
                    End If
                End If
            Else '(equivale a la condicion resto = 0)
                If rs.Fields("Rdo_si_Af_igual_Min_y_Resto_igual_0").Value = "A" Then
                    CalculoResultado = "AFIRMATIVO"
                Else
                    If rs.Fields("Rdo_si_Af_igual_Min_y_Resto_igual_0").Value = "E" Then
                        If pAfirmativos = pNegativos Then
                            CalculoResultado = "EMPATE"
                        Else
                            CalculoResultado = "AFIRMATIVO"
                        End If
                    Else
                        CalculoResultado = "NEGATIVO"
                    End If
                End If
            End If
        End If
    End If
    If pAfirmativos = 0 And pNegativos = 0 Then
         CalculoResultado = "NEGATIVO"
    End If
    pResto = xResto
    pMin_p_afirmativa_Calculo = pMin_p_afirmativa_Calculo + IIf(LCase(rs.Fields("Rdo_si_Af_igual_Min_y_Resto_mayor_0").Value) = "n", 1, 0)

    ' Control caso vota el presidente
    If CalculoResultado = "EMPATE" Then
        If pVotaElPresidente > 0 Then
            If (pVotoPresidente = "s") Then
                CalculoResultado = "AFIRMATIVO"
            End If
            If (pVotoPresidente = "n") Then
                CalculoResultado = "NEGATIVO"
            End If
        End If
    End If
Exit Function
TrapError:
    Select Case err.Number
        Case Else
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            Resume
    End Select
    CalculoResultado = Trim(CalculoResultado)
End Function

Private Sub CargarComboResultados()
    With cmbResultados
        If (Trim(UCase(txtTipoOperacion)) = "VOTACIÓN NUMÉRICA") Or (Trim(UCase(txtTipoOperacion)) = "VOTACIÓN NOMINAL") Then
            .AddItem "ABSTENCION", 0
            .AddItem "AFIRMATIVO", 1
            .AddItem "AUSENTE", 2
            .AddItem "NEGATIVO", 3
            mPaseLista = False
        Else
            mPaseLista = True
            .AddItem "AUSENTE", 0
            .AddItem "PRESENTE", 1
        End If
    End With
    vsGrilla.RowHeightMin = cmbResultados.Height
    cmbResultados.Visible = False
End Sub
Private Sub ponerTitulosGrilla()
    With vsGrilla
        .TextMatrix(0, 1) = "Banca"
        .TextMatrix(0, 2) = "Legislador"
        .TextMatrix(0, 3) = "Apellido y Nombre"
        .TextMatrix(0, 4) = "Resultado"
        .ColWidth(0) = 0
        .ColWidth(1) = 0 '1500 para que aparezca la columna banca
        .ColWidth(2) = 0
        .ColWidth(3) = 6000
        .ColWidth(4) = 1700
        .ColWidth(5) = 0 'id
    End With
End Sub
Private Sub establecerPermisos()
    Select Case gTipoUsuario
        Case 0, 1
            cmdMOdificar.Enabled = True
        Case Else
            cmdMOdificar.Enabled = False
    End Select
End Sub
Public Sub MostrarDatos(pActa As Integer, pPeriodo As String, pSesion As Integer, pVersion As Integer)
    mActa = pActa
    mPeriodo = pPeriodo
    mSesion = pSesion
    mVersion = pVersion
    MostrarDatosSesion
    mCambios = False
End Sub
Public Sub imprimirActaFiltrada(strTipoOperacion As String, pPeriodo As String, pSesion As Integer, pActa As Integer, pVersion As Integer, CantidadAbstenciones As Integer, CantidadAfirmativos As Integer, CantidadNegativos As Integer, CantidadAusentes As Integer, SonCopias As Boolean)
'    'On Error GoTo TrapError
    'IMPRESION ACTA 2009
    Dim Vuelta1 As Integer
    Dim VueltasTotales As Integer
    Dim Contador As Integer
    Dim Abstenciones_Procesadas As Boolean
    Dim Afirmativos_Procesados As Boolean
    Dim Negativos_Procesados As Boolean
    Dim Ausencias_Procesadas As Boolean
    Dim ImprimirActa As Boolean
    Dim TipoTitulo As String
    Dim IDPresidente As Integer
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    SetearRs "SELECT Presidente FROM actas WHERE " & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ")", RsTemp
    If RsTemp.EOF Then
        IDPresidente = 0
    Else
        If Trim(RsTemp.Fields(0)) = "" Then
            IDPresidente = 0
        Else
            IDPresidente = Val(Trim(RsTemp.Fields(0)))
        End If
    End If
    RsTemp.Close
    Set RsTemp = Nothing
    Abstenciones_Procesadas = False
    Afirmativos_Procesados = False
    Negativos_Procesados = False
    Ausencias_Procesadas = False
    Contador = 0
    VueltasTotales = CantidadAbstenciones + CantidadAfirmativos + CantidadNegativos + CantidadAusentes
    For Vuelta1 = 1 To VueltasTotales
        ImprimirActa = True
        Dim filtro As String 'filtro para los sql
        Contador = Contador + 1
        If (Abstenciones_Procesadas = False) Then
            filtro = FiltroResultado("ABSTENCION", SonCopias)
            If Tipo_PreActa <> "consulta" Then
                TipoTitulo = "(Abstenciones)"
            End If
            If CantidadAbstenciones = 0 Then
                Vuelta1 = Vuelta1 - 1
                ImprimirActa = False
            End If
            If Contador >= CantidadAbstenciones Then
                Contador = 0
                Abstenciones_Procesadas = True
            End If
        ElseIf (Afirmativos_Procesados = False) Then
            filtro = FiltroResultado("AFIRMATIVO", SonCopias)
            TipoTitulo = "(Afirmativos)"
            If CantidadAfirmativos = 0 Then
                Vuelta1 = Vuelta1 - 1
                ImprimirActa = False
            End If
            If Contador >= CantidadAfirmativos Then
                Contador = 0
                Afirmativos_Procesados = True
            End If
        ElseIf (Negativos_Procesados = False) Then
            filtro = FiltroResultado("NEGATIVO", SonCopias)
            TipoTitulo = "(Negativos)"
            If CantidadNegativos = 0 Then
                Vuelta1 = Vuelta1 - 1
                ImprimirActa = False
            End If
            If Contador >= CantidadNegativos Then
                Contador = 0
                Negativos_Procesados = True
            End If
        ElseIf (Ausencias_Procesadas = False) Then
            filtro = FiltroResultado("AUSENTE", SonCopias)
            TipoTitulo = "(Ausencias)"
            If CantidadAusentes = 0 Then
                Vuelta1 = Vuelta1 - 1
                ImprimirActa = False
            End If
            If Contador >= CantidadAusentes Then
                Contador = 0
                Ausencias_Procesadas = True
            End If
        End If
        If ImprimirActa = True Then
            If PermisosTotales.ConsultaActas = 0 Then
                MsgBox "No posee permisos para consulta de actas", vbInformation + vbOKOnly, "Usuario sin permisos!!"
                Exit Sub
            End If
            
            Dim m_Report As New rptActas
            Dim rstActa  As New ADODB.Recordset
            Dim sql      As String
            Dim sql_voto_afirmativo_presidente As String
            Dim sql_voto_negativo_presidente As String
            Dim sql_voto_abstencion_presidente As String
            'Nueva Version para cálculo de presidente
            Dim rsPresi As ADODB.Recordset
            Dim PresidenteVotosAfirmativos As Integer
            Dim PresidenteVotosNegativos As Integer
            Dim PresidenteAbstenciones As Integer
            Dim PresidenteEstabaHabilitado As Boolean
            PresidenteVotosAfirmativos = 0
            PresidenteVotosNegativos = 0
            PresidenteAbstenciones = 0
            Set rsPresi = New ADODB.Recordset
            SetearRs "SELECT resultado_voto_presidente FROM Actas WHERE presidente_habilitado_votar = 1 AND " & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ")", rsPresi
            If rsPresi.EOF Then 'Reviso si votó el presidente, y qué votó
                PresidenteEstabaHabilitado = False
                PresidenteVotosAfirmativos = 0
                PresidenteVotosNegativos = 0
                PresidenteAbstenciones = 0
            Else
                PresidenteEstabaHabilitado = True
                Select Case Trim(LCase(rsPresi.Fields(0)))
                Case "s"
                    PresidenteVotosAfirmativos = 1
                Case "n"
                    PresidenteVotosNegativos = 1
                Case ""
                    PresidenteAbstenciones = 1
                End Select
            End If
            rsPresi.Close
            Set rsPresi = Nothing
            sql_voto_afirmativo_presidente = Trim(Str(PresidenteVotosAfirmativos))
            sql_voto_abstencion_presidente = Trim(Str(PresidenteAbstenciones))
            sql_voto_negativo_presidente = Trim(Str(PresidenteVotosNegativos))
             'Aumento contador de copias
            If strTipoOperacion = "votnum" Then
                sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, 0 AS Votos_Afirm_Identificables, (actas.Votos_Afirm_Identificables + actas.Votos_Afirm_No_Identificables " & _
                "+ " & sql_voto_afirmativo_presidente & ") AS Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, 0 AS Votos_Neg_Identificables , (actas.Votos_Neg_Identificables + actas.Votos_Neg_No_Identificables + " & sql_voto_negativo_presidente & ") AS Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion" & _
                      ", actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                      " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa,actas.Reunion, actas.Tipo_de_operación , Legisladores.distrito,tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                      " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' WHEN 'I' THEN 'Informativa' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                      " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                      " AND actas.Sesión = " & _
                      " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                      " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
                      " WHERE " & filtro & "(Actas.Período_Legislativo='" & pPeriodo & "')" & _
                      " AND (Actas.Sesión=" & pSesion & ")" & _
                      " AND (Actas.Número_de_Acta=" & pActa & ")" & _
                      " AND (Actas.Versión_Acta=" & pVersion & ")" & _
                      " AND (DetalleActas.Legislador_asignado <> '" & IDPresidente & "')" & _
                      " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
                
                'm_Report.Section5.Suppress = True
                m_Report.Detail.Visible = False
                'm_Report.Cuadro1.Suppress = True
                m_Report.delimitador1.Visible = False
                m_Report.delimitador2.Visible = False
                m_Report.delimitador3.Visible = False
                m_Report.delimitador4.Visible = False
                m_Report.lblVotacion.Caption = "Votación Numérica"
                'm_Report.Texto25.Suppress = True
                m_Report.Texto25.Visible = False
                'm_Report.Línea5.Suppress = True
                m_Report.Línea5.Visible = False
                'm_Report.Texto23.Suppress = True
                m_Report.Texto23.Visible = False
                'm_Report.Texto1.Suppress = True
                m_Report.Texto1.Visible = False
                'm_Report.Línea6.Suppress = True
                m_Report.Línea6.Visible = False
                'm_Report.Línea7.Suppress = True
                m_Report.Línea7.Visible = False
                'm_Report.Línea8.Suppress = True
                'm_Report.Línea8.Visible = False
                'm_Report.Línea9.Suppress = True
                'm_Report.Línea9.Visible = False
                'm_Report.Texto22.Suppress = True
                m_Report.Texto22.Visible = False
                                
                With m_Report
                    '.Campo14.Visible = False
                    '.Campo18.Visible = False
                    '.Campo22.Visible = False
                    '.Texto18.Visible = False
                    .Texto19.Caption = "Diputados"
                    .lblPresidente.Visible = False
                    .vap.Visible = False
                    .vnp.Visible = False
                    .vabsp.Visible = False
                End With
                    
            Else
                    'Version restando el voto de la columna identificados
                'sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - " & sql_voto_afirmativo_presidente & " as Votos_Afirm_Identificables " & _
                        ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables - " & sql_voto_negativo_presidente & " as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables - " & sql_voto_abstencion_presidente & " as Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                        " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                        " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                        "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                        "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                        "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                        " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                        " AND actas.Sesión = " & _
                        " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                        " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE " & filtro & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
                sql = " SELECT " & TotalPaginas & " AS Ultima_Pagina, DetalleActas.estado, DetalleActas.Grupo_Politico,DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - 0 as Votos_Afirm_Identificables " & _
                        ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, actas.Votos_Neg_Identificables - 0 as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total, actas.Abstenciones_Identificables + '" & sql_voto_abstencion_presidente & "' as Abstenciones_Identificables, actas.Abstenciones_Identificables AS Absten_Sin_Presi, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total + " & sql_voto_abstencion_presidente & " AS Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                        " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                        " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + " & _
                        "(CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'H' THEN 'Homenajes' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                        "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                        "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                        "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                        " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                        " AND actas.Sesión = " & _
                        " detalleactas.Sesión AND DetalleActas.Legislador_asignado <> '" & IDPresidente & "'" & _
                        IIf(ImpresionDeConsola And SonCopias = False, " AND LTrim(RTrim(DetalleActas.Resultado)) <> 'AUSENTE' ", "") & _
                        " AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                        " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE " & filtro & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ")" & _
                        " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
                    m_Report.Campo23.DataField = "Absten_Sin_Presi"
                m_Report.Campo15.DataField = "Votos_Afirm_Identificables"
                m_Report.Campo19.DataField = "Votos_Neg_Identificables"
                m_Report.Campo24.DataField = "Abstenciones_Identificables"
            End If
            If Trim(txtTipoOperacion.Text) = "Pase de Lista" Or strTipoOperacion = "paslis" Then
                Const Corrimiento As Integer = 3250
                With m_Report
                .lblVotacion.Caption = "Pase de Lista " & TipoTitulo
                .Texto12.Caption = "Resultado :"
                'Se eliminan todos los textos
                '.Texto18.Visible = False
                .Texto19.Visible = False
                .Texto20.Visible = False
                .Texto21.Visible = False
                .Texto13.Visible = False
                .Texto14.Visible = False
                .Texto15.Visible = False
                .Texto2.Visible = False
                .Texto8.Visible = False
                .Texto7.Visible = False
                '.Label1.Visible = False
                'Se eliminan todos los campos con Datafields
                '.Campo14.Visible = False
                .Campo15.Visible = False
                .Campo16.Visible = False
                .Campo17.Visible = False
                '.Campo18.Visible = False
                .Campo19.Visible = False
                .Campo20.Visible = False
                .Campo21.Visible = False
                '.Campo22.Visible = False
                .Campo23.Visible = False
                .Campo24.Visible = False
                .Campo33.Visible = False
                .Campo24.Visible = False
                .Campo33.Visible = False
                .Campo30.Visible = False
                .Campo35.Visible = False
                .Campo6.Visible = False
                'Se elimina el delimitador
                .Line7.Visible = False
                'Se acomodan los controles
                .Label2.Left = .Label2.Left + Corrimiento
                .Label3.Left = .Label3.Left + Corrimiento
                .Label4.Left = .Label4.Left + Corrimiento
                .Texto16.Left = .Texto16.Left + Corrimiento
                .Texto17.Left = .Texto17.Left + Corrimiento
                .Campo25.Left = .Campo25.Left + Corrimiento
                .Campo26.Left = .Campo26.Left + Corrimiento
                .Campo27.Left = .Campo27.Left + Corrimiento
                .Campo28.Left = .Campo28.Left + Corrimiento
                .lblPresidente.Visible = False
                .vap.Visible = False
                .vnp.Visible = False
                .vabsp.Visible = False
                End With
            End If
            SetearRs sql, rstActa
            Dim cPatch As Boolean
            cPatch = False
            If (rstActa.RecordCount <= 0) Then
                If SonCopias = True Then
                    sql = " SELECT TOP 1 " & TotalPaginas & " AS Ultima_Pagina, DetalleActas.estado, DetalleActas.Grupo_Politico,DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - 0 as Votos_Afirm_Identificables " & _
                            ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, actas.Votos_Neg_Identificables - 0 as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total, actas.Abstenciones_Identificables + '" & sql_voto_abstencion_presidente & "' as Abstenciones_Identificables, actas.Abstenciones_Identificables AS Absten_Sin_Presi, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total + " & sql_voto_abstencion_presidente & " AS Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                            " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                            " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + " & _
                            "(CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'H' THEN 'Homenajes' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                            "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                            "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                            "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                            " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                            " AND actas.Sesión = " & _
                            " detalleactas.Sesión AND DetalleActas.Legislador_asignado <> '" & IDPresidente & "'" & _
                            IIf(ImpresionDeConsola And SonCopias = False, " AND LTrim(RTrim(DetalleActas.Resultado)) <> 'AUSENTE' ", "") & _
                            " AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                            " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ")" & _
                            " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
                    SetearRs sql, rstActa
                    cPatch = True
                End If
            End If
            If (rstActa.RecordCount <> 0) Then
            'Para probar si funciona
            'Sin ir a la DB, eliminar la condicion
            'Asi se abriran los reportes aunque
            'No haya ninguna abstencion/ausencia/etc
            Set m_Report.DataControl1.Recordset = rstActa
            If cPatch = True Then
                m_Report.Detail.Visible = False
                m_Report.GroupHeader3.Visible = False
            End If
            If SonCopias = True Then
                m_Report.lblVotacion.Caption = "Detalle de Votación"
            End If
            If True Then
                If VistaPrevia = True Then
                    m_Report.Show vbModal
                Else
                    m_Report.Run False
                    'If strTipoOperacion <> "votnum" Then
                        Dim X As Integer
                        For X = 0 To m_Report.Pages.Count - 1
                            m_Report.Pages(X).Width = m_Report.Pages(X).Width - 1000
                            m_Report.Pages.Commit
                        Next X
                        m_Report.Pages.Commit
                    'End If
                    If ImpresionDeConsola = True Then
                        m_Report.PrintReport False
                    Else
                        If Tipo_PreActa = "consulta" Then
                            m_Report.PrintReport True
                        Else
                            m_Report.PrintReport False
                        End If
                    End If
                End If
            Else
                'm_Report.Printer.Copies = 1
                'm_Report.Printer.StartJob "acta"
                'm_Report.Show vbModal
                m_Report.PrintReport False
            End If
            Set m_Report = Nothing
            End If
            rstActa.Close
            Set rstActa = Nothing
        End If
    Next Vuelta1
Exit Sub
TrapError:
    Select Case err.Number
        Case Else
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            End
    End Select
Return
End Sub
Private Function FiltroResultado(tipoFiltro As String, SonCopias As Boolean)
If SonCopias = True Then
    FiltroResultado = " detalleactas.Resultado = '" & tipoFiltro & "' AND "
Else
    FiltroResultado = ""
End If
End Function
Public Sub imprimirActaAuditoria(strTipoOperacion As String, pPeriodo As String, pSesion As Integer, pActa As Integer, pVersion As Integer, CantidadAbstenciones As Integer, CantidadAfirmativos As Integer, CantidadNegativos As Integer, CantidadAusentes As Integer, SonCopias As Boolean)
    'On Error GoTo TrapError
    'IMPRESION ACTA 2009
    Dim Vuelta1 As Integer
    Dim VueltasTotales As Integer
    Dim Contador As Integer
    Dim Abstenciones_Procesadas As Boolean
    Dim Afirmativos_Procesados As Boolean
    Dim Negativos_Procesados As Boolean
    Dim Ausencias_Procesadas As Boolean
    Dim ImprimirActa As Boolean
    Dim TipoTitulo As String
    Abstenciones_Procesadas = False
    Afirmativos_Procesados = False
    Negativos_Procesados = False
    Ausencias_Procesadas = False
    Contador = 0
    VueltasTotales = CantidadAbstenciones + CantidadAfirmativos + CantidadNegativos + CantidadAusentes
    For Vuelta1 = 1 To VueltasTotales
        ImprimirActa = True
        Dim filtro As String 'filtro para los sql
        Contador = Contador + 1
        If (Abstenciones_Procesadas = False) Then
            filtro = FiltroResultado("ABSTENCION", SonCopias)
            TipoTitulo = "(Abstenciones)"
            If CantidadAbstenciones = 0 Then
                Vuelta1 = Vuelta1 - 1
                ImprimirActa = False
            End If
            If Contador >= CantidadAbstenciones Then
                Contador = 0
                Abstenciones_Procesadas = True
            End If
        ElseIf (Afirmativos_Procesados = False) Then
            filtro = FiltroResultado("AFIRMATIVO", SonCopias)
            TipoTitulo = "(Afirmativos)"
            If CantidadAfirmativos = 0 Then
                Vuelta1 = Vuelta1 - 1
                ImprimirActa = False
            End If
            If Contador >= CantidadAfirmativos Then
                Contador = 0
                Afirmativos_Procesados = True
            End If
        ElseIf (Negativos_Procesados = False) Then
            filtro = FiltroResultado("NEGATIVO", SonCopias)
            TipoTitulo = "(Negativos)"
            If CantidadNegativos = 0 Then
                Vuelta1 = Vuelta1 - 1
                ImprimirActa = False
            End If
            If Contador >= CantidadNegativos Then
                Contador = 0
                Negativos_Procesados = True
            End If
        ElseIf (Ausencias_Procesadas = False) Then
            filtro = FiltroResultado("AUSENTE", SonCopias)
            TipoTitulo = "(Ausencias)"
            If CantidadAusentes = 0 Then
                Vuelta1 = Vuelta1 - 1
                ImprimirActa = False
            End If
            If Contador >= CantidadAusentes Then
                Contador = 0
                Ausencias_Procesadas = True
            End If
        End If
        If ImprimirActa = True Then
            If PermisosTotales.ConsultaActas = 0 Then
                MsgBox "No posee permisos para consulta de actas", vbInformation + vbOKOnly, "Usuario sin permisos!!"
                Exit Sub
            End If
            
            Dim m_Report As New rptAuditoria
            Dim rstActa  As New ADODB.Recordset
            Dim sql      As String
            Dim sql_voto_afirmativo_presidente As String
            Dim sql_voto_negativo_presidente As String
            Dim sql_voto_abstencion_presidente As String
            'Nueva Version para cálculo de presidente
            Dim rsPresi As ADODB.Recordset
            Dim PresidenteVotosAfirmativos As Integer
            Dim PresidenteVotosNegativos As Integer
            Dim PresidenteAbstenciones As Integer
            Dim PresidenteEstabaHabilitado As Boolean
            PresidenteVotosAfirmativos = 0
            PresidenteVotosNegativos = 0
            PresidenteAbstenciones = 0
            Set rsPresi = New ADODB.Recordset
            SetearRs "SELECT resultado_voto_presidente FROM Actas WHERE presidente_habilitado_votar = 1 AND " & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ")", rsPresi
            If rsPresi.EOF Then 'Reviso si votó el presidente, y qué votó
                PresidenteEstabaHabilitado = False
                PresidenteVotosAfirmativos = 0
                PresidenteVotosNegativos = 0
                PresidenteAbstenciones = 0
            Else
                PresidenteEstabaHabilitado = True
                Select Case Trim(LCase(rsPresi.Fields(0)))
                Case "s"
                    PresidenteVotosAfirmativos = 1
                Case "n"
                    PresidenteVotosNegativos = 1
                Case ""
                    PresidenteAbstenciones = 1
                End Select
            End If
            rsPresi.Close
            Set rsPresi = Nothing
            sql_voto_afirmativo_presidente = Trim(Str(PresidenteVotosAfirmativos))
            sql_voto_abstencion_presidente = Trim(Str(PresidenteAbstenciones))
            sql_voto_negativo_presidente = Trim(Str(PresidenteVotosNegativos))
             'Aumento contador de copias
            If strTipoOperacion = "votnum" Then
                sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, 0 AS Votos_Afirm_Identificables, (actas.Votos_Afirm_Identificables + actas.Votos_Afirm_No_Identificables " & _
                "+ " & sql_voto_afirmativo_presidente & ") AS Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, 0 AS Votos_Neg_Identificables , (actas.Votos_Neg_Identificables + actas.Votos_Neg_No_Identificables + " & sql_voto_negativo_presidente & ") AS Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion" & _
                      ", actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                      " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa,actas.Reunion, actas.Tipo_de_operación , Legisladores.distrito,tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                      " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' WHEN 'I' THEN 'Informativa' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                      " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                      " AND actas.Sesión = " & _
                      " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                      " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
                      " WHERE " & filtro & "(Actas.Período_Legislativo='" & pPeriodo & "')" & _
                      " AND (Actas.Sesión=" & pSesion & ")" & _
                      " AND (Actas.Número_de_Acta=" & pActa & ")" & _
                      " AND (Actas.Versión_Acta=" & pVersion & ")" & _
                      " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
                
                'm_Report.Section5.Suppress = True
                m_Report.Detail.Visible = False
                'm_Report.Cuadro1.Suppress = True
                    
            Else
                    'Version restando el voto de la columna identificados
                'sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - " & sql_voto_afirmativo_presidente & " as Votos_Afirm_Identificables " & _
                        ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables - " & sql_voto_negativo_presidente & " as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables - " & sql_voto_abstencion_presidente & " as Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                        " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                        " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                        "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                        "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                        "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                        " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                        " AND actas.Sesión = " & _
                        " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                        " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE " & filtro & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
                'sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - 0 as Votos_Afirm_Identificables " & _
                        ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, actas.Votos_Neg_Identificables - 0 as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total, actas.Abstenciones_Identificables + '" & sql_voto_abstencion_presidente & "' as Abstenciones_Identificables, actas.Abstenciones_Identificables AS Absten_Sin_Presi, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total + " & sql_voto_abstencion_presidente & " AS Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                        " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                        " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                        "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                        "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                        "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                        " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                        " AND actas.Sesión = " & _
                        " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                        " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE " & filtro & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
                sql = " SELECT " & Paginas_Fijas_Auditoria & " AS Paginas_Fijas, " & _
                "Legisladores.apellido + ', ' + Legisladores.nombre AS DetalleLegislador, Legisladores.bloque_politico AS Bloque_político, distritos.distrito, nueva.Legislador_asignado, CASE vieja.Resultado WHEN 'AUSENTE' THEN 'Presente No Identificado' ELSE vieja.resultado END AS Voto_Viejo, CASE nueva.Resultado WHEN 'AUSENTE' THEN 'Presente No Identificado' ELSE nueva.resultado END AS Voto_Nuevo FROM detalleactas nueva INNER JOIN detalleactas vieja ON nueva.Legislador_asignado = vieja.Legislador_asignado AND nueva.Versión_Acta = " & pVersion & " AND vieja.Nro_de_Acta = " & pActa & " AND vieja.Sesión = nueva.Sesión AND nueva.Período_Legislativo = vieja.Período_Legislativo AND nueva.Resultado <> vieja.Resultado AND nueva.Nro_de_Acta = " & pActa & " And vieja.Versión_Acta = 1 And nueva.Período_Legislativo = '" & pPeriodo & "' INNER JOIN Legisladores ON nueva.Legislador_asignado = Legisladores.id INNER JOIN distritos ON Legisladores.distrito = distritos.id_distrito WHERE nueva.Sesión = " & pSesion
                sql = sql & " ORDER BY Legisladores.apellido"
                'vieja.Versión_Acta = 1 Siempre va a ser la original
            End If
            If Trim(txtTipoOperacion.Text) = "Pase de Lista" Or strTipoOperacion = "paslis" Then
                Const Corrimiento As Integer = 3250
                With m_Report
                .lblVotacion.Caption = "Pase de Lista " & TipoTitulo
                End With
            End If
            SetearRs sql, rstActa
            If strTipoOperacion = "paslis" Then
                m_Report.lblVotacion.Caption = Trim(Replace(m_Report.lblVotacion.Caption, "(Abstenciones)", ""))
                m_Report.Texto2.Visible = False
                m_Report.Texto8.Visible = False
                m_Report.Campo30.Visible = False
                m_Report.Campo35.Visible = False
            End If
            
            If (rstActa.RecordCount <> 0 Or TieneManifestaciones()) Then
                If (rstActa.RecordCount = 0) Then
                    'Si entró acá y no tiene modificaciones de votos, lo preparo para Manifestaciones
                    m_Report.GroupHeader2.Height = 450
                    m_Report.GroupHeader3.Height = 0
                    m_Report.GroupHeader2.Controls("Label7").Top = 170
                    m_Report.PageHeader.Controls("Campo31").DataField = ""
                    If (Paginas_Fijas_Auditoria <> "") Then
                        m_Report.PageHeader.Controls("Campo31").Text = Trim(Str(Paginas_Fijas_Auditoria))
                    End If
                End If
                Set m_Report.DataControl1.Recordset = rstActa
                If False Then
                    m_Report.Show vbModal, Me
                Else
                    'Manifestaciones
                    Call Me.imprimirReporteManifestaciones(m_Report)
                    m_Report.Run False
                    Dim X As Integer
                    Dim conta As Integer
                    conta = TodoElReporte.Pages.Count - 1
                    Dim TotalReporte
                    For X = 0 To m_Report.Pages.Count - 1
                        conta = conta + 1
                        TodoElReporte.Pages.Insert conta, m_Report.Pages(X)
                    Next X
                    TodoElReporte.Pages.Commit
                End If
                Set m_Report = Nothing
            End If
            rstActa.Close
            Set rstActa = Nothing
        End If
    Next Vuelta1
Exit Sub
TrapError:
    Select Case err.Number
        Case Else
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            End
    End Select
Return
End Sub
Public Function ObtenerCantidadAuditoria(strTipoOperacion As String, pPeriodo As String, pSesion As Integer, pActa As Integer, pVersion As Integer, CantidadAbstenciones As Integer, CantidadAfirmativos As Integer, CantidadNegativos As Integer, CantidadAusentes As Integer, SonCopias As Boolean) As Integer
    'On Error GoTo TrapError
    'IMPRESION ACTA 2009
    Dim Vuelta1 As Integer
    Dim VueltasTotales As Integer
    Dim Contador As Integer
    Dim Abstenciones_Procesadas As Boolean
    Dim Afirmativos_Procesados As Boolean
    Dim Negativos_Procesados As Boolean
    Dim Ausencias_Procesadas As Boolean
    Dim ImprimirActa As Boolean
    Dim TipoTitulo As String
    Abstenciones_Procesadas = False
    Afirmativos_Procesados = False
    Negativos_Procesados = False
    Ausencias_Procesadas = False
    Contador = 0
    VueltasTotales = CantidadAbstenciones + CantidadAfirmativos + CantidadNegativos + CantidadAusentes
    For Vuelta1 = 1 To VueltasTotales
        ImprimirActa = True
        Dim filtro As String 'filtro para los sql
        Contador = Contador + 1
        If (Abstenciones_Procesadas = False) Then
            filtro = FiltroResultado("ABSTENCION", SonCopias)
            TipoTitulo = "(Abstenciones)"
            If CantidadAbstenciones = 0 Then
                Vuelta1 = Vuelta1 - 1
                ImprimirActa = False
            End If
            If Contador >= CantidadAbstenciones Then
                Contador = 0
                Abstenciones_Procesadas = True
            End If
        ElseIf (Afirmativos_Procesados = False) Then
            filtro = FiltroResultado("AFIRMATIVO", SonCopias)
            TipoTitulo = "(Afirmativos)"
            If CantidadAfirmativos = 0 Then
                Vuelta1 = Vuelta1 - 1
                ImprimirActa = False
            End If
            If Contador >= CantidadAfirmativos Then
                Contador = 0
                Afirmativos_Procesados = True
            End If
        ElseIf (Negativos_Procesados = False) Then
            filtro = FiltroResultado("NEGATIVO", SonCopias)
            TipoTitulo = "(Negativos)"
            If CantidadNegativos = 0 Then
                Vuelta1 = Vuelta1 - 1
                ImprimirActa = False
            End If
            If Contador >= CantidadNegativos Then
                Contador = 0
                Negativos_Procesados = True
            End If
        ElseIf (Ausencias_Procesadas = False) Then
            filtro = FiltroResultado("AUSENTE", SonCopias)
            TipoTitulo = "(Ausencias)"
            If CantidadAusentes = 0 Then
                Vuelta1 = Vuelta1 - 1
                ImprimirActa = False
            End If
            If Contador >= CantidadAusentes Then
                Contador = 0
                Ausencias_Procesadas = True
            End If
        End If
        If ImprimirActa = True Then
            If PermisosTotales.ConsultaActas = 0 Then
                MsgBox "No posee permisos para consulta de actas", vbInformation + vbOKOnly, "Usuario sin permisos!!"
                Exit Function
            End If
            
            Dim m_Report As New rptAuditoria
            Dim rstActa  As New ADODB.Recordset
            Dim sql      As String
            Dim sql_voto_afirmativo_presidente As String
            Dim sql_voto_negativo_presidente As String
            Dim sql_voto_abstencion_presidente As String
            'Nueva Version para cálculo de presidente
            Dim rsPresi As ADODB.Recordset
            Dim PresidenteVotosAfirmativos As Integer
            Dim PresidenteVotosNegativos As Integer
            Dim PresidenteAbstenciones As Integer
            Dim PresidenteEstabaHabilitado As Boolean
            PresidenteVotosAfirmativos = 0
            PresidenteVotosNegativos = 0
            PresidenteAbstenciones = 0
            Set rsPresi = New ADODB.Recordset
            SetearRs "SELECT resultado_voto_presidente FROM Actas WHERE presidente_habilitado_votar = 1 AND " & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ")", rsPresi
            If rsPresi.EOF Then 'Reviso si votó el presidente, y qué votó
                PresidenteEstabaHabilitado = False
                PresidenteVotosAfirmativos = 0
                PresidenteVotosNegativos = 0
                PresidenteAbstenciones = 0
            Else
                PresidenteEstabaHabilitado = True
                Select Case Trim(LCase(rsPresi.Fields(0)))
                Case "s"
                    PresidenteVotosAfirmativos = 1
                Case "n"
                    PresidenteVotosNegativos = 1
                Case ""
                    PresidenteAbstenciones = 1
                End Select
            End If
            rsPresi.Close
            Set rsPresi = Nothing
            sql_voto_afirmativo_presidente = Trim(Str(PresidenteVotosAfirmativos))
            sql_voto_abstencion_presidente = Trim(Str(PresidenteAbstenciones))
            sql_voto_negativo_presidente = Trim(Str(PresidenteVotosNegativos))
             'Aumento contador de copias
           If True Then
                    'Version restando el voto de la columna identificados
                'sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - " & sql_voto_afirmativo_presidente & " as Votos_Afirm_Identificables " & _
                        ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables - " & sql_voto_negativo_presidente & " as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables - " & sql_voto_abstencion_presidente & " as Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                        " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                        " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                        "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                        "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                        "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                        " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                        " AND actas.Sesión = " & _
                        " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                        " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE " & filtro & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
                'sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - 0 as Votos_Afirm_Identificables " & _
                        ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, actas.Votos_Neg_Identificables - 0 as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total, actas.Abstenciones_Identificables + '" & sql_voto_abstencion_presidente & "' as Abstenciones_Identificables, actas.Abstenciones_Identificables AS Absten_Sin_Presi, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total + " & sql_voto_abstencion_presidente & " AS Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                        " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                        " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                        "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                        "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                        "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                        " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                        " AND actas.Sesión = " & _
                        " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                        " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE " & filtro & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
                sql = " SELECT Legisladores.apellido + ', ' + Legisladores.nombre AS DetalleLegislador, Legisladores.bloque_politico AS Bloque_político, distritos.distrito, nueva.Legislador_asignado, RTRIM(CASE vieja.Resultado WHEN 'AUSENTE' THEN 'Presente No Identificado' ELSE vieja.resultado END) + ' -----------> ' + RTRIM(CASE nueva.Resultado WHEN 'AUSENTE' THEN 'Presente No Identificado' ELSE nueva.resultado END) AS Resultado FROM detalleactas nueva INNER JOIN detalleactas vieja ON nueva.Legislador_asignado = vieja.Legislador_asignado AND nueva.Versión_Acta = " & pVersion & " AND vieja.Nro_de_Acta = " & pActa & " AND vieja.Sesión = nueva.Sesión AND nueva.Período_Legislativo = vieja.Período_Legislativo AND nueva.Resultado <> vieja.Resultado AND nueva.Nro_de_Acta = " & pActa & " And vieja.Versión_Acta = 1 And nueva.Período_Legislativo = '" & pPeriodo & "' INNER JOIN Legisladores ON nueva.Legislador_asignado = Legisladores.id INNER JOIN distritos ON Legisladores.distrito = distritos.id_distrito WHERE " & _
                " nueva.Sesión = " & pSesion
                'vieja.Versión_Acta = 1 Siempre va a ser la original
            End If
            SetearRs sql, rstActa
            If (rstActa.RecordCount <> 0 Or Me.TieneManifestaciones()) Then
                If (rstActa.RecordCount = 0) Then
                    'Si entró acá y no tiene modificaciones de votos, lo preparo para Manifestaciones
                    m_Report.GroupHeader2.Height = 450
                    m_Report.GroupHeader3.Height = 0
                    m_Report.GroupHeader2.Controls("Label7").Top = 170
                    m_Report.PageHeader.Controls("Campo31").DataField = ""
                    If (Paginas_Fijas_Auditoria <> "") Then
                        m_Report.PageHeader.Controls("Campo31").Text = Trim(Str(Paginas_Fijas_Auditoria))
                    End If
                End If
                Set m_Report.DataControl1.Recordset = rstActa
                If False Then
                    m_Report.Show vbModal, Me
                Else
                    'Manifestaciones
                    Call Me.imprimirReporteManifestaciones(m_Report)
                    m_Report.Run False
                    ObtenerCantidadAuditoria = m_Report.Pages.Count
                End If
                Set m_Report = Nothing
            End If
        End If
    Next Vuelta1
Exit Function
TrapError:
    Select Case err.Number
        Case Else
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            End
    End Select
Return
End Function
Public Sub imprimirActaConPagina(strTipoOperacion As String, pPeriodo As String, pSesion As Integer, pActa As Integer, pVersion As Integer, CantidadAbstenciones As Integer, CantidadAfirmativos As Integer, CantidadNegativos As Integer, CantidadAusentes As Integer, SonCopias As Boolean)

'    'On Error GoTo TrapError
'    'IMPRESION ACTA 2009
'    Dim Vuelta1 As Integer
'    Dim VueltasTotales As Integer
'    Dim Contador As Integer
'    Dim Abstenciones_Procesadas As Boolean
'    Dim Afirmativos_Procesados As Boolean
'    Dim Negativos_Procesados As Boolean
'    Dim Ausencias_Procesadas As Boolean
'    Dim ImprimirActa As Boolean
'    Dim TipoTitulo As String
'    Abstenciones_Procesadas = False
'    Afirmativos_Procesados = False
'    Negativos_Procesados = False
'    Ausencias_Procesadas = False
'    Contador = 0
'    VueltasTotales = CantidadAbstenciones + CantidadAfirmativos + CantidadNegativos + CantidadAusentes
'    For Vuelta1 = 1 To VueltasTotales
'        ImprimirActa = True
'        Dim filtro As String 'filtro para los sql
'        Contador = Contador + 1
'        If (Abstenciones_Procesadas = False) Then
'            filtro = FiltroResultado("ABSTENCION", SonCopias)
'            TipoTitulo = "(Abstenciones)"
'            If CantidadAbstenciones = 0 Then
'                Vuelta1 = Vuelta1 - 1
'                ImprimirActa = False
'            End If
'            If Contador >= CantidadAbstenciones Then
'                Contador = 0
'                Abstenciones_Procesadas = True
'            End If
'        ElseIf (Afirmativos_Procesados = False) Then
'            filtro = FiltroResultado("AFIRMATIVO", SonCopias)
'            TipoTitulo = "(Afirmativos)"
'            If CantidadAfirmativos = 0 Then
'                Vuelta1 = Vuelta1 - 1
'                ImprimirActa = False
'            End If
'            If Contador >= CantidadAfirmativos Then
'                Contador = 0
'                Afirmativos_Procesados = True
'            End If
'        ElseIf (Negativos_Procesados = False) Then
'            filtro = FiltroResultado("NEGATIVO", SonCopias)
'            TipoTitulo = "(Negativos)"
'            If CantidadNegativos = 0 Then
'                Vuelta1 = Vuelta1 - 1
'                ImprimirActa = False
'            End If
'            If Contador >= CantidadNegativos Then
'                Contador = 0
'                Negativos_Procesados = True
'            End If
'        ElseIf (Ausencias_Procesadas = False) Then
'            filtro = FiltroResultado("AUSENTE", SonCopias)
'            TipoTitulo = "(Ausencias)"
'            If CantidadAusentes = 0 Then
'                Vuelta1 = Vuelta1 - 1
'                ImprimirActa = False
'            End If
'            If Contador >= CantidadAusentes Then
'                Contador = 0
'                Ausencias_Procesadas = True
'            End If
'        End If
'        If ImprimirActa = True Then
'            If PermisosTotales.ConsultaActas = 0 Then
'                MsgBox "No posee permisos para consulta de actas", vbInformation + vbOKOnly, "Usuario sin permisos!!"
'                Exit Sub
'            End If
'
'            Dim m_Report As New rptAuditoria
'            Dim rstActa  As New ADODB.Recordset
'            Dim sql      As String
'            Dim sql_voto_afirmativo_presidente As String
'            Dim sql_voto_negativo_presidente As String
'            Dim sql_voto_abstencion_presidente As String
'            'Nueva Version para cálculo de presidente
'            Dim rsPresi As ADODB.Recordset
'            Dim PresidenteVotosAfirmativos As Integer
'            Dim PresidenteVotosNegativos As Integer
'            Dim PresidenteAbstenciones As Integer
'            Dim PresidenteEstabaHabilitado As Boolean
'            PresidenteVotosAfirmativos = 0
'            PresidenteVotosNegativos = 0
'            PresidenteAbstenciones = 0
'            Set rsPresi = New ADODB.Recordset
'            SetearRs "SELECT resultado_voto_presidente FROM Actas WHERE presidente_habilitado_votar = 1 AND " & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ")", rsPresi
'            If rsPresi.EOF Then 'Reviso si votó el presidente, y qué votó
'                PresidenteEstabaHabilitado = False
'                PresidenteVotosAfirmativos = 0
'                PresidenteVotosNegativos = 0
'                PresidenteAbstenciones = 0
'            Else
'                PresidenteEstabaHabilitado = True
'                Select Case Trim(LCase(rsPresi.Fields(0)))
'                Case "s"
'                    PresidenteVotosAfirmativos = 1
'                Case "n"
'                    PresidenteVotosNegativos = 1
'                Case ""
'                    PresidenteAbstenciones = 1
'                End Select
'            End If
'            rsPresi.Close
'            Set rsPresi = Nothing
'            sql_voto_afirmativo_presidente = Trim(Str(PresidenteVotosAfirmativos))
'            sql_voto_abstencion_presidente = Trim(Str(PresidenteAbstenciones))
'            sql_voto_negativo_presidente = Trim(Str(PresidenteVotosNegativos))
'             'Aumento contador de copias
'            If strTipoOperacion = "votnum" Then
'                sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, 0 AS Votos_Afirm_Identificables, (actas.Votos_Afirm_Identificables + actas.Votos_Afirm_No_Identificables " & _
'                "+ " & sql_voto_afirmativo_presidente & ") AS Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, 0 AS Votos_Neg_Identificables , (actas.Votos_Neg_Identificables + actas.Votos_Neg_No_Identificables + " & sql_voto_negativo_presidente & ") AS Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion" & _
'                      ", actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
'                      " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa,actas.Reunion, actas.Tipo_de_operación , Legisladores.distrito,tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
'                      " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' WHEN 'I' THEN 'Informativa' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
'                      " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
'                      " AND actas.Sesión = " & _
'                      " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
'                      " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
'                      " WHERE " & filtro & "(Actas.Período_Legislativo='" & pPeriodo & "')" & _
'                      " AND (Actas.Sesión=" & pSesion & ")" & _
'                      " AND (Actas.Número_de_Acta=" & pActa & ")" & _
'                      " AND (Actas.Versión_Acta=" & pVersion & ")" & _
'                      " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
'
'                'm_Report.Section5.Suppress = True
'                m_Report.Detail.Visible = False
'                'm_Report.Cuadro1.Suppress = True
'                m_Report.delimitador1.Visible = False
'                m_Report.delimitador2.Visible = False
'                m_Report.delimitador3.Visible = False
'                m_Report.delimitador4.Visible = False
'                m_Report.lblVotacion.Caption = "Votación Numérica"
'                'm_Report.Texto25.Suppress = True
'                m_Report.Texto25.Visible = False
'                'm_Report.Línea5.Suppress = True
'                m_Report.Línea5.Visible = False
'                'm_Report.Texto23.Suppress = True
'                'm_Report.Texto23.Visible = False
'                'm_Report.Texto1.Suppress = True
'                'm_Report.Texto1.Visible = False
'                'm_Report.Línea6.Suppress = True
'                'm_Report.Línea6.Visible = False
'                'm_Report.Línea7.Suppress = True
'                'm_Report.Línea7.Visible = False
'                'm_Report.Línea8.Suppress = True
'                'm_Report.Línea8.Visible = False
'                'm_Report.Línea9.Suppress = True
'                'm_Report.Línea9.Visible = False
'                'm_Report.Texto22.Suppress = True
'                m_Report.Texto22.Visible = False
'
'                With m_Report
'                    '.Campo14.Visible = False
'                    '.Campo18.Visible = False
'                    '.Campo22.Visible = False
'                    '.Texto18.Visible = False
'                    '.Texto19.Caption = "Diputados"
'                    '.lblPresidente.Visible = False
'                    '.vap.Visible = False
'                    '.vnp.Visible = False
'                    '.vabsp.Visible = False
'                End With
'
'            Else
'                    'Version restando el voto de la columna identificados
'                'sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - " & sql_voto_afirmativo_presidente & " as Votos_Afirm_Identificables " & _
'                        ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables - " & sql_voto_negativo_presidente & " as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables - " & sql_voto_abstencion_presidente & " as Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
'                        " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
'                        " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
'                        "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
'                        "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
'                        "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
'                        " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
'                        " AND actas.Sesión = " & _
'                        " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
'                        " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE " & filtro & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
'                sql = " SELECT '" & TotalPaginas & "' AS Ultima_Pagina, DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - 0 as Votos_Afirm_Identificables " & _
'                        ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, actas.Votos_Neg_Identificables - 0 as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total, actas.Abstenciones_Identificables + '" & sql_voto_abstencion_presidente & "' as Abstenciones_Identificables, actas.Abstenciones_Identificables AS Absten_Sin_Presi, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total + " & sql_voto_abstencion_presidente & " AS Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
'                        " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
'                        " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
'                        "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
'                        "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
'                        "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
'                        " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
'                        " AND actas.Sesión = " & _
'                        " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
'                        " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE " & filtro & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
'                m_Report.Campo23.DataField = "Absten_Sin_Presi"
'                m_Report.Campo15.DataField = "Votos_Afirm_Identificables"
'                m_Report.Campo19.DataField = "Votos_Neg_Identificables"
'                m_Report.Campo24.DataField = "Abstenciones_Identificables"
'            End If
'            If Trim(txtTipoOperacion.Text) = "Pase de Lista" Or strTipoOperacion = "paslis" Then
'                Const Corrimiento As Integer = 3250
'                With m_Report
'                .lblVotacion.Caption = "Pase de Lista " & TipoTitulo
'                .Texto12.Caption = "Resultado :"
'                'Se eliminan todos los textos
'                '.Texto18.Visible = False
'                .Texto19.Visible = False
'                .Texto20.Visible = False
'                .Texto21.Visible = False
'                .Texto13.Visible = False
'                .Texto14.Visible = False
'                .Texto15.Visible = False
'                .Texto2.Visible = False
'                .Texto8.Visible = False
'                .Texto7.Visible = False
'                '.Label1.Visible = False
'                'Se eliminan todos los campos con Datafields
'                '.Campo14.Visible = False
'                .Campo15.Visible = False
'                .Campo16.Visible = False
'                .Campo17.Visible = False
'                '.Campo18.Visible = False
'                .Campo19.Visible = False
'                .Campo20.Visible = False
'                .Campo21.Visible = False
'                '.Campo22.Visible = False
'                .Campo23.Visible = False
'                .Campo24.Visible = False
'                .Campo33.Visible = False
'                .Campo24.Visible = False
'                .Campo33.Visible = False
'                .Campo30.Visible = False
'                .Campo35.Visible = False
'                .Campo6.Visible = False
'                'Se elimina el delimitador
'                .Line7.Visible = False
'                'Se acomodan los controles
'                .Label2.Left = .Label2.Left + Corrimiento
'                .Label3.Left = .Label3.Left + Corrimiento
'                .Label4.Left = .Label4.Left + Corrimiento
'                .Texto16.Left = .Texto16.Left + Corrimiento
'                .Texto17.Left = .Texto17.Left + Corrimiento
'                .Campo25.Left = .Campo25.Left + Corrimiento
'                .Campo26.Left = .Campo26.Left + Corrimiento
'                .Campo27.Left = .Campo27.Left + Corrimiento
'                .Campo28.Left = .Campo28.Left + Corrimiento
'                .lblPresidente.Visible = False
'                .vap.Visible = False
'                .vnp.Visible = False
'                .vabsp.Visible = False
'                End With
'            End If
'            SetearRs sql, rstActa
'            If (rstActa.RecordCount <> 0) Then
'            'Para probar si funciona
'            'Sin ir a la DB, eliminar la condicion
'            'Asi se abriran los reportes aunque
'            'No haya ninguna abstencion/ausencia/etc
'            Set m_Report.DataControl1.Recordset = rstActa
'            If True Then
'                m_Report.Show vbModal, Me
'            Else
'                'm_Report.Printer.Copies = 1
'                'm_Report.Printer.StartJob "acta"
'                'm_Report.Show vbModal
'                m_Report.PrintReport False
'            End If
'            Set m_Report = Nothing
'            End If
'            rstActa.Close
'            Set rstActa = Nothing
'        End If
'    Next Vuelta1
'Exit Sub
'TrapError:
'    Select Case Err.Number
'        Case Else
'            MsgBox "Error Nº " & Err.Number & Chr(10) & Err.Description & Chr(10) & "Originado en " & Err.Source
'            End
'    End Select
Return
End Sub

Public Sub imprimirUnActaCompleta(strTipoOperacion As String, pPeriodo As String, pSesion As Integer, pActa As Integer, pVersion As Integer)
    'On Error GoTo TrapError
    'IMPRESION ACTA 2009
    Dim IDPresidente As Integer
    If PermisosTotales.ConsultaActas = 0 Then
        MsgBox "No posee permisos para consulta de actas", vbInformation + vbOKOnly, "Usuario sin permisos!!"
        Exit Sub
    End If
    
    Dim m_Report As New rptActas
    Dim rstActa  As New ADODB.Recordset
    'Dim fViewer  As frmVisor
    Dim sql      As String
    Dim sql_voto_afirmativo_presidente As String
    Dim sql_voto_negativo_presidente As String
    Dim sql_voto_abstencion_presidente As String
    Dim rsPresi As ADODB.Recordset
    Dim PresidenteVotosAfirmativos As Integer
    Dim PresidenteVotosNegativos As Integer
    Dim PresidenteAbstenciones As Integer
    Dim RsTemp As ADODB.Recordset
    Dim rsAC As ADODB.Recordset
    Dim Cons As String
    Set rsAC = New ADODB.Recordset
    Cons = "Select Observaciones FROM actas WHERE (Actas.Período_Legislativo='" & pPeriodo & "')" & _
              " AND (Actas.Sesión=" & pSesion & ")" & _
              " AND (Actas.Número_de_Acta=" & pActa & ")" & _
              " AND (Actas.Versión_Acta=" & pVersion & ")"
    SetearRs Cons, rsAC
    ObsActa = Trim(rsAC.Fields("Observaciones"))
    rsAC.Close
    Set rsAC = Nothing
    
    Set RsTemp = New ADODB.Recordset
    SetearRs "SELECT Presidente FROM actas WHERE " & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ")", RsTemp
    If RsTemp.EOF Then
        IDPresidente = 0
    Else
        If Trim(RsTemp.Fields(0)) = "" Then
            IDPresidente = 0
        Else
            IDPresidente = Val(Trim(RsTemp.Fields(0)))
        End If
    End If
    RsTemp.Close
    Set RsTemp = Nothing
    PresidenteVotosAfirmativos = 0
    PresidenteVotosNegativos = 0
    PresidenteAbstenciones = 0
    Set rsPresi = New ADODB.Recordset
    SetearRs "SELECT resultado_voto_presidente FROM Actas WHERE presidente_habilitado_votar = 1 AND " & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ")", rsPresi
    If rsPresi.EOF Then 'Reviso si votó el presidente, y qué votó
        PresidenteVotosAfirmativos = 0
        PresidenteVotosNegativos = 0
        PresidenteAbstenciones = 0
    Else
        Select Case Trim(LCase(rsPresi.Fields(0)))
        Case "s"
            PresidenteVotosAfirmativos = 1
        Case "n"
            PresidenteVotosNegativos = 1
        Case ""
            PresidenteAbstenciones = 1
        End Select
    End If
    rsPresi.Close
    Set rsPresi = Nothing
    sql_voto_afirmativo_presidente = Trim(Str(PresidenteVotosAfirmativos))
    sql_voto_abstencion_presidente = Trim(Str(PresidenteAbstenciones))
    sql_voto_negativo_presidente = Trim(Str(PresidenteVotosNegativos))
    'Set fViewer = New frmVisor
    'leyenda tipo periodo legislativo. Cambiar los CASE WHEN
    If strTipoOperacion = "votnum" Then
'        sql = "SELECT * From actas " & _
'              " WHERE Período_Legislativo = '" & pPeriodo & "' " & _
'              " AND sesión = " & pSesion & _
'              " AND Número_de_Acta = " & pActa & _
'              " AND Versión_Acta =" & pVersion
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables, actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa,actas.Reunion, actas.Tipo_de_operación , Legisladores.distrito,tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
              " WHERE (Actas.Período_Legislativo='" & pPeriodo & "')" & _
              " AND (Actas.Sesión=" & pSesion & ")" & _
              " AND (Actas.Número_de_Acta=" & pActa & ")" & _
              " AND (Actas.Versión_Acta=" & pVersion & ")" & _
              " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
        
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, 0 AS Votos_Afirm_Identificables, (actas.Votos_Afirm_Identificables + actas.Votos_Afirm_No_Identificables) AS Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, O AS Votos_Neg_Identificables , (actas.Votos_Neg_Identificables + actas.Votos_Neg_No_Identificables) AS Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, " & _
              " actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa,actas.Reunion, actas.Tipo_de_operación , Legisladores.distrito,tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
              " WHERE (Actas.Período_Legislativo='" & pPeriodo & "')" & _
              " AND (Actas.Sesión=" & pSesion & ")" & _
              " AND (Actas.Número_de_Acta=" & pActa & ")" & _
              " AND (Actas.Versión_Acta=" & pVersion & ")" & _
              " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
        
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, 0 AS Votos_Afirm_Identificables, (actas.Votos_Afirm_Identificables + actas.Votos_Afirm_No_Identificables + " & _
        sql_voto_afirmativo_presidente & ") AS Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, 0 AS Votos_Neg_Identificables , (actas.Votos_Neg_Identificables + actas.Votos_Neg_No_Identificables + " & sql_voto_negativo_presidente & ") AS Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total , actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion" & _
              ", actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa,actas.Reunion, actas.Tipo_de_operación , Legisladores.distrito,tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Ordinario' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Prórroga' WHEN 'L' THEN 'Legislativo' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'P' THEN 'Preparatoria 'WHEN 'I' THEN 'Informativa' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
              " WHERE (Actas.Período_Legislativo='" & pPeriodo & "')" & _
              " AND (Actas.Sesión=" & pSesion & ")" & _
              " AND (Actas.Número_de_Acta=" & pActa & ")" & _
              " AND (Actas.Versión_Acta=" & pVersion & ")" & _
              " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
        
        'm_Report.Section5.Suppress = True
        
        m_Report.Detail.Visible = False
        'm_Report.Cuadro1.Suppress = True
        m_Report.delimitador1.Visible = False
        m_Report.delimitador2.Visible = False
        m_Report.delimitador3.Visible = False
        m_Report.delimitador4.Visible = False
        m_Report.Label2.Visible = False
        m_Report.Label3.Visible = False
        m_Report.Campo25.Visible = False
        m_Report.Campo26.Visible = False
        m_Report.Campo23.DataField = "Abstenciones_Total"
        m_Report.lblVotacion.Caption = "Votación Numérica"
        'm_Report.Texto25.Suppress = True
        m_Report.Texto25.Visible = False
        'm_Report.Línea5.Suppress = True
        m_Report.Línea5.Visible = False
        'm_Report.Texto23.Suppress = True
        m_Report.Texto23.Visible = False
        'm_Report.Texto1.Suppress = True
        m_Report.Texto1.Visible = False
        'm_Report.Línea6.Suppress = True
        m_Report.Línea6.Visible = False
        'm_Report.Línea7.Suppress = True
        m_Report.Línea7.Visible = False
        'm_Report.Línea8.Suppress = True
        'm_Report.Línea8.Visible = False
        'm_Report.Línea9.Suppress = True
        'm_Report.Línea9.Visible = False
        'm_Report.Texto22.Suppress = True
        m_Report.Texto22.Visible = False
                        
        With m_Report
            '.Campo14.Visible = False
            '.Campo18.Visible = False
            '.Campo22.Visible = False
            '.Texto18.Visible = False
            .Texto19.Caption = "Diputados"
            .lblPresidente.Visible = False
            .vap.Visible = False
            .vnp.Visible = False
            .vabsp.Visible = False
        End With
            
    Else
        m_Report.Width = m_Report.Width - 3000
        'version sin voto presidente
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables, actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
              
        'version con voto presidente, pero sin restar de la columna de identificados
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables, actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                " AND actas.Sesión = " & _
                " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
    
        'Version restando el voto de la columna identificados
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - " & sql_voto_afirmativo_presidente & " as Votos_Afirm_Identificables " & _
                ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables - " & sql_voto_negativo_presidente & " as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables - " & sql_voto_abstencion_presidente & " as Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                " AND actas.Sesión = " & _
                " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
    
'        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - " & sql_voto_afirmativo_presidente & " as Votos_Afirm_Identificables " & _
                ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables - " & sql_voto_negativo_presidente & " as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables - " & sql_voto_abstencion_presidente & " as Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                " AND actas.Sesión = " & _
                " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (RTrim(DetalleActas.Resultado) <> 'AUSENTE') AND (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
                
        sql = " SELECT " & TotalPaginas & " AS Ultima_Pagina, DetalleActas.Grupo_Politico, DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - 0 as Votos_Afirm_Identificables " & _
                ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, actas.Votos_Neg_Identificables - 0 as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total, actas.Abstenciones_Identificables - " & sql_voto_abstencion_presidente & " as Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total + " & sql_voto_abstencion_presidente & " AS Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                " AND actas.Sesión = " & _
                " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") AND LTrim(DetalleActas.Legislador_Asignado) <> " & IDPresidente & _
                " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "

        m_Report.Campo23.DataField = "Abstenciones_Identificables"
        m_Report.Campo15.DataField = "Votos_Afirm_Identificables"
        m_Report.Campo19.DataField = "Votos_Neg_Identificables"
        m_Report.Campo24.DataField = "Abstenciones_Identificables"
    End If
    If Trim(txtTipoOperacion.Text) = "Pase de Lista" Or strTipoOperacion = "paslis" Then
        Const Corrimiento As Integer = 3250
        With m_Report
        .lblVotacion.Caption = "Pase de Lista"
        .Texto12.Caption = "Resultado :"
        'Se eliminan todos los textos
        '.Texto18.Visible = False
        .Texto19.Visible = False
        .Texto20.Visible = False
        .Texto21.Visible = False
        .Texto13.Visible = False
        .Texto14.Visible = False
        .Texto15.Visible = False
        .Texto2.Visible = False
        .Texto8.Visible = False
        .Texto7.Visible = False
        '.Label1.Visible = False
        'Se eliminan todos los campos con Datafields
        '.Campo14.Visible = False
        .Campo15.Visible = False
        .Campo16.Visible = False
        .Campo17.Visible = False
        '.Campo18.Visible = False
        .Campo19.Visible = False
        .Campo20.Visible = False
        .Campo21.Visible = False
        '.Campo22.Visible = False
        .Campo23.Visible = False
        .Campo24.Visible = False
        .Campo33.Visible = False
        .Campo24.Visible = False
        .Campo33.Visible = False
        .Campo30.Visible = False
        .Campo35.Visible = False
        .Campo6.Visible = False
        'Se elimina el delimitador
        .Line7.Visible = False
        'Se acomodan los controles
        .Label2.Left = .Label2.Left + Corrimiento
        .Label3.Left = .Label3.Left + Corrimiento
        .Label4.Left = .Label4.Left + Corrimiento
        .Texto16.Left = .Texto16.Left + Corrimiento
        .Texto17.Left = .Texto17.Left + Corrimiento
        .Campo25.Left = .Campo25.Left + Corrimiento
        .Campo26.Left = .Campo26.Left + Corrimiento
        .Campo27.Left = .Campo27.Left + Corrimiento
        .Campo28.Left = .Campo28.Left + Corrimiento
        .lblPresidente.Visible = False
        .vap.Visible = False
        .vnp.Visible = False
        .vabsp.Visible = False
        End With
    End If
    SetearRs sql, rstActa
    'm_Report.Database.SetDataSource rstActa
    Set m_Report.DataControl1.Recordset = rstActa
    If False Then
        m_Report.Show vbModal, Me
    Else
        m_Report.Run False
        Dim X As Integer
        TodoElReporte.Run
        For X = 0 To m_Report.Pages.Count - 1
            TodoElReporte.Pages.Insert X, m_Report.Pages(X)
            TodoElReporte.Pages.Commit
        Next X
        TodoElReporte.Pages.Commit
    End If
'    fViewer.CRViewer1.ReportSource = m_Report
'    If PermisosTotales.ImprimeActas = 1 Then
'        fViewer.CRViewer1.EnablePrintButton = True
'    Else
'        fViewer.CRViewer1.EnablePrintButton = False
'    End If
'
'    fViewer.CRViewer1.ViewReport
'    fViewer.CRViewer1.Zoom 100
'    fViewer.Show vbModal
    
    Set rstActa = Nothing
    'Set fViewer = Nothing
    Set m_Report = Nothing
    ObsActa = ""
Exit Sub
TrapError:
    Select Case err.Number
        Case Else
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            End
    End Select
Return
End Sub

Public Function ObtenerCantidadDePaginas(strTipoOperacion As String, pPeriodo As String, pSesion As Integer, pActa As Integer, pVersion As Integer) As String
    'On Error GoTo TrapError
    'IMPRESION ACTA 2009
    Dim IDPresidente As Integer
    If PermisosTotales.ConsultaActas = 0 Then
        MsgBox "No posee permisos para consulta de actas", vbInformation + vbOKOnly, "Usuario sin permisos!!"
        Exit Function
    End If
    
    Dim m_Report As New rptActas
    Dim rstActa  As New ADODB.Recordset
    'Dim fViewer  As frmVisor
    Dim sql      As String
    Dim sql_voto_afirmativo_presidente As String
    Dim sql_voto_negativo_presidente As String
    Dim sql_voto_abstencion_presidente As String
    Dim rsPresi As ADODB.Recordset
    Dim PresidenteVotosAfirmativos As Integer
    Dim PresidenteVotosNegativos As Integer
    Dim PresidenteAbstenciones As Integer
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    SetearRs "SELECT Presidente FROM actas WHERE " & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ")", RsTemp
    If RsTemp.EOF Then
        IDPresidente = 0
    Else
        If Trim(RsTemp.Fields(0)) = "" Then
            IDPresidente = 0
        Else
            IDPresidente = Val(Trim(RsTemp.Fields(0)))
        End If
    End If
    RsTemp.Close
    Set RsTemp = Nothing
    PresidenteVotosAfirmativos = 0
    PresidenteVotosNegativos = 0
    PresidenteAbstenciones = 0
    Set rsPresi = New ADODB.Recordset
    SetearRs "SELECT resultado_voto_presidente FROM Actas WHERE presidente_habilitado_votar = 1 AND " & "(Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ")", rsPresi
    If rsPresi.EOF Then 'Reviso si votó el presidente, y qué votó
        PresidenteVotosAfirmativos = 0
        PresidenteVotosNegativos = 0
        PresidenteAbstenciones = 0
    Else
        Select Case Trim(LCase(rsPresi.Fields(0)))
        Case "s"
            PresidenteVotosAfirmativos = 1
        Case "n"
            PresidenteVotosNegativos = 1
        Case ""
            PresidenteAbstenciones = 1
        End Select
    End If
    rsPresi.Close
    Set rsPresi = Nothing
    sql_voto_afirmativo_presidente = Trim(Str(PresidenteVotosAfirmativos))
    sql_voto_abstencion_presidente = Trim(Str(PresidenteAbstenciones))
    sql_voto_negativo_presidente = Trim(Str(PresidenteVotosNegativos))
    'Set fViewer = New frmVisor
    'leyenda tipo periodo legislativo. Cambiar los CASE WHEN
    If strTipoOperacion = "votnum" Then
'        sql = "SELECT * From actas " & _
'              " WHERE Período_Legislativo = '" & pPeriodo & "' " & _
'              " AND sesión = " & pSesion & _
'              " AND Número_de_Acta = " & pActa & _
'              " AND Versión_Acta =" & pVersion
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables, actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa,actas.Reunion, actas.Tipo_de_operación , Legisladores.distrito,tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
              " WHERE (Actas.Período_Legislativo='" & pPeriodo & "')" & _
              " AND (Actas.Sesión=" & pSesion & ")" & _
              " AND (Actas.Número_de_Acta=" & pActa & ")" & _
              " AND (Actas.Versión_Acta=" & pVersion & ")" & _
              " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
        
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, 0 AS Votos_Afirm_Identificables, (actas.Votos_Afirm_Identificables + actas.Votos_Afirm_No_Identificables) AS Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, O AS Votos_Neg_Identificables , (actas.Votos_Neg_Identificables + actas.Votos_Neg_No_Identificables) AS Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, " & _
              " actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa,actas.Reunion, actas.Tipo_de_operación , Legisladores.distrito,tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
              " WHERE (Actas.Período_Legislativo='" & pPeriodo & "')" & _
              " AND (Actas.Sesión=" & pSesion & ")" & _
              " AND (Actas.Número_de_Acta=" & pActa & ")" & _
              " AND (Actas.Versión_Acta=" & pVersion & ")" & _
              " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
        
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, 0 AS Votos_Afirm_Identificables, (actas.Votos_Afirm_Identificables + actas.Votos_Afirm_No_Identificables + " & _
        sql_voto_afirmativo_presidente & ") AS Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, 0 AS Votos_Neg_Identificables , (actas.Votos_Neg_Identificables + actas.Votos_Neg_No_Identificables + " & sql_voto_negativo_presidente & ") AS Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total , actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion" & _
              ", actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa,actas.Reunion, actas.Tipo_de_operación , Legisladores.distrito,tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Ordinario' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Prórroga' WHEN 'L' THEN 'Legislativo' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'P' THEN 'Preparatoria 'WHEN 'I' THEN 'Informativa' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
              " WHERE (Actas.Período_Legislativo='" & pPeriodo & "')" & _
              " AND (Actas.Sesión=" & pSesion & ")" & _
              " AND (Actas.Número_de_Acta=" & pActa & ")" & _
              " AND (Actas.Versión_Acta=" & pVersion & ")" & _
              " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
        
        'm_Report.Section5.Suppress = True
        m_Report.Detail.Visible = False
        'm_Report.Cuadro1.Suppress = True
        m_Report.delimitador1.Visible = False
        m_Report.delimitador2.Visible = False
        m_Report.delimitador3.Visible = False
        m_Report.delimitador4.Visible = False
        m_Report.Label2.Visible = False
        m_Report.Label3.Visible = False
        m_Report.Campo25.Visible = False
        m_Report.Campo26.Visible = False
        m_Report.Campo23.DataField = "Abstenciones_Total"
        m_Report.lblVotacion.Caption = "Votación Numérica"
        'm_Report.Texto25.Suppress = True
        m_Report.Texto25.Visible = False
        'm_Report.Línea5.Suppress = True
        m_Report.Línea5.Visible = False
        'm_Report.Texto23.Suppress = True
        m_Report.Texto23.Visible = False
        'm_Report.Texto1.Suppress = True
        m_Report.Texto1.Visible = False
        'm_Report.Línea6.Suppress = True
        m_Report.Línea6.Visible = False
        'm_Report.Línea7.Suppress = True
        m_Report.Línea7.Visible = False
        'm_Report.Línea8.Suppress = True
        'm_Report.Línea8.Visible = False
        'm_Report.Línea9.Suppress = True
        'm_Report.Línea9.Visible = False
        'm_Report.Texto22.Suppress = True
        m_Report.Texto22.Visible = False
                        
        With m_Report
            '.Campo14.Visible = False
            '.Campo18.Visible = False
            '.Campo22.Visible = False
            '.Texto18.Visible = False
            .Texto19.Caption = "Diputados"
            .lblPresidente.Visible = False
            .vap.Visible = False
            .vnp.Visible = False
            .vabsp.Visible = False
        End With
            
    Else
        'version sin voto presidente
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables, actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
              
        'version con voto presidente, pero sin restar de la columna de identificados
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables, actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                " AND actas.Sesión = " & _
                " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
    
        'Version restando el voto de la columna identificados
        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - " & sql_voto_afirmativo_presidente & " as Votos_Afirm_Identificables " & _
                ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables - " & sql_voto_negativo_presidente & " as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables - " & sql_voto_abstencion_presidente & " as Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                " AND actas.Sesión = " & _
                " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
    
'        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - " & sql_voto_afirmativo_presidente & " as Votos_Afirm_Identificables " & _
                ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables - " & sql_voto_negativo_presidente & " as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables - " & sql_voto_abstencion_presidente & " as Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ ' Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'Ordinaria' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                " AND actas.Sesión = " & _
                " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN distritos ON DetalleActas.Departamento = distritos.id_distrito LEFT OUTER JOIN secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (RTrim(DetalleActas.Resultado) <> 'AUSENTE') AND (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "
                
        sql = " SELECT " & TotalPaginas & " AS Ultima_Pagina, DetalleActas.Grupo_Politico, DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables - 0 as Votos_Afirm_Identificables " & _
                ", actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total + " & sql_voto_afirmativo_presidente & " AS Votos_Afirm_Total, actas.Votos_Neg_Identificables - 0 as Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total + " & sql_voto_negativo_presidente & " AS Votos_Neg_Total, actas.Abstenciones_Identificables - " & sql_voto_abstencion_presidente & " as Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total + " & sql_voto_abstencion_presidente & " AS Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
                " CASE CONVERT(varchar(5),actas.Ultima_Versión_Acta) WHEN '0' THEN 'Original' ELSE 'Ult.Mod.Ver ' + CONVERT(varchar(5),actas.Ultima_Versión_Acta + 1) END AS VersionActa, actas.Reunion, actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
                " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, " & SQL_Provincia & " AS Distrito,RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador,RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + Case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN ' Período Ordinario' WHEN 'E' THEN ' Período Extraordinario' WHEN 'P' THEN ' Prórroga Ordinario' END + ' - ' + (CASE actas.Sesión WHEN -1 THEN 'Sesión ' ELSE CAST(actas.Sesión AS Varchar(5)) + 'ª Sesión ' END) + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - ' + CONVERT(varchar(10),actas.Reunion) + 'º Reunión' END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
                "," & sql_voto_afirmativo_presidente & " AS voto_afirmativo_presidente " & _
                "," & sql_voto_negativo_presidente & " AS voto_negativo_presidente " & _
                "," & sql_voto_abstencion_presidente & " AS voto_abstencion_presidente " & _
                " FROM actas INNER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
                " AND actas.Sesión = " & _
                " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_De_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
                " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes INNER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado WHERE (Actas.Período_Legislativo='" & pPeriodo & "') AND (Actas.Sesión=" & pSesion & ") AND (Actas.Número_de_Acta=" & pActa & ") AND (Actas.Versión_Acta=" & pVersion & ") AND LTrim(DetalleActas.Legislador_Asignado) <> " & IDPresidente & " ORDER BY P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) "

        m_Report.Campo23.DataField = "Abstenciones_Identificables"
        m_Report.Campo15.DataField = "Votos_Afirm_Identificables"
        m_Report.Campo19.DataField = "Votos_Neg_Identificables"
        m_Report.Campo24.DataField = "Abstenciones_Identificables"
    End If
    If Trim(txtTipoOperacion.Text) = "Pase de Lista" Or strTipoOperacion = "paslis" Then
        Const Corrimiento As Integer = 3250
        With m_Report
        .lblVotacion.Caption = "Pase de Lista"
        .Texto12.Caption = "Resultado :"
        'Se eliminan todos los textos
        '.Texto18.Visible = False
        .Texto19.Visible = False
        .Texto20.Visible = False
        .Texto21.Visible = False
        .Texto13.Visible = False
        .Texto14.Visible = False
        .Texto15.Visible = False
        .Texto2.Visible = False
        .Texto8.Visible = False
        .Texto7.Visible = False
        '.Label1.Visible = False
        'Se eliminan todos los campos con Datafields
        '.Campo14.Visible = False
        .Campo15.Visible = False
        .Campo16.Visible = False
        .Campo17.Visible = False
        '.Campo18.Visible = False
        .Campo19.Visible = False
        .Campo20.Visible = False
        .Campo21.Visible = False
        '.Campo22.Visible = False
        .Campo23.Visible = False
        .Campo24.Visible = False
        .Campo33.Visible = False
        .Campo24.Visible = False
        .Campo33.Visible = False
        .Campo30.Visible = False
        .Campo35.Visible = False
        .Campo6.Visible = False
        'Se elimina el delimitador
        .Line7.Visible = False
        'Se acomodan los controles
        .Label2.Left = .Label2.Left + Corrimiento
        .Label3.Left = .Label3.Left + Corrimiento
        .Label4.Left = .Label4.Left + Corrimiento
        .Texto16.Left = .Texto16.Left + Corrimiento
        .Texto17.Left = .Texto17.Left + Corrimiento
        .Campo25.Left = .Campo25.Left + Corrimiento
        .Campo26.Left = .Campo26.Left + Corrimiento
        .Campo27.Left = .Campo27.Left + Corrimiento
        .Campo28.Left = .Campo28.Left + Corrimiento
        .lblPresidente.Visible = False
        .vap.Visible = False
        .vnp.Visible = False
        .vabsp.Visible = False
        End With
    End If
    SetearRs sql, rstActa
    'm_Report.Database.SetDataSource rstActa
    Set m_Report.DataControl1.Recordset = rstActa
    If True Then
        m_Report.Run False
        ObtenerCantidadDePaginas = UltimaPaginaReporte
    Else
        'm_Report.Printer.Copies = 1
        'm_Report.Printer.StartJob "acta"
        m_Report.PrintReport False
    End If
'    fViewer.CRViewer1.ReportSource = m_Report
'    If PermisosTotales.ImprimeActas = 1 Then
'        fViewer.CRViewer1.EnablePrintButton = True
'    Else
'        fViewer.CRViewer1.EnablePrintButton = False
'    End If
'
'    fViewer.CRViewer1.ViewReport
'    fViewer.CRViewer1.Zoom 100
'    fViewer.Show vbModal
    
    Set rstActa = Nothing
    'Set fViewer = Nothing
    Set m_Report = Nothing
Exit Function
TrapError:
    Select Case err.Number
        Case Else
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            End
    End Select
Return
End Function

Public Function TieneManifestaciones() As Boolean
Dim s As String
Dim rs As New ADODB.Recordset
s = "SELECT * FROM manifestaciones_vivavoz" & _
" WHERE manifestaciones_vivavoz.periodo = '" & mPeriodo & "'" & _
" AND manifestaciones_vivavoz.sesion = " & mSesion & _
" AND manifestaciones_vivavoz.nro_acta = " & mActa & _
" AND manifestaciones_vivavoz.version_acta = " & Str(xUltimaVersionActa + 1)
Call SetearRs(s, rs)
If (rs.EOF) Then
    TieneManifestaciones = False
    Exit Function
End If
TieneManifestaciones = True
End Function

Public Sub GuardarManifestaciones()
Dim s As String
s = "UPDATE manifestaciones_vivavoz SET version_acta = " & Str(xUltimaVersionActa + 2) & _
" WHERE manifestaciones_vivavoz.periodo = '" & mPeriodo & "'" & _
" AND manifestaciones_vivavoz.sesion = " & mSesion & _
" AND manifestaciones_vivavoz.nro_acta = " & mActa & _
" AND manifestaciones_vivavoz.version_acta = -1"
Call InsertSQL(s)
mVivavozChanged = False
End Sub

Public Sub CopiarManifestaciones()
Dim s As String
s = "UPDATE manifestaciones_vivavoz SET version_acta = " & Str(xUltimaVersionActa + 2) & _
" WHERE manifestaciones_vivavoz.periodo = '" & mPeriodo & "'" & _
" AND manifestaciones_vivavoz.sesion = " & mSesion & _
" AND manifestaciones_vivavoz.nro_acta = " & mActa & _
" AND manifestaciones_vivavoz.version_acta = " & Str(xUltimaVersionActa + 1)
Call InsertSQL(s)
mVivavozChanged = False
End Sub

