VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "SQVServer"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   Visible         =   0   'False
   Begin VB.Frame FrameSQVGeneral 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Caption         =   "SQV General"
      ForeColor       =   &H80000008&
      Height          =   12000
      Left            =   0
      TabIndex        =   244
      Top             =   0
      Width           =   15360
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   11520
         Left            =   270
         ScaleHeight     =   11520
         ScaleWidth      =   15360
         TabIndex        =   245
         Top             =   270
         Width           =   15360
         Begin VB.PictureBox picA 
            BorderStyle     =   0  'None
            Height          =   5555
            Index           =   0
            Left            =   0
            ScaleHeight     =   5550
            ScaleWidth      =   18000
            TabIndex        =   246
            Top             =   0
            Width           =   18000
            Begin MSWinsockLib.Winsock Ws 
               Left            =   270
               Top             =   210
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin VB.Shape shpRecuadroOrador 
               BorderColor     =   &H00FFFFFF&
               Height          =   1575
               Left            =   3540
               Top             =   3780
               Width           =   8775
            End
            Begin VB.Shape shpTitulo 
               BorderColor     =   &H00000080&
               BorderWidth     =   3
               FillColor       =   &H00FFFFFF&
               Height          =   915
               Left            =   1440
               Top             =   1500
               Width           =   4335
            End
            Begin VB.Label lblCAusentes 
               BackStyle       =   0  'Transparent
               Caption         =   "AUSENTES"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   39.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   975
               Left            =   7920
               TabIndex        =   372
               Top             =   2700
               Width           =   4695
            End
            Begin VB.Label lblCPresentes 
               BackStyle       =   0  'Transparent
               Caption         =   "PRESENTES"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   39.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   1095
               Left            =   8700
               TabIndex        =   371
               Top             =   1260
               Width           =   4395
            End
            Begin VB.Shape shpRecuadroQuorum 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               FillColor       =   &H00FFFFFF&
               Height          =   975
               Left            =   8580
               Top             =   240
               Width           =   4395
            End
            Begin VB.Shape shpHora 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   3
               Height          =   975
               Left            =   4740
               Top             =   60
               Width           =   2415
            End
            Begin VB.Shape shpRecuadroFecha 
               BorderColor     =   &H00FFFFFF&
               Height          =   495
               Left            =   1380
               Top             =   300
               Width           =   2895
            End
            Begin VB.Label lblGeneralSesionDato 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Datos de la Sesión"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   795
               Index           =   5
               Left            =   300
               TabIndex        =   368
               Top             =   3360
               Visible         =   0   'False
               Width           =   8925
            End
            Begin VB.Label lblGeneralSesionDato 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Datos de la Sesión"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   795
               Index           =   4
               Left            =   300
               TabIndex        =   367
               Top             =   2520
               Visible         =   0   'False
               Width           =   8925
            End
            Begin VB.Label lblGeneralSesionDato 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Datos de la Sesión"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   795
               Index           =   3
               Left            =   300
               TabIndex        =   366
               Top             =   1620
               Visible         =   0   'False
               Width           =   8925
            End
            Begin VB.Label lblVersionCartel 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   735
               Left            =   2160
               TabIndex        =   363
               Top             =   840
               Width           =   3855
            End
            Begin VB.Label lblGeneralLeyendaQuorumDato 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "NO HAY QUORUMX"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   48
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   1155
               Left            =   7080
               TabIndex        =   251
               Top             =   0
               Width           =   8055
            End
            Begin VB.Label lblGeneralHoraDato 
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "00:00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   36
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   915
               Left            =   4920
               TabIndex        =   250
               Top             =   0
               Width           =   2505
            End
            Begin VB.Label lblGeneralFechaDato 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "00/00/00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   36
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   795
               Left            =   180
               TabIndex        =   249
               Top             =   0
               Width           =   4290
            End
            Begin VB.Label lblGeneralAusentesDato 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   48
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   1155
               Left            =   14040
               TabIndex        =   248
               Top             =   3000
               Width           =   1410
            End
            Begin VB.Label lblGeneralPresentesDato 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   48
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   1275
               Left            =   14040
               TabIndex        =   247
               Top             =   1440
               Width           =   1410
            End
            Begin VB.Image imgA 
               Height          =   4500
               Index           =   0
               Left            =   240
               Picture         =   "frmMainAEB.frx":0000
               Top             =   -120
               Width           =   15360
            End
         End
         Begin VB.PictureBox picC 
            BackColor       =   &H80000008&
            BorderStyle     =   0  'None
            Height          =   7290
            Index           =   1
            Left            =   240
            ScaleHeight     =   7290
            ScaleWidth      =   18675
            TabIndex        =   262
            Top             =   3960
            Visible         =   0   'False
            Width           =   18675
            Begin MSComctlLib.ImageList iml 
               Left            =   9600
               Top             =   1080
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   24
               ImageHeight     =   24
               MaskColor       =   12632256
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   5
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmMainAEB.frx":23B5
                     Key             =   "afirmativo"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmMainAEB.frx":2AC9
                     Key             =   "negativo"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmMainAEB.frx":31DD
                     Key             =   "abstencion"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmMainAEB.frx":38F1
                     Key             =   "presente"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmMainAEB.frx":45E2
                     Key             =   "ausente"
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.ListView lsv 
               Height          =   480
               Left            =   7800
               TabIndex        =   263
               Top             =   600
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   847
               View            =   2
               Sorted          =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   393217
               Icons           =   "iml"
               SmallIcons      =   "iml"
               ForeColor       =   0
               BackColor       =   -2147483643
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
            Begin VB.Label Label73 
               Caption         =   "Label73"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   0
               TabIndex        =   377
               Top             =   0
               Visible         =   0   'False
               Width           =   2415
            End
            Begin VB.Label lblOperador1 
               Caption         =   "Operador1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4020
               TabIndex        =   376
               Top             =   5760
               Width           =   2415
            End
            Begin VB.Label lblLeyendaVotoAfirmativo 
               BackStyle       =   0  'Transparent
               Caption         =   "AFIRMATIVOS NEGATIVOS ABSTENCIONES"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   39.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   1335
               Left            =   480
               TabIndex        =   373
               Top             =   2640
               Width           =   5835
            End
            Begin VB.Shape shpBanca 
               FillColor       =   &H000000FF&
               FillStyle       =   0  'Solid
               Height          =   345
               Index           =   0
               Left            =   2400
               Shape           =   2  'Oval
               Top             =   6360
               Width           =   405
            End
            Begin VB.Label lblTituloBaseYTipoDeMayoria 
               BackColor       =   &H80000007&
               Caption         =   "Tipo y Base de mayoría"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   32.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   1545
               Left            =   240
               TabIndex        =   361
               Top             =   1440
               Width           =   5295
            End
            Begin VB.Label lblGeneralMayoriaDato 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Más de la mitad de los Legisladores Presentes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   32.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   1425
               Index           =   2
               Left            =   7440
               TabIndex        =   349
               Top             =   1440
               Width           =   7305
            End
            Begin VB.Label lblOrador03 
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Orador03"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   855
               Left            =   240
               TabIndex        =   355
               Top             =   2160
               Width           =   11535
            End
            Begin VB.Label lblOrador04 
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Orador04"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   855
               Left            =   240
               TabIndex        =   356
               Top             =   3120
               Width           =   14535
            End
            Begin VB.Label lblOrador02 
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Caption         =   "Orador02"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   855
               Left            =   240
               TabIndex        =   354
               Top             =   1200
               Width           =   14295
            End
            Begin VB.Label lblGeneralTituloTiempo 
               BackColor       =   &H80000007&
               Caption         =   "TIEMPO"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   56.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   1305
               Left            =   2640
               TabIndex        =   362
               Top             =   3600
               Width           =   4575
            End
            Begin VB.Label lblGeneralTiempoDato 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "CANCELADA"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   56.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   1305
               Left            =   7320
               TabIndex        =   350
               Top             =   3600
               Width           =   7500
            End
            Begin VB.Label lblPresentesIdentificados 
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Caption         =   "PI"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   705
               Left            =   12120
               TabIndex        =   360
               Top             =   1200
               Width           =   1935
            End
            Begin VB.Label lblTituloPresentesIdentificados 
               BackColor       =   &H80000007&
               Caption         =   "Presentes identificados:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   705
               Left            =   1920
               TabIndex        =   359
               Top             =   1200
               Width           =   9855
            End
            Begin VB.Label lblOrador01 
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Orador01"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   855
               Left            =   240
               TabIndex        =   353
               Top             =   240
               Width           =   12015
            End
            Begin VB.Label lblGeneralSesionDato 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Datos de la Sesión"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   675
               Index           =   2
               Left            =   840
               TabIndex        =   347
               Top             =   120
               Visible         =   0   'False
               Width           =   14805
            End
            Begin VB.Label lblGeneralTituloDato 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Asunto en votacion Exp. 2030S/09 - Orden del Dia 48 1234567890 abcdefghijklmno"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   32.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   855
               Index           =   4
               Left            =   240
               TabIndex        =   348
               Top             =   240
               Width           =   14505
            End
            Begin VB.Image imgC 
               Height          =   6810
               Index           =   1
               Left            =   8040
               Picture         =   "frmMainAEB.frx":4CF6
               Top             =   -5160
               Width           =   15360
            End
            Begin VB.Label lblTituloOcupadosNoIdentificados 
               BackColor       =   &H80000007&
               Caption         =   "Pendientes de identificarse:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   705
               Index           =   0
               Left            =   1920
               TabIndex        =   358
               Top             =   2040
               Width           =   9855
            End
            Begin VB.Label lblOcupadosNoIdentificados 
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "OCNI"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   705
               Index           =   1
               Left            =   12120
               TabIndex        =   357
               Top             =   2040
               Width           =   1935
            End
         End
         Begin VB.PictureBox picC 
            BorderStyle     =   0  'None
            Height          =   8145
            Index           =   0
            Left            =   0
            ScaleHeight     =   8145
            ScaleWidth      =   15930
            TabIndex        =   264
            Top             =   3000
            Width           =   15930
            Begin VB.Label lblOcupadosNoIdentificados 
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "OCNI"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   36
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   825
               Index           =   2
               Left            =   13680
               TabIndex        =   369
               Top             =   4050
               Width           =   1455
            End
            Begin VB.Label lblTituloOcupadosNoIdentificados 
               BackColor       =   &H80000007&
               Caption         =   "Legisladores sin identificar:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   32.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   1545
               Index           =   1
               Left            =   9840
               TabIndex        =   370
               Top             =   3360
               Width           =   5655
            End
            Begin VB.Label lblGeneralMayoriaDato 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "MITAD DE LOS PRESENTES"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   585
               Index           =   3
               Left            =   9000
               TabIndex        =   352
               Top             =   5760
               Visible         =   0   'False
               Width           =   5745
            End
            Begin VB.Label lblGeneralTipoOperacionDato 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "PASAJE DE LISTA"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   36
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   915
               Left            =   840
               TabIndex        =   345
               Top             =   5520
               Visible         =   0   'False
               Width           =   6735
            End
            Begin VB.Label lblGeneralResultadoDato 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   48
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   1035
               Left            =   4560
               TabIndex        =   344
               Top             =   240
               Width           =   6255
            End
            Begin VB.Label lblGeneralAbstencionesDato 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   48
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   1005
               Left            =   8040
               TabIndex        =   343
               Top             =   3720
               Width           =   1560
            End
            Begin VB.Label lblGeneralNegativosDato 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   48
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   1125
               Left            =   8040
               TabIndex        =   342
               Top             =   2640
               Width           =   1560
            End
            Begin VB.Label lblGeneralAfirmativosDato 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   48
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   1005
               Left            =   8040
               TabIndex        =   341
               Top             =   1560
               Width           =   1560
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   2
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   335
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   3
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   334
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   4
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   333
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   5
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   332
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   6
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   331
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   7
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   330
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   8
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   329
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   9
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   328
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   10
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   327
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   11
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   326
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   12
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   325
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   13
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   324
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   14
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   323
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   15
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   322
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   16
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   321
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   17
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   320
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   18
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   319
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   19
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   318
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   20
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   317
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   21
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   316
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   22
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   315
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   23
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   314
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   24
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   313
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   25
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   312
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   26
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   311
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   27
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   310
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   28
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   309
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   29
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   308
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   30
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   307
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   31
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   306
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   32
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   305
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   33
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   304
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   34
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   303
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   35
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   302
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   36
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   301
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   37
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   300
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   38
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   299
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   39
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   298
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   40
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   297
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   41
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   296
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   42
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   295
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   43
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   294
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   44
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   293
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   45
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   292
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   46
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   291
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   47
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   290
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   48
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   289
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   49
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   288
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   50
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   287
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   51
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   286
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   52
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   285
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   53
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   284
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   54
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   283
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   55
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   282
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   56
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   281
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   57
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   280
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   58
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   279
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   59
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   278
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   60
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   277
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   61
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   276
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   62
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   275
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   63
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   274
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   64
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   273
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   65
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   272
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   66
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   271
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   67
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   270
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   68
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   269
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   69
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   268
               Top             =   105
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   70
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   195
               TabIndex        =   267
               Top             =   105
               Width           =   420
            End
            Begin VB.Label lblBanca 
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
               Left            =   240
               TabIndex        =   266
               Top             =   540
               Width           =   480
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               Height          =   765
               Index           =   1
               Left            =   0
               Shape           =   3  'Circle
               Top             =   0
               Width           =   795
            End
            Begin VB.Label lblBanca 
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
               Left            =   150
               TabIndex        =   265
               Top             =   120
               Width           =   420
            End
            Begin VB.Shape ctrBanca 
               BorderStyle     =   6  'Inside Solid
               BorderWidth     =   3
               FillColor       =   &H00FFFFFF&
               FillStyle       =   0  'Solid
               Height          =   765
               Index           =   0
               Left            =   60
               Shape           =   3  'Circle
               Top             =   420
               Width           =   795
            End
            Begin VB.Label lblGeneralTituloDato 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Asunto en votación"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   1035
               Index           =   3
               Left            =   0
               TabIndex        =   346
               Top             =   0
               Width           =   9585
            End
            Begin VB.Image imgC 
               Height          =   6810
               Index           =   0
               Left            =   0
               Picture         =   "frmMainAEB.frx":8D0A
               Top             =   0
               Width           =   15360
            End
         End
         Begin VB.PictureBox picB 
            BorderStyle     =   0  'None
            Height          =   3915
            Index           =   4
            Left            =   150
            ScaleHeight     =   3915
            ScaleWidth      =   15360
            TabIndex        =   256
            Top             =   3390
            Width           =   15360
            Begin VB.Label lblGeneralInformacion 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Pulse SI para afirmativo. NO para negativo. No presione ningún botón para abstención."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   23.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   1215
               Left            =   240
               TabIndex        =   257
               Top             =   600
               Width           =   14625
            End
            Begin VB.Image imgB 
               Height          =   1935
               Index           =   4
               Left            =   0
               Picture         =   "frmMainAEB.frx":139CD
               Top             =   0
               Width           =   15360
            End
         End
         Begin VB.PictureBox picB 
            BorderStyle     =   0  'None
            Height          =   2235
            Index           =   1
            Left            =   1080
            ScaleHeight     =   2235
            ScaleWidth      =   15360
            TabIndex        =   258
            Top             =   3390
            Width           =   15360
            Begin VB.Label lblGeneralTituloDato 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Asunto en Tratamiento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   1065
               Index           =   0
               Left            =   300
               TabIndex        =   259
               Top             =   660
               Width           =   14775
            End
            Begin VB.Image imgB 
               Height          =   1935
               Index           =   1
               Left            =   0
               Picture         =   "frmMainAEB.frx":17672
               Top             =   0
               Width           =   15360
            End
         End
         Begin VB.PictureBox picB 
            BorderStyle     =   0  'None
            Height          =   1935
            Index           =   0
            Left            =   0
            ScaleHeight     =   1935
            ScaleWidth      =   15360
            TabIndex        =   254
            Top             =   1320
            Width           =   15360
            Begin VB.Timer TimerPic 
               Interval        =   5000
               Left            =   510
               Top             =   1350
            End
            Begin VB.Label lblGeneralSesionDato 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Datos de la Sesión"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   675
               Index           =   0
               Left            =   315
               TabIndex        =   255
               Top             =   600
               Width           =   14805
            End
            Begin VB.Image imgB 
               Height          =   1935
               Index           =   0
               Left            =   480
               Picture         =   "frmMainAEB.frx":1AF43
               Top             =   0
               Width           =   15360
            End
         End
         Begin VB.PictureBox picB 
            BorderStyle     =   0  'None
            Height          =   1935
            Index           =   3
            Left            =   0
            ScaleHeight     =   1935
            ScaleWidth      =   15360
            TabIndex        =   252
            Top             =   1320
            Width           =   15360
            Begin VB.Label lblGeneralMayoriaDato 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "MITAD DE LOS PRESENTES"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   945
               Index           =   0
               Left            =   7800
               TabIndex        =   253
               Top             =   960
               Width           =   7305
            End
            Begin VB.Image imgB 
               Height          =   1935
               Index           =   3
               Left            =   240
               Picture         =   "frmMainAEB.frx":1E802
               Top             =   0
               Width           =   15360
            End
         End
         Begin VB.PictureBox picB 
            BorderStyle     =   0  'None
            Height          =   1935
            Index           =   2
            Left            =   0
            ScaleHeight     =   1935
            ScaleWidth      =   15360
            TabIndex        =   260
            Top             =   1365
            Width           =   15360
            Begin VB.Label lblGeneralTituloDato 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Asunto en votación"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   1035
               Index           =   1
               Left            =   315
               TabIndex        =   261
               Top             =   645
               Width           =   9585
            End
            Begin VB.Label lblGeneralMayoriaDato 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Mas de la mitad de los legisladores votantes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   1215
               Index           =   1
               Left            =   10140
               TabIndex        =   337
               Top             =   630
               Width           =   5025
            End
            Begin VB.Image imgB 
               Height          =   1935
               Index           =   2
               Left            =   120
               Picture         =   "frmMainAEB.frx":25DA9
               Top             =   0
               Width           =   15360
            End
         End
      End
      Begin VB.PictureBox picA 
         BorderStyle     =   0  'None
         Height          =   1365
         Index           =   1
         Left            =   240
         ScaleHeight     =   1365
         ScaleWidth      =   15360
         TabIndex        =   336
         Top             =   270
         Width           =   15360
         Begin VB.Label lblGeneralTituloDato 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Asunto en Tratamiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   645
            Index           =   2
            Left            =   330
            TabIndex        =   338
            Top             =   720
            Width           =   14775
         End
         Begin VB.Image imgA 
            Height          =   1410
            Index           =   1
            Left            =   120
            Picture         =   "frmMainAEB.frx":2AA5F
            Top             =   0
            Width           =   15360
         End
      End
      Begin VB.PictureBox picA 
         BorderStyle     =   0  'None
         Height          =   1365
         Index           =   2
         Left            =   240
         ScaleHeight     =   1365
         ScaleWidth      =   15360
         TabIndex        =   339
         Top             =   270
         Width           =   15360
         Begin VB.Label lblGeneralSesionDato 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Datos de la sesión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   675
            Index           =   1
            Left            =   330
            TabIndex        =   340
            Top             =   720
            Width           =   14775
         End
         Begin VB.Image imgA 
            Height          =   1410
            Index           =   2
            Left            =   240
            Picture         =   "frmMainAEB.frx":2DC87
            Top             =   0
            Width           =   15360
         End
      End
      Begin VB.Label lblOperador4 
         Caption         =   "Operador4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   380
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lblOperador3 
         Caption         =   "Operador3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         TabIndex        =   379
         Top             =   780
         Width           =   2415
      End
      Begin VB.Label lblOperador2 
         Caption         =   "Operador2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   378
         Top             =   180
         Width           =   2415
      End
   End
   Begin VB.Frame FrameControl 
      Caption         =   "Control de SQV"
      Height          =   7815
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9855
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DrawMode        =   1  'Blackness
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         Picture         =   "frmMainAEB.frx":33F39
         ScaleHeight     =   148.536
         ScaleMode       =   0  'User
         ScaleWidth      =   182.667
         TabIndex        =   364
         Top             =   240
         Width           =   2085
         Begin VB.Label lblSQV 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SQV 4.1"
            Height          =   255
            Left            =   120
            TabIndex        =   365
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.CommandButton bCartelApagado 
         Caption         =   "Apagar Cartel"
         Height          =   375
         Left            =   4200
         TabIndex        =   351
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Frame frmEstadoCartel 
         Caption         =   "Estado de Cartel : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   0
         TabIndex        =   76
         Top             =   1680
         Width           =   9675
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Presentes : "
            Height          =   195
            Left            =   1920
            TabIndex        =   94
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label xx 
            Alignment       =   1  'Right Justify
            Caption         =   "Ausentes : "
            Height          =   195
            Left            =   4320
            TabIndex        =   93
            Top             =   255
            Width           =   1005
         End
         Begin VB.Label rr 
            Alignment       =   1  'Right Justify
            Caption         =   "Resultado : "
            Height          =   195
            Left            =   6720
            TabIndex        =   92
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label cc 
            Alignment       =   1  'Right Justify
            Caption         =   "Afirmativos : "
            Height          =   195
            Left            =   1920
            TabIndex        =   91
            Top             =   525
            Width           =   1005
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Negativos : "
            Height          =   195
            Left            =   4320
            TabIndex        =   90
            Top             =   540
            Width           =   1005
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Abstenciones : "
            Height          =   195
            Left            =   6720
            TabIndex        =   89
            Top             =   555
            Width           =   1005
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Minimo de Votos Para Afirmativo : "
            Height          =   255
            Left            =   1800
            TabIndex        =   88
            Top             =   930
            Width           =   2595
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Leyenda Quorum : "
            Height          =   255
            Left            =   5280
            TabIndex        =   87
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label lblcrt_Presentes 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3120
            TabIndex        =   86
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label lblcrt_Ausentes 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   5520
            TabIndex        =   85
            Top             =   255
            Width           =   1005
         End
         Begin VB.Label lblcrt_Resultado 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7920
            TabIndex        =   84
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label lblcrt_Afirmativos 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3120
            TabIndex        =   83
            Top             =   525
            Width           =   1005
         End
         Begin VB.Label lblcrt_Negativos 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   5520
            TabIndex        =   82
            Top             =   540
            Width           =   1005
         End
         Begin VB.Label lblcrt_Abstenciones 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7920
            TabIndex        =   81
            Top             =   555
            Width           =   1005
         End
         Begin VB.Label lblcrt_MinimoParaAfirmativo 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4440
            TabIndex        =   80
            Top             =   930
            Width           =   765
         End
         Begin VB.Label lblcrt_LeyendaQuorum 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   6720
            TabIndex        =   79
            Top             =   960
            Width           =   2205
         End
         Begin VB.Label lblcrt_LeyendaTiempo 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   720
            Width           =   1605
         End
         Begin VB.Label lblcrt_Tiempox 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   1080
            Width           =   1605
         End
      End
      Begin VB.Frame frmEstadoRecinto 
         Caption         =   "Estado del Recinto : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   0
         TabIndex        =   12
         Top             =   3240
         Width           =   9675
         Begin VB.Label lblVersionSQV 
            Alignment       =   1  'Right Justify
            Caption         =   "Merge040225a:"
            Height          =   195
            Left            =   5160
            TabIndex        =   243
            Top             =   4095
            Width           =   2175
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Vector Color() :"
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Vector Presencia() :"
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   735
            Width           =   2655
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Vector Identificacion() :"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   990
            Width           =   2655
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Vector Resultados() :"
            Height          =   255
            Left            =   240
            TabIndex        =   72
            Top             =   1245
            Width           =   2655
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Ocupados No Identificados :"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   1500
            Width           =   2655
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Pendientes de Emitir Votos :"
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   1755
            Width           =   2655
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Sesión :"
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   2010
            Width           =   2655
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Periodo Legislativo :"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   2265
            Width           =   2655
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Nº de Acta :"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   2520
            Width           =   2655
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Título Del Acta :"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   2775
            Width           =   2655
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Identificador De Formulario :"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   3030
            Width           =   2655
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "IP Consola :"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   3285
            Width           =   2655
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Base Mayoria :"
            Height          =   195
            Left            =   5160
            TabIndex        =   63
            Top             =   1515
            Width           =   2205
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Mayoria :"
            Height          =   195
            Left            =   5160
            TabIndex        =   62
            Top             =   1785
            Width           =   2205
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Mayoria Quórum :"
            Height          =   195
            Left            =   5160
            TabIndex        =   61
            Top             =   2040
            Width           =   2205
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo De Operación :"
            Height          =   195
            Left            =   5160
            TabIndex        =   60
            Top             =   2295
            Width           =   2205
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "Tiempo Para Votación :"
            Height          =   195
            Left            =   5160
            TabIndex        =   59
            Top             =   2565
            Width           =   2205
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Error :"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   3540
            Width           =   2655
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Estado de Votación y Pase de Lista :"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   3795
            Width           =   2655
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Modalidad de Votación :"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   4050
            Width           =   2655
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Mensaje Al Operador :"
            Height          =   195
            Left            =   5160
            TabIndex        =   55
            Top             =   480
            Width           =   2205
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Modo Mantenimiento de Bancas :"
            Height          =   195
            Left            =   4920
            TabIndex        =   54
            Top             =   735
            Width           =   2445
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Modo Normal Mant Sistema :"
            Height          =   195
            Left            =   5160
            TabIndex        =   53
            Top             =   1005
            Width           =   2205
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Cartel Encendido :"
            Height          =   195
            Left            =   5160
            TabIndex        =   52
            Top             =   1260
            Width           =   2205
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "Grabar Automáticamente :"
            Height          =   195
            Left            =   5160
            TabIndex        =   51
            Top             =   2820
            Width           =   2205
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "Listar Automáticamente :"
            Height          =   195
            Left            =   5160
            TabIndex        =   50
            Top             =   3075
            Width           =   2205
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   "Acta Grabada :"
            Height          =   195
            Left            =   5160
            TabIndex        =   49
            Top             =   3345
            Width           =   2205
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "Solicitud Grabar Manual :"
            Height          =   195
            Left            =   5160
            TabIndex        =   48
            Top             =   3600
            Width           =   2205
         End
         Begin VB.Label lblVectorColor 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   47
            Top             =   480
            Width           =   2000
         End
         Begin VB.Label llVectorPresencia 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   46
            Top             =   735
            Width           =   2000
         End
         Begin VB.Label lblVectorIdentificacion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   45
            Top             =   990
            Width           =   2000
         End
         Begin VB.Label lblVectorResultado 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   44
            Top             =   1245
            Width           =   2000
         End
         Begin VB.Label lblOcupadosNoIdentificados 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   2880
            TabIndex        =   43
            Top             =   1500
            Width           =   2000
         End
         Begin VB.Label lblPendientesEmitirVotos 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   42
            Top             =   1755
            Width           =   2000
         End
         Begin VB.Label lblSesion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   41
            Top             =   2010
            Width           =   2000
         End
         Begin VB.Label lblPeriodoLegislativo 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   40
            Top             =   2265
            Width           =   2000
         End
         Begin VB.Label lblNumeroActa 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   39
            Top             =   2520
            Width           =   2000
         End
         Begin VB.Label lblTituloActa 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   38
            Top             =   2775
            Width           =   2000
         End
         Begin VB.Label lblIdentificadorFormulario 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   37
            Top             =   3030
            Width           =   2000
         End
         Begin VB.Label lblIpConsola 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   36
            Top             =   3285
            Width           =   2000
         End
         Begin VB.Label lblError 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   35
            Top             =   3540
            Width           =   2000
         End
         Begin VB.Label lblEstadoVotacionyPaseDeLista 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   34
            Top             =   3795
            Width           =   2000
         End
         Begin VB.Label lblModalidadVotacion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   33
            Top             =   4050
            Width           =   2000
         End
         Begin VB.Label lblMensajeAlOperador 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   32
            Top             =   480
            Width           =   1485
         End
         Begin VB.Label lblModoMantenimientoBancas 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   31
            Top             =   735
            Width           =   1485
         End
         Begin VB.Label lblModoMantenimientosistema 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   30
            Top             =   990
            Width           =   1485
         End
         Begin VB.Label lblCartelEncendido 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   29
            Top             =   1245
            Width           =   1485
         End
         Begin VB.Label lblBaseMayoria 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   28
            Top             =   1500
            Width           =   1485
         End
         Begin VB.Label lblTipoMayoria 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   27
            Top             =   1755
            Width           =   1485
         End
         Begin VB.Label lbltipoMayoriaQuorum 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   26
            Top             =   2010
            Width           =   1485
         End
         Begin VB.Label lblTipoOperacion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   25
            Top             =   2265
            Width           =   1485
         End
         Begin VB.Label lblTiempoParaVotacion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   24
            Top             =   2520
            Width           =   1485
         End
         Begin VB.Label lblGrabarAutomaticamente 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   23
            Top             =   2775
            Width           =   1485
         End
         Begin VB.Label lblListarAutomaticamente 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   22
            Top             =   3030
            Width           =   1485
         End
         Begin VB.Label lblActaGrabada 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   21
            Top             =   3285
            Width           =   1485
         End
         Begin VB.Label lblSolicituGrabarManual 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   20
            Top             =   3540
            Width           =   1485
         End
         Begin VB.Label lblEstadoSesion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   19
            Top             =   3795
            Width           =   1485
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            Caption         =   "Estado de Sesión :"
            Height          =   195
            Left            =   5160
            TabIndex        =   18
            Top             =   3840
            Width           =   2205
         End
         Begin VB.Label lblAbsAut 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   17
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label LblAbstencionesAut 
            Alignment       =   1  'Right Justify
            Caption         =   "Abs.Aut"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label lblAppMayor 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   15
            Top             =   4080
            Width           =   285
         End
         Begin VB.Label lblAppMinor 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7800
            TabIndex        =   14
            Top             =   4080
            Width           =   525
         End
         Begin VB.Label lblAppRevision 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   8400
            TabIndex        =   13
            Top             =   4080
            Width           =   525
         End
      End
      Begin VB.Timer Timer 
         Interval        =   100
         Left            =   5880
         Top             =   600
      End
      Begin VB.TextBox txtVecesPorSegundo 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         TabIndex        =   11
         Text            =   "1"
         Top             =   960
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Height          =   1050
         Left            =   7920
         ScaleHeight     =   990
         ScaleWidth      =   1695
         TabIndex        =   8
         Top             =   360
         Width           =   1750
         Begin VB.CommandButton cmdSalir 
            Caption         =   "&Salir"
            Height          =   495
            Left            =   0
            TabIndex        =   10
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton cmdTerminar 
            Caption         =   "&Iniciar Server"
            Enabled         =   0   'False
            Height          =   495
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   810
         Left            =   5760
         ScaleHeight     =   750
         ScaleWidth      =   2055
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   2120
         Begin VB.CommandButton HabilitarSeguimientoPizarraRecinto 
            Caption         =   "&Ocultar Estado de Recinto"
            Height          =   375
            Left            =   0
            TabIndex        =   7
            Top             =   375
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CommandButton HabilitarSeguimientoPizarraCartel 
            Caption         =   "&Ocultar Estado de Cartel"
            Height          =   375
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin VB.CheckBox chkLog_Mensajes 
         Caption         =   "Guardar Copia de Mensajes"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1365
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdResetarVectores 
         Caption         =   "Reset"
         Height          =   315
         Left            =   7080
         TabIndex        =   3
         ToolTipText     =   "Resetear vectores de estado"
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Config"
         Height          =   315
         Left            =   6360
         TabIndex        =   2
         ToolTipText     =   "Resetear vectores de estado"
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton bCartelGeneral 
         Caption         =   "Ver Cartel General"
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Inicio Server :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   100
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblFechaInicioServer 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   99
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ciclos por segundo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   98
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label37 
         Caption         =   "Ciclos Totales : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   97
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblCiclos 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4080
         TabIndex        =   96
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblVersion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   95
         Top             =   885
         Width           =   3135
      End
   End
   Begin VB.Frame FrameSQVActa 
      BorderStyle     =   0  'None
      Height          =   12000
      Left            =   0
      TabIndex        =   121
      Top             =   0
      Width           =   15360
      Begin VB.TextBox txtPagina 
         Height          =   285
         Left            =   5685
         TabIndex        =   242
         Text            =   "Text1"
         Top             =   11640
         Width           =   1770
      End
      Begin VB.Frame frameActaDatos 
         Height          =   11595
         Left            =   45
         TabIndex        =   122
         Top             =   0
         Width           =   15285
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   27
            Left            =   13110
            TabIndex        =   241
            Top             =   11100
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   27
            Left            =   7905
            TabIndex        =   240
            Top             =   11100
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   26
            Left            =   13110
            TabIndex        =   239
            Top             =   10620
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   26
            Left            =   7890
            TabIndex        =   238
            Top             =   10620
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   25
            Left            =   13110
            TabIndex        =   237
            Top             =   10170
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   25
            Left            =   7890
            TabIndex        =   236
            Top             =   10170
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   24
            Left            =   13095
            TabIndex        =   235
            Top             =   9690
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   24
            Left            =   7875
            TabIndex        =   234
            Top             =   9690
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   23
            Left            =   13095
            TabIndex        =   233
            Top             =   9195
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   23
            Left            =   7875
            TabIndex        =   232
            Top             =   9195
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   22
            Left            =   13080
            TabIndex        =   231
            Top             =   8685
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   22
            Left            =   7860
            TabIndex        =   230
            Top             =   8715
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   21
            Left            =   13065
            TabIndex        =   229
            Top             =   8220
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   21
            Left            =   7845
            TabIndex        =   228
            Top             =   8220
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   20
            Left            =   13050
            TabIndex        =   227
            Top             =   7740
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   20
            Left            =   7830
            TabIndex        =   226
            Top             =   7740
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   19
            Left            =   13050
            TabIndex        =   225
            Top             =   7245
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   19
            Left            =   7830
            TabIndex        =   224
            Top             =   7245
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   18
            Left            =   13035
            TabIndex        =   223
            Top             =   6750
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   18
            Left            =   7815
            TabIndex        =   222
            Top             =   6765
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   17
            Left            =   13035
            TabIndex        =   221
            Top             =   6270
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   17
            Left            =   7815
            TabIndex        =   220
            Top             =   6270
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   16
            Left            =   13020
            TabIndex        =   219
            Top             =   5790
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   16
            Left            =   7800
            TabIndex        =   218
            Top             =   5790
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   15
            Left            =   13020
            TabIndex        =   217
            Top             =   5295
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   15
            Left            =   7800
            TabIndex        =   216
            Top             =   5295
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   14
            Left            =   13005
            TabIndex        =   215
            Top             =   4800
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   14
            Left            =   7770
            TabIndex        =   214
            Top             =   4815
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   13
            Left            =   5475
            TabIndex        =   213
            Top             =   11085
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   13
            Left            =   270
            TabIndex        =   212
            Top             =   11085
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   12
            Left            =   5475
            TabIndex        =   211
            Top             =   10605
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   12
            Left            =   255
            TabIndex        =   210
            Top             =   10605
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   11
            Left            =   5475
            TabIndex        =   209
            Top             =   10155
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   11
            Left            =   255
            TabIndex        =   208
            Top             =   10155
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   10
            Left            =   5460
            TabIndex        =   207
            Top             =   9675
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   10
            Left            =   240
            TabIndex        =   206
            Top             =   9675
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   9
            Left            =   5460
            TabIndex        =   205
            Top             =   9180
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   9
            Left            =   240
            TabIndex        =   204
            Top             =   9180
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   8
            Left            =   5445
            TabIndex        =   203
            Top             =   8670
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   8
            Left            =   225
            TabIndex        =   202
            Top             =   8700
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   7
            Left            =   5430
            TabIndex        =   201
            Top             =   8205
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   7
            Left            =   210
            TabIndex        =   200
            Top             =   8205
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   6
            Left            =   5430
            TabIndex        =   199
            Top             =   7725
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   6
            Left            =   195
            TabIndex        =   198
            Top             =   7725
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   5
            Left            =   5415
            TabIndex        =   197
            Top             =   7230
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   5
            Left            =   195
            TabIndex        =   196
            Top             =   7230
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   4
            Left            =   5400
            TabIndex        =   195
            Top             =   6735
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   4
            Left            =   180
            TabIndex        =   194
            Top             =   6750
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   3
            Left            =   5400
            TabIndex        =   193
            Top             =   6255
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   3
            Left            =   180
            TabIndex        =   192
            Top             =   6255
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   2
            Left            =   5385
            TabIndex        =   191
            Top             =   5775
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   2
            Left            =   165
            TabIndex        =   190
            Top             =   5775
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   1
            Left            =   5385
            TabIndex        =   189
            Top             =   5280
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   1
            Left            =   165
            TabIndex        =   188
            Top             =   5280
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   0
            Left            =   5370
            TabIndex        =   187
            Top             =   4785
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   0
            Left            =   135
            TabIndex        =   186
            Top             =   4800
            Width           =   5055
         End
         Begin VB.CommandButton cmdPresidente 
            Caption         =   "Cam&biar presidente"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9150
            TabIndex        =   157
            Top             =   1950
            Width           =   1515
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   315
            Left            =   1830
            TabIndex        =   156
            Top             =   2370
            Width           =   7185
         End
         Begin VB.Frame Frame3 
            Height          =   45
            Left            =   510
            TabIndex        =   155
            Top             =   4410
            Width           =   11235
         End
         Begin VB.TextBox txtAbstencionesTotales 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   11100
            Locked          =   -1  'True
            TabIndex        =   154
            Top             =   4020
            Width           =   615
         End
         Begin VB.TextBox txtAbstencionesNoId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   11100
            Locked          =   -1  'True
            TabIndex        =   153
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox txtAbstencionesId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   11100
            Locked          =   -1  'True
            TabIndex        =   152
            Top             =   2850
            Width           =   615
         End
         Begin VB.TextBox txtNegativoTotales 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   8100
            Locked          =   -1  'True
            TabIndex        =   151
            Top             =   4020
            Width           =   615
         End
         Begin VB.TextBox txtNegativoDesempate 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   8100
            Locked          =   -1  'True
            TabIndex        =   150
            Top             =   3630
            Width           =   615
         End
         Begin VB.TextBox txtNegativoNoId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   8100
            Locked          =   -1  'True
            TabIndex        =   149
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox txtNegativoID 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   8100
            Locked          =   -1  'True
            TabIndex        =   148
            Top             =   2850
            Width           =   615
         End
         Begin VB.TextBox txtAfirmativosTotal 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   147
            Top             =   4020
            Width           =   615
         End
         Begin VB.TextBox txtAfirmativosDesempate 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   146
            Top             =   3630
            Width           =   615
         End
         Begin VB.TextBox txtAfirmativosNoId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   145
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox txtAfirmativosId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   144
            Top             =   2850
            Width           =   615
         End
         Begin VB.TextBox txtAusentesTotal 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   143
            Top             =   4020
            Width           =   615
         End
         Begin VB.TextBox txtPresentesTotal 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   142
            Top             =   3630
            Width           =   615
         End
         Begin VB.TextBox txtPresentesNoId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   141
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox txtPresentesId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   140
            Top             =   2850
            Width           =   615
         End
         Begin VB.Frame Frame2 
            Height          =   45
            Left            =   510
            TabIndex        =   139
            Top             =   2730
            Width           =   11235
         End
         Begin VB.TextBox txtNombrePresidente 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   2850
            Locked          =   -1  'True
            TabIndex        =   138
            Top             =   1980
            Width           =   6165
         End
         Begin VB.TextBox txtCodigoPresidente 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   137
            Top             =   1980
            Width           =   945
         End
         Begin VB.TextBox txtVotacion 
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10350
            Locked          =   -1  'True
            TabIndex        =   136
            Top             =   1590
            Width           =   1365
         End
         Begin VB.TextBox txtBase 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   135
            Top             =   1590
            Width           =   2115
         End
         Begin VB.TextBox txtDesempate 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   10350
            Locked          =   -1  'True
            TabIndex        =   134
            Top             =   1200
            Width           =   1365
         End
         Begin VB.TextBox txtMiembros 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   133
            Top             =   1200
            Width           =   2115
         End
         Begin VB.TextBox txtTipoMayoria 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   132
            Top             =   1590
            Width           =   3105
         End
         Begin VB.TextBox txtTipoQuorum 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   131
            Top             =   1200
            Width           =   3105
         End
         Begin VB.Frame Frame1 
            Height          =   45
            Left            =   510
            TabIndex        =   130
            Top             =   1110
            Width           =   11235
         End
         Begin VB.TextBox txtHora 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   10350
            Locked          =   -1  'True
            TabIndex        =   129
            Top             =   750
            Width           =   1365
         End
         Begin VB.TextBox txtFecha 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   9180
            Locked          =   -1  'True
            TabIndex        =   128
            Top             =   750
            Width           =   1035
         End
         Begin VB.TextBox txtNombre 
            Height          =   315
            Left            =   1830
            TabIndex        =   127
            Top             =   750
            Width           =   7155
         End
         Begin VB.TextBox txtVersion 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   10350
            Locked          =   -1  'True
            TabIndex        =   126
            Top             =   360
            Width           =   1365
         End
         Begin VB.TextBox txtNroActa 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   7500
            Locked          =   -1  'True
            TabIndex        =   125
            Top             =   360
            Width           =   1485
         End
         Begin VB.TextBox txtSesion 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5730
            Locked          =   -1  'True
            TabIndex        =   124
            Top             =   360
            Width           =   1035
         End
         Begin VB.TextBox txtTipoOperacion 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   123
            Top             =   360
            Width           =   3105
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Left            =   480
            TabIndex        =   185
            Top             =   2430
            Width           =   1065
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            Caption         =   "Abstenciones Totales"
            Height          =   195
            Left            =   8880
            TabIndex        =   184
            Top             =   4080
            Width           =   1530
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Abstenciones No Identificables"
            Height          =   195
            Left            =   8880
            TabIndex        =   183
            Top             =   3300
            Width           =   2190
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            Caption         =   "Abstenciones identificables"
            Height          =   195
            Left            =   8880
            TabIndex        =   182
            Top             =   2910
            Width           =   1920
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            Caption         =   "Votos Negativos Totales"
            Height          =   195
            Left            =   6030
            TabIndex        =   181
            Top             =   4080
            Width           =   1740
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            Caption         =   "Votos Neg. Desempate"
            Height          =   195
            Left            =   6030
            TabIndex        =   180
            Top             =   3690
            Width           =   1650
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            Caption         =   "Votos Neg. No Identificables"
            Height          =   195
            Left            =   6030
            TabIndex        =   179
            Top             =   3300
            Width           =   2025
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "Votos Neg. Identificables"
            Height          =   195
            Left            =   6030
            TabIndex        =   178
            Top             =   2910
            Width           =   1770
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            Caption         =   "Votos Afirmativos Total"
            Height          =   195
            Left            =   3180
            TabIndex        =   177
            Top             =   4080
            Width           =   1620
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "Votos Afirm. Desempate"
            Height          =   195
            Left            =   3180
            TabIndex        =   176
            Top             =   3690
            Width           =   1695
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            Caption         =   "Votos Afirm. No Identificables"
            Height          =   195
            Left            =   3180
            TabIndex        =   175
            Top             =   3300
            Width           =   2070
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "Votos Afirm. Identificables"
            Height          =   195
            Left            =   3180
            TabIndex        =   174
            Top             =   2910
            Width           =   1815
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            Caption         =   "Ausentes Total"
            Height          =   195
            Left            =   480
            TabIndex        =   173
            Top             =   4080
            Width           =   1065
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "Presentes Total"
            Height          =   195
            Left            =   480
            TabIndex        =   172
            Top             =   3690
            Width           =   1110
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "Presentes no identificables"
            Height          =   195
            Left            =   480
            TabIndex        =   171
            Top             =   3300
            Width           =   1890
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Presentes identificables"
            Height          =   195
            Left            =   480
            TabIndex        =   170
            Top             =   2910
            Width           =   1665
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "Presidente"
            Height          =   195
            Left            =   480
            TabIndex        =   169
            Top             =   2040
            Width           =   750
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Votación"
            Height          =   195
            Left            =   9180
            TabIndex        =   168
            Top             =   1650
            Width           =   630
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Base"
            Height          =   195
            Left            =   5190
            TabIndex        =   167
            Top             =   1650
            Width           =   360
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Desempate"
            Height          =   195
            Left            =   9180
            TabIndex        =   166
            Top             =   1260
            Width           =   810
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Miembros del cuerpo"
            Height          =   195
            Left            =   5190
            TabIndex        =   165
            Top             =   1260
            Width           =   1470
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de mayoría"
            Height          =   195
            Left            =   480
            TabIndex        =   164
            Top             =   1650
            Width           =   1155
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de quorum"
            Height          =   195
            Left            =   480
            TabIndex        =   163
            Top             =   1260
            Width           =   1110
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Nombre del acta"
            Height          =   195
            Left            =   480
            TabIndex        =   162
            Top             =   810
            Width           =   1170
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Versión"
            Height          =   195
            Left            =   9180
            TabIndex        =   161
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Nº acta"
            Height          =   195
            Left            =   6870
            TabIndex        =   160
            Top             =   420
            Width           =   540
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Sesión"
            Height          =   195
            Left            =   5190
            TabIndex        =   159
            Top             =   420
            Width           =   480
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de operación"
            Height          =   195
            Left            =   480
            TabIndex        =   158
            Top             =   420
            Width           =   1290
         End
      End
   End
   Begin VB.Frame FrameMantenimiento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12000
      Left            =   0
      TabIndex        =   102
      Top             =   0
      Width           =   15360
      Begin VB.Line Line3 
         X1              =   60
         X2              =   10995
         Y1              =   5070
         Y2              =   5070
      End
      Begin VB.Line Line2 
         X1              =   45
         X2              =   11010
         Y1              =   3075
         Y2              =   3075
      End
      Begin VB.Line Line1 
         X1              =   15
         X2              =   11145
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label44 
         BackColor       =   &H80000009&
         Caption         =   "A revisar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   195
         TabIndex        =   120
         Top             =   6780
         Width           =   1545
      End
      Begin VB.Label Label43 
         BackColor       =   &H80000009&
         Caption         =   "A registrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   180
         TabIndex        =   119
         Top             =   5070
         Width           =   1545
      End
      Begin VB.Label Label42 
         BackColor       =   &H80000009&
         Caption         =   "Identificaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   195
         TabIndex        =   118
         Top             =   4035
         Width           =   2040
      End
      Begin VB.Label Label41 
         BackColor       =   &H80000009&
         Caption         =   "Switches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   180
         TabIndex        =   117
         Top             =   3120
         Width           =   1545
      End
      Begin VB.Label Label40 
         BackColor       =   &H80000009&
         Caption         =   "Presentes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   90
         TabIndex        =   116
         Top             =   0
         Width           =   1545
      End
      Begin VB.Label Label39 
         BackColor       =   &H80000009&
         Caption         =   "Ausentes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5430
         TabIndex        =   115
         Top             =   0
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Paneles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   180
         TabIndex        =   114
         Top             =   585
         Width           =   1545
      End
      Begin VB.Label lblMantenimientostrMantListaFallas 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Lfallas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   200
         TabIndex        =   113
         Top             =   8040
         Width           =   11500
      End
      Begin VB.Label lblMantenimientostrMantListaPendientes 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "LPendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   195
         TabIndex        =   112
         Top             =   6075
         Width           =   11505
      End
      Begin VB.Label lblMantenimientostrPendientes 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CantPendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   180
         TabIndex        =   111
         Top             =   5595
         Width           =   1215
      End
      Begin VB.Label lblMantenimientostrFallas 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CantFallas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   240
         TabIndex        =   110
         Top             =   7305
         Width           =   1215
      End
      Begin VB.Label lblMantenimientostrId 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   195
         TabIndex        =   109
         Top             =   4440
         Width           =   510
      End
      Begin VB.Label lblMantenimientostrPresencias 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Presencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   180
         TabIndex        =   108
         Top             =   3540
         Width           =   14025
      End
      Begin VB.Label lblMantenimientostrPanel3 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Panel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   195
         TabIndex        =   107
         Top             =   2565
         Width           =   14265
      End
      Begin VB.Label lblMantenimientostrPanel2 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Panel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   195
         TabIndex        =   106
         Top             =   1860
         Width           =   13785
      End
      Begin VB.Label lblMantenimientostrPanel1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Panel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   195
         TabIndex        =   105
         Top             =   1095
         Width           =   13905
      End
      Begin VB.Label lblMantenimientostrAusentes 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   7650
         TabIndex        =   104
         Top             =   30
         Width           =   1215
      End
      Begin VB.Label lblMantenimientostrPresentes 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   2355
         TabIndex        =   103
         Top             =   -15
         Width           =   1215
      End
   End
   Begin VB.Frame FrameSQVApagado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12000
      Left            =   0
      TabIndex        =   101
      Top             =   0
      Width           =   15360
   End
   Begin VB.Label lblLeyendaVotoAbstencion 
      BackStyle       =   0  'Transparent
      Caption         =   "AFIRMATIVOS NEGATIVOS ABSTENCIONES"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3435
      Left            =   0
      TabIndex        =   375
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label lblLeyendaVotoNegativo 
      BackStyle       =   0  'Transparent
      Caption         =   "AFIRMATIVOS NEGATIVOS ABSTENCIONES"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3435
      Left            =   0
      TabIndex        =   374
      Top             =   0
      Width           =   8055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Shape shpTitulo2 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Cn          As ADODB.Connection
Attribute Cn.VB_VarHelpID = -1
Private WithEvents rs          As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private WithEvents RsOtro      As ADODB.Recordset
Attribute RsOtro.VB_VarHelpID = -1
Private WithEvents RsW         As ADODB.Recordset
Attribute RsW.VB_VarHelpID = -1
Private WithEvents RsWrite     As ADODB.Recordset
Attribute RsWrite.VB_VarHelpID = -1
Private WithEvents RsLocal     As ADODB.Recordset
Attribute RsLocal.VB_VarHelpID = -1
Private rsTemp                 As ADODB.Recordset
Private rstActa                As ADODB.Recordset

Private blServerPrendido       As Boolean
Private CantidadEidrxh(256)    As Long
Private TiempoEidrxh(256)    As Long
Private xIntervalo             As Long
Private xUltimoMensajeSB       As Long        ' Control de ultimo mensaje enviado por el servidor de bancas
Private xUltimoMensajeCosola   As Long        ' Control de ultimo mensaje enviado por la consola
Private blMostrarEstadoCartel  As Boolean     ' flag que habilita o no el frame de estado de cartel
Private blMostrarEstadoRecinto As Boolean     ' flag que habilita o no el frame de estado de recinto
Private xBanca                 As Long
Private xCiclosTotales         As Double
Private Mensaje2Banca          As MensajeSistema
Private strPath                As String
Private xFechaArranque         As Date
Private xFechaUltimoReset      As Date
Private xFechaInicioProceso As Date
Private blBanderaTimer         As Boolean
Private xFileSqv As Long
Private xNroMensajeSB           As Long
Private xPrimerMensajeSB As Long
Private PrimerRecuento As Boolean
Public EstabaEnIdentificacion As Boolean
Public Tick_InicioPasLis As Long

'A02
Private TimerCounter As Integer
Private colLeg As New colLegisladores
Private imagePath As String
Private xVersionFormularios As String

Private Const OffsetOrador As Integer = 0
Private lAbstencionPresidente As Boolean
Private Const MiBlanco As String = "&H00E0E0E0"
Private Const MiRojo As String = "&H000000C0"
Private Sub AbrirDB()
    On Error GoTo TrapError
    Set Cn = New ADODB.Connection
    With Cn
        .ConnectionString = strConexion
        .CommandTimeout = 30
        ' .CursorLocation = adUseServer
        .CursorLocation = adUseClient
        .Open
    End With
    Set rs = New ADODB.Recordset
    Set RsWrite = New ADODB.Recordset
    VotoRemoto.DatabaseOpened = True
Exit Sub
TrapError:
    MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
    'Resume
End Sub
Private Sub SetVersion()
    lblVersionCartel.Caption = "v6.0c (110330)" '"v6.0a (031211)"
End Sub
Public Sub SetearRs(strCadena As String)
    On Error GoTo TrapError
    'Set Rs = New ADODB.Recordset
    With rs
        If .State = adStateOpen Then
            .Close
            .Source = strCadena
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open
        Else
            .Source = strCadena
            .ActiveConnection = Cn
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .CursorLocation = adUseClient
            .Open
        End If
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
Exit Sub
TrapError:
    MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
    'Resume
End Sub
Private Sub SetearOtroRs(strCadena As String)
    On Error GoTo TrapError
    'Set Rs = New ADODB.Recordset
    With RsOtro
        If .State = adStateOpen Then
            .Close
            .Source = strCadena
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open
        Else
            .Source = strCadena
            .ActiveConnection = Cn
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .CursorLocation = adUseClient
            .Open
        End If
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
Exit Sub
TrapError:
    MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
    'Resume
End Sub
Private Sub SetearRsW(strCadena As String)
    On Error GoTo TrapError
    'Set Rs = New ADODB.Recordset
    With RsWrite
        If .State = adStateOpen Then
            .Close
            .Source = strCadena
            .Open
        Else
            .Source = strCadena
            .ActiveConnection = Cn
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End If
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
Exit Sub
TrapError:
    MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
    'Resume
End Sub
Public Function SetearRsAux(pCadena As String, ByRef pRst As ADODB.Recordset) As Boolean
    SetearRsAux = False
    pRst.CursorLocation = adUseClient
    pRst.Open pCadena, Cn, adOpenForwardOnly, adLockReadOnly
    If Not pRst.BOF And Not pRst.EOF Then
         SetearRsAux = True
    End If
End Function

Public Function EjecutaSQLCartel(pCadena As String) As Boolean
    'Call Cn.Execute(pCadena)
End Function

Private Sub bCartelApagado_Click()
    EstadoActual.CartelEncendido = 0
    FrameSQVApagado.ZOrder 0
    txtTipoOperacion = "d"
End Sub

Private Sub bCartelGeneral_Click()
    EstadoActual.CartelEncendido = 2
    FrameSQVGeneral.ZOrder 0
End Sub

Private Sub chkLog_Mensajes_Click()
    If chkLog_Mensajes.Value = 0 Then
        Call AltaLogGeneral("Operador del sistema", "Operador termina grabación de log de mensajes")
    ElseIf chkLog_Mensajes.Value = 1 Then
        Call AltaLogGeneral("Operador del sistema", "Operador inicia grabación de log de mensajes")
    End If
End Sub
Private Sub cmdConfig_Click()
    frmConfig.Show 1
End Sub
Public Function EjecutarSQL(sql As String)
Cn.Execute sql, , adCmdText
End Function
Private Sub ResetearVectores()
    Dim X      As Long
    With EstadoActual
        For X = 0 To (xUltimaBanca)
            If Trim(EstadoActual.strError) <> "cambio?mantenimiento" Then
                .VectorPresencia(X) = BANCA_INHABILITADA
                .VectorIdentificacion(X) = NO_IDENTIFICADO
                .VectorColor(X) = AsignarColor(X)
                .VectorResultados(X) = ABSTENCION
                .VMantEstado(X) = ABSTENCION
                .VTipoIdentificacion(X) = TIPO_IDENTIFICACION_HUELLA
            Else
                If X > 0 Then
                    .VectorPresencia(X) = BANCA_INHABILITADA
                    .VectorIdentificacion(X) = NO_IDENTIFICADO
                    .VectorColor(X) = AsignarColor(X)
                    .VectorResultados(X) = ABSTENCION
                    .VMantEstado(X) = ABSTENCION
                    .VTipoIdentificacion(X) = TIPO_IDENTIFICACION_HUELLA
                End If
            End If
        Next X
        If Trim(EstadoActual.strError) <> "cambio?mantenimiento" Then
            Call ResetearPresidente
        End If
        xPresidenteLegislador = False
        'Vectores de mantenimiento
        For X = 0 To (cUltimoPanelMant)
            .VMantBanca(X) = 99999
            .VMantInfo(X) = " "
            .VMantIdentificacion(X) = " "
        Next X
    End With
End Sub
Private Sub CargarVectorIdentificacionHabilitados()
    Dim strSql As String
    Dim X As Long
    strSql = "SELECT deskid, id FROM legisladores_activos ORDER BY deskid"
    
    For X = 0 To (xUltimaBanca)
        EstadoActual.VectorIdentificacionHabilitados(X) = NO_IDENTIFICADO
    Next X
    Call SetearRs(strSql)
    With rs
        While Not .EOF
            X = Int(.Fields("deskid").Value) - 1
            EstadoActual.VectorIdentificacionHabilitados(X) = .Fields("id").Value
            .MoveNext
        Wend
    End With
    xVectorIdentificacionHabilitados = EstadoActual.VectorIdentificacionHabilitados
End Sub

Private Sub cmdGeneralSalir_Click()
    Call AltaLogGeneral("Saliendo de SQV Server por solicitud Server SQV", Now)
    Call Salir_SQV
End Sub


Private Sub cmdResetarVectores_Click()
    Dim X As Long
    With EstadoActual
        For X = 0 To (xUltimaBanca)
            .VectorPresencia(X) = BANCA_INHABILITADA
            .VectorIdentificacion(X) = NO_IDENTIFICADO
            .VectorColor(X) = AsignarColor(X)
            .VectorResultados(X) = ABSTENCION
        Next X
        Call ResetearPresidente
        .ActaGrabada = 0
        .Ausentes = xMiembrosDelCuerpo
        .BaseMayoria = "votemi"
        
        'cartel visible al iniciar
        .CartelEncendido = 2 ' general poner 1 para control
        FrameControl.ZOrder 0
        
        .EstadoVotacion_y_PasList = "espera"
        .GrabarAutomaticamente = 0
        .IdentificadorDeFormulario = ""
        .IP_Consola = ""
        .ListarAutomaticamente = 0
        .MensajeAlOperador = ""
        .ModoMantenimientoBancas = 0
        .ModoNormalMantSistema = 0
        .NroActa = 0
        .OcupadosNoIdentificados = 0
        .PendientesEmitirVotos = 0
        .PeriodoLegislativo = ""
        .Presentes = 0
        .Sesion = 0
        .SolicitudGrabarManual = 0
        .TiempoParaVotacion = 30
        .TipoDeOperacion = "quorum"
        .TipoMayoria = "120"
        .TipoMayoriaQuorum = "MAN"
        .TituloDelActa = ""
        .strError = ""
        .EstadoSesion = ""
        .TipoDeAbstencion = "votlar"
        .Reunion = 0
        .Orador = "0"
        .OradorNombre = ""
        .OradorAgrupacionPolitica = ""
        .OradorDistrito = ""
        .OradorSexo = ""
    End With
    With CartelActual
        .Abstenciones = 0
        .Afirmativos = 0
        .Ausentes = xMiembrosDelCuerpo
        .LeyendaQuorum = "NO HAY QUORUM"
        .MinimoVotosParaAfirmativo = 0
        .Negativos = 0
        .Presentes = 0
        .Resultado = 0
    End With
    Call CalcularMinimoParaQuorum
    Call ActualizarVector_enBD
End Sub
Private Sub cmdSalir_Click()
    Call AltaLogGeneral("Saliendo de SQV Server por solicitud Server SQV", Now)
    Call Salir_SQV
End Sub
Private Sub Salir_SQV()
    
    With EstadoActual
        .CartelEncendido = 0
        .IP_Consola = ""
        .ModoMantenimientoBancas = 0
        .ModoNormalMantSistema = 0
    End With
    Call ReinicioSistema
    'MsgBox "AGREGAR COMANDO DE SALIDA DE SERVER BANCAS AQUI"
    Screen.MousePointer = vbDefault
    ShowCursor True
    Unload Me
End Sub
Private Sub Mantenimiento_SQV()
    With EstadoActual
        .CartelEncendido = 3
        FrameMantenimiento.ZOrder 0
        .ModoMantenimientoBancas = 1
        .ModoNormalMantSistema = 0
    End With
    Dim Presi As String
    Presi = EstadoActual.VectorIdentificacion(0)
    Call ReinicioSistema
    EstadoActual.VectorIdentificacion(0) = Presi
    EstadoActual.TipoDeOperacion = "votnom"
    Call InicializarVotacion
    EstadoActual.EstadoVotacion_y_PasList = "votando"
    EstadoActual.TiempoParaVotacion = 9999
    EstadoActual.FechaVotacion = DateAdd("s", xtiempoInicioVotac, Now)
    EstadoActual.TituloDelActa = "MANTENIMIENTO DEL SISTEMA SQV"
    xPresidenteLegislador = True
    EstadoActual.ModoVotaPresidente = True 'habilita votacion para el presidente tambien
    SolicitarHabilitarVotoPresidente
End Sub
Private Sub Fin_Mantenimiento_SQV()
    With EstadoActual
        .CartelEncendido = 0
        FrameMantenimiento.ZOrder 1
        .ModoMantenimientoBancas = 0
        .ModoNormalMantSistema = 0
        .TituloDelActa = " "
        .ModoVotaPresidente = False 'deshabilita votacion para el presidente
        .PresidenteHabilitadoParaVotar = False
    End With
    Call ReinicioSistema 'inicializa votacion tambien
    EstadoActual.TiempoParaVotacion = 15
    xPresidenteLegislador = False
    Call EnviarMensajesFinAuth("brc", "BRC de Cancel manten a normal")
End Sub
Private Sub NormalMantenimiento_SQV()
    With EstadoActual
        .CartelEncendido = 2
        FrameSQVGeneral.ZOrder 0
        .ModoMantenimientoBancas = 0
        .ModoNormalMantSistema = 1
    End With
    Call ReinicioSistema
    Call InicializarVotacion
    EstadoActual.TituloDelActa = "MANTENIMIENTO DEL SISTEMA SQV"
End Sub
Private Sub DeterminarValoresInicioServer()
    On Error GoTo err
    Dim rstCfg As New ADODB.Recordset
    Dim strSql As String
    ' ------------------------------------------------------------------------------
    ' Esta funcion levanta los valores de configuracion default
    ' para levantar los servicios de sqv server y sb
    ' ------------------------------------------------------------------------------
    strSql = "SELECT Ejecutable_sqv, Ejecutable_sb, IP_levanta_ap, " _
           & "puerto_levanta_ap From config"
    SetearRsAux strSql, rstCfg
        
    If rstCfg.EOF = False Then
        If Not IsNull(rstCfg.Fields("puerto_levanta_ap").Value) Then
            strPuerto = rstCfg.Fields("puerto_levanta_ap").Value
        End If
        If Not IsNull(rstCfg.Fields("IP_levanta_ap").Value) Then
            strIpServer = rstCfg.Fields("IP_levanta_ap").Value
        End If
        If Not IsNull(rstCfg.Fields("Ejecutable_sqv").Value) Then
            strExeSqv = rstCfg.Fields("Ejecutable_sqv").Value
        End If
        If Not IsNull(rstCfg.Fields("Ejecutable_sb").Value) Then
            strExeSb = rstCfg.Fields("Ejecutable_sb").Value
        End If
    Else
        strPuerto = ""
        strIpServer = ""
        strExeSqv = ""
        strExeSb = ""
    End If
    rstCfg.Close: Set rstCfg = Nothing
    Exit Sub
err:
    Call AltaLogGeneral("SQV SERVER", "Error en valores inicio server")
    'MsgBox "error"
End Sub
Private Sub Levanta_Banca()
    Dim strPathToProgram As String
On Error GoTo TrapError:
If True Then 'habilitar
    
    'strPathToProgram = Environ("windir") & "\System32\calc.exe"
    strPathToProgram = Environ("sqv")
    If strPathToProgram = "" Then strPathToProgram = "e:\Sistemas\siprevo"
    strPathToProgram = strPathToProgram & "\SB\SB.exe"
    Shell ("" & strPathToProgram & "")
    'MsgBox "x"
Else
    Ws.Close
    Ws.RemoteHost = Trim(strIpServer)
    Ws.RemotePort = strPuerto
    Ws.Connect
    strPath = strExeSb
    DoEvents
    While Ws.State = 6
        DoEvents
    Wend
End If
Exit Sub
TrapError:
    Call AltaLogGeneral("SQV SERVER", "No se pudo levantar Servidor de bancas " & strPathToProgram)
End Sub


Private Sub Command3_Click()

End Sub



Private Sub Form_Activate()

BorrarImpresion
Imprimio = False
End Sub

Private Sub imgA_Click(Index As Integer)
MsgBox ("imgA(" & Index & ")")
End Sub

Private Sub imgB_Click(Index As Integer)
MsgBox ("imgB(" & Index & ")")
End Sub

Private Sub imgC_Click(Index As Integer)
MsgBox ("imgC(" & Index & ")")
End Sub

Private Sub lblGeneralLeyendaQuorumDato_Change()
If Trim(lblGeneralLeyendaQuorumDato.Caption) = "QUORUM" Then
    lblGeneralLeyendaQuorumDato.ForeColor = &HFFFF&
    shpRecuadroQuorum.Width = Trim(lblGeneralLeyendaQuorumDato.Width) - 3000
    shpRecuadroQuorum.Left = lblGeneralLeyendaQuorumDato.Left + 2800
    shpRecuadroQuorum.BorderWidth = 3
    shpRecuadroQuorum.Height = shpRecuadroFecha.Height + 200
    shpRecuadroQuorum.top = shpRecuadroFecha.top - 120
    shpRecuadroQuorum.BorderColor = &HFFFF&
Else
    lblGeneralLeyendaQuorumDato.ForeColor = MiRojo
    shpRecuadroQuorum.Width = lblGeneralLeyendaQuorumDato.Width - 800
    shpRecuadroQuorum.Left = lblGeneralLeyendaQuorumDato.Left + 900
    shpRecuadroQuorum.BorderWidth = 3
    shpRecuadroQuorum.Height = shpRecuadroFecha.Height + 200
    shpRecuadroQuorum.top = shpRecuadroFecha.top - 120
    shpRecuadroQuorum.BorderColor = MiRojo
End If
End Sub

Private Sub lblOcupadosNoIdentificados_Click(Index As Integer)
MsgBox Str(Index)
End Sub

Private Sub picA_Click(Index As Integer)
MsgBox ("picA(" & Index & ")")
End Sub

Private Sub picB_Click(Index As Integer)
MsgBox ("picB(" & Index & ")")
End Sub

Private Sub picC_Click(Index As Integer)
MsgBox ("picC(" & Index & ")")
End Sub

Private Sub Picture1_Click()
MsgBox ("Picture1")
End Sub

Private Sub Picture2_Click()
MsgBox ("Picture 2")
End Sub

Private Sub Picture3_Click()
MsgBox ("Picture3")
End Sub

Private Sub Ws_Connect()
    If Ws.State = sckConnected Then
       Ws.SendData strPath & vbCrLf
       DoEvents
    Else
        Call AltaLogGeneral("SQV SERVER", "Error : No se puedo levantar la aplicacion, Reintente")
        'MsgBox "Error : No se puedo levantar la aplicacion, Reintente"
    End If
    Ws.Close
End Sub
Private Function EnviarMensajesConsolaSqv(strTipo As String, _
                                    strComponente As String, _
                                        strObjeto As String, _
                                      strAtributo As String, _
                                         strValor As String, _
                                    strComentario As String) As Long
    On Error Resume Next
    Dim strSql As String
    strSql = "INSERT INTO consola_sqv_mensajes (Tipo, Componente, Objeto, Atributo, Valor, Comentario) VALUES " _
           & "('" & strTipo & "','" & strComponente & "', '" & strObjeto & "', '" & strAtributo & "', '" & strValor & "','" & strComentario & "')"
    Cn.Execute (strSql)
    EnviarMensajesConsolaSqv = err.Number
End Function

Private Sub AlmacenarActa() ' grabar acta / guardar acta

    Dim strSql As String
    Dim strDesempate                  As String
    Dim blHayDesempate                As Boolean
    ' Valores para encontrar el acta
    Dim xBancaActual                  As Long
    Dim X                             As Long
    Dim strResultado                  As String
    Dim strIdLegislador               As String
    Dim rsTemp                        As ADODB.Recordset
    Dim strBuscarLegislador           As String
    Dim strBloquePolitico             As String
    Dim strDepartamento               As String
    Dim strGrupoPolitico              As String
    Dim IdEstado                      As Integer
    Dim PVez                          As Boolean
    Dim xZonaAsignada                 As Long
    Dim strPresidente                 As String
    ' Presentes
    Dim xPresentesTotales             As Long
    Dim xPresentesIdentificables      As Long
    Dim xPresentesNOIdentificables    As Long
    ' Ausentes
    Dim xAusentesTotales              As Long
    ' Votos Afirmativos
    Dim xVotosAfirmIdentificables     As Long
    Dim xVotosAfirmNOIdentificables   As Long
    Dim xVotosAfirmTotales            As Long
    ' Votos negativos
    Dim xVotosNegatIdentificables     As Long
    Dim xVotosNegatNOIdentificables   As Long
    Dim xVotosNegatTotales            As Long
    ' Abstenciones
    Dim xAbstencionesIdentificables   As Long
    Dim xAbstencionesNOIdentificables As Long
    Dim xAbstencionesTotales          As Long
    
    Dim xInicio                       As Long
    Dim xUltimaActaSesion             As Long
    Set rsTemp = New ADODB.Recordset
    frmCartel2011.Update
    frmCartel2011.ActualizarColoresVotos
    DoEvents
    PVez = True
    xPresentesTotales = 0
    xPresentesIdentificables = 0
    xPresentesNOIdentificables = 0
    xAusentesTotales = 0
    xVotosAfirmIdentificables = 0
    xVotosAfirmNOIdentificables = 0
    xVotosAfirmTotales = 0
    xVotosNegatIdentificables = 0
    xVotosNegatNOIdentificables = 0
    xVotosNegatTotales = 0
    xAbstencionesIdentificables = 0
    xAbstencionesNOIdentificables = 0
    xAbstencionesTotales = 0
    
    ' --------------------------------------------------------------------------------
    ' Averiguar el nombre del presidente
    ' --------------------------------------------------------------------------------
    strSql = "SELECT apellido + ', ' + nombre AS Presidente FROM Legisladores WHERE id = '" & Trim(EstadoActual.VectorIdentificacion(0)) & "'"
    Call SetearOtroRs(strSql)
    If RsOtro.EOF <> True Then
        strPresidente = RsOtro.Fields(0).Value
    Else
        strPresidente = "NO DISPONIBLE"
    End If
    RsOtro.Close
    If (EstadoActual.TipoDeOperacion = "votnom" And xTipoVotacion <> "votnum") Or EstadoActual.TipoDeOperacion = "paslis" Then
        ' --------------------------------------------------------------------------------
        ' Escribir tabla DetalleActa
        ' --------------------------------------------------------------------------------
        strSql = "SELECT Período_Legislativo, Sesión, Nro_de_Acta, Versión_Acta, " _
               & "Operación, Numero_de_banca, Resultado, Legislador_asignado, Año_inicio_mandato, Año_fin_mandato, " _
               & "Zona_asignada , Grupo_Politico,Bloque_político, Departamento, estado " _
               & "From DetalleActas WHERE 1 = 0"
        Call SetearRsW(strSql)
        ' Determinar si se debe considerar al presidente como legislador o como vicegobernador
        Dim SesionAux As Long
        Dim NroActaAux As Long
        SesionAux = 0
        NroActaAux = 0
        If Not SesionValida(EstadoActual.PeriodoLegislativo, EstadoActual.Sesion) Then
            Dim r As Integer
            SesionAux = EstadoActual.Sesion 'Para volver a este numero de sesion
            NroActaAux = EstadoActual.NroActa 'Mismo acta
            EstadoActual.Sesion = 9999
            EstadoActual.NroActa = PAS(EstadoActual.PeriodoLegislativo, EstadoActual.Sesion, "usoint")
        End If
    End If
    xInicio = IIf(xPresidenteLegislador And EstadoActual.PresidenteHabilitadoParaVotar, 0, 1) 'revisar
'A02
    PanelResultadosInicializar
'A02 END
    For X = xInicio To xUltimaBanca
        ' VECTOR IDENTIFICACION
        strIdLegislador = ""
        If EstadoActual.VectorPresencia(X) = PRESENTE _
            And (((EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis") And EstadoActual.VectorIdentificacion(X) <> NO_IDENTIFICADO) _
                Or (EstadoActual.TipoDeOperacion = "votnum")) Then
            If (((EstadoActual.TipoDeOperacion = "votnom" And xTipoVotacion <> "votnum") Or EstadoActual.TipoDeOperacion = "paslis") And EstadoActual.VectorIdentificacion(X) <> NO_IDENTIFICADO) Then
                ' Identificar al legislador
                strIdLegislador = Trim(LCase(EstadoActual.VectorIdentificacion(X)))
                strBuscarLegislador = "SELECT Legisladores.*,distritos.distrito as Provincia,legisladores_para_actualizar.estado FROM Legisladores LEFT OUTER JOIN distritos ON Legisladores.distrito = distritos.id_distrito LEFT OUTER JOIN legisladores_para_actualizar ON Legisladores.id = legisladores_para_actualizar.id WHERE Legisladores.id = '" & strIdLegislador & "'"
                rsTemp.Open strBuscarLegislador, Cn, adOpenForwardOnly, adLockReadOnly
                If Not rsTemp.EOF Then
                    strBloquePolitico = GetCadena(rsTemp.Fields("bloque_Politico").Value)
                    strDepartamento = GetCadena(rsTemp.Fields("Provincia").Value)
                    strGrupoPolitico = GetCadena(rsTemp.Fields("Grupo_Politico").Value)
                    If IsNull(rsTemp.Fields("estado")) Then
                        IdEstado = 8
                    Else
                        IdEstado = rsTemp.Fields("estado")
                    End If
                    xZonaAsignada = -1 'GetNumero(rsTemp.Fields("Zona").Value)
                Else
                    strBloquePolitico = "BLOQUE NO IDENTIFICADO"
                    strDepartamento = "ND"
                    IdEstado = 8
                    xZonaAsignada = -1
                End If
                rsTemp.Close
            End If
            Select Case EstadoActual.TipoDeOperacion
            Case "votnom", "votnum"
                Select Case EstadoActual.VectorResultados(X)
                    Case AFIRMATIVO
                        strResultado = IIf(X = 0 And xVotoSenadorEmpate > "", IIf(xVotoSenadorEmpate = "s", "AFIRMATIVO", "NEGATIVO"), "AFIRMATIVO")
                    Case NEGATIVO
                        strResultado = IIf(X = 0 And xVotoSenadorEmpate > "", IIf(xVotoSenadorEmpate = "s", "AFIRMATIVO", "NEGATIVO"), "NEGATIVO")
                    Case ABSTENCION, ABSTENCION_AUTORIZADA
                        strResultado = "ABSTENCION"
                End Select
'A02
                'Cargo el Panel de resultados
                PanelResultadosCargar EstadoActual.TipoDeOperacion, strIdLegislador, strResultado
'A02 END
            Case "paslis"
                strResultado = "PRESENTE"
'A02
                'Cargo el Panel de resultados
                PanelResultadosCargar EstadoActual.TipoDeOperacion, strIdLegislador, strResultado
'A02 END
            End Select
        Else
            strResultado = "AUSENTE"
        End If
        ' Escribir tabla solo en votnom, paslis
        If (EstadoActual.TipoDeOperacion = "votnom" And xTipoVotacion <> "votnum") Or EstadoActual.TipoDeOperacion = "paslis" Then
            If PVez = True Then
                Dim rsTemp2 As ADODB.Recordset
                Set rsTemp2 = New ADODB.Recordset
                frmMain.SetearRsAux "SELECT Legisladores.bloque_politico,Legisladores.grupo_politico,distritos.distrito AS Provincia, legisladores_para_actualizar.estado FROM Legisladores LEFT OUTER JOIN distritos ON Legisladores.distrito = distritos.id_distrito LEFT OUTER JOIN legisladores_para_actualizar ON legisladores_para_actualizar.id = Legisladores.id WHERE legisladores.id = '" & EstadoActual.VectorIdentificacion(0) & "'", rsTemp2
                RsWrite.AddNew
                RsWrite.Fields("Período_Legislativo").Value = EstadoActual.PeriodoLegislativo
                RsWrite.Fields("Sesión").Value = Int(EstadoActual.Sesion)
                RsWrite.Fields("Nro_de_Acta").Value = Int(EstadoActual.NroActa)
                RsWrite.Fields("Versión_Acta").Value = 0
                RsWrite.Fields("Operación").Value = EstadoActual.TipoDeOperacion
                RsWrite.Fields("Numero_de_banca").Value = 0
                RsWrite.Fields("Resultado").Value = "AUSENTE"
                RsWrite.Fields("Legislador_asignado").Value = EstadoActual.VectorIdentificacion(0)
                RsWrite.Fields("Año_inicio_mandato").Value = Date
                RsWrite.Fields("Año_fin_mandato").Value = Date
                RsWrite.Fields("Zona_asignada").Value = 0
                RsWrite.Fields("Bloque_político").Value = rsTemp2.Fields("bloque_politico")
                RsWrite.Fields("Grupo_politico").Value = rsTemp2.Fields("grupo_politico")
                RsWrite.Fields("Departamento").Value = rsTemp2.Fields("Provincia")
                If IsNull(rsTemp2.Fields("estado")) Then
                    RsWrite.Fields("estado").Value = 1
                Else
                    RsWrite.Fields("estado").Value = rsTemp2.Fields("estado")
                End If
                RsWrite.Update
                PVez = False
            End If
            RsWrite.AddNew
            RsWrite.Fields("Período_Legislativo").Value = EstadoActual.PeriodoLegislativo
            RsWrite.Fields("Sesión").Value = Int(EstadoActual.Sesion)
            RsWrite.Fields("Nro_de_Acta").Value = Int(EstadoActual.NroActa)
            RsWrite.Fields("Versión_Acta").Value = 0
            RsWrite.Fields("Operación").Value = EstadoActual.TipoDeOperacion
            RsWrite.Fields("Numero_de_banca").Value = X
            RsWrite.Fields("Resultado").Value = strResultado
            RsWrite.Fields("Legislador_asignado").Value = strIdLegislador
            RsWrite.Fields("Año_inicio_mandato").Value = Date
            RsWrite.Fields("Año_fin_mandato").Value = Date
            RsWrite.Fields("Zona_asignada").Value = xZonaAsignada
            RsWrite.Fields("Bloque_político").Value = Trim(strBloquePolitico)
            RsWrite.Fields("Grupo_politico").Value = Trim(strGrupoPolitico)
            RsWrite.Fields("Departamento").Value = Trim(strDepartamento)
            RsWrite.Fields("estado").Value = IdEstado
            RsWrite.Update
        End If
        'Totalizar
        ' AUSENTES TOTALES Y PRESENTES TOTALES
        If EstadoActual.VectorPresencia(X) = PRESENTE Then
            ' PRESENTES IDENTIFICABLES
            If EstadoActual.VectorIdentificacion(X) <> NO_IDENTIFICADO Then
                xPresentesIdentificables = xPresentesIdentificables + 1 ' Contar Legisladores identificados
                If EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum" Then
                    ' Contar votos afirmativos identificables
                    If (X = 0 And xVotoSenadorEmpate = "s") Or ((xVotoSenadorEmpate = "" Or X > 0) And EstadoActual.VectorResultados(X) = AFIRMATIVO) Then
                        xVotosAfirmIdentificables = xVotosAfirmIdentificables + 1
                    ElseIf (X = 0 And xVotoSenadorEmpate = "n") Or ((xVotoSenadorEmpate = "" Or X > 0) And EstadoActual.VectorResultados(X) = NEGATIVO) Then
                        ' Contar votos negativos identificables
                        xVotosNegatIdentificables = xVotosNegatIdentificables + 1
                    ' ElseIf EstadoActual.VectorResultados(X) = ABSTENCION Then
                    ElseIf EstadoActual.VectorResultados(X) = ABSTENCION Or EstadoActual.VectorResultados(X) = ABSTENCION_AUTORIZADA Then
                        xAbstencionesIdentificables = xAbstencionesIdentificables + 1
                    End If
                End If
            Else ' GENTE NO IDENTIFICADA
                ' PRESENTES NO IDENTIFICADOS Y PRESENTES IDENTIFICADOS
                xPresentesNOIdentificables = xPresentesNOIdentificables + 1
                If EstadoActual.TipoDeOperacion = "votnum" Then
                    If EstadoActual.VectorResultados(X) = NEGATIVO Then
                        ' Contar votos afirmativos Identificables
                        xVotosNegatNOIdentificables = xVotosNegatNOIdentificables + 1
                    ElseIf EstadoActual.VectorResultados(X) = AFIRMATIVO Then
                        ' Contar votos afirmativos NO identificables
                        xVotosAfirmNOIdentificables = xVotosAfirmNOIdentificables + 1
                    ' ElseIf EstadoActual.VectorResultados(X) = ABSTENCION Then
                    ElseIf EstadoActual.VectorResultados(X) = ABSTENCION Or EstadoActual.VectorResultados(X) = ABSTENCION_AUTORIZADA Then
                        xAbstencionesNOIdentificables = xAbstencionesNOIdentificables + 1
                    End If
                End If
            End If
        Else
            xAusentesTotales = xAusentesTotales + 1
        End If
        If (X = 10 Or X = 20 Or X = 30 Or X = 40 Or X = 50 Or X = 60 Or X = 70 Or X = 80 Or X = 90 Or X = 100 Or X = 120 Or X = 140 Or X = 160 Or X = 180 Or X = 200 Or X = 220 Or X = 230 Or X = 240) Then
            DoEvents
        End If
    Next X
    'If xPresidenteLegislador Then
    If EstadoActual.PresidenteHabilitadoParaVotar Then
        xAusentesTotales = xAusentesTotales - 1
    End If
    If (EstadoActual.TipoDeOperacion = "votnom" And xTipoVotacion <> "votnum") Or EstadoActual.TipoDeOperacion = "paslis" Then
        RsWrite.Close
    End If
    
    
    'FIX PARA RESULTADOS
    xVotosAfirmIdentificables = 0
    xVotosNegatIdentificables = 0
    xAbstencionesIdentificables = 0
    
    Dim i As Integer
    
    'FIX Afirmativos Identificabless
    
    For i = 1 To 256
        If (EstadoActual.VectorIdentificacion(i) <> NO_IDENTIFICADO) Then
            If (EstadoActual.VectorResultados(i) = AFIRMATIVO) Then
                xVotosAfirmIdentificables = xVotosAfirmIdentificables + 1
            End If
        End If
    Next i
    
    'FIX Negativos Identificables
    
    For i = 1 To 256
        If (EstadoActual.VectorIdentificacion(i) <> NO_IDENTIFICADO) Then
            If (EstadoActual.VectorResultados(i) = NEGATIVO) Then
                xVotosNegatIdentificables = xVotosNegatIdentificables + 1
            End If
        End If
    Next i
    
    'FIX Abstenciones Identificables
    
    For i = 1 To 256
        If (EstadoActual.VectorIdentificacion(i) <> NO_IDENTIFICADO) Then
            If (EstadoActual.VectorResultados(i) = ABSTENCION) Then
                xAbstencionesIdentificables = xAbstencionesIdentificables + 1
            End If
        End If
    Next i
    
    'Totales generales
    xVotosAfirmTotales = xVotosAfirmIdentificables + xVotosAfirmNOIdentificables
    xVotosNegatTotales = xVotosNegatNOIdentificables + xVotosNegatIdentificables
    xAbstencionesTotales = xAbstencionesIdentificables + xAbstencionesNOIdentificables
    xPresentesTotales = xPresentesIdentificables + xPresentesNOIdentificables
    
        
    'strDesempate = IIf((xVotosAfirmTotales = xVotosNegatTotales) And xVotosAfirmTotales > 0, "Si", "No")
    strDesempate = IIf(xHuboDesempate, "Si", "No")
    
    ' --------------------------------------------------------------------------------
    ' Escribir tabla Actas
    ' --------------------------------------------------------------------------------
    strSql = "SELECT Tipo_de_operación, Período_Legislativo, Sesión, " _
           & "Número_de_Acta, Versión_Acta, Ultima_Versión_Acta, " _
           & "Nombre_del_Acta, Fecha, Hora, Tipo_de_Quorum, " _
           & "Base_de_Mayoria, Tipo_de_Mayoria, Miembros_del_cuerpo, " _
           & "Desempate, Votacion, Presidente, Presentes_Identificables, " _
           & "Presentes_No_Identificables, Presentes_Total, " _
           & "Ausentes_Total, Votos_Afirm_Identificables, " _
           & "Votos_Afirm_No_Identificables, Votos_Afirm_Desempate, " _
           & "Votos_Afirm_Total, Votos_Neg_Identificables, " _
           & "Votos_Neg_No_Identificables, Votos_Neg_Desempate, " _
           & "Votos_Neg_Total, Abstenciones_Identificables, " _
           & "Abstenciones_No_Identificables, Abstenciones_Total, " _
           & "Fecha_Modificacion, Hora_Modificacion, " _
           & "Usuario_Modificacion, IP_Modificacion, Observaciones, " _
           & "NroOrdenDia , Tipo, Origen, Destino, vota_presidente, Reunion, " _
           & "presidente_habilitado_votar, resultado_voto_presidente " _
           & " From Actas " _
           & " WHERE 1=0 "
    Call SetearRsW(strSql)
    
    RsWrite.AddNew
    RsWrite.Fields("Tipo_de_operación").Value = IIf(EstadoActual.TipoDeOperacion = "votnom" And xTipoVotacion = "votnum", "votnum", EstadoActual.TipoDeOperacion)
    RsWrite.Fields("Período_Legislativo").Value = EstadoActual.PeriodoLegislativo
    RsWrite.Fields("Sesión").Value = EstadoActual.Sesion
    RsWrite.Fields("Número_de_Acta").Value = EstadoActual.NroActa
    RsWrite.Fields("Versión_Acta").Value = 0
    RsWrite.Fields("Ultima_Versión_Acta").Value = 0
    RsWrite.Fields("Reunion").Value = EstadoActual.Reunion
    If (EstadoActual.TipoDeOperacion = "paslis") Then
        RsWrite.Fields("Nombre_del_Acta").Value = ""
    Else
        RsWrite.Fields("Nombre_del_Acta").Value = EstadoActual.TituloDelActa
    End If
    RsWrite.Fields("Fecha").Value = EstadoActual.FechaVotacion
    RsWrite.Fields("Hora").Value = EstadoActual.HoraVotacion
    RsWrite.Fields("Tipo_de_Quorum").Value = EstadoActual.TipoMayoriaQuorum
    RsWrite.Fields("Base_de_Mayoria").Value = EstadoActual.BaseMayoria
    RsWrite.Fields("Tipo_de_Mayoria").Value = EstadoActual.TipoMayoria
    RsWrite.Fields("Miembros_del_cuerpo").Value = xMiembrosDelCuerpo
    RsWrite.Fields("Desempate").Value = strDesempate
    RsWrite.Fields("Votacion").Value = IIf(EstadoActual.TipoDeOperacion = "paslis", CartelActual.LeyendaQuorum, CartelActual.Resultado)
    RsWrite.Fields("Presidente").Value = EstadoActual.VectorIdentificacion(0)
    RsWrite.Fields("Presentes_Identificables").Value = GetIdentificados 'NICO
    RsWrite.Fields("Presentes_No_Identificables").Value = GetNoIdentificadosSobrePresentes
    RsWrite.Fields("Presentes_Total").Value = getPresentes
    RsWrite.Fields("Ausentes_Total").Value = GetAusentes
    RsWrite.Fields("Votos_Afirm_Identificables").Value = xVotosAfirmIdentificables
    RsWrite.Fields("Votos_Afirm_No_Identificables").Value = xVotosAfirmNOIdentificables
    blHayDesempate = (UCase(strDesempate) = "SI")
    RsWrite.Fields("Votos_Afirm_Desempate").Value = IIf(blHayDesempate And LCase(CartelActual.Resultado) = "afirmativo", 1, 0)
    RsWrite.Fields("Votos_Afirm_Total").Value = xVotosAfirmTotales + IIf(blHayDesempate And LCase(CartelActual.Resultado) = "afirmativo", 1, 0)
    RsWrite.Fields("Votos_Neg_Identificables").Value = xVotosNegatIdentificables
    RsWrite.Fields("Votos_Neg_No_Identificables").Value = xVotosNegatNOIdentificables
    RsWrite.Fields("Votos_Neg_Desempate").Value = IIf(blHayDesempate And LCase(CartelActual.Resultado) = "negativo", 1, 0)
    RsWrite.Fields("Votos_Neg_Total").Value = xVotosNegatTotales + IIf(blHayDesempate And LCase(CartelActual.Resultado) = "negativo", 1, 0)
    RsWrite.Fields("Abstenciones_Identificables").Value = xAbstencionesIdentificables
    RsWrite.Fields("Abstenciones_No_Identificables").Value = xAbstencionesNOIdentificables
    RsWrite.Fields("Abstenciones_Total").Value = xAbstencionesTotales
    RsWrite.Fields("Fecha_Modificacion").Value = Date
    RsWrite.Fields("Hora_Modificacion").Value = Time
    RsWrite.Fields("Usuario_Modificacion").Value = "SQV"
    RsWrite.Fields("IP_Modificacion").Value = EstadoActual.IP_Consola
    RsWrite.Fields("Observaciones").Value = " "
    RsWrite.Fields("NroOrdenDia").Value = 0
    RsWrite.Fields("Tipo").Value = ""
    RsWrite.Fields("Origen").Value = ""
    RsWrite.Fields("Destino").Value = ""
    RsWrite.Fields("vota_presidente").Value = IIf(xPresidenteLegislador = True, 1, 0)
    RsWrite.Fields("vota_presidente").Value = IIf(EstadoActual.PresidenteHabilitadoParaVotar, 1, 0)
    '110227
    RsWrite.Fields("presidente_habilitado_votar").Value = IIf(EstadoActual.PresidenteHabilitadoParaVotar, 1, 0)
    RsWrite.Fields("resultado_voto_presidente").Value = EstadoActual.ResultadoVotoPresidente
    '
    RsWrite.Update
    RsWrite.Close
    
    If (EstadoActual.TipoDeOperacion = "votnom" And xTipoVotacion <> "votnum") Or EstadoActual.TipoDeOperacion = "paslis" Then
        Call AuditarLegisladoresAusentes
    End If
    ' Incrementar en 1 el # de sesion de proxima acta en tabla Sesion
    strSql = "SELECT * From Sesion WHERE Período_Legislativo = '" & Trim(EstadoActual.PeriodoLegislativo) & "' AND sesión = " & Trim(EstadoActual.Sesion)
    Call SetearRsW(strSql)
    xUltimaActaSesion = RsWrite.Fields("Próximo_Acta").Value + 1
    RsWrite.Fields("Próximo_Acta").Value = xUltimaActaSesion
    RsWrite.Update
    RsWrite.Close
    EstadoActual.NroActa = xUltimaActaSesion
    If SesionAux <> 0 Then 'Si guardo sesion 99999
        EstadoActual.NroActa = NroActaAux
        EstadoActual.Sesion = SesionAux
    End If
End Sub
Private Sub AuditarLegisladoresAusentes()
    ' ---------------------------------------------------------------------------
    ' Dejar constancia en detalles actas de los legisladores que no fueron a sesionar
    ' ---------------------------------------------------------------------------
    Dim strSql            As String
    Dim strSql2           As String
    Dim strIdAusente      As String
    Dim strSqlDetalleActa As String
    Dim rsDet             As ADODB.Recordset
    Dim strPeriodoLeg     As String
    Dim strSesion         As String
    Dim strNroActa        As String
    
    Set rsDet = New ADODB.Recordset
    strPeriodoLeg = LCase(Trim(EstadoActual.PeriodoLegislativo))
    strSesion = LCase(Trim(Str(EstadoActual.Sesion)))
    strNroActa = LCase(Trim(Str(EstadoActual.NroActa)))
    
    ' Listado de todas las bancas cuyos legisladores estuvieron AUSENTES en el acta
    strSql = "SELECT numero_de_banca, resultado, legislador_asignado, Bloque_político,Grupo_Politico, Departamento,estado " _
                      & "From detalleactas WHERE resultado = 'AUSENTE' AND Legislador_asignado = '' AND período_legislativo = '" & strPeriodoLeg & "' AND sesión = " & strSesion & " AND " _
                      & "nro_de_acta = " & strNroActa & " And versión_acta = 0"
    
    Call SetearRsW(strSql)
    If Not RsWrite.EOF Then
    
        ' Legisladores encontrados como ausente para el acta en cuestion
        strSql2 = "SELECT legisladores_activos.id as id, Legisladores.bloque_politico,Legisladores.grupo_politico, distritos.distrito AS Provincia, legisladores_para_actualizar.estado From legisladores_activos inner join legisladores on legisladores.id = legisladores_activos.id LEFT OUTER JOIN distritos ON Legisladores.distrito = distritos.id_distrito LEFT OUTER JOIN legisladores_para_actualizar ON legisladores_para_actualizar.id = Legisladores.id WHERE legisladores_activos.ID <> '" & EstadoActual.VectorIdentificacion(0) & "' AND (legisladores_activos.ID NOT IN " _
               & "(SELECT Legislador_asignado From detalleactas " _
               & "WHERE (Sesión = " & strSesion & ") AND (Nro_de_Acta = " & strNroActa & ") AND " _
               & "(Versión_Acta = 0) AND (Período_Legislativo = '" & strPeriodoLeg & "') AND " _
               & "(Resultado <> 'AUSENTE'))) AND (Legisladores.es_legislador = 1)" 'AP 091031 Faltaria considerar el caso de pruebas de mantenimiento. En este caso posiblemente seria necesario manejar el 1 variable segun si es mantenimiento o no
        rsDet.CursorLocation = adUseClient
        rsDet.Open strSql2, Cn, adOpenDynamic, adLockOptimistic
        ' SetearRsAux strSql2, rsDet
        
        If rsDet.RecordCount > 0 Then
            rsDet.MoveFirst
            RsWrite.MoveFirst
            While Not RsWrite.EOF And Not rsDet.EOF
                If RsWrite.Fields("numero_de_banca").Value <> 0 Then
                    strIdAusente = Trim(RsWrite.Fields(0).Value)
                    ' Si encontro legisladores ausentes en un acta, se debe registralos
                    RsWrite.Fields("Legislador_Asignado").Value = Trim(rsDet.Fields(0).Value)
                    If IsNull(rsDet.Fields("bloque_politico").Value) Then
                        RsWrite.Fields("Bloque_político").Value = ""
                    Else
                        RsWrite.Fields("Bloque_político").Value = Trim(rsDet.Fields("bloque_politico").Value)
                    End If
                    If IsNull(rsDet.Fields("Provincia").Value) Then
                        RsWrite.Fields("Departamento").Value = ""
                    Else
                        RsWrite.Fields("Departamento").Value = Trim(rsDet.Fields("Provincia").Value)
                    End If
                    If IsNull(rsDet.Fields("Grupo_Politico").Value) Then
                        RsWrite.Fields("Grupo_Politico").Value = ""
                    Else
                        RsWrite.Fields("Grupo_Politico").Value = Trim(rsDet.Fields("Grupo_Politico").Value)
                    End If
                    If IsNull(rsDet.Fields("estado").Value) Then
                        RsWrite.Fields("estado") = 8
                    Else
                        RsWrite.Fields("estado") = rsDet.Fields("estado")
                    End If
                    'Cargo el Panel de resultados con Legisladores Ausentes
                    PanelResultadosCargar EstadoActual.TipoDeOperacion, Trim(rsDet.Fields(0).Value), "AUSENTE"
                End If
    'A02 END
                RsWrite.MoveNext
                rsDet.MoveNext
            Wend
        End If
        If Not RsWrite.EOF And Not rsDet.EOF Then RsWrite.Update
        rsDet.Close
        Set rsDet = Nothing
    End If
    RsWrite.Close
End Sub

Private Sub AuditarLegisladoresAusentesANTERIORE_MAL_1()
    ' ---------------------------------------------------------------------------
    ' Dejar constancia en detalles actas de los legisladores que no fueron a sesionar
    ' ---------------------------------------------------------------------------
    Dim strSql            As String
    Dim strSql2           As String
    Dim strIdAusente      As String
    Dim strSqlDetalleActa As String
    Dim rsDet             As ADODB.Recordset
    Dim strPeriodoLeg     As String
    Dim strSesion         As String
    Dim strNroActa        As String
    
    Set rsDet = New ADODB.Recordset
    strPeriodoLeg = LCase(Trim(EstadoActual.PeriodoLegislativo))
    strSesion = LCase(Trim(Str(EstadoActual.Sesion)))
    strNroActa = LCase(Trim(Str(EstadoActual.NroActa)))
    
    ' Listado de todas las bancas cuyos legisladores estuvieron AUSENTES en el acta
    strSql = "SELECT numero_de_banca, resultado, legislador_asignado " _
                      & "From detalleactas WHERE resultado = 'AUSENTE' AND período_legislativo = '" & strPeriodoLeg & "' AND sesión = " & strSesion & " AND " _
                      & "nro_de_acta = " & strNroActa & " And versión_acta = 0"
    
    Call SetearRsW(strSql)
    ' Legisladores encontrados como ausente para el acta en cuestion
    strSql2 = "SELECT ID From legisladores_activos WHERE ID NOT IN " _
           & "(SELECT Legislador_asignado From detalleactas " _
           & "WHERE (Sesión = " & strSesion & ") AND (Nro_de_Acta = " & strNroActa & ") AND " _
           & "(Versión_Acta = 0) AND (Período_Legislativo = '" & strPeriodoLeg & "') AND " _
           & "(Resultado <> 'AUSENTE'))"
    rsDet.CursorLocation = adUseClient
    rsDet.Open strSql2, Cn, adOpenDynamic, adLockOptimistic
    ' SetearRsAux strSql2, rsDet
    
    If rsDet.RecordCount > 0 Then
        rsDet.MoveFirst
        RsWrite.MoveFirst
        While Not RsWrite.EOF And Not rsDet.EOF
            strIdAusente = Trim(RsWrite.Fields(0).Value)
            ' Si encontro legisladores ausentes en un acta, se debe registralos
            RsWrite.Fields("Legislador_Asignado").Value = Trim(rsDet.Fields(0).Value)
'A02
            'Cargo el Panel de resultados con Legisladores Ausentes
            PanelResultadosCargar EstadoActual.TipoDeOperacion, Trim(rsDet.Fields(0).Value), "AUSENTE"
'A02 END
            RsWrite.MoveNext
            rsDet.MoveNext
        Wend
    End If
    If Not RsWrite.EOF And Not rsDet.EOF Then RsWrite.Update 'CONSULTA MAXIMILIANO
    RsWrite.Close
    rsDet.Close
    Set rsDet = Nothing
End Sub

' Esta fucion envia un mensaje al servidor de bancas
Private Sub EnviarMensajesBancas(MensajeBanca As MensajeSistema)
    On Error Resume Next
    Dim strTipo       As String
    Dim strComponente As String
    Dim strObjeto     As String
    Dim strAtributo   As String
    Dim strValor      As String
    Dim strComentario As String
    Dim strSql        As String
    With MensajeBanca
        strTipo = Trim(.sTipo)
        strComponente = Trim(.sComponente)
        strObjeto = Trim(.sObjeto)
        strAtributo = Trim(.sAtributo)
        strValor = Trim(.sValor)
        strComentario = "SB:" & Str(xNroMensajeSB) & "|" & .sComentario
    End With
    ' strSql = "INSERT INTO sqv_sb_mensajes (Tipo, Componente, Objeto, Atributo, Valor, Comentario) VALUES " _
           & "('" & strTipo & "','" & strComponente & "', '" & strObjeto & "', '" & strAtributo & "', '" & strValor & "','" & strComentario & "')"
    ' Mismo SQL de arriba, pero disparado desde SP
    strSql = "insert_sqv_sb_mensajes('" & strTipo & "', '" & strComponente & "', '" & strObjeto & "', '" & strValor & "', '" & strComentario & "', '" & strAtributo & "')"
    strUltimoMensaje_SQV_SB = strTipo & ";" & strComponente & ";" & strObjeto & ";" & strValor & ";" & strComentario & ";" & strAtributo
    With Cn
        .Execute (strSql)
    End With
    ' Log
    nLogSQVPrueba = nLogSQVPrueba + 1
    xLogSQVPrueba = xLogSQVPrueba & "    " & Format(nLogSQVPrueba, "0000000") & "¦ " & Now & "¦ " & "Atiende Msj SB¦" & Str(xNroMensajeSB) & "¦" & strSql & vbCrLf
    Call AltaLogGeneral("SQV SERVER", " " & Format(nLogSQVPrueba, "0000000") & "¦ " & Now & "¦ " & "Atiende Msj SB¦" & Str(xNroMensajeSB) & "¦" & strSql)
    
End Sub
Private Sub ServerOnOff()
    On Error GoTo TrapError
    If False Then
        cmdConfig.Enabled = Not blServerPrendido
        If blServerPrendido = False Then
            cmdTerminar.Caption = "&Iniciar Server"
            lblFechaInicioServer.Caption = "SERVER DETENIDO"
        Else
            cmdTerminar.Caption = "&Detener Server"
            lblFechaInicioServer.Caption = Now
        End If
        Timer.Interval = 1000 / xIntervalo
        Timer.Enabled = blServerPrendido '(Para permitir prender y apagar el server cambiar por blServerPrendido. Sino dejarlo en True).
    Else
        cmdTerminar.Enabled = False
        cmdTerminar.Caption = "&Detener Server"
        Timer.Interval = 1000 / xIntervalo
        Timer.Enabled = True
    End If
Exit Sub
TrapError:
    Select Case err.Number
        Case 11
            xIntervalo = 2
             txtVecesPorSegundo.text = ""
            'Resume
        Case Else
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            'Resume
    End Select
End Sub
Private Sub cmdTerminar_Click()
    ' Prender o apagar el servidor
    Exit Sub
    If blServerPrendido Then
        blServerPrendido = False
        xCiclosTotales = 0
        Call AltaLogGeneral("SQV SERVER", " Deteniendo SQV Server " & Now)
    Else
        blServerPrendido = True
    End If
    Call ServerOnOff
End Sub
Private Sub BorrarMensajesTotales()
    On Error GoTo TrapError
    Dim strSql As String
    
    Screen.MousePointer = 11
        ' Eliminar todos los mensajes de SQV al SB
        ' strSql = "DELETE FROM sqv_sb_mensajes"
        strSql = "TRUNCATE TABLE sqv_sb_mensajes"
        Cn.Execute (strSql)
        ' Eliminar todos los mensahes del SB al SQV
        strSql = "TRUNCATE TABLE sb_sqv_mensajes"
        Cn.Execute (strSql)
        ' Eliminar todos los mensajes de la consola al SQV
        strSql = "TRUNCATE TABLE consola_sqv_mensajes"
        Cn.Execute (strSql)
    Screen.MousePointer = 0
Exit Sub
TrapError:
    Call AltaLogGeneral("SQV SERVER", "BorrarMensajesTotales" & "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source)
    'MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
    End
    'Resume
End Sub
Private Sub LeerEstadoRecinto()

    Dim strSql As String
    
    strSql = "SELECT * From vector"
    Call SetearRs(strSql)
    With CartelActual
        .Abstenciones = rs.Fields("Abstenciones").Value
        .Afirmativos = rs.Fields("Afirmativos").Value
        .Ausentes = GetNumero(rs.Fields("Ausentes").Value)
        .LeyendaQuorum = rs.Fields("Leyenda_Quorum").Value
        .MinimoVotosParaAfirmativo = rs.Fields("Minimo_de_votos_para_afirmativa").Value
        .Negativos = rs.Fields("Negativos").Value
        .Presentes = rs.Fields("presentes").Value
        .Resultado = rs.Fields("Resultado").Value
    End With
    
    With EstadoActual
        .ActaGrabada = GetNumero(rs.Fields("Acta_Grabada"))
        .Ausentes = GetNumero(rs.Fields("Ausentes").Value)
        .BaseMayoria = GetCadena(rs.Fields("Base_de_Mayoría").Value)
        .CartelEncendido = GetNumero(rs.Fields("Encender_Carteles").Value)
        .EstadoVotacion_y_PasList = GetCadena(rs.Fields("Estado_de_votacion").Value)
        .GrabarAutomaticamente = GetNumero(rs.Fields("Grabar_automaticamente").Value)
        .IdentificadorDeFormulario = GetCadena(rs.Fields("Identificador_de_Formulario").Value)
        .IP_Consola = GetCadena(rs.Fields("IP_Consola_Habilitada").Value)
        .ListarAutomaticamente = GetNumero(rs.Fields("Listar_automaticamente").Value)
        .MensajeAlOperador = GetCadena(rs.Fields("Mensaje_al_operador").Value)
        .TipoDeAbstencion = GetCadena(rs.Fields("tipo_de_abstención").Value)
        .ModoMantenimientoBancas = GetNumero(rs.Fields("Modo_Mantenimiento_Bancas").Value)
        .ModoNormalMantSistema = GetNumero(rs.Fields("Modo_Normal_Mant_Sistema").Value)
        .NroActa = GetNumero(rs.Fields("Nro_de_Acta").Value)
        .Reunion = GetNumero(rs.Fields("Reunion").Value)
        .OcupadosNoIdentificados = GetNumero(rs.Fields("Ocupadas_no_identificadas").Value)
        .PendientesEmitirVotos = GetNumero(rs.Fields("Pendientes_Emitir_Voto").Value)
        '.PeriodoLegislativo = GetCadena(rs.Fields("Período_Legislativo").Value)
        '.Presentes = GetNumero(Rs.Fields("Presentes").Value)
        .Sesion = GetNumero(rs.Fields("Sesión").Value)
        .SolicitudGrabarManual = GetNumero(rs.Fields("Solicitud_Grabacion_Manual").Value)
        .TiempoParaVotacion = GetNumero(rs.Fields("Tiempo_de_votación").Value)
        .TipoDeOperacion = GetCadena(rs.Fields("Identificador_tipo_de_operacion").Value)
        .TipoMayoria = GetCadena(rs.Fields("Tipo_de_Mayoría").Value)
        .TipoMayoriaQuorum = GetCadena(rs.Fields("Tipo_Mayoria_Quorum").Value)
        .TituloDelActa = GetCadena(rs.Fields("Titulo_del_Acta").Value)
        .strError = GetCadena(rs.Fields("strError").Value)
        .EstadoSesion = GetCadena(rs.Fields("Estado_sesion").Value)
        .FechaVotacion = GetCadena(rs.Fields("FechaVotacion").Value)
        .HoraVotacion = GetCadena(rs.Fields("HoraVotacion").Value)
        If .TipoDeAbstencion = "" Then
                .TipoDeAbstencion = "votlar"
        End If
        If .BaseMayoria = "" Then
            .BaseMayoria = "votemi"
        End If
        If .TipoMayoria = "" Then
            .TipoMayoria = "120"
        End If
        If .TipoMayoriaQuorum = "" Then
            .TipoMayoriaQuorum = "120"
        End If
    End With
    Dim rsPer As ADODB.Recordset
    Set rsPer = New ADODB.Recordset
    SetearRsAux "SELECT TOP 1 Período_Legislativo From Sesion WHERE (LTRIM(Estado_sesión) = 'abierta') ORDER BY Fecha_de_inicio DESC", rsPer
    If rsPer.EOF Then
        Call MsgBox("Error de integridad de datos de sesiones", vbCritical)
        End
    End If
    EstadoActual.PeriodoLegislativo = Trim(rsPer.Fields(0))
    rsPer.Close
    Set rsPer = Nothing
    Call PublicarEstadoRecinto
End Sub

Private Sub Command1_Click()
        
    Dim i As Long
    
    'CartelActual.Presentes = 0
    'CartelActual.Ausentes = 0
    'EstadoActual.Presentes = 0
    'EstadoActual.Ausentes = 0
    'For i = 0 To xUltimaBanca
    '    If i Mod 2 = 1 Then
    '        EstadoActual.VectorPresencia(i) = AUSENTE
    '        EstadoActual.Ausentes = EstadoActual.Ausentes + 1
    '        CartelActual.Ausentes = CartelActual.Ausentes + 1
    '    Else
    '        EstadoActual.VectorPresencia(i) = PRESENTE
    '        EstadoActual.Presentes = EstadoActual.Presentes + 1
    '        CartelActual.Presentes = CartelActual.Presentes + 1
    '    End If
    '    Call PintarVectorColor(i)
    'Next i
    'Call ActualizarVector_enBD
    ''EstadoActual.TipoMayoriaQuorum = "120"
    'Call CalcularMinimoParaQuorum
End Sub
Private Sub SetearSesionActiva()
    Dim strSql As String
    ' Verifico que el periodo legislativo exista.
    strSql = "SELECT Período_Legislativo, Nro_de_Período_Legislativo, " _
    & "Tipo_de_período_sesión, Fecha_de_comienzo, Tipo_de_Sesión , Nro_de_Sesion_actual, Histórico, Ultima_Reunion " _
    & "FROM perparl WHERE Rtrim(Período_Legislativo) = '" & Trim(EstadoActual.PeriodoLegislativo) & "' ORDER BY Orden"
    Call SetearRs(strSql)
    ' Si no existe, se selecciona el ultimo disponible
    If rs.RecordCount <= 0 Or rs.EOF = True Or rs.BOF = True Then
        ' si no esta definida, selecciono la ultima disponible
        rs.Close
        DoEvents
        strSql = "SELECT * FROM perparl ORDER BY Orden DESC"
        Call SetearRs(strSql)
        rs.MoveFirst
    End If
    EstadoActual.Reunion = rs.Fields("Ultima_Reunion")
    EstadoActual.PeriodoLegislativo = rs.Fields("Período_Legislativo").Value
    'sesiones
    rs.Close
    DoEvents
    strSql = "SELECT * FROM Sesion WHERE Rtrim(Período_Legislativo) = '" & Trim(EstadoActual.PeriodoLegislativo) & "' AND Sesión = " & EstadoActual.Sesion & "AND Sesión <> 9999"
    Call SetearRs(strSql)
    If rs.RecordCount <= 0 Or rs.EOF = True Or rs.BOF = True Then
        strSql = "SELECT * FROM Sesion WHERE Rtrim(Período_Legislativo) = '" & Trim(EstadoActual.PeriodoLegislativo) & "' AND Sesión <> 9999 ORDER BY Sesión DESC"
        Call SetearRs(strSql)
        If rs.RecordCount <= 0 Or rs.EOF = True Or rs.BOF = True Then
            strSql = "SELECT Período_Legislativo, Sesión, Fecha_de_inicio, Próximo_Acta, Estado_sesión, Prorroga " _
                   & "FROM  Sesion WHERE 1=0"
            Call SetearRsW(strSql)
            RsWrite.AddNew
            RsWrite.Fields("Período_Legislativo").Value = EstadoActual.PeriodoLegislativo
            RsWrite.Fields("Sesión").Value = 1
            RsWrite.Fields("Fecha_de_inicio").Value = Date
            RsWrite.Fields("Próximo_Acta").Value = 1
            RsWrite.Fields("Estado_sesión").Value = "nueva"
            RsWrite.Fields("Prorroga").Value = "0"
            RsWrite.Update
            RsWrite.Close
        Else
            rs.MoveFirst
        End If
    End If
    EstadoActual.Sesion = IIf(IsNull(rs.Fields("Sesión").Value), 0, rs.Fields("Sesión").Value)
    EstadoActual.NroActa = rs.Fields("Próximo_Acta").Value
    EstadoActual.EstadoSesion = Trim(rs.Fields("Estado_sesión").Value)
    rs.Close
    ' En ambos casos: tomo el numero de "proximo acta" y lo pongo en estadoactual.nroacta
End Sub

Private Sub InicializarServer()
    Dim strSql As String
    'lblVersion.Caption = strVersion  ' Mostrar versión de sqv
    Call AbrirDB                     ' establece conexion con la base de datos
    Call BorrarMensajesTotales       ' Borrar mensajes relacionados con la consola y el servidor de bancas
    Call DeterminarValoresInicioServer
    EjecutarSQL "UPDATE vector SET Base_de_Mayoría = 'votemi', Tipo_de_Mayoría = '120'"
    
    'AP 080905 se cancela el levantador de bancas, pero tiene que habilitarse para produccion ATENCION IMPORTANTE
    If False Then Call Levanta_Banca 'deshabilitado, se hara por levantador de aplicaciones

    ' fin ap 080905
    ' ------------------------------------------------------------------
    ' Interfaz de usuario del servidor
    ' ------------------------------------------------------------------
    blServerPrendido = True
    Call ServerOnOff
    ' ------------------------------------------------------------------
    ' Setear estado inicial del recinto
    ' ------------------------------------------------------------------
    Call LeerEstadoRecinto
    EtiquetasCartel.strBase = DevolverLeyendaBase(EstadoActual.BaseMayoria)
    EtiquetasCartel.strTipo = DevolverLeyendaTipo(EstadoActual.TipoMayoria)
    Call cargarColores
'A02
    imagePath = App.Path & "\imagenes" & Trim(cFORMULARIO_VERSION) & "\"
    Call CargarImagenes
    Call CargarColoresFuente
    Call ConfigurarFrames
    Call ArmarBancasCartel
    Call CrearColeccionLegisladores
'A02 END
    With EstadoActual
        .Presentes = 0 '1 'inicializa con presidente
        .Ausentes = xUltimaBanca + 1 - .Presentes
        .strError = ""
        .IP_Consola = 0
        .TipoDeOperacion = "quorum"
        .ModoMantenimientoBancas = 0
        .ModoNormalMantSistema = 0
        .TituloDelActa = ""
        .CartelEncendido = 3
        FrameControl.ZOrder 0
        .ActaGrabada = 0
        .SolicitudGrabarManual = 0
    End With
    With CartelActual
        .Abstenciones = 0
        .Afirmativos = 0
        .Negativos = 0
        .Presentes = 1
        .Ausentes = xUltimaBanca + 1 - .Presentes
        .LeyendaQuorum = "NO HAY QUORUM"
    End With
    Screen.MousePointer = vbIbeam
    'CMBShowCursor False
    ShowCursor True
    Call CargarVectorIdentificacionHabilitados
    Call CalcularMinimoParaQuorum
    Call AltaLogGeneral("SQV SERVER", "Inicializando SQV Server " & Now)
    Call SetearSesionActiva
    Call ReinicioSistema
    'Se inicializa el tiempo para votación por default en 15
    EstadoActual.TiempoParaVotacion = 15
End Sub

Private Sub MensajeDisplayTerminal(xBanca As String, xMensaje As String)
    Dim MensajeBanca As MensajeSistema

    With MensajeBanca
        .sTipo = "mset"
        .sObjeto = xBanca
        .sComponente = "term.display"
        .sAtributo = "text"
        .sValor = xMensaje
        .sComentario = ""
    End With

    Call EnviarMensajesBancas(MensajeBanca)
End Sub

Private Sub ProbarCalculoResultado()

    Dim strSql    As String
    Dim strResult As String
    
    Dim xpMin_p_af_Calc As Long
    
    strSql = "SELECT base_mayoria, tipo_mayoria, miecue, presentes, Afirmativos , " _
           & "Negativos, Presidente, VotoPresidente, NuevoRes, NuevoMinAfirmativo From prueba_resultados_d WHERE clase='HCDN'"
    Call SetearRsW(strSql)
    
    With RsWrite
        While Not .EOF
            If .Fields("base_mayoria").Value = "miecue" And .Fields("tipo_mayoria").Value = 121 Then
                'Stop
            End If
            'If .Fields("Afirmativos").Value = 17 Then Stop
            
            'If .Fields("base_mayoria").Value = "legpre" And Str(.Fields("tipo_mayoria").Value) = "120" And _
               .Fields("Afirmativos").Value = 17 Then Stop
            If .Fields("presentes").Value < 10 Then Stop
            
            .Fields("NuevoRes").Value = CalculoResultado(.Fields("base_mayoria").Value, Str(.Fields("tipo_mayoria").Value), .Fields("miecue").Value, .Fields("presentes").Value, .Fields("Afirmativos").Value, .Fields("Negativos").Value, "w", 0, 0, xpMin_p_af_Calc, .Fields("VotoPresidente").Value, IIf(IsNull(.Fields("Presidente").Value), 0, .Fields("Presidente").Value))
            ' .Fields("NuevoRes").Value = CalculoResultado("miecue", "120", 70, 50, 20, 10, "", 0, 0, 0, "n", 0)
            ' MsgBox CalculoResultado("miecue", "120", 70, 50, 20, 10, " ", 0, 0, 0, "n", 0)
            .Fields("NuevoMinAfirmativo").Value = xpMin_p_af_Calc
            .Update
            .MoveNext
        Wend
    End With
    End
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

    On Error GoTo TrapError
    Dim strSql As String
    Dim pAuxMinParaAfirmativa As Long
    Dim xNumerador     As Long
    Dim xDenominador   As Long
    Dim xVotosEmitidos As Long
    Dim xResto As Long
    Dim xBase_para_Mayoria As Long

    If pTipo_Mayoria = "" Then
        pTipo_Mayoria = "120"
    End If
    If pBase_de_Mayoria = "" Then
        pBase_de_Mayoria = "votemi"
    End If

    pTipo_Mayoria = Trim(LCase(pTipo_Mayoria))

    strSql = "SELECT * From tipmay WHERE Tipo_de_Mayoria = '" & pTipo_Mayoria & "'"
    Call SetearOtroRs(strSql)

    xVotosEmitidos = pAfirmativos + pNegativos

    'legpre' : pPresentes (parametro: presentes en el momento del cierre de la votacion)
    'miecue' : pMiembros del cuerpo (parametro: total de miembros definidos)
    'votemi' : Votos Emitidos (calculados anteriormente)

    xNumerador = RsOtro.Fields("Numerador").Value ' NUMERADOR DE LA TABLA
    xDenominador = RsOtro.Fields("Denominador").Value  ' MODIFICAR POR DENOMINADOR DE LA TABLA
    
    Dim presTemp As Integer
    presTemp = getPresentes
    xBase_para_Mayoria = IIf(pBase_de_Mayoria = "legpre", presTemp, IIf(pBase_de_Mayoria = "miecue", pMiembros_del_cuerpo, IIf(pBase_de_Mayoria = "votemi", xVotosEmitidos, 0)))
    xResto = xBase_para_Mayoria * xNumerador Mod xDenominador
    pAuxMinParaAfirmativa = Fix(xBase_para_Mayoria * xNumerador / xDenominador)
    pMin_p_afirmativa_Calculo = IIf(xResto > 0, pAuxMinParaAfirmativa + 1, pAuxMinParaAfirmativa)
    pMin_p_afirmativa_Calculo = pMin_p_afirmativa_Calculo + IIf(LCase(RsOtro.Fields("Rdo_si_Af_igual_Min_y_Resto_mayor_0").Value) = "n", 1, 0)
    pMin_p_afirmativa_Calculo = pMin_p_afirmativa_Calculo + IIf(xResto = 0, LCase(RsOtro.Fields("SumarMinAfSiRestoIgual0").Value), 0)

    'AGREGADO DE NICO
    If pTipo_Mayoria = "121" Then
        'Si da 130 de minimo afirmativos y hay 130 afirmativos no entra a la condicion de AFIRMATIVO.
        'Se le resta 1 para que si entre, solo en el caso de que sea la mitad +1
        pMin_p_afirmativa_Calculo = pMin_p_afirmativa_Calculo - 1 'Offset para correccion
    End If
    'FIN DE AGREGADO DE NICO
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
                If RsOtro.Fields("Rdo_si_Af_igual_Min_y_Resto_mayor_0").Value = "A" Then
                    CalculoResultado = "AFIRMATIVO"
                Else
                    If RsOtro.Fields("Rdo_si_Af_igual_Min_y_Resto_mayor_0").Value = "E" Then ' Caso 1
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
                If RsOtro.Fields("Rdo_si_Af_igual_Min_y_Resto_igual_0").Value = "A" Then
                    CalculoResultado = "AFIRMATIVO"
                Else
                    If RsOtro.Fields("Rdo_si_Af_igual_Min_y_Resto_igual_0").Value = "E" Then
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
    
    ' Control caso vota el presidente
    If CalculoResultado = "EMPATE" Then
        If pVotaElPresidente > 0 And xHuboEmpate Then
            If (LCase(pVotoPresidente) = "s") Then
                CalculoResultado = "AFIRMATIVO"
                xHuboDesempate = True
            End If
            If (LCase(pVotoPresidente) = "n") Then
                CalculoResultado = "NEGATIVO"
                xHuboDesempate = True
            End If
        Else
            xVotoSenadorEmpate = pVotoPresidente
            xHuboDesempate = False
        End If
        xHuboEmpate = True
    Else
        xHuboEmpate = False
        xHuboDesempate = False
    End If
Exit Function
TrapError:
    Select Case err.Number
        Case Else
            Call AltaLogGeneral("SQV SERVER", "CalculoResultado" & "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source)
            Resume
    End Select
End Function

Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To 256
        EstadoActual.EnIdentificacion(i) = False
    Next i
    For i = 0 To 256
        VectorControlDoble(i) = 0
        VectorControlDobleTick(i) = 0
    Next i
    lblGeneralLeyendaQuorumDato.Width = lblGeneralLeyendaQuorumDato.Width - 500
    lblGeneralLeyendaQuorumDato.Left = lblGeneralLeyendaQuorumDato.Left + 800
    lblGeneralLeyendaQuorumDato.Alignment = vbLeftJustify
    lblGeneralSesionDato(3).top = lblGeneralAusentesDato.top - 1400
    lblGeneralSesionDato(4).top = lblGeneralSesionDato(3).top + 770 'sesion de tablas
    lblGeneralSesionDato(5).top = lblGeneralAusentesDato.top + 140
    lblCAusentes.top = lblGeneralAusentesDato.top
    lblCPresentes.top = lblGeneralSesionDato(3).top - 50
    lblCAusentes.Left = 7900
    lblCPresentes.Left = 7900
    lblGeneralAusentesDato.top = lblCAusentes.top - 100
    lblGeneralAusentesDato.Left = lblGeneralAusentesDato.Left - 300
    lblGeneralPresentesDato.Left = lblGeneralAusentesDato.Left
    lblGeneralPresentesDato.top = lblGeneralPresentesDato.top
    lblGeneralInformacion.Left = lblGeneralSesionDato(4).Left + 50
    lblGeneralInformacion.Height = lblGeneralInformacion.Height + 2000
    lblGeneralInformacion.top = lblGeneralInformacion.top - 750
    lblGeneralInformacion.top = lblGeneralInformacion.top + 100
    lblGeneralFechaDato.Left = lblGeneralFechaDato.Left + 200
    lblGeneralHoraDato.Left = shpRecuadroFecha.Left + 4200
    lblGeneralFechaDato.top = lblGeneralFechaDato.top + 200
    lblGeneralHoraDato.top = lblGeneralFechaDato.top
    shpRecuadroFecha.Width = shpRecuadroFecha.Width + 4500
    shpRecuadroFecha.Left = lblGeneralFechaDato.Left - 150
    lblGeneralFechaDato.Left = lblGeneralFechaDato.Left - 100
    shpRecuadroFecha.top = lblGeneralFechaDato.top + 120
    shpRecuadroFecha.BorderWidth = 1
    shpRecuadroFecha.Width = shpRecuadroFecha.Width + 40
    shpRecuadroFecha.Height = shpRecuadroFecha.Height + 400
    shpHora.Left = lblGeneralHoraDato.Left
    shpHora.Height = shpRecuadroFecha.Height
    shpHora.top = shpRecuadroFecha.top
    shpHora.Width = 2100
    lblGeneralHoraDato.Left = lblGeneralHoraDato.Left + 80
    lblGeneralFechaDato.top = lblGeneralFechaDato.top + 50
    lblGeneralHoraDato.top = lblGeneralFechaDato.top
    shpRecuadroQuorum.Width = lblGeneralLeyendaQuorumDato.Width - 800
    shpRecuadroQuorum.Left = lblGeneralLeyendaQuorumDato.Left + 900
    shpRecuadroQuorum.BorderWidth = 1
    shpRecuadroQuorum.Height = shpRecuadroFecha.Height + 200
    shpRecuadroQuorum.top = shpRecuadroFecha.top - 120
    shpRecuadroQuorum.BorderColor = MiRojo
    lblVersionCartel.Visible = False
    lblGeneralLeyendaQuorumDato.top = lblGeneralFechaDato.top
    Set lblGeneralTituloDato(4).Container = lblGeneralInformacion.Container
    lblGeneralTituloDato(4).ZOrder 0
    lblGeneralInformacion.top = lblGeneralInformacion.top + 350
    'lblGeneralInformacion.ZOrder 0
    EstabaEnIdentificacion = False
    PrimerRecuento = True
    If App.PrevInstance = True Then ' Si se esta ejecutando una instancia previa del server, se apaga!
        End
    End If
    Call SetVersion
    
    Set RsLocal = New ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rstActa = New ADODB.Recordset
    Set RsOtro = New ADODB.Recordset
    
    ' Indicar si se esta utilizando la base de pruebas o la de produccion
    If blBanderaPruebas = True Then
        Frame1.Caption = Frame1.Caption & " BASE PRUEBAS"
        Frame1.Caption = Frame2.Caption & " BASE PRUEBAS"
        Frame1.Caption = Frame3.Caption & " BASE PRUEBAS"
'        lblPruebas.Visible = False
    Else
        'lblPruebas.Visible = False
    End If
    
    
    xFechaArranque = Now
    xFechaUltimoReset = Now
    xFechaInicioProceso = Now
    blBanderaTimer = True
    xIntervalo = 2 ' en segundos
    txtVecesPorSegundo = xIntervalo
    xUltimoMensajeSB = 0 ' ultimo mensaje leido del servidor de banca
    xUltimoMensajeCosola = 0 ' ultimo mensaje leido de la consola
    blMostrarEstadoCartel = True
    blMostrarEstadoRecinto = True
    flSwitchExitoso = False
    xBancaPruebaScan = 0
    xCiclosTotales = 0
    cmdConfig.Enabled = False
    xBancaDuplicada = -1
    Beep
    Call InicializarServer
'    Servidor.Show
    lblAppMayor.Caption = App.Major
    lblAppRevision.Caption = App.Revision
    lblAppMinor.Caption = App.Minor
    lblVersionSQV.Caption = "091012:"
    'Call ProbarCalculoResultado
    
'A02
        Picture3.Left = 0
        Picture3.top = 0
'A02 END

    EstadoActual.CartelEncendido = 2 ' general poner 1 para control
    FrameSQVGeneral.ZOrder 0
    Set picB(4).Container = Me
    picB(4).Height = picB(4).Height + 1500
    picB(4).top = picB(4).top - 500
    lblGeneralInformacion.top = lblGeneralInformacion.top + 1400
    lblGeneralTituloDato(4).top = lblGeneralTituloDato(4).top
    lblGeneralTituloDato(4).ZOrder 0
    lblGeneralTituloDato(4).Left = lblGeneralInformacion.Left + 50
    Set shpTitulo.Container = lblGeneralInformacion.Container
    Set shpTitulo2.Container = shpTitulo.Container
    shpTitulo2.ZOrder 0
    lblGeneralTituloDato(4).Height = lblGeneralTituloDato(4).Height + 500
    Set lblTituloBaseYTipoDeMayoria.Container = lblGeneralInformacion.Container
    Set lblGeneralMayoriaDato(2).Container = lblGeneralInformacion.Container
    lblTituloBaseYTipoDeMayoria.ZOrder 0
    lblGeneralMayoriaDato(2).ZOrder 0
    lblTituloBaseYTipoDeMayoria.Left = shpRecuadroQuorum.Left - 200 '+ 500
    lblGeneralMayoriaDato(2).Left = shpRecuadroQuorum.Left - 600
    lblTituloBaseYTipoDeMayoria.Width = lblTituloBaseYTipoDeMayoria.Width + 1000
    lblTituloBaseYTipoDeMayoria.top = shpTitulo.top + 200
    lblGeneralMayoriaDato(2).top = lblTituloBaseYTipoDeMayoria.top + 600
    'Set lblLeyendaVoto.Container = lblGeneralNegativosDato.Container
    'lblLeyendaVoto.ZOrder 0
    'lblLeyendaVoto.Left = shpTitulo.Left
    'lblLeyendaVoto.Top = lblGeneralAfirmativosDato.Top
    lblLeyendaVotoAfirmativo.Caption = "AFIRMATIVOS"
    lblLeyendaVotoAfirmativo.Alignment = vbLeftJustify
    lblLeyendaVotoAfirmativo.top = lblGeneralAfirmativosDato.top + 10
    Set lblLeyendaVotoAfirmativo.Container = lblGeneralAfirmativosDato.Container
    lblLeyendaVotoAfirmativo.ZOrder 0
    lblLeyendaVotoNegativo.Caption = "NEGATIVOS"
    lblLeyendaVotoNegativo.Alignment = vbLeftJustify
    lblLeyendaVotoNegativo.top = lblGeneralNegativosDato.top + 20
    Set lblLeyendaVotoNegativo.Container = lblGeneralNegativosDato.Container
    lblLeyendaVotoNegativo.ZOrder 0
    lblLeyendaVotoAbstencion.Caption = "ABSTENCIONES"
    lblLeyendaVotoAbstencion.Alignment = vbLeftJustify
    lblLeyendaVotoAbstencion.top = lblGeneralAbstencionesDato.top + 20
    Set lblLeyendaVotoAbstencion.Container = lblGeneralAbstencionesDato.Container
    lblLeyendaVotoAbstencion.ZOrder 0
    lblGeneralAfirmativosDato.top = lblGeneralAfirmativosDato.top - 120
    lblGeneralNegativosDato.top = lblGeneralNegativosDato.top - 100
    lblGeneralAbstencionesDato.top = lblGeneralAbstencionesDato.top - 100
    Set lblGeneralResultadoDato.Container = lblTituloOcupadosNoIdentificados(1).Container
    lblGeneralResultadoDato.ZOrder 0
    lblGeneralResultadoDato.top = lblGeneralResultadoDato.top + 2300
    lblGeneralResultadoDato.Left = lblTituloBaseYTipoDeMayoria.Left - 300
    lblGeneralResultadoDato.Alignment = vbLeftJustify
    lblTituloOcupadosNoIdentificados(1).Width = lblTituloOcupadosNoIdentificados(1).Width + 500
    lblTituloOcupadosNoIdentificados(1).Left = lblTituloBaseYTipoDeMayoria.Left
    lblTituloOcupadosNoIdentificados(1).top = lblTituloOcupadosNoIdentificados(1).top + 600
    lblOcupadosNoIdentificados(2).top = lblOcupadosNoIdentificados(2).top - 180
    shpTitulo.Visible = False
    shpRecuadroQuorum.Visible = False
    lblOrador04.top = 4200
    lblOrador03.top = lblOrador04.top - 650
    lblOrador02.top = lblOrador03.top - 650
    lblOrador01.top = lblOrador02.top - 650
    Set shpRecuadroOrador.Container = lblOrador01.Container
    shpRecuadroOrador.Width = 14800
    shpRecuadroOrador.Height = 2850
    shpRecuadroOrador.Left = shpRecuadroFecha.Left + OffsetOrador
    lblOrador04.Left = lblOrador04.Left + OffsetOrador ' + 200
    lblOrador03.Left = lblOrador03.Left + OffsetOrador ' + 200
    lblOrador02.Left = lblOrador02.Left + OffsetOrador ' + 200
    lblOrador01.Left = lblOrador01.Left + OffsetOrador ' + 200
    shpRecuadroOrador.top = 2200
    shpRecuadroOrador.ZOrder 0
    shpRecuadroOrador.Visible = False
    lblOrador04.Height = lblOrador04.Height - 50 'Para que no tape la linea y evitar asi el titileo
    shpHora.Visible = False
    lblGeneralHoraDato.Width = lblGeneralHoraDato.Width - 500
    shpTitulo2.BorderWidth = 1
    lblGeneralInformacion.Left = 1500
    lblGeneralResultadoDato.Alignment = vbCenter
    lblGeneralTiempoDato.Height = 3305
    lblGeneralTiempoDato.top = lblGeneralTiempoDato.top - 2000
        '/***SE PONE POR DEFAULT IMPRESION AUTOMATICA EN 1***
    'Call EjecutarSQL("UPDATE vector SET Listar_automaticamente = 1")
    EstadoActual.ListarAutomaticamente = 1
    '/***FIN Impresion Automatica***
    EstadoActual.TipoMayoriaQuorum = "120"
    EstadoActual.TipoDeOperacion = "quorum"
    EstadoActual.Expresiones_Minoria = False
    PrimeraVezCeros = True
    EstadoActual.BancaEnPrueba = -1
    For i = 0 To 256
        BancasDeshabilitadas(i) = False 'Todas las bancas arrancan habilitadas
    Next i
'****************FIX PARA LOS CONTROLES DE MANTENIMIENTO****************
    Dim cTemp1 As Integer
    Dim cTemp2 As Integer
    Dim cTempSTR As String
    cTemp1 = Label1.top
    Label1.top = Label43.top
    Label43.top = cTemp1 - 400
    Line1.X2 = Line1.X2 + 4000
    cTemp1 = lblMantenimientostrPanel1.top
    lblMantenimientostrPanel1.top = lblMantenimientostrPendientes.top
    lblMantenimientostrPendientes.top = cTemp1 - 300
    lblMantenimientostrMantListaPendientes.top = lblMantenimientostrPendientes.top + 400
    lblMantenimientostrPanel2.top = lblMantenimientostrPanel1.top + 300
    lblMantenimientostrPanel3.top = lblMantenimientostrPanel2.top + 300
    lblMantenimientostrMantListaPendientes.Width = 6000
    lblMantenimientostrMantListaPendientes.Height = 3200
    Label44.Left = Label1.Left + 8000
    Label44.top = Label43.top
    lblMantenimientostrFallas.top = lblMantenimientostrPendientes.top
    lblMantenimientostrFallas.Left = lblMantenimientostrPendientes.Left + 8000
    lblMantenimientostrMantListaFallas.top = lblMantenimientostrMantListaPendientes.top
    lblMantenimientostrMantListaFallas.Left = lblMantenimientostrMantListaPendientes.Left + 8000
'****************COLORES DE MANTENIMIENTO****************
    FrameMantenimiento.BackColor = vbBlack
    Line1.BorderColor = vbWhite
    Label1.ForeColor = vbWhite
    Label1.BackColor = vbBlack
    lblMantenimientostrPanel1.ForeColor = vbWhite
    lblMantenimientostrPanel1.BackColor = vbBlack
    lblMantenimientostrPanel2.ForeColor = vbWhite
    lblMantenimientostrPanel2.BackColor = vbBlack
    lblMantenimientostrPanel3.ForeColor = vbWhite
    lblMantenimientostrPanel3.BackColor = vbBlack
    Line3.BorderColor = vbWhite
    Label43.ForeColor = vbWhite
    Label43.BackColor = vbBlack
    lblMantenimientostrPendientes.ForeColor = vbWhite
    lblMantenimientostrPendientes.BackColor = vbBlack
    lblMantenimientostrMantListaPendientes.ForeColor = vbWhite
    lblMantenimientostrMantListaPendientes.BackColor = vbBlack
    Label44.ForeColor = vbWhite
    Label44.BackColor = vbBlack
    lblMantenimientostrFallas.ForeColor = vbWhite
    lblMantenimientostrFallas.BackColor = vbBlack
    lblMantenimientostrMantListaFallas.ForeColor = vbWhite
    lblMantenimientostrMantListaFallas.BackColor = vbBlack
    Set lblOperador1.Container = lblMantenimientostrMantListaFallas.Container
    Set lblOperador2.Container = lblMantenimientostrMantListaFallas.Container
    Set lblOperador3.Container = lblMantenimientostrMantListaFallas.Container
    Set lblOperador4.Container = lblMantenimientostrMantListaFallas.Container
    lblOperador1.top = lblMantenimientostrPanel1.top
    lblOperador1.Left = lblMantenimientostrPanel1.Left + 1000
    lblOperador2.Visible = True
    lblOperador3.Visible = True
    lblOperador4.Visible = True
    lblMantenimientostrPanel1.Visible = False
    lblMantenimientostrPanel2.Visible = False
    lblMantenimientostrPanel3.Visible = False
    lblOperador1.ForeColor = vbWhite
    lblOperador1.BackColor = vbBlack
    lblOperador2.ForeColor = vbWhite
    lblOperador2.BackColor = vbBlack
    lblOperador3.ForeColor = vbWhite
    lblOperador3.BackColor = vbBlack
    lblOperador4.ForeColor = vbWhite
    lblOperador4.BackColor = vbBlack
    lblOperador1.FontSize = 32
    lblOperador2.FontSize = 32
    lblOperador3.FontSize = 32
    lblOperador4.FontSize = 32
    lblOperador1.Width = 6000
    lblOperador2.Width = 6000
    lblOperador3.Width = 6000
    lblOperador4.Width = 6000
    lblOperador1.Height = 1000
    lblOperador2.Height = 1000
    lblOperador3.Height = 1000
    lblOperador4.Height = 1000
    lblOperador2.Left = Label44.Left
    lblOperador2.top = lblOperador1.top
    lblOperador3.Left = lblOperador1.Left
    lblOperador3.top = lblOperador1.top + 4500
    lblOperador4.Left = lblOperador2.Left
    lblOperador4.top = lblOperador3.top
    Line3.Visible = False
    Label1.Visible = False
    CuentaSQL = 0
    Tick_InicioPasLis = 0
    For i = 0 To 256
        VectorDesconectadas(i) = False
    Next i
    For i = 0 To 256
        CantidadEidrxh(i) = 0
    Next i
    For i = 0 To 256
        TiempoEidrxh(i) = 0
    Next i
End Sub
Private Sub ActualizarTiempoCartel()
Dim xTiempoVotacionTranscurrido As Long
Dim xTiempoRestanteVotacion As Long
If IsNumeric(lblGeneralTiempoDato.Caption) Then
    lblGeneralTituloTiempo.Visible = True
End If
If EstadoActual.EstadoVotacion_y_PasList = "votando" Then
    xTiempoVotacionTranscurrido = DateDiff("s", EstadoActual.FechaVotacion, Now)
    xTiempoRestanteVotacion = EstadoActual.TiempoParaVotacion - xTiempoVotacionTranscurrido
    lblGeneralTituloTiempo.Visible = True
    CartelActual.LeyendaTiempo = _
        IIf(EstadoActual.EstadoVotacion_y_PasList = "espera", "", _
            IIf(EstadoActual.EstadoVotacion_y_PasList = "votando", _
                IIf(xTiempoRestanteVotacion > EstadoActual.TiempoParaVotacion, "", _
                    IIf(xTiempoRestanteVotacion > 59, Str(xTiempoRestanteVotacion), _
                        IIf(xTiempoVotacionTranscurrido > EstadoActual.TiempoParaVotacion, " 0", Right(Str(xTiempoRestanteVotacion), 2)))), _
            IIf(EstadoActual.EstadoVotacion_y_PasList = "larga", " 0", _
            IIf(EstadoActual.EstadoVotacion_y_PasList = "cancelada", "VOTACION CANCELADA", " 0"))))
    If lblGeneralTiempoDato.Caption <> CartelActual.LeyendaTiempo Then
        lblGeneralTiempoDato.Caption = CartelActual.LeyendaTiempo
    End If
    DoEvents
End If
End Sub
Private Sub MostrarCartel()
    Dim strMensajeLog As String
    Dim strInfoMant(cUltimoPanelMant) As String
    Dim i As Long
    
    ' Mostrar datos de estado de cartel
    With CartelActual
    
        If EstadoActual.CartelEncendido = 1 Then  'Frame Control
                lblcrt_Presentes.Caption = .Presentes
                lblcrt_Ausentes.Caption = .Ausentes
                lblcrt_Resultado.Caption = .Resultado
                lblcrt_Afirmativos.Caption = .Afirmativos
                lblcrt_Negativos.Caption = .Negativos
                lblcrt_Abstenciones.Caption = .Abstenciones
                lblcrt_MinimoParaAfirmativo.Caption = .MinimoVotosParaAfirmativo
                lblcrt_LeyendaQuorum.Caption = .LeyendaQuorum
                lblcrt_LeyendaTiempo.Caption = .LeyendaTiempo
                lblPendientesEmitirVotos = EstadoActual.PendientesEmitirVotos
                lblAbsAut.Caption = EstadoActual.AbstencionistasAutorizados
                lblOcupadosNoIdentificados(0) = EstadoActual.OcupadosNoIdentificados
        End If
        If EstadoActual.CartelEncendido = 2 Then  'Frame General separar apagado Separar luego el 0 acaa acaa accaca
                'control de visualizacion, solo si hubo cambios de tipo o estado de operacion
            If xControlCartelTipoOperacion <> EstadoActual.TipoDeOperacion Or xControlCartelEstadoOperacion <> EstadoActual.EstadoVotacion_y_PasList Then
                xControlCartelTipoOperacion = EstadoActual.TipoDeOperacion
                xControlCartelEstadoOperacion = EstadoActual.EstadoVotacion_y_PasList

'                lblGeneralTituloDato.Visible = (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom")
'                lblGeneralOrdenDiaDato.Visible = False
'                lblGeneralTipoOperacionDato.Visible = (EstadoActual.TipoDeOperacion <> "quorum")
'                lblGeneralTiempo.Visible = (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom")
'                lblGeneralTiempoDato.Visible = (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom")
'                lblGeneralNegativos.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
'                lblGeneralNegativosDato.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
'                lblGeneralAfirmativos.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
'                lblGeneralAfirmativosDato.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
'                lblGeneralAbstenciones.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
'                lblGeneralAbstencionesDato.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
'                lblGeneralResultadoDato.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
'                lblGeneralMayoriaDato.Visible = (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom")
'            End If
                    
'                lblGeneralPresentesDato.Caption = Min(.Presentes, xMiembrosDelCuerpo)
'                lblGeneralAusentesDato.Caption = max(.Ausentes, IIf(xPresidenteLegislador, 1, 0))
'                lblGeneralHoraDato.Caption = Format(Now, "HH:MM")
'                lblGeneralFechaDato.Caption = Format(Now, FORMATOFECHA )
'                lblGeneralLeyendaQuorumDato.ForeColor = IIf(.LeyendaQuorum = "QUORUM", &HFFFFFF, &HC0C0FF)
'                lblGeneralLeyendaQuorumDato.Caption = .LeyendaQuorum
                
'                lblGeneralSesionDato.Caption = LeyendaSesion()
'                lblGeneralTituloDato.Caption = EstadoActual.TituloDelActa
'                lblGeneralOrdenDiaDato.Caption = ""
'                lblGeneralTipoOperacionDato.Caption = LeyendaTipoOperacion
'                lblGeneralMayoriaDato.Caption = "Base y tipo de mayoria: " & LeyendaTipoMayoria & " de los " & LeyendaBaseMayoria
                
'                If (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") Then
'                    lblGeneralTiempoDato.Caption = .LeyendaTiempo
'                    lblGeneralNegativosDato.Caption = .Negativos
'                    lblGeneralAfirmativosDato.Caption = .Afirmativos
'                    lblGeneralAbstencionesDato.Caption = .Abstenciones
'                    lblGeneralResultadoDato.Caption = .Resultado
'                End If
                

'A02
                Select Case EstadoActual.TipoDeOperacion
                    Case "quorum"
                        'Cuando entra en modo Quorum da 15 segundos al operador para que cambie
                        'El titulo del acta en tratamiento antes de visualizarlo
                        MostrarPIC "A", 0 'Muestro Standard
                        MostrarPIC "B", 0 'Muestro Datos de Sesion alternando con Asunto en Tratamiento
                        MostrarPIC "C", 1 '091011 Muestra sin resultados  'Muestro Mapa de Bancas
                    Case "paslis"
                        Debug.Print EstadoActual.EstadoVotacion_y_PasList
                        Select Case EstadoActual.EstadoVotacion_y_PasList
                            Case "espera"
                                'lblGeneralInformacion.Caption = "PASE DE LISTA" 'AQUI puede indicarse una leyenda para el comienzo de la operacion
                                lblGeneralInformacion.Caption = ""
                                MostrarPIC "A", 0 'Muestro Standard
                                MostrarPIC "B", 4 'Muestro Informacion
                                MostrarPIC "C", 1 '091011 Muestra sin resultados  'Muestro Mapa de Bancas
                                PanelResultadosInicializar
                            Case "finalizada"
                                lblGeneralInformacion.Caption = "PASE DE LISTA FINALIZADO" 'AQUI puede indicarse una leyenda para el comienzo de la operacion
                                MostrarPIC "A", 0 'Muestro Standard
                                MostrarPIC "B", 4 'Muestro Informacion
                                MostrarPIC "C", 1 'Muestro Panel de Resultados de Pase de Lista
                        End Select
                    Case "votnum"
                        Select Case EstadoActual.EstadoVotacion_y_PasList
                            Case "espera"
                                lblGeneralInformacion.Caption = "VOTACION" & vbCrLf & "NUMERICA"
                                MostrarPIC "A", 0 'Muestro Standard
                                MostrarPIC "B", 4 'Muestro Informacion - ***Ver si aqui no es conveniente mostrar el procedimiento de voto
                                MostrarPIC "C", 1 '091011 Muestra sin resultados  'Muestro mapa de bancas
                                lblLeyendaVotoAfirmativo.Left = shpRecuadroOrador.Left
                                lblLeyendaVotoNegativo.Left = shpRecuadroOrador.Left
                                lblLeyendaVotoAbstencion.Left = shpRecuadroOrador.Left
                                lblGeneralAfirmativosDato.Left = lblLeyendaVotoAbstencion.Left + lblLeyendaVotoAbstencion.Width - 2100
                                lblGeneralNegativosDato.Left = lblGeneralAfirmativosDato.Left
                                lblGeneralAbstencionesDato.Left = lblGeneralAfirmativosDato.Left
                            Case "votando", "larga"
                                MostrarPIC "A", 0 'Muestro Standard
                                MostrarPIC "B", 4 'tipo de operacion
                                MostrarPIC "C", 1 '091011 Muestra sin resultados  'Muestro mapa de Bancas
                            Case "finalizada", "empate" 'en empate muestra el resultado
                                If EstadoActual.TituloDelActa = "" Then
                                    MostrarPIC "A", 0 'Muestro Standard
                                Else
                                    MostrarPIC "A", 0 'Standard... o 1 Muestro Asunto
                                End If
                                MostrarPIC "B", 4 'Muestro Resultado de la Votación
                                MostrarPIC "C", 0 '091011 Muestra con resultados  'Muestro mapa de Bancas
                        End Select
                    Case "votnom"
                        Select Case EstadoActual.EstadoVotacion_y_PasList
                            Case "espera"
                                lblGeneralInformacion.Caption = "VOTACION" & vbCrLf & "  NOMINAL"
                                shpTitulo.Left = Me.shpRecuadroFecha + 350
                                shpTitulo.top = lblGeneralInformacion.top - 50
                                shpTitulo.BorderColor = MiRojo
                                shpTitulo.Height = lblGeneralInformacion.Height - 1400
                                shpTitulo.Width = lblGeneralInformacion.Width - 7500
                                shpTitulo.ZOrder 0
                                'Para titulo
                                shpTitulo2.Height = shpTitulo.Height / 2 + 500
                                shpTitulo2.Left = shpTitulo.Left
                                shpTitulo2.BorderColor = MiRojo
                                shpTitulo2.top = shpTitulo.top - 1500
                                shpTitulo2.Width = shpTitulo.Width * 2 + 700
                                lblLeyendaVotoAfirmativo.Left = shpTitulo.Left
                                lblLeyendaVotoNegativo.Left = shpTitulo.Left
                                lblLeyendaVotoAbstencion.Left = shpTitulo.Left
                                lblGeneralAfirmativosDato.Left = lblLeyendaVotoAbstencion.Left + lblLeyendaVotoAbstencion.Width - 2100
                                lblGeneralNegativosDato.Left = lblGeneralAfirmativosDato.Left
                                lblGeneralAbstencionesDato.Left = lblGeneralAfirmativosDato.Left
                                MostrarPIC "A", 0 'Muestro Standard
                                MostrarPIC "B", 4 'Muestro Informacion - ***Ver si aqui no es conveniente mostrar el procedimiento de voto
                                MostrarPIC "C", 1 '091011 Muestra sin resultados  'Muestro mapa de bancas
                            Case "votando", "larga"
                                MostrarPIC "A", 0 'Muestro Standard
                                MostrarPIC "B", 4 'tipo de operacion
                                MostrarPIC "C", 1 '091011 Muestra sin resultados  'Muestro mapa de Bancas
                                PanelResultadosInicializar
                            Case "finalizada", "empate" 'en empate muestra el resultado
                                MostrarPIC "A", 0 'Muestro Standard
                                MostrarPIC "B", 4 'Muestro Resultado
                                MostrarPIC "C", 0 'Muestro Panel de Resultados de Votación
                        End Select
                End Select
            End If
'A02 END

            lblGeneralPresentesDato.Caption = max(Min(CartelActual.Presentes, xMiembrosDelCuerpo - 1), 1)
            lblGeneralAusentesDato.Caption = Min(max(CartelActual.Ausentes, 1), xMiembrosDelCuerpo - 1) 'max(.Ausentes, IIf(xPresidenteLegislador, 1, 0))
            
            
            lblGeneralHoraDato.Caption = Format(Now, "HH:MM")
            lblGeneralFechaDato.Caption = Format(Now, FORMATOFECHA)
            
            'lblGeneralLeyendaQuorumDato.ForeColor = IIf(.LeyendaQuorum = "QUORUM", &HFFFFFF, &HC0C0FF)
            If EstadoActual.Expresiones_Minoria = True Then
                'frmTapaQuorum.Show
                frmTapaQuorum.Left = 8000
                frmTapaQuorum.top = 200
                frmTapaQuorum.Width = 7200
                'frmTapaSesion.Show
                frmTapaSesion.Left = 300
                frmTapaSesion.top = 2400
                'frmExpresionesMinoria.Show
                frmExpresionesMinoria.Left = 0
                frmExpresionesMinoria.top = 5000
            Else
                If frmTapaQuorum.Visible = True Then
                    Unload frmTapaQuorum
                End If
                If frmTapaSesion.Visible = True Then
                    Unload frmTapaSesion
                End If
                If frmExpresionesMinoria.Visible = True Then
                    Unload frmExpresionesMinoria
                End If
            End If
            lblGeneralLeyendaQuorumDato.FontSize = IIf(.LeyendaQuorum = "QUORUM", 40, 40)
            If .LeyendaQuorum = "QUORUM" And Trim(lblGeneralLeyendaQuorumDato.Caption) <> "QUORUM" Then
                lblGeneralLeyendaQuorumDato.ForeColor = &HFFFF&
                lblGeneralLeyendaQuorumDato.Caption = "         QUORUM"
            ElseIf .LeyendaQuorum <> "QUORUM" And Trim(lblGeneralLeyendaQuorumDato.Caption) <> "NO HAY QUORUM" Then
                lblGeneralLeyendaQuorumDato.ForeColor = MiRojo
                lblGeneralLeyendaQuorumDato.Caption = " NO HAY QUORUM"
            End If
            'lblGeneralLeyendaQuorumDato.Caption = IIf(.LeyendaQuorum = "QUORUM", "QUORUM     ", "NO HAY QUORUM")
            lblGeneralSesionDato(0).Caption = "" 'LeyendaSesion()
            lblGeneralSesionDato(1).Caption = "" 'LeyendaSesion()
            lblGeneralSesionDato(2).Caption = "" 'LeyendaSesion()
            lblGeneralSesionDato(3).Caption = LeyendaSesion(1)
            lblGeneralSesionDato(4).Caption = LeyendaSesion(2)
            lblGeneralSesionDato(5).Caption = LeyendaSesion(3)
            lblGeneralSesionDato(0).Visible = False
            lblGeneralSesionDato(1).Visible = False
            lblGeneralSesionDato(2).Visible = False
            lblGeneralSesionDato(3).Visible = True
            lblGeneralSesionDato(4).Visible = True
            lblGeneralSesionDato(5).Visible = True
            lblGeneralTituloDato(0).Caption = EstadoActual.TituloDelActa
            lblGeneralTituloDato(1).Caption = EstadoActual.TituloDelActa
            lblGeneralTituloDato(2).Caption = EstadoActual.TituloDelActa
            lblGeneralTituloDato(3).Caption = EstadoActual.TituloDelActa
            If Trim(EstadoActual.TituloDelActa) = "" Or EstadoActual.TipoDeOperacion = "paslis" Then
                shpTitulo2.Visible = False
                lblGeneralTituloDato(4).Caption = ""
            Else
                shpTitulo2.Visible = True
                lblGeneralTituloDato(4).Caption = EstadoActual.TituloDelActa
            End If
            If EstadoActual.TipoDeOperacion = "paslis" Then
                lblTituloBaseYTipoDeMayoria.Visible = False
                If lblGeneralInformacion.Left <> 1000 Then
                    lblGeneralInformacion.Left = 1000
                End If
            Else
                lblTituloBaseYTipoDeMayoria.Visible = True
                If lblGeneralInformacion.Left <> 1500 Then
                    lblGeneralInformacion.Left = 1500
                End If
            End If
            lblGeneralTipoOperacionDato.Caption = LeyendaTipoOperacion 'No es el seteo oficial
            lblGeneralMayoriaDato(0).Caption = LeyendaTipoMayoria & " de los " & LeyendaBaseMayoria
            lblGeneralMayoriaDato(1).Caption = LeyendaTipoMayoria & " de los " & LeyendaBaseMayoria
            lblGeneralMayoriaDato(2).Caption = LeyendaTipoMayoria & " de los " & LeyendaBaseMayoria
            lblGeneralMayoriaDato(3).Caption = LeyendaTipoMayoria & " de los " & LeyendaBaseMayoria
            lblGeneralMayoriaDato(2).Visible = EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom"
            lblGeneralMayoriaDato(3).Visible = EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom"
            lblGeneralTituloDato(3).Visible = False
            If EstadoActual.EstadoVotacion_y_PasList = "espera" And Trim(EstadoActual.Orador) > "" Then
                lblOrador01.Caption = "Diputad" & IIf(EstadoActual.OradorSexo = "F", "a", "o") & " en uso de la palabra:"
                lblOrador02.Caption = EstadoActual.OradorNombre
                lblOrador03.Caption = EstadoActual.OradorAgrupacionPolitica
                lblOrador04.Caption = EstadoActual.OradorDistrito
                shpRecuadroOrador.Visible = False
                lblOrador01.Visible = True
                lblOrador02.Visible = True
                lblOrador03.Visible = True
                lblOrador04.Visible = True
                'lblGeneralTituloDato(4).Visible = False
                'shpTitulo2.Visible = False
                'lblTituloBaseYTipoDeMayoria.Visible = False
            Else
                lblOrador01.Caption = ""
                lblOrador02.Caption = ""
                lblOrador03.Caption = ""
                lblOrador04.Caption = ""
                lblOrador01.Visible = False
                lblOrador02.Visible = False
                lblOrador03.Visible = False
                lblOrador04.Visible = False
                shpRecuadroOrador.Visible = False
                lblGeneralTituloDato(4).Visible = True
                'lblTituloBaseYTipoDeMayoria.Visible = EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom"
            End If
            
            If (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") Then
                lblGeneralTiempoDato.Caption = .LeyendaTiempo
                If IsNumeric(lblGeneralTiempoDato.Caption) Then
                    lblGeneralTituloTiempo.Visible = True
                End If
                If EstadoActual.EstadoVotacion_y_PasList = "cancelada" Then
                    lblGeneralTituloTiempo.Visible = False
                Else
                    If .LeyendaTiempo <> "" Then
                        lblGeneralTituloTiempo.Visible = True
                    Else
                        lblGeneralTituloTiempo.Visible = False
                    End If
                    'lblGeneralTituloTiempo.Visible = (.LeyendaTiempo <> "")
                End If
                lblGeneralNegativosDato.Caption = .Negativos
                lblGeneralAfirmativosDato.Caption = .Afirmativos
                lblGeneralAbstencionesDato.Caption = .Abstenciones
                lblGeneralResultadoDato.Caption = .Resultado
            Else
                lblGeneralTiempoDato.Caption = ""
                lblGeneralTituloTiempo.Visible = False
            End If

            If (EstadoActual.TipoDeOperacion = "paslis" Or EstadoActual.TipoDeOperacion = "votnom") Then
                lblTituloPresentesIdentificados = "Presentes identificados:"
                lblTituloOcupadosNoIdentificados(0) = "Pendientes de identificarse:"
                lblTituloOcupadosNoIdentificados(1) = "Legis. sin identif:"
                lblOcupadosNoIdentificados(0) = GetNoIdentificadosSobrePresentes 'EstadoActual.OcupadosNoIdentificados
                If EstadoActual.EstadoVotacion_y_PasList = "finalizada" Then
                    lblOcupadosNoIdentificados(2) = EstadoActual.OcupadosNoIdentificadosCongelados 'GetNoIdentificadosSobrePresentes 'CuentaOcupadosNoIdentificadosCongelados()
                Else
                    lblOcupadosNoIdentificados(2) = GetNoIdentificadosSobrePresentes 'CuentaOcupadosNoIdentificadosCongelados()
                End If
                lblPresentesIdentificados = GetIdentificados 'EstadoActual.Presentes - EstadoActual.OcupadosNoIdentificados
                lblTituloOcupadosNoIdentificados(0).Visible = False 'True
                lblTituloOcupadosNoIdentificados(1).Visible = (EstadoActual.TipoDeOperacion = "votnom")
                lblOcupadosNoIdentificados(1).Visible = False 'True
                lblOcupadosNoIdentificados(2).Visible = (EstadoActual.TipoDeOperacion = "votnom")

                lblTituloPresentesIdentificados.Visible = False 'True
                lblPresentesIdentificados.Visible = False 'True
            Else
                lblTituloOcupadosNoIdentificados(0).Visible = False
                lblTituloOcupadosNoIdentificados(1).Visible = False
                lblOcupadosNoIdentificados(0).Visible = False
                lblOcupadosNoIdentificados(1).Visible = False
                lblOcupadosNoIdentificados(2).Visible = False
                lblTituloPresentesIdentificados.Visible = False
                lblPresentesIdentificados.Visible = False
            End If

            
            
            If xCiclosTotales Mod 10 = 0 Then
                Call PintarBancasCartel
            End If
            'fin cartel encendido (tipo 2)
        End If
        If EstadoActual.CartelEncendido = 3 Then  'Actualizacion cartel mantenimiento
             If lblMantenimientostrPresentes.Visible Then
                lblMantenimientostrPresentes.Visible = False
                lblMantenimientostrAusentes.Visible = False
                lblMantenimientostrPanel1.FontSize = 44
                lblMantenimientostrPanel2.FontSize = 44
                lblMantenimientostrPanel3.FontSize = 44
                lblMantenimientostrPanel1.Font = "Arial"
                lblMantenimientostrPanel2.Font = "Arial"
                lblMantenimientostrPanel3.Font = "Arial"
                lblMantenimientostrPanel1.Height = 1000
                lblMantenimientostrPanel2.Height = 1000
                lblMantenimientostrPanel3.Height = 1000
                lblMantenimientostrPanel2.Height = lblMantenimientostrPanel1.top + 1100
                lblMantenimientostrPanel3.Height = lblMantenimientostrPanel2.top + 1100
                lblMantenimientostrPresencias.Visible = False
                lblMantenimientostrId.Visible = False
                lblMantenimientostrMantListaPendientes.FontSize = 24
                lblMantenimientostrMantListaFallas.FontSize = 24
                Label40.Visible = False
                Label39.Visible = False
                Label41.Visible = False
                Label42.Visible = False
                Line2.Visible = False
             End If
             lblMantenimientostrPresentes.Caption = Str(.Presentes)
             lblMantenimientostrAusentes.Caption = Str(.Ausentes)
             
            For i = 1 To cUltimoPanelMant
                strInfoMant(i) = IIf(EstadoActual.VMantBanca(i) <= xUltimaBanca, Str(EstadoActual.VMantBanca(i)), "") & " " & Left(EstadoActual.VMantInfo(i), 10)
            Next i
             i = 1
             lblMantenimientostrPanel1.Caption = Trim(strInfoMant(i)) & Space(14 - Len(Trim(strInfoMant(i)))) & Left(Trim(strInfoMant(i + 1)), 13)
             i = i + 2
             lblMantenimientostrPanel2.Caption = Trim(strInfoMant(i)) & Space(14 - Len(Trim(strInfoMant(i)))) & Left(Trim(strInfoMant(i + 1)), 13)
             i = i + 2
             Dim xTemp As String
             xTemp = Trim(strInfoMant(i)) & Space(14 - Len(Trim(strInfoMant(i)))) & Left(Trim(strInfoMant(i + 1)), 13)
             If PrimeraVezCeros = True Then
                xTemp = Replace(xTemp, "0", "")
                lblOperador1.Caption = ""
                lblOperador2.Caption = ""
                lblOperador3.Caption = ""
                lblOperador4.Caption = ""
                PrimeraVezCeros = False
             End If
             If Trim(xTemp) <> "" Then
                lblMantenimientostrPanel3.Caption = xTemp
                lblOperador1.Caption = strInfoMant(6)
                If InStr(lblOperador1.Caption, "??") > 0 Then
                    lblOperador1.ForeColor = vbRed
                ElseIf InStr(lblOperador1.Caption, "OK") > 0 Then
                    lblOperador1.ForeColor = vbGreen
                Else
                    lblOperador1.ForeColor = vbWhite
                End If
                lblOperador2.Caption = strInfoMant(5)
                If InStr(lblOperador2.Caption, "??") > 0 Then
                    lblOperador2.ForeColor = vbRed
                ElseIf InStr(lblOperador2.Caption, "OK") > 0 Then
                    lblOperador2.ForeColor = vbGreen
                Else
                    lblOperador2.ForeColor = vbWhite
                End If
                lblOperador3.Caption = strInfoMant(4)
                If InStr(lblOperador3.Caption, "??") > 0 Then
                    lblOperador3.ForeColor = vbRed
                ElseIf InStr(lblOperador3.Caption, "OK") > 0 Then
                    lblOperador3.ForeColor = vbGreen
                Else
                    lblOperador3.ForeColor = vbWhite
                End If
                If Trim(strInfoMant(3)) <> "" Then
                    lblOperador4.Caption = strInfoMant(3)
                End If
                If InStr(lblOperador4.Caption, "??") > 0 Then
                    lblOperador4.ForeColor = vbRed
                ElseIf InStr(lblOperador4.Caption, "OK") > 0 Then
                    lblOperador4.ForeColor = vbGreen
                Else
                    lblOperador4.ForeColor = vbWhite
                End If
             End If
             lblMantenimientostrPresencias.Caption = Trim(EstadoActual.MantPresencias)
             lblMantenimientostrId.Caption = Trim(EstadoActual.MantPresencias)
             lblMantenimientostrFallas.Caption = EstadoActual.MantCantFallas
        
             lblMantenimientostrPendientes.Caption = EstadoActual.MantCantPendientes
        
             lblMantenimientostrMantListaPendientes.Caption = EstadoActual.MantListaPendientes
             lblMantenimientostrMantListaFallas.Caption = EstadoActual.MantListaFallas & GetDesconectadas
             lblMantenimientostrMantListaFallas.Height = lblMantenimientostrMantListaPendientes.Height
        End If
        
        If False And EstadoActual.CartelEncendido >= 1 Then  'Actualizacion cartel Serial. Separar luego el 0 acaa acaa accaca
                
                sCartel.strPresentes = Str(Min(.Presentes, xMiembrosDelCuerpo))
                'sCartel.strAusentes = Str(max(.Ausentes, IIf(EstadoActual.PresidenteHabilitadoParaVotar, 1, 0)))
                sCartel.strAusentes = Str(max(.Ausentes, IIf(xPresidenteLegislador, 1, 0)))
                'fecha automatica
                'hora automatica
                If Not (EstadoActual.ModoMantenimientoBancas = 1) Then
                    
                    sCartel.strAtributo03 = "^V^3" '"^L" & "^V^8"
                    sCartel.strAtributo04 = "^V^3"
                    sCartel.strAtributo05 = "^V^3"
                    sCartel.strAtributo10 = "^V^3"
                
                    sCartel.strQuorum = IIf(.LeyendaQuorum = "QUORUM", "QUORUM", " ")
                                    
                    sCartel.strSesion = LeyendaSesionCartelSerial
                    sCartel.strTitulo = EstadoActual.TituloDelActa
                    sCartel.strOrdenDia = " "
                    sCartel.strMayoria = LeyendaTipoMayoria & " " & LeyendaBaseMayoriaCartelSerial
                    
                    
                    If (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") Then
                        'cartel serial
                        sCartel.strTipoVota = LeyendaTipoOperacionCartelSerial
                        sCartel.strTiempoVota = .LeyendaTiempo
                        sCartel.strNegativos = .Negativos
                        sCartel.strAfirmativos = .Afirmativos
                        sCartel.strAbtenciones = .Abstenciones
                        sCartel.strResultado = .Resultado
                    Else
                        sCartel.strTipoVota = " "
                    End If
                    'control de visibilidad del cartel serial
                    If Not (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") Then
                        sCartel.strTitulo = " "
                        sCartel.strTiempoVota = " "
                        sCartel.strMayoria = " "
                    End If
                    If Not (((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))) Then
                        sCartel.strAbtenciones = " "
                        sCartel.strAfirmativos = " "
                        sCartel.strNegativos = " "
                        sCartel.strResultado = " "
                    End If
                    If Not (EstadoActual.TipoDeOperacion <> "quorum") Then
                        sCartel.strTipoVota = " "
                    End If
                            
                    sCartel.strLineaCartel10 = " "
                    
                    sCartel.strAtributo11 = "^L" & "^V^8"
                    If EstadoActual.TipoDeOperacion = "paslis" Then
                        sCartel.strLineaCartel11 = LeyendaTipoOperacionCartelSerial
                    Else
                        sCartel.strLineaCartel11 = " "
                    End If
                Else 'Mantenimiento
                    
                    sCartel.strQuorum = "PRUEBA"
                                                        
                    For i = 1 To cUltimoPanelMant
                        strInfoMant(i) = IIf(EstadoActual.VMantBanca(i) <= xUltimaBanca, Str(EstadoActual.VMantBanca(i)), "") & " " & Left(EstadoActual.VMantInfo(i), 9)
                        'strInfoMant(i) = Format(EstadoActual.VMantBanca(i), "00") & " " & Left(EstadoActual.VMantInfo(i), 10)
                    Next i
                    
                    i = 1
                    sCartel.strAtributo03 = "^L" & "^V^8" ' "^V^3" '"^L" & "^V^8"
                    sCartel.strSesion = Left(Trim(strInfoMant(i)), 11) & Space(12 - Len(Trim(strInfoMant(i)))) & Left(Trim(strInfoMant(i + 1)), 12)
                                        'Trim (strInfoMant(i)) & Space(13 - Len(Trim(strInfoMant(i)))) & Trim(strInfoMant(i + 1))
                    'sCartel.strAtributo03 = "^V^3"
                                                            
                    i = i + 2
                    sCartel.strAtributo04 = "^L" & "^V^8" '
                    sCartel.strTitulo = Left(Trim(strInfoMant(i)), 11) & Space(12 - Len(Trim(strInfoMant(i)))) & Left(Trim(strInfoMant(i + 1)), 12) 'Trim(strInfoMant(i)) & Space(14 - Len(Trim(strInfoMant(i)))) & Left(Trim(strInfoMant(i + 1)), 13)
                    i = i + 2
                    sCartel.strAtributo05 = "^L" & "^V^8" '
                    sCartel.strOrdenDia = Left(Trim(strInfoMant(i)), 11) & Space(12 - Len(Trim(strInfoMant(i)))) & Left(Trim(strInfoMant(i + 1)), 12)
                    
                    sCartel.strTipoVota = " "
                    sCartel.strTiempoVota = " "
                        
                    sCartel.strMayoria = Left(Trim(EstadoActual.MantPresencias), 13) & "-" & Left(Trim(EstadoActual.MantIdentificaciones), 26 - Len(Trim(Left(Trim(EstadoActual.MantIdentificaciones), 13))))
                    sCartel.strNegativos = EstadoActual.MantCantFallas
                    sCartel.strAfirmativos = " "
                    sCartel.strAbtenciones = EstadoActual.MantCantPendientes
                    sCartel.strResultado = " "
                    sCartel.strAtributo10 = IIf(Len(Trim(EstadoActual.MantListaPendientes)) > 24, "^V^3", "^L" & "^V^8")
                    sCartel.strLineaCartel10 = EstadoActual.MantListaPendientes
                    sCartel.strAtributo11 = IIf(Len(Trim(EstadoActual.MantListaFallas)) > 24, "^V^3", "^L" & "^V^8")
                    sCartel.strLineaCartel11 = EstadoActual.MantListaFallas
                End If
        End If
        If False And EstadoActual.CartelEncendido = 0 Then ' renombrar por caso 0 aca acaaa caccac
                sCartel.strPresentes = " "
                sCartel.strAusentes = " "
                'fecha automatica
                'hora automatica
                sCartel.strQuorum = " "
                                
                sCartel.strSesion = " "
                sCartel.strTitulo = " "
                sCartel.strOrdenDia = " "
                sCartel.strTipoVota = " "
                sCartel.strMayoria = " "
                
                sCartel.strTitulo = " "
                sCartel.strTiempoVota = " "
                sCartel.strMayoria = " "
                sCartel.strAbtenciones = " "
                sCartel.strAfirmativos = " "
                sCartel.strNegativos = " "
                sCartel.strResultado = " "
                sCartel.strTipoVota = " "
                sCartel.strLineaCartel10 = " "
                sCartel.strLineaCartel11 = " "
        End If
        DoEvents
        ' Armo mensaje para log
        strMensajeLog = "Presentes: " & .Presentes & "; Ausentes: " & .Ausentes _
                        & "; Resultado: " & .Resultado & "; Afirmativos: " & .Afirmativos _
                        & "; Negativos: " & .Negativos & "; Abstenciones: " & .Abstenciones _
                        & "; Minimo Para Afirmativo: " & .MinimoVotosParaAfirmativo _
                        & "; Leyenda Quorum: " & .LeyendaQuorum
        ' Verificar si hay que dejar log de mensajes
        If chkLog_Mensajes.Value = 1 Then
            Call AltaLogGeneral("Cartel", strMensajeLog)
        End If
    End With
    Call CartelSerial
    'Call PintarBancasCartel
End Sub
Private Sub CartelSerial(Optional strTipo As String)
    If EstadoActual.CartelEncendido > 0 Then
       'Enviar datos al cartel de leds
    End If
End Sub
Private Function ProcesoDeMensajesQuorum() As Long
    
    Dim strSql                      As String
    Dim MensajeActual               As MensajeSistema
    Dim Mensaje2Banca               As MensajeSistema
    Dim strMensajeLog               As String
    Dim StrTempCadena               As String
    Dim xActualBanca                As Long
    Dim xMax                        As Long
    Dim X                           As Long
    Dim xTiempoRestanteVotacion     As Long
    Dim xTiempoVotacionTranscurrido As Long
    Dim i                           As Long
    Dim j                           As Long
    Dim xStrVector                  As String
    Dim xVotacionReconsideracion    As Boolean
    Dim xVotoOperador               As String
    Dim xNuevoID                    As String
    Dim flIdDupOperador             As Boolean
    Dim xSesionReconsideracion      As Long
    Dim xActaReconsideracion        As Long
    Dim strVector                   As String
    Dim nNuevosAbstenidos           As Long
    Dim nNuevosCancelados           As Long
    Dim nTotalAbstenciones          As Long
    Dim vTemp                       As Variant
    Dim xEsLegislador               As Boolean
    Dim CntInterno As Long
     
    Dim nCBanca As Integer
    
    If (EstadoActual.EstadoVotacion_y_PasList = "espera") Then
        For nCBanca = 1 To 256
            If (EstadoActual.VectorIdentificacion(nCBanca) <> NO_IDENTIFICADO) Then
                If (EstadoActual.VectorColor(nCBanca) <> cCELESTE) Then
                    EstadoActual.VectorColor(nCBanca) = cCELESTE
                End If
            End If
        Next nCBanca
    End If
    
    ' Atender a todos los mensajes nuevos Emitidos por las consolas
   strSql = "SELECT * FROM consola_sqv_mensajes WHERE serial > " & Str(xUltimoMensajeCosola)
   CntInterno = 0
   Call SetearRs(strSql)
   'Call SetearRsCadena(xUltimoMensajeCosola)
    With rs
        While Not .EOF
            ' ----------------------------------------------------------------
            ' Leer mensaje de la consola
            ' ----------------------------------------------------------------
            MensajeActual.sTipo = LCase(Trim(.Fields("Mensaje").Value))
            MensajeActual.sComponente = LCase(Trim(.Fields("Parametro1").Value))
            MensajeActual.sObjeto = LCase(Trim(.Fields("Parametro2").Value))
            MensajeActual.sAtributo = LCase(Trim(.Fields("Ip").Value))
            MensajeActual.sValor = LCase(Trim(.Fields("timestamp").Value))
            ' -------------------------------------------------------------------------------------
            ' Armo mensaje para log
            ' -------------------------------------------------------------------------------------
            strMensajeLog = "Tipo: " & MensajeActual.sTipo & "; Componente: " & MensajeActual.sComponente _
                        & "Objeto: " & MensajeActual.sObjeto & "; Atributo: " & MensajeActual.sAtributo _
                        & "Valor: " & MensajeActual.sValor
            ' -------------------------------------------------------------------------------------
            ' Verificar si hay que dejar log de mensajes
            ' -------------------------------------------------------------------------------------
            If chkLog_Mensajes.Value = 1 Then
                Call AltaLogGeneral("Consola", strMensajeLog, MensajeActual.sObjeto)
            End If
            
            With MensajeActual
                EstadoActual.EstadoVotacion_y_PasList = LTrim(LCase(EstadoActual.EstadoVotacion_y_PasList))
                Select Case .sTipo
                    ' ---------------------------------------------------------------------------------
                    ' Pedido de identificaicon por teclado
                    ' ---------------------------------------------------------------------------------
                    Case Is = "inicia_teclado"
                        vTemp = MensajeActual.sObjeto
                        If vTemp = "brc" Then
                            'okooo
                            '>> el siguiente programa envia a TODOS los que en el vector presencia tengan un 1 y que no esten aun identificados; el comando de identificarse.
                            xStrVector = "0" & SEPARADOR_VECTOR 'presidente no se identifica
                            For i = 1 To UBound(EstadoActual.VectorPresencia)
                                EstadoActual.VTipoIdentificacion(i) = TIPO_IDENTIFICACION_TECLADO
                                xStrVector = xStrVector & IIf(EstadoActual.VectorPresencia(i) = PRESENTE And (EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO), "1", "0") & SEPARADOR_VECTOR
                            Next i
                            Call EnviarMensajesComienzoAuth(xStrVector, "brc procesado a solo presentes no identificados")
                            EstadoActual.strError = "**error"
                            EstadoActual.MensajeAlOperador = "Modo Id. por TECLADO para todas las bancas"
                        Else 'toggle entre opciones teclado y huella
                            EstadoActual.strError = "**error"
                            If EstadoActual.VTipoIdentificacion(MensajeActual.sObjeto) = TIPO_IDENTIFICACION_TECLADO Then
                                EstadoActual.VTipoIdentificacion(MensajeActual.sObjeto) = TIPO_IDENTIFICACION_HUELLA
                                EstadoActual.MensajeAlOperador = "Banca " & MensajeActual.sObjeto & ". Modo Id. por HUELLA"
                            Else
                                EstadoActual.VTipoIdentificacion(MensajeActual.sObjeto) = TIPO_IDENTIFICACION_TECLADO
                                EstadoActual.MensajeAlOperador = "Banca " & MensajeActual.sObjeto & ". Modo Id. por TECLADO"
                            End If
                            If EstadoActual.VectorPresencia(MensajeActual.sObjeto) = PRESENTE And (EstadoActual.VectorIdentificacion(MensajeActual.sObjeto) = NO_IDENTIFICADO) Then
                                Call EnviarMensajesComienzoAuth(MensajeActual.sObjeto, "")
                            End If
                        End If
                    ' ---------------------------------------------------------------------------------
                    ' Levantar legisladores en abstencion
                    ' ---------------------------------------------------------------------------------
                    Case Is = "abstenciones"
                        ' If (LCase(EstadoActual.TipoDeOperacion) = "votnom" And InStr("votando larga espera", EstadoActual.EstadoVotacion_y_PasList) > 0) Or LCase(EstadoActual.TipoDeOperacion) = "quorum" Then
                        If ((LCase(EstadoActual.TipoDeOperacion) = "votnom" Or LCase(EstadoActual.TipoDeOperacion) = "votnum") And InStr("votando larga espera", EstadoActual.EstadoVotacion_y_PasList) > 0) Then
                            If Not IsNull(.sComponente) Then
                                .sComponente = Trim(.sComponente)
                                strVector = Replace(.sComponente, " ", "")
                                EstadoActual.VectorAbstencion = Split(strVector, SEPARADOR_VECTOR)
                                Call AbstenerVector(strVector, nNuevosAbstenidos, nNuevosCancelados, nTotalAbstenciones)
                                EstadoActual.strError = "**error"
                                EstadoActual.MensajeAlOperador = "Solicitud de abstencion. " & vbCrLf & "Total Autorizados:" & Str(nTotalAbstenciones) & _
                                    vbCrLf & "Nuevos Autorizados:" & Str(nNuevosAbstenidos) & _
                                    vbCrLf & "Nuevos Cancelados:" & Str(nNuevosCancelados)
                            Else
                                EstadoActual.strError = "**error"
                                EstadoActual.MensajeAlOperador = "Solicitud de abstencion. No se recibio lista de autorizados a abstenerse"
                            End If
                        Else
                            EstadoActual.strError = "**error"
                            EstadoActual.MensajeAlOperador = "Solicitud de abstencion fuera de quorum nominal o votacion nominal a iniciar o en curso"
                        End If
                    ' ---------------------------------------------------------------------------------
                    ' Usuario enciende o apaga carteles
                    ' ---------------------------------------------------------------------------------
                    Case Is = "cambio?cartelencendido"
                        EstadoActual.strError = "cambio?cartelencendido"
                        If .sComponente = "s" Or .sComponente = "2" Then
                            EstadoActual.CartelEncendido = 2
                            Unload frmNegro
                            'frmMain.Visible = True
                            If EstadoActual.ModoMantenimientoBancas = 1 Then
                                frmCartel2011.Visible = False
                                frmMain.Visible = True
                            Else
                                frmCartel2011.Visible = True
                                frmMain.Visible = False
                            End If
                            'FrameSQVGeneral.ZOrder 0
                        ElseIf .sComponente = "1" Then
                            EstadoActual.CartelEncendido = 1
                            frmSexto.shpHora.Left = frmCartel2011.shpHora.Left
                            frmSexto.shpHora.top = frmCartel2011.shpHora.top
                            frmSexto.lblFecha.Left = frmCartel2011.lblFecha.Left
                            frmSexto.lblFecha.top = frmCartel2011.lblFecha.top
                            frmSexto.lblHora.Left = frmCartel2011.lblHora.Left
                            frmSexto.lblHora.top = frmCartel2011.lblHora.top
                            frmSexto.lblLeyendaReunion.Left = frmCartel2011.lblLeyendaReunion.Left
                            frmSexto.lblLeyendaReunion.top = frmCartel2011.lblLeyendaReunion.top
                            frmSexto.lblNumeroPeriodo.Left = frmCartel2011.lblNumeroPeriodo.Left
                            frmSexto.lblNumeroPeriodo.top = frmCartel2011.lblNumeroPeriodo.top
                            frmSexto.lblNumeroReunion.Left = frmCartel2011.lblNumeroReunion.Left
                            frmSexto.lblNumeroReunion.top = frmCartel2011.lblNumeroReunion.top
                            frmSexto.lblNumeroSesion.Left = frmCartel2011.lblNumeroSesion.Left
                            frmSexto.lblNumeroSesion.top = frmCartel2011.lblNumeroSesion.top
                            frmSexto.lblSeparacionPeriodo.Left = frmCartel2011.lblSeparacionPeriodo.Left
                            frmSexto.lblSeparacionPeriodo.top = frmCartel2011.lblSeparacionPeriodo.top
                            frmSexto.lblSeparadorReunion.Left = frmCartel2011.lblSeparadorReunion.Left
                            frmSexto.lblSeparadorReunion.top = frmCartel2011.lblSeparadorReunion.top
                            frmSexto.lblSeparadorSesion.Left = frmCartel2011.lblSeparadorSesion.Left
                            frmSexto.lblSeparadorSesion.top = frmCartel2011.lblSeparadorSesion.top
                            frmSexto.lblTipoPeriodo.Left = frmCartel2011.lblTipoPeriodo.Left
                            frmSexto.lblTipoPeriodo.top = frmCartel2011.lblTipoPeriodo.top
                            frmSexto.lblTipoSesion.Left = frmCartel2011.lblTipoSesion.Left
                            frmSexto.lblTipoSesion.top = frmCartel2011.lblTipoSesion.top
                            'Datos
                            frmSexto.lblFecha.Caption = frmCartel2011.lblFecha.Caption
                            frmSexto.lblHora.Caption = frmCartel2011.lblHora.Caption
                            frmSexto.lblLeyendaReunion.Caption = frmCartel2011.lblLeyendaReunion.Caption
                            frmSexto.lblNumeroPeriodo.Caption = frmCartel2011.lblNumeroPeriodo.Caption
                            frmSexto.lblNumeroReunion.Caption = frmCartel2011.lblNumeroReunion.Caption
                            frmSexto.lblNumeroSesion.Caption = frmCartel2011.lblNumeroSesion.Caption
                            frmSexto.lblTipoPeriodo.Caption = frmCartel2011.lblTipoPeriodo.Caption
                            frmSexto.lblTipoSesion.Caption = frmCartel2011.lblTipoSesion.Caption
                            frmCartel2011.Visible = False
                            frmSexto.Show
                        ElseIf .sComponente = "3" Then 'Mantenimiento
                            EstadoActual.CartelEncendido = 3
                            'FrameMantenimiento.ZOrder 0
                        ElseIf .sComponente = "4" Then
                            EstadoActual.CartelEncendido = 4
                            FrameSQVActa.ZOrder 0
                        Else '.sComponente = "n" Or .sComponente = "0"
                            EstadoActual.CartelEncendido = 0
                            'frmMain.Visible = False
                            frmCartel2011.Visible = False
                            frmNegro.Show
                            'FrameSQVApagado.ZOrder 0
                            txtTipoOperacion = "d"
                        End If
                    Case Is = "cambio?expresiones_minoria"
                        If EstadoActual.Expresiones_Minoria = True Then
                            EstadoActual.Expresiones_Minoria = False
                        Else
                            EstadoActual.Expresiones_Minoria = True
                        End If
                        ' ---------------------------------------------------------------------------------
                        ' Usuario inicia prueba scan de una banca
                        ' ---------------------------------------------------------------------------------
                    Case Is = "banca?deshabilitar"
                        Dim i_banca As String
                        i_banca = .sComponente
                        Call EnviarMensajesFinAuth(i_banca, "Cancelacion de LED por deshabilitacion")
                        Dim nT As Long
                        nT = GetTickCount
                        While GetTickCount - nT < 2000
                            DoEvents 'Doy tiempo para que se ejecute el FinAuth antes de cerrar socket
                        Wend
                        With Mensaje2Banca ' Mensaje para SB
                            .sTipo = "mget"
                            .sComponente = "term.mon"
                            .sComentario = "modo_deshabilitar"
                            .sObjeto = i_banca
                            .sAtributo = "action"
                            .sValor = "reset"
                        End With
                        Call EnviarMensajesBancas(Mensaje2Banca)
                    Case Is = "scan?prueba"
                        EstadoActual.BancaEnPrueba = Val(.sComponente)
                        Call EnviarMensajesComienzoAuth(.sComponente, "Prueba de Scan 2011")
                    Case Is = "scan?finprueba"
                        Call EnviarMensajesFinAuth(.sComponente, "FIN Prueba de Scan 2011")
                        EstadoActual.BancaEnPrueba = -1
                    Case Is = "pruebascan"
                        If EstadoActual.TipoDeOperacion = "quorum" Then
                            If xBancaPruebaScan > 0 Then 'Ya estaba otra banca en prueba
                                'Call EnviarMensajesFinAuth(Str(xBancaPruebaScan), "Prueba Scan Fin por nuevo pedido de prueba scan")
                                Call MensajeDisplayTerminal(Str(xBancaPruebaScan), "Fin de prueba de scan (1)")
                            End If
                            xBancaPruebaScan = Int(IIf(.sObjeto = "", 0, .sObjeto)) ' Nueva banca en prueba de scan
                            If xBancaPruebaScan > 0 Then
                                Call EnviarMensajesComienzoAuth(Str(xBancaPruebaScan), "Prueba Scan comienzo por nuevo pedido de prueba scan")
                                Call MensajeDisplayTerminal(Str(xBancaPruebaScan), "Comienzo de prueba de scan")
                            End If
                        End If
                        ' ////////////////////////////////////////////////////////////////////////////////////
                        ' Usuario finaliza la prueba scan de una banca
                    Case Is = "pruebascanfin"
                        If Not IsNull(.sComponente) Then
                            If xBancaPruebaScan > 0 Then 'Ya estaba una banca en prueba
                                Call EnviarMensajesFinAuth(Str(xBancaPruebaScan), "Prueba Scan Fin por pedido de fin de prueba scan")
                                Call MensajeDisplayTerminal(Str(xBancaPruebaScan), "Fin de prueba de scan (0)")
                                xBancaPruebaScan = 0
                                EstadoActual.strError = ""
                            End If
                        End If
                    Case Is = "pruebascanlimpiar"
                        EstadoActual.strError = ""
                        xBancaPruebaScan = 0
                    Case Is = "limpiaridpruebascan"
                        Call EnviarMensajesFinAuth(.sComponente, "Prueba Scan Fin por pedido de fin de prueba scan")
                        ' ---------------------------------------------------------------------------------
                        ' Usuario cambia el tipo de operacion
                        ' ---------------------------------------------------------------------------------
                    Case Is = "limpieza_individual"
                        If EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis" Or EstadoActual.Modo_Ident_Nom Then
                            Call EnviarMensajesComienzoAuth(Trim(.sComponente), "Operador limpio ID y vuelve a pedir Idenficacion")
                        Else
                            Call EnviarMensajesFinAuth(Trim(.sComponente), "Operador Limpió ID")
                        End If
                        EstadoActual.VectorIdentificacion(Val(.sComponente)) = NO_IDENTIFICADO
                        EstadoActual.VectorColor(Val(.sComponente)) = AsignarColor(Val(.sComponente))
                    Case Is = "cambio?tipoop"
                        Call DeshabilitarVotoPresidente
                        'FIX de Colores
                        'Nos aseguramos que al pasar de modo, se ignoran todos
                        'los que se estaban intentando identificar
                        'por falta de mensaje de TIMEOUT
                        For nCBanca = 1 To 256
                            If (EstadoActual.VectorIdentificacion(nCBanca) <> NO_IDENTIFICADO) Then
                                'Si esta identficado
                                If (EstadoActual.VectorColor(nCBanca) <> cCELESTE) Then
                                    EstadoActual.VectorColor(nCBanca) = cCELESTE
                                End If
                            End If
                        Next nCBanca
                        For nCBanca = 1 To 256
                            If (EstadoActual.VectorColor(nCBanca) = cMarronClaro) Then
                                EstadoActual.VectorColor(nCBanca) = cAMARILLO
                            End If
                        Next nCBanca
                        'Termina FIX de colores
                        ModoMant = False
                        PrimeraVezControl = True
                        EstadoActual.ExtensionDeTiempoPorPresidente = False
                        EstadoActual.ExtensionDeTiempoPorPresidente = False
                        EstadoActual.strError = "cambio?tipoop"
                        If EstabaEnIdentificacion = True Then
                            EstadoActual.Modo_Ident_Nom = 1
                            EstabaEnIdentificacion = False
                        End If
                        If Not IsNull(.sComponente) Then
                            If (InStr("votnom;votnum", Trim(EstadoActual.TipoDeOperacion)) > 0) Then
                                'Elimina las abstenciones al cambiar el modo
                                For i = 0 To Min(xUltimaBanca, UBound(EstadoActual.VectorAbstencion))
                                    EstadoActual.VectorAbstencion(i) = 0
                                Next
                                Call AbstenerVector(Join(EstadoActual.VectorAbstencion, SEPARADOR_VECTOR), 0, 0, 0)
                            End If
                            '>> IMPORTANTE: si pasa de un modo de no identifiacion a uno que permite la identificacion, debe habilitar a identificarse a todos los presentes
                            'cambio de no identificacion a identificacion  de (quorum o votacion numerica ) a (vnominal o pase de lista)
                            If (EstadoActual.Modo_Ident_Nom = 1 And EstadoActual.TipoDeOperacion <> "votnum") Then
                                Call SolicitarIdentificacionPendientes("Comienzo modo nominal desde " & EstadoActual.TipoDeOperacion & "a " & .sComponente, "start")
                                EstadoActual.OcupadosNoIdentificados = CuentaOcupadosNoIdentificadosCong  'EstadoActual.Presentes - IIf(xPresidenteLegislador, 1, 0) 'aca5
                            ElseIf (Trim(EstadoActual.TipoDeOperacion) = "votnum" Or (Trim(EstadoActual.TipoDeOperacion) = "quorum" And Not (EstadoActual.Modo_Ident_Nom = 1))) And InStr("votnom;paslis", Trim(.sComponente)) > 0 Then
                                'Caso: Pasa de un modo no nominal a uno nominal: deben identificarse todos. 091018
                                'antes nominal If (InStr("quorum;votnum", Trim(EstadoActual.TipoDeOperacion)) > 0) And InStr("votnom;paslis", Trim(.sComponente)) > 0 Then
                                'enviar broadcast identificarse Msj mset/term.auth?ACTION=AUTH_START
                                '>> el siguiente programa envia a TODOS los que en el vector presencia tengan un 1 el comando de identificarse.
                                'multifruta
                                xStrVector = "0" & SEPARADOR_VECTOR 'presidente
                                For i = 1 To UBound(EstadoActual.VectorPresencia)
                                    'xStrVector = xStrVector & IIf(EstadoActual.VectorPresencia(i) = "1", "1", "0") & SEPARADOR_VECTOR
                                    xStrVector = xStrVector & IIf(EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO And EstadoActual.VectorPresencia(i) <> AUSENTE, "1", "0") & SEPARADOR_VECTOR
                                Next i
                                Call EnviarMensajesComienzoAuth(xStrVector, "Comienzo modo nominal desde " & EstadoActual.TipoDeOperacion & "a " & .sComponente)
'                                For i = 1 To (xUltimaBanca)
'                                    EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO
'                                    EstadoActual.VectorColor(i) = AsignarColor(i)
'                                Next i
                                EstadoActual.OcupadosNoIdentificados = EstadoActual.Presentes - IIf(xPresidenteLegislador, 1, 0) 'aca5
                                'EstadoActual.OcupadosNoIdentificados = EstadoActual.Presentes - IIf(EstadoActual.PresidenteHabilitadoParaVotar, 1, 0) 'aca5
                            '>> A la inversa, cancela las identificaciones y la posibilidad de identificarse a todos los presentes.
                            ElseIf (InStr("votnom;paslis", Trim(EstadoActual.TipoDeOperacion)) > 0) And _
                                 (Trim(.sComponente) = "votnum" Or (Trim(.sComponente) = "quorum" And Not (EstadoActual.Modo_Ident_Nom = 1))) Then
                                'Caso: Pasa de un modo nominal a uno no nominal: deben perderse las identificaciones de todos.
                                'antes nominal If (InStr("votnom;paslis", Trim(EstadoActual.TipoDeOperacion)) > 0) And InStr("quorum;votnum", Trim(.sComponente)) > 0 Then
                                'TODO COMENTADO PARA QUE NO SE PIERDAN LAS IDENTIFICACIONES 14FEB
                                xStrVector = "0" & SEPARADOR_VECTOR 'presidente
                                For i = 1 To UBound(EstadoActual.VectorPresencia)
                                    'xStrVector = xStrVector & IIf(EstadoActual.VectorPresencia(i) = "1", "1", "0") & SEPARADOR_VECTOR
                                    xStrVector = xStrVector & IIf(EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO, "1", "0") & SEPARADOR_VECTOR
                                Next i
                                Call EnviarMensajesFinAuth(xStrVector, "Fin mod<o nominal desde " & EstadoActual.TipoDeOperacion & "a " & .sComponente)
'                                For i = 1 To (xUltimaBanca)
'                                    EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO
'                                    EstadoActual.VectorColor(i) = AsignarColor(i)
'                                Next i
                                EstadoActual.OcupadosNoIdentificados = 0
                            End If
                            If (Trim(.sComponente) = "votnum") Then
                                'Modo nominal, votacion numerica. Se trata como nominal pero se guarda el tipo de operacion en el auxiliar
                                EstadoActual.TipoDeOperacion = "votnum" 'HCDN 2011
                                xTipoVotacion = "votnum"
                                If EstadoActual.Modo_Ident_Nom = 1 Then
                                    EstabaEnIdentificacion = True
                                    EstadoActual.Modo_Ident_Nom = 0
                                End If
                                xStrVector = "0" & SEPARADOR_VECTOR 'presidente
                                For i = 1 To UBound(EstadoActual.VectorPresencia)
                                    xStrVector = xStrVector & IIf(EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO, "1", "0") & SEPARADOR_VECTOR
                                Next i
                                Call EnviarMensajesFinAuth(xStrVector, "Fin mod<o nominal desde " & EstadoActual.TipoDeOperacion & "a " & "")
                            Else
                                ' Modo no nominal (o la operacion no es votacion numerica) , el tipo de operacion coincide con el modo o seleccionado
                                EstadoActual.TipoDeOperacion = LCase(Trim(.sComponente))
                                xTipoVotacion = LCase(Trim(.sComponente))
                                If EstadoActual.TipoDeOperacion = "votnom" Then
                                    Call AbstenerVector(Join(EstadoActual.VectorAbstencion, SEPARADOR_VECTOR), 0, 0, 0)
                                End If
                            End If
                            'antes nominal EstadoActual.TipoDeOperacion = .sComponente
                            
                            'inicializacion parametros votacion
                            'If (InStr("votnom;votnum", Trim(EstadoActual.TipoDeOperacion)) > 0) Then
                            '    Call InicializarVotacion
                            'End If
                            'Dim maxbanca As Long
                            If (InStr("votnom;votnum", Trim(EstadoActual.TipoDeOperacion)) > 0) Then
                                Call InicializarVotacion
                            End If
                            
                            Call PintarTodasLasBancas
                            EstadoActual.LimpiarResultados = 1
                       End If
                    ' Usuario cambia el tiempo de votacion
                Case Is = "cambio?tiempo"
                    If Not IsNull(.sComponente) Then
                        EstadoActual.strError = "cambio?tiempo"
                        EstadoActual.TiempoParaVotacion = Int(.sComponente)
                    End If
                ' ---------------------------------------------------------------------------------
                ' Usuario cambia el presidente
                ' ---------------------------------------------------------------------------------
                Case Is = "cambio?presidente" 'And EstadoActual.TipoDeOperacion = "quorum"
                    EstadoActual.strError = "cambio?presidente"
                    If Not IsNull(.sComponente) Then
                        If EstadoActual.VectorIdentificacion(0) = NO_IDENTIFICADO Then
                            If EstadoActual.VectorPresencia(0) <> PRESENTE Then 'habria una inconsistencia en el caso en que no se tuviera el presidente anterior, pero este hubiera sido legislador.
                                'este caso vale tanto para AUSENTE como para BANCA_INHABILITADA
                                xPresidenteAnteriorLegislador = False
                            Else
                                xPresidenteAnteriorLegislador = True 'no se sabe quien era, pero era legislador
                            End If
                        Else
                            'buscar si es legislador el presidente anterior
                            strSql = "SELECT Es_Legislador FROM Legisladores WHERE id = '" & Trim(EstadoActual.VectorIdentificacion(0)) & "'"
                            rsTemp.CursorLocation = adUseClient
                            rsTemp.Open strSql, Cn, adOpenForwardOnly, adLockReadOnly
                            If rsTemp.RecordCount > 0 And (rsTemp.EOF = False Or rsTemp.BOF = False) Then
                                If rsTemp("Es_Legislador").Value = 0 Then
                                    xPresidenteAnteriorLegislador = False
                                Else
                                    xPresidenteAnteriorLegislador = True
                                End If
                            Else ' si hubo error, tomar del vector presencia. Aqui, habria una inconsistencia en el caso en que no se tuviera el presidente anterior, pero este hubiera sido legislador.
                                If EstadoActual.VectorPresencia(0) <> PRESENTE Then
                                    xPresidenteAnteriorLegislador = False
                                Else
                                    xPresidenteAnteriorLegislador = True
                                End If
                            End If
                            rsTemp.Close
                        End If
                        ' Buscar en la base si es legislador el nuevo presidente
                        strSql = "SELECT Es_Legislador FROM Legisladores WHERE id = '" & Trim(.sComponente) & "'"
                        rsTemp.CursorLocation = adUseClient
                        rsTemp.Open strSql, Cn, adOpenForwardOnly, adLockReadOnly
                        If rsTemp.RecordCount > 0 And (rsTemp.EOF = False Or rsTemp.BOF = False) Then
                            If rsTemp("Es_Legislador").Value = 0 Or Not PermitirVotarAlPresidente Then
                                xPresidenteLegislador = False
                            Else
                                xPresidenteLegislador = True
                            End If
                            If xPresidenteLegislador = True Then
                                If Not xPresidenteAnteriorLegislador Then
                                    EstadoActual.VectorPresencia(0) = PRESENTE
                                    If False Then
                                        EstadoActual.Presentes = EstadoActual.Presentes + 1
                                        EstadoActual.Ausentes = EstadoActual.Ausentes - 1
                                    End If
                                End If
                            Else
                                If xPresidenteAnteriorLegislador Then
                                    EstadoActual.VectorPresencia(0) = AUSENTE
                                    If False Then
                                        EstadoActual.Presentes = EstadoActual.Presentes - 1
                                        EstadoActual.Ausentes = EstadoActual.Ausentes + 1
                                    End If
                                End If
                            End If
                            Call PintarVectorColor(0)
                            EstadoActual.VectorIdentificacion(0) = Trim(.sComponente)
                        Else
                            EstadoActual.strError = "**error"
                            EstadoActual.MensajeAlOperador = "ID Presidente " & Trim(.sComponente) & " inválido. No está registrado en el sistema"
                        End If
                        rsTemp.Close
                    Else
                        EstadoActual.strError = "**error"
                        EstadoActual.MensajeAlOperador = "ID Presidente Nulo"
                    End If
                ' ---------------------------------------------------------------------------------
                ' Usuario cambia el orador
                ' ---------------------------------------------------------------------------------
                Case Is = "cambio?orador"
                    EstadoActual.strError = "cambio?orador"
                    If Not IsNull(.sComponente) Then
                        If Not .sComponente = "0" Then
                            ' Buscar en la base si es legislador
                            If DISTRITO_HABILITADO Then
                                strSql = _
                                    "SELECT     Legisladores.es_legislador, Legisladores.apellido, Legisladores.nombre, Legisladores.grupo_politico, Legisladores.bloque_politico, " + _
                                    "secciones.seccion , distritos.distrito, Legisladores.sexo " + _
                                    "FROM         Legisladores LEFT OUTER JOIN " + _
                                    "distritos ON Legisladores.distrito = distritos.id_distrito LEFT OUTER JOIN " + _
                                    "secciones ON distritos.seccion = secciones.id_seccion " + _
                                    "WHERE id = '" & Trim(.sComponente) & "'"
                            Else
                                strSql = _
                                    "SELECT     Legisladores.es_legislador, Legisladores.apellido, Legisladores.nombre, Legisladores.grupo_politico, Legisladores.bloque_politico, " + _
                                    "0 as seccion , ' ' as distrito, Legisladores.sexo " + _
                                    "FROM         Legisladores " & _
                                    "WHERE id = '" & Trim(.sComponente) & "'"
                            End If
                            rsTemp.CursorLocation = adUseClient
                            rsTemp.Open strSql, Cn, adOpenForwardOnly, adLockReadOnly
                            If rsTemp.RecordCount > 0 And (rsTemp.EOF = False Or rsTemp.BOF = False) Then
                                If rsTemp("Es_Legislador").Value = 0 Then
                                    xEsLegislador = False
                                Else
                                    xEsLegislador = True
                                End If
                                If xEsLegislador Then
                                    EstadoActual.Orador = Trim(.sComponente)
                                    EstadoActual.OradorNombre = rsTemp("apellido") + " " + rsTemp("nombre")
                                    EstadoActual.OradorSexo = IIf(Val(rsTemp("sexo")) = 0, "F", "M")
                                    EstadoActual.OradorAgrupacionPolitica = IIf(AGRUPACION_POLITICA_HABILITADA, Trim(rsTemp("grupo_politico")) & " - ", "") & Trim(rsTemp("bloque_politico"))
                                    If rsTemp.Fields("seccion") <> vbNull Then
                                        EstadoActual.OradorDistrito = IIf(DISTRITO_HABILITADO, Trim(rsTemp("seccion")), "")
                                    Else
                                        EstadoActual.OradorDistrito = "Sin Provincia"
                                    End If
                                Else
                                    EstadoActual.strError = "**error"
                                    EstadoActual.MensajeAlOperador = "ID Orador" & Trim(.sComponente) & " inválido. No está registrado en el sistema como Legislador"
                                End If
                            Else
                                EstadoActual.strError = "**error"
                                EstadoActual.MensajeAlOperador = "ID Orador " & Trim(.sComponente) & " inválido. No está registrado en el sistema."
                            End If
                            rsTemp.Close
                        Else
                            EstadoActual.Orador = "" ' blanquea orador
                        End If
                    Else
                        EstadoActual.Orador = "" ' blanquea orador
                    End If
                Case Is = "listapendientes?siguiente"
                    Pendientes.paginaActualPendientes = Pendientes.paginaActualPendientes + 1
                Case Is = "listapendientes?anterior"
                    If Pendientes.paginaActualPendientes > 0 Then
                        Pendientes.paginaActualPendientes = Pendientes.paginaActualPendientes - 1
                    End If
                    ' ---------------------------------------------------------------------------------
                    ' Usuario cambia reunion
                    ' ---------------------------------------------------------------------------------
                Case Is = "cambio?reunion"
                    EstadoActual.strError = "cambio?reunion"
                    If Not IsNull(.sComponente) Then
                        EstadoActual.Reunion = Val(Trim(.sComponente))
                        EjecutarSQL ("UPDATE perparl SET Ultima_Reunion = " & EstadoActual.Reunion & " WHERE Período_Legislativo LIKE '" & Left(EstadoActual.PeriodoLegislativo, 3) & "%'")
                    Else
                        EstadoActual.strError = "**error"
                        EstadoActual.MensajeAlOperador = "Reunion invalida"
                    End If
                    ' ---------------------------------------------------------------------------------
                    ' sincronizar
                    ' ---------------------------------------------------------------------------------
                Case Is = "accion?sincronizar"
                    EstadoActual.strError = "accion?sincronizar"
                    If Not IsNull(.sComponente) Then
                        Call SincronizarBancas((Trim(.sComponente)))
                    Else
                        EstadoActual.strError = "**error"
                        EstadoActual.MensajeAlOperador = "Reunion invalida"
                    End If
                    ' ---------------------------------------------------------------------------------
                    ' Usuario cambia el tipo de abstencion para la votacion
                    ' ---------------------------------------------------------------------------------
                Case Is = "cambio?modvot"
                    If Not IsNull(.sComponente) Then
                        EstadoActual.strError = "cambio?modvot"
                        EstadoActual.TipoDeAbstencion = .sComponente
                    End If
                    ' ---------------------------------------------------------------------------------
                    ' Usuario cambia el tipo de Quórum para votacion
                    ' ---------------------------------------------------------------------------------
                Case Is = "cambio?tipoquorum"
                    If Not IsNull(.sComponente) Then
                        EstadoActual.TipoMayoriaQuorum = .sComponente
                        EstadoActual.strError = "cambio?tipoquorum"
                        Call CalcularMinimoParaQuorum
                    End If
                    ' ---------------------------------------------------------------------------------
                    ' Usuario cambia la base de la mayoría para la votación
                    ' ---------------------------------------------------------------------------------
                Case Is = "cambio?basemayoria"
                    If Not IsNull(.sComponente) Then
                        EstadoActual.BaseMayoria = .sComponente
                        EtiquetasCartel.strBase = DevolverLeyendaBase(.sComponente)
                        EstadoActual.strError = "cambio?basemayoria"
                    End If
                    ' ---------------------------------------------------------------------------------
                    ' Usuario cambia el tipo de la mayoría para la votación
                    ' ---------------------------------------------------------------------------------
                Case Is = "cambio?tipomayoriavotacion"
                    If Not IsNull(.sComponente) Then
                        EstadoActual.TipoMayoria = .sComponente
                        EtiquetasCartel.strTipo = DevolverLeyendaTipo(.sComponente)
                        EstadoActual.strError = "cambio?tipomayoriavotacion"
                    End If
                    ' ---------------------------------------------------------------------------------
                    ' Usuario cambia la sesión actual
                    ' ---------------------------------------------------------------------------------
                Case Is = "cambiosesion"
                    'FALTA:
                    '>> Busca una sesion de estado abierta o nueva en la tabla de sesiones, del periodo parlamentario actual, y cuyo numero de sesion coincida con la solicitada por el operador.
                    '>> luego actualiza el numero de proximo acta con el proximo acta de esa sesion y tambien el estadoactual.estadosesion con es estado de la sesion
                    If Not IsNull(.sComponente) Then
                        ' Buscar en Sesion ultima acta y estado sesion
                        strSql = "SELECT Próximo_Acta, Estado_sesión From Sesion Where Sesión = '" & .sComponente & "' And Período_Legislativo = '" & EstadoActual.PeriodoLegislativo & "' And lower(Estado_sesión) in ('abierta','nueva') "
                        Call SetearOtroRs(strSql)
                        If RsOtro.RecordCount = 0 Or RsOtro.EOF = True Or RsOtro.BOF = True Then
                            ' no se puede seleccionar la sesion pedida, no se hace nada, pero devuelve el error.
                            EstadoActual.strError = "**error"
                            EstadoActual.MensajeAlOperador = "Nro Sesion no encontrada o no valida para seleccionar. Periodo " & EstadoActual.PeriodoLegislativo
                        Else
                            EstadoActual.NroActa = RsOtro.Fields("Próximo_Acta").Value
                            EstadoActual.EstadoSesion = Trim(RsOtro.Fields("Estado_sesión").Value)
                            EstadoActual.Sesion = .sComponente
                        End If
                        RsOtro.Close
                    End If
                    
                    ' ---------------------------------------------------------------------------------
                    ' Simulación masiva de votos negativos
                    ' ---------------------------------------------------------------------------------
                Case Is = "simulacion?votonegativo"
                        With Mensaje2Banca
                            .sTipo = "simulacion_voto"
                            .sComponente = "s_term.keyb"
                            .sObjeto = "0"
                            .sAtributo = "s_voto"
                            .sValor = "s_negativo"
                        End With
                        Call EnviarMensajesBancas(Mensaje2Banca)
                        
                    ' ---------------------------------------------------------------------------------
                    ' Usuario inicia una votación
                    ' ---------------------------------------------------------------------------------
                Case Is = "accion?iniciovotacion"
                    Imprimio = False
                    PrimeraVezControl = True
                    EstadoActual.ExtensionDeTiempoPorPresidente = False
                    If EstadoActual.EstadoVotacion_y_PasList = "espera" And (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum") Then
                        'EstadoActual.PendientesEmitirVotos = EstadoActual.Presentes - EstadoActual.AbstencionistasAutorizados 'Como el presidente solo suma para presente si es legislador no hace falta restarlo:  - IIf(xPresidenteLegislador, 0, 1) ' Si el presidente no es legislador no se espera su voto
                        EstadoActual.PendientesEmitirVotos = EstadoActual.Presentes - EstadoActual.AbstencionistasAutorizados - IIf(EstadoActual.PresidenteHabilitadoParaVotar, 0, 1) ' Si el presidente no es legislador no se espera su voto'Como el presidente solo suma para presente si es legislador no hace falta restarlo:  - IIf(xPresidenteLegislador, 0, 1) ' Si el presidente no es legislador no se espera su voto
                        xHuboEmpate = False
                        xHuboDesempate = False
                        xCierreEmpateOperador = False
                        xVotoSenadorEmpate = ""
                        
                        xMax = UBound(EstadoActual.VectorPresencia)
                        'strTempCadena = IIf(xPresidenteLegislador, "1", "0") & SEPARADOR_VECTOR
                        
                        Call AltaLogGeneral("SQV SERVER: consola accion?iniciovotacion 1", "EstadoActual.PendientesEmitirVotos = EstadoActual.Presentes - EstadoActual.AbstencionistasAutorizados - IIf(xPresidenteLegislador, 0, 1): " & EstadoActual.PendientesEmitirVotos, 0, "0")
                        
                        If EstadoActual.TipoDeOperacion = "votnom" Then
                            ' Pasar el vector presencia a un string
                            StrTempCadena = ("0") & SEPARADOR_VECTOR
                            For X = 1 To xMax
                                If EstadoActual.VectorPresencia(X) = PRESENTE Then
                                    If (EstadoActual.TipoDeOperacion = "votnum" _
                                        Or Not (EstadoActual.VectorIdentificacion(X) = NO_IDENTIFICADO)) _
                                        And EstadoActual.VectorResultados(X) <> ABSTENCION_AUTORIZADA Then 'Solo los no identif
                                        StrTempCadena = StrTempCadena & PRESENTE & SEPARADOR_VECTOR
                                    Else
                                        StrTempCadena = StrTempCadena & AUSENTE & SEPARADOR_VECTOR
                                    End If
                                Else
                                    StrTempCadena = StrTempCadena & AUSENTE & SEPARADOR_VECTOR
                                End If
                            Next X
                            With Mensaje2Banca ' Mensaje para SB
                                .sTipo = "mset"
                                .sComponente = "term.keyb"
                                .sObjeto = StrTempCadena
                                .sAtributo = "state"
                                .sValor = "on" & EstadoActual.TipoDeOperacion
                            End With
                            Call EnviarMensajesBancas(Mensaje2Banca)
                        Else
                            '**********IMPLEMENTACION PARA BANCAS IDENTIFICADAS EN VOTACION NUMERICA**********
                            StrTempCadena = ("0") & SEPARADOR_VECTOR
                            For X = 1 To xMax
                                If EstadoActual.VectorPresencia(X) = PRESENTE Then
                                    If (EstadoActual.TipoDeOperacion = "votnum" _
                                        And (EstadoActual.VectorIdentificacion(X) <> NO_IDENTIFICADO)) _
                                        And EstadoActual.VectorResultados(X) <> ABSTENCION_AUTORIZADA Then 'Solo los identif
                                        StrTempCadena = StrTempCadena & PRESENTE & SEPARADOR_VECTOR
                                    Else
                                        StrTempCadena = StrTempCadena & AUSENTE & SEPARADOR_VECTOR
                                    End If
                                Else
                                    StrTempCadena = StrTempCadena & AUSENTE & SEPARADOR_VECTOR
                                End If
                            Next X
                            With Mensaje2Banca ' Mensaje para SB
                                .sTipo = "mset"
                                .sComponente = "term.keyb"
                                .sObjeto = StrTempCadena
                                .sAtributo = "state"
                                .sValor = "onvotnom" 'SE TRATA COMO NOMINAL (SVOTAR)
                            End With
                            Call EnviarMensajesBancas(Mensaje2Banca)
                            'AHORA MANDO MENSAJE A LAS BANCAS QUE NO ESTAN IDENTIFICADAS (SVOTNU)
                            StrTempCadena = ("0") & SEPARADOR_VECTOR
                            For X = 1 To xMax
                                If EstadoActual.VectorPresencia(X) = PRESENTE Then
                                    If (EstadoActual.TipoDeOperacion = "votnum" _
                                        And (EstadoActual.VectorIdentificacion(X) = NO_IDENTIFICADO)) _
                                        And EstadoActual.VectorResultados(X) <> ABSTENCION_AUTORIZADA Then 'Solo los no identif
                                        StrTempCadena = StrTempCadena & PRESENTE & SEPARADOR_VECTOR
                                    Else
                                        StrTempCadena = StrTempCadena & AUSENTE & SEPARADOR_VECTOR
                                    End If
                                Else
                                    StrTempCadena = StrTempCadena & AUSENTE & SEPARADOR_VECTOR
                                End If
                            Next X
                            With Mensaje2Banca ' Mensaje para SB
                                .sTipo = "mset"
                                .sComponente = "term.keyb"
                                .sObjeto = StrTempCadena
                                .sAtributo = "state"
                                .sValor = "onvotnum"
                            End With
                            Call EnviarMensajesBancas(Mensaje2Banca)
                        End If
                        EstadoActual.EstadoVotacion_y_PasList = "votando"
                        If "VOTO PRESIDENTE" = "SIEMPRE" Then 'alternativa deshabilitada para hcdn 11
                            If xPresidenteLegislador Then
                                With Mensaje2Banca ' Mensaje para SB
                                    .sTipo = "mset"

                                    .sComponente = "term.keyb"
                                    .sObjeto = "0"
                                    .sAtributo = "state"
                                    .sValor = "onvotnum"
                                End With
                                Call EnviarMensajesBancas(Mensaje2Banca)
                            End If
                        Else 'solo cuando se habilita votar al presidente
                            Call ComenzarVotacionPresidente
                        End If
                        'MsgBox "validar bien el tema de los tiempos..."
                        EstadoActual.FechaVotacion = DateAdd("s", xtiempoInicioVotac, Now)
                        tFinVotacion = DateAdd("s", EstadoActual.TiempoParaVotacion, EstadoActual.FechaVotacion)
                        ' EstadoActual.HoraVotacion = Hour(EstadoActual.FechaVotacion) & ":" & Minute(EstadoActual.FechaVotacion) & ":" & Second(EstadoActual.FechaVotacion)
                        EstadoActual.LimpiarResultados = 1
                    End If
                    ' ------------------------------------------------------------------------------------
                    ' Usuario cierra una votación
                    ' SI EL OPERADOR PRESIONA EL BOTON DE CIERRE CUANDO ES UNA VOTACION LARGA Y TODAVIA SE ESTA VOTANDO.
                    ' SI EN LA V.1.0 LO PERMITE EN EMPATE, DEBE ELIMINARSE EN CASO DE NO PERMITIRSE EL EMPATE COMO RESULTADO
                    ' FINAL DE UNA VOTACION. Por ello solo lo permite si esta en 'larga'
                    ' ------------------------------------------------------------------------------------
                Case Is = "accion?cierrevotacion"
                    If EstadoActual.EstadoVotacion_y_PasList = "larga" Or EstadoActual.EstadoVotacion_y_PasList = "empate" Then
                        VL.log "--------------------------- Llamada 2 ---------------------------"
                        EstadoActual.EstadoVotacion_y_PasList = "cierre"
                        Call FinVotacionBrc("cierre operador")
                        xCierreEmpateOperador = True
                    End If
                    
                    ' ------------------------------------------------------------------------------------
                    ' Usuario cancela una votaciónUsuario cancela una votación
                    ' ------------------------------------------------------------------------------------
                Case Is = "accion?cancelavotacion"
                    lAbstencionPresidente = False
                    With EstadoActual
                        If .EstadoVotacion_y_PasList = "larga" Or .EstadoVotacion_y_PasList = "empate" Or .EstadoVotacion_y_PasList = "votando" Then
                            ' Apagar todos los teclados de las bancas
                            Mensaje2Banca.sTipo = "mset"
                            If EstadoActual.EstadoVotacion_y_PasList = "larga" Or EstadoActual.EstadoVotacion_y_PasList = "votando" Then 'hcdn 2011 - ap 110211
                                VL.log "--------------------------- Llamada 3 ---------------------------"
                                EstadoActual.EstadoVotacion_y_PasList = "cierre"
                                Call FinVotacionBrc("cancela operador")
                            End If
                            .EstadoVotacion_y_PasList = "cancelada"
                            Call AltaLogGeneral("Operador del sistema", "Cancelacion votacion. PL: " & EstadoActual.PeriodoLegislativo & " S " & EstadoActual.Sesion & " A " & EstadoActual.NroActa & " Estado " & EstadoActual.EstadoVotacion_y_PasList, , "1")
                        End If
                        If .EstadoVotacion_y_PasList = "inipas" Then
                            .EstadoVotacion_y_PasList = "canpas"
                        End If
                    End With
                    ' ------------------------------------------------------------------------
                    ' Usuario cambia el titulo de un acta
                    ' ------------------------------------------------------------------------
                Case Is = "cambio?tacta"
                    If Not IsNull(.sObjeto) Then ' .sComponente tiene el Id del titulo del acta
                        EstadoActual.strError = "cambio?tacta"
                        EstadoActual.TituloDelActa = rs.Fields("Parametro2").Value ' titulo del acta: Lo toma directamente de record set para mantener mayusculas.
                    End If
                    ' -------------------------------------------------------------------------------------
                    ' Usuario cambia el periodo legislativo
                    ' -------------------------------------------------------------------------------------
                Case Is = "cambioperiodo"
                    If Not IsNull(.sComponente) Then
                        EstadoActual.strError = "cambioperiodo"
                        EstadoActual.PeriodoLegislativo = .sComponente
                        Dim rsA As ADODB.Recordset
                        Set rsA = New ADODB.Recordset
                        strSql = "SELECT Ultima_Reunion From dbo.perparl Where Período_Legislativo = '" & EstadoActual.PeriodoLegislativo & "'"
                        frmMain.SetearRsAux strSql, rsA
                        If Not rsA.EOF Then
                            EstadoActual.Reunion = rsA.Fields(0)
                        End If
                        rsA.Close
                        Set rsA = Nothing
                        ' Buscar en Periodo Legislativo la ultima sesion
                        strSql = "SELECT Nro_de_Sesion_actual From dbo.perparl Where Período_Legislativo = '" & EstadoActual.PeriodoLegislativo & "'"
                        strSql = "SELECT max(Sesión) as maximo FROM sesion WHERE (Sesión <> 9999) AND (Sesión <> -1) AND Período_Legislativo = '" & .sComponente & "'"
                        Call SetearOtroRs(strSql)
                        EstadoActual.Sesion = IIf(IsNull(RsOtro.Fields(0).Value), 0, RsOtro.Fields(0).Value)
                        RsOtro.Close
                        ' Buscar en Sesion ultima acta
                        strSql = "SELECT Próximo_Acta, Estado_sesión From Sesion Where Sesión = '" & EstadoActual.Sesion & "' And Período_Legislativo = '" & EstadoActual.PeriodoLegislativo & "'"
                        Call SetearOtroRs(strSql)
                        If RsOtro.RecordCount = 0 Or rs.EOF = True Or rs.BOF = True Then
                            ' Dar de alta la sesion
                            RsOtro.Close
                            strSql = "SELECT Período_Legislativo, Sesión, Fecha_de_inicio, Próximo_Acta, Estado_sesión FROM sesion WHERE 0 = 1"
                            Call SetearRsW(strSql)
                            With RsWrite
                                .AddNew
                                .Fields("Período_Legislativo").Value = EstadoActual.PeriodoLegislativo
                                .Fields("Sesión").Value = EstadoActual.Sesion
                                .Fields("Fecha_de_inicio").Value = Date
                                .Fields("Próximo_Acta").Value = 1
                                .Fields("Estado_sesión").Value = "nueva"
                                .Update
                                EstadoActual.NroActa = 1
                                EstadoActual.EstadoSesion = "nueva"
                            End With
                        Else
                            EstadoActual.NroActa = RsOtro.Fields("Próximo_Acta").Value
                            EstadoActual.EstadoSesion = RsOtro.Fields("Estado_sesión").Value
                        End If
                    End If
                    ' ------------------------------------------------------------------------------------------------
                    ' Usuario abre la sesión actual
                    ' ------------------------------------------------------------------------------------------------
                Case Is = "abrirsesion"
                    strSql = "SELECT Estado_Sesión From Sesion WHERE período_legislativo = '" & EstadoActual.PeriodoLegislativo & "' AND sesión = '" & EstadoActual.Sesion & "'"
                    Call SetearRsW(strSql)
                    If Not (RsWrite.EOF And RsWrite.BOF) Then
                        If Trim(LCase(RsWrite.Fields("Estado_Sesión").Value)) = "nueva" Then
                            RsWrite.Fields("Estado_Sesión").Value = "abierta"
                            RsWrite.Update
                            EstadoActual.EstadoSesion = "abierta"
                        End If
                    End If
                    RsWrite.Close
                    ' ------------------------------------------------------------------------------------------------
                    ' Usuario corta el pase de lista
                    ' ------------------------------------------------------------------------------------------------
                Case Is = "accion?cortepaselista"
                    Imprimio = False
                    EstadoActual.ActaGrabada = 0
                    ActualizarVector_enBD
                    If EstadoActual.TipoDeOperacion = "paslis" And EstadoActual.EstadoVotacion_y_PasList = "espera" Then 'EstadoActual.EstadoVotacion_y_PasList <> "votando" And EstadoActual.EstadoVotacion_y_PasList <> "larga" And EstadoActual.EstadoVotacion_y_PasList <> "cierre" Then
                        EstadoActual.EstadoVotacion_y_PasList = "inipas"
                        Call AltaLogGeneral("Operador del sistema", "Pase de lista iniciado: " & Trim(EstadoActual.PeriodoLegislativo) & " sesion" & Trim(Str(EstadoActual.Sesion)) & " acta" & Trim(Str(EstadoActual.NroActa)))
                        lblGeneralInformacion.Caption = "INICIANDO PASE DE LISTA"
                        Dim tempTick As Long
                        tempTick = GetTickCount
                        Tick_InicioPasLis = GetTickCount
                    End If
                Case Is = "accion?borrarcacheacta"
                    EstadoActual.ActaGrabada = 0
                    ActualizarVector_enBD
                    ' Usuario actualiza los datos de las bancas
                Case Is = "accion?recargardatossb"
                    EstadoActual.strError = "accion?recargardatossb"
                    Call AltaLogGeneral("Consola", "accion?recargardatossb")
                Case Is = "habilitarconsola"
                    EstadoActual.strError = "habilitarconsola"
                    If EstadoActual.IP_Consola = "0" Then
                        EstadoActual.IP_Consola = .sAtributo
                    Else
                        EstadoActual.MensajeAlOperador = "La consola en la dirección IP " & EstadoActual.IP_Consola & " ya tiene el control de SQV"
                        EstadoActual.strError = "**error"
                    End If
                    ' Usuario libera una consola
                Case Is = "liberarconsola"
                    EstadoActual.strError = "liberarconsola"
                    If EstadoActual.IP_Consola = .sAtributo Then
                        EstadoActual.IP_Consola = "0"
                    End If
                    ' Usuario cancela la consola habilitada y habilita la suya propia
                Case Is = "cancelarconsola"
                    EstadoActual.strError = "cancelarconsola"
                    EstadoActual.IP_Consola = .sAtributo
                    ' Usuario solicita al SB el estado de una banca
                Case Is = "estadobanca"
                    EstadoActual.strError = "estadobanca"
                    xBanca = -1
                    If Not IsNull(.sObjeto) Then
                        If Trim(LCase(.sObjeto)) = "presidente" Then
                            xBanca = 0
                        ElseIf .sObjeto <> "" Then
                            xBanca = Int(.sObjeto)
                        End If
                    End If
                    'xBanca = IIf(Not IsNull(.sObjeto), IIf(.sObjeto <> "", (.sObjeto), -1), -1)
                    If xBanca >= 0 And (xBanca) <= xUltimaBanca Then
                        With Mensaje2Banca ' Mensaje para SB
                            .sTipo = "mget"
                            .sComponente = "term.mon"
                            .sObjeto = Str(xBanca)
                            .sAtributo = "state"
                            .sValor = ""
                        End With
                        Call EnviarMensajesBancas(Mensaje2Banca)
                    End If
                    'MsgBox "Solicitud de estado Banca"
                    ' Usuario resetea el SO de la terminal (por falla de componente)
                Case Is = "estadoioc"
                    If .sObjeto <> "brc" Then
                        EstadoActual.strError = "estadoioc"
                        xBanca = -1
                        If Not IsNull(.sObjeto) Then
                            If Trim(LCase(.sObjeto)) = "presidente" Then
                                xBanca = 0
                            ElseIf .sObjeto <> "" Then
                                xBanca = Int(.sObjeto)
                            End If
                        End If
                        'xBanca = IIf(Not IsNull(.sObjeto), IIf(Trim(LCase(.sObjeto)) = "presidente", 0, IIf(.sObjeto <> "", .sObjeto, -1)), -1)
                        If xBanca >= 0 And (xBanca) <= xUltimaBanca Then
                            With Mensaje2Banca ' Mensaje para SB
                                .sTipo = "mget"
                                .sComponente = "term.mon"
                                .sComentario = "CONSOLA ESTADO IOC"
                                .sObjeto = Str(xBanca)
                                .sAtributo = "action"
                                .sValor = "reset"
                            End With
                            Call EnviarMensajesBancas(Mensaje2Banca)
                        Else
                            EstadoActual.MensajeAlOperador = "Banca invalida"
                        End If
                    Else
                        With Mensaje2Banca ' Mensaje para SB
                            .sTipo = "mget"
                            .sComponente = "term.mon"
                            .sComentario = "CONSOLA RESET"
                            .sObjeto = "brc"
                            .sAtributo = "action"
                            .sValor = "reset"
                        End With
                        Call EnviarMensajesBancas(Mensaje2Banca)
                        End If
                    'MsgBox "Resetear SO de la banca " & Str(xActualBanca)
                    ' Usuario solicita al Sb se apague una banca y que no se le envíen más mensajes
                
                Case Is = "resethard"
                    EstadoActual.strError = "resethard"
                    xBanca = -1
                    If Not IsNull(.sObjeto) Then
                        If Trim(LCase(.sObjeto)) = "presidente" Then
                            xBanca = 0
                        ElseIf .sObjeto = "brc" Then
                            xBanca = 9999
                        ElseIf .sObjeto <> "" Then
                            xBanca = Int(.sObjeto)
                        End If
                    End If
                    'xBanca = IIf(Not IsNull(.sObjeto), IIf(Trim(LCase(.sObjeto)) = "presidente", 0, IIf(.sObjeto <> "", .sObjeto, -1)), -1)
                    If (xBanca >= 0 And (xBanca) <= xUltimaBanca) Or xBanca = 9999 Then
                        With Mensaje2Banca ' Mensaje para SB
                            .sTipo = "mget"
                            .sComponente = "term.mon"
                            .sObjeto = IIf(xBanca <= xUltimaBanca, Str(xBanca), "brc")
                            .sAtributo = "action"
                            .sValor = "resethard"
                        End With
                        Call EnviarMensajesBancas(Mensaje2Banca)
                    Else
                        EstadoActual.MensajeAlOperador = "Banca invalida"
                    End If
                Case Is = "actualizarips"
                        With Mensaje2Banca ' Mensaje para SB
                            .sTipo = "ips"
                            .sComponente = "update"
                            .sObjeto = "brc"
                            .sAtributo = "update"
                            .sValor = "update"
                        End With
                        Call EnviarMensajesBancas(Mensaje2Banca)
                Case Is = "apagarbanca"
                    EstadoActual.strError = "apagarbanca"
                    If Not IsNull(.sComponente) Then
                        xActualBanca = Int(.sComponente)
                        'MsgBox "Apagar banca " & Str(xActualBanca)
                        Call AltaLogGeneral("Consola", "apagarbanca " & xActualBanca)
                        Call PintarVectorColor(xActualBanca)
                    End If
                    ' Usuario solicita se limpie la cola de mensajes sqv->consola
                Case Is = "limpia"
                        Call AltaLogGeneral("Consola", "Limpia Definir regla de negocio")
                    ' Usuario decide realizar una grabación manual No implementado
                Case Is = "accion?grabarmanual"
                        Call AltaLogGeneral("Consola", "Usuario decide hacer grabacion manual")
                    ' Usuario decide pasar al modo de mantenimiento
                Case Is = "cambio?mantenimiento"
                    PresidenteEstuvoMantenimiento = False
                    Me.lblMantenimientostrPanel1.Caption = ""
                    Me.lblMantenimientostrPanel2.Caption = ""
                    Me.lblMantenimientostrPanel3.Caption = ""
                    EstadoActual.strError = "cambio?mantenimiento"
                    If EstadoActual.ModoMantenimientoBancas = 0 Then
                        frmCartel2011.Visible = False
                        frmMain.Show
                        Call Mantenimiento_SQV
                        ModoMant = True
                    Else
                        frmMain.Visible = False
                        frmCartel2011.Visible = True
                        Call Fin_Mantenimiento_SQV
                        ModoMant = False
                        EstadoActual.CartelEncendido = 2
                        EstadoActual.strError = "cancelarconsola"
                    End If
                    'MsgBox "Pasar a modo mantenimiento"
                    ' Usuario decide pasar al modo normal con mantenimiento
                    ' En este modo se puede operar el sistema en forma normal pero se valida la identificacion con registros de personas definidas como Personal de Mantenimiento (tipo = 0)
                ' ---------------------------------------------------------------------------------
                ' Usuario cancela las identificaciones
                ' ---------------------------------------------------------------------------------
                Case Is = "cambio?cancelarids"
                        For i = 1 To (xUltimaBanca)
                            EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO
                            EstadoActual.VectorColor(i) = AsignarColor(i)
                            Call EnviarMensajesFinAuth(Str(i), "Cancelar Ids - Fin modo habilitar scanners")
                        Next i
                        'EstadoACtual.Modo_Ident_Nom = 0
                        If EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis" Or EstadoActual.Modo_Ident_Nom Then
                            xStrVector = "0" & SEPARADOR_VECTOR 'presidente
                            For i = 1 To UBound(EstadoActual.VectorPresencia)
                                xStrVector = xStrVector & IIf(EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO And EstadoActual.VectorPresencia(i) <> AUSENTE, "1", "0") & SEPARADOR_VECTOR
                            Next i
                            Call EnviarMensajesComienzoAuth(xStrVector, "Comienzo modo nominal desde " & EstadoActual.TipoDeOperacion & "a " & .sComponente)
                        End If
                Case Is = "cambio?forzarids"
                        xStrVector = "0" & SEPARADOR_VECTOR 'presidente
                        For i = 1 To UBound(EstadoActual.VectorPresencia)
                            'xStrVector = xStrVector & IIf(EstadoActual.VectorPresencia(i) = "1", "1", "0") & SEPARADOR_VECTOR
                            xStrVector = xStrVector & IIf(EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO And EstadoActual.VectorPresencia(i) <> AUSENTE, "1", "0") & SEPARADOR_VECTOR
                        Next i
                        Call EnviarMensajesComienzoAuth(xStrVector, "Comienzo modo nominal desde " & EstadoActual.TipoDeOperacion & "a " & .sComponente)
                        EstadoActual.Modo_Ident_Nom = 1
                Case Is = "cambio?forzaroffids"
                        xStrVector = "0" & SEPARADOR_VECTOR 'presidente
                        For i = 1 To UBound(EstadoActual.VectorPresencia)
                            'xStrVector = xStrVector & IIf(EstadoActual.VectorPresencia(i) = "1", "1", "0") & SEPARADOR_VECTOR
                            xStrVector = xStrVector & IIf(EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO, "1", "0") & SEPARADOR_VECTOR
                        Next i
                        Call EnviarMensajesFinAuth(xStrVector, "Fin mod<o nominal desde " & EstadoActual.TipoDeOperacion & "a " & .sComponente)
                        EstadoActual.Modo_Ident_Nom = 0
                Case Is = "cambio?nominal" 'Habilitar identificación (sin Presencia con identificación)
                    EstadoActual.strError = "cambio?nominal"
                    EstadoActual.Modo_Presencia_Nom = 0 'Para poder permitir presencia nominal es necesario aplicar el comando cambio?presenciaidentificacion. En todo caso, al seleccionar habilitar identificacion, se deja de contar presencia si, es que antes se hubiera solicitado
                    If EstadoActual.Modo_Ident_Nom = 0 Then
                        EstadoActual.Modo_Ident_Nom = 1
                        Call SolicitarIdentificacionPendientes("Comienzo modo nominal desde " & EstadoActual.TipoDeOperacion & "a " & .sComponente, "start")
                    Else
                        EstadoActual.Modo_Ident_Nom = 0
                        For i = 1 To (xUltimaBanca)
                            EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO
                            Call EnviarMensajesFinAuth(Str(i), "Fin modo habilitar scanners")
                        Next i
                        'Call Fin_Mantenimiento_SQV 'revisar si es necesario tambien enviarlo al setear este comando a 1
                    End If
                Case Is = "cambio?presenciaidentificacion" 'Habilitar identificación con Presencia con identificación: se cuenta como presente sólo si esta identificado (esto afecta al calculo de quorum).
                    'verificar
                    EstadoActual.strError = "cambio?presenciaidentificacion"
                    EstadoActual.Modo_Ident_Nom = 1 'Encendiendo o apagando presencia con identificacion, se mantiene el modo de identificación.
                    If EstadoActual.Modo_Presencia_Nom = 0 Then
                        EstadoActual.Modo_Presencia_Nom = 1
                    Else
                        EstadoActual.Modo_Presencia_Nom = 0
                    End If
                    Call SolicitarIdentificacionPendientes("Comienzo modo nominal desde " & EstadoActual.TipoDeOperacion & "a " & .sComponente, "start")
                    EstadoActual.OcupadosNoIdentificados = CuentaOcupadosNoIdentificadosCong  'EstadoActual.Presentes - IIf(xPresidenteLegislador, 1, 0) 'aca5
                    'Para poder permitir presencia nominal es necesario armar el control en la consola y actualizarlo con 1 o 0
                    'Call Fin_Mantenimiento_SQV ' Se elimino 091031 porque sino perderia los id...
                Case Is = "cambio?normalmantenimiento"
                    EstadoActual.strError = "cambio?normalmantenimiento"
                    If EstadoActual.ModoNormalMantSistema = 0 Then
                        Call NormalMantenimiento_SQV
                    Else
                        Call Fin_Mantenimiento_SQV
                    End If
                    'MsgBox "Pasar a modo mantenimiento"
                    ' Usuario cambia el Modo de presentacion de informacion fija
                Case Is = "cambio?formulario"
                    Call AltaLogGeneral("Consola", "Usuario cambia el Modo de presentacion de informacion fija")
                    ' Usuario reinicia el Sistema
                Case Is = "accion?reiniciosistema"
                    EstadoActual.strError = "reiniciosistema"
                    Call ReinicioSistema
                    'MsgBox "Usuario reinicia el Sistema"
                    ' Usuario reinicia el Sistema
                Case Is = "reiniciarbancas"
                    EstadoActual.strError = "reiniciarBancas"
                    Call ReinicioSistema
                    'MsgBox "Usuario reinicia el Sistema"
                     ' Usuario reinicia SQV en modo mantenimiento
                Case Is = "accion?reiniciarserverman"
                    Call AltaLogGeneral("Consola", "Usuario reinicia SQV en modo mantenimiento")
                    ' 'MsgBox "Usuario reinicia SQV en modo mantenimiento"
                    ' Usuario sale del SQV
                Case Is = "accion?salirserver"
                    EstadoActual.strError = "accion?salirserver"
                    Call AltaLogGeneral("Saliendo de SQV Server por solicitud Consola", Now)
                    Call Salir_SQV
                    ' Usuario reinicia el server de quorum inmediatamente
                Case Is = "accion?inicioconsola"
                    Call AltaLogGeneral("Consola", "Usuario reinicia el server de quorum inmediatamente")
                    'MsgBox "Usuario reinicia el server de quorum inmediatamente"
                    ' Usuario consulta un acta grabada
                Case Is = "mostrar?periodo"
                    EstadoActual.strError = "mostrar?periodo"
                    Call AltaLogGeneral("Consola", "Usuario consulta un acta grabada")
                    'MsgBox "Usuario consulta un acta grabada"
                    ' Usuario cambia la modalidad de grabación automática
                Case Is = "cambio?grabar"
                    Call AltaLogGeneral("Consola", "Usuario cambia la modalidad de grabación automática")
                    'MsgBox "Usuario cambia la modalidad de grabación automática"
                    ' Usuario cambia la modalidad de lista automática
                Case Is = "cambio?listar"
                    EstadoActual.strError = "cambio?listar"
                    If EstadoActual.ListarAutomaticamente = 0 Then
                        EstadoActual.ListarAutomaticamente = 1
                    Else
                        EstadoActual.ListarAutomaticamente = 0
                    End If
                    'MsgBox "Usuario cambia la modalidad de lista automática"
                    ' Usuario reinicia la votación
                Case Is = "accion?reiniciovotacion"
                    Dim z As Integer
                    sinIdentificarCongelado = False
                    Pendientes.paginaActualPendientes = 0
                    For z = 1 To 256
                        If (Not bancaValida(z)) Then 'Si formo parte de la vot
                            If EstadoActual.VectorPresencia(z) <> VL.PresenciaReal(z) Then
                                EstadoActual.VectorPresencia(z) = VL.PresenciaReal(z)
                            End If
                            If VL.PerdioIdentificacion(z) = True Then
                                EstadoActual.VectorIdentificacion(z) = NO_IDENTIFICADO
                            End If
                        End If
                    Next z
                    VL.modoExtendido = False
                    PrimerControlLarga = False
                    ultimoResultadoEvaluado = ""
                    If EstadoActual.TipoDeOperacion = "paslis" Then
                        EstadoActual.EstadoVotacion_y_PasList = "espera"
                    Else
                        If (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "cancelada") Then
                            '>> REINCIO DE VOTACION. CUANDO EL OPERADOR PRESIONA EL BOTON INICIALIZAR, SE PREPARA EL SISTEMA PARA UNA NUEVA VOTACION.
                            '>> en el caso de que la ultima haya sido una votacion de reconsideracion (nominal), a todos los presentes no identificados les permitira identificars enuevamente
                            If EstadoActual.TipoDeOperacion = "votnom" Then
                                xStrVector = ""
                                For i = 0 To xUltimaBanca
                                    If EstadoActual.VectorPresencia(i) = PRESENTE And Val(EstadoActual.VectorIdentificacion(i)) = NO_IDENTIFICADO Then
                                        xStrVector = xStrVector & "1" & SEPARADOR_VECTOR
                                    Else
                                        xStrVector = xStrVector & "0" & SEPARADOR_VECTOR
                                    End If
                                Next i
                                Call EnviarMensajesComienzoAuth(xStrVector, "Permitir identificacion tras reconsideracion")
                            End If 'fin si es reconsideracion
                            Call DeshabilitarVotoPresidente ' Se cancela la autorizacion para votar. en votacion autorizada por operador funciona como nominal, mientras que por empate funciona como numerica.
                            If (InStr("votnom;votnum", Trim(EstadoActual.TipoDeOperacion)) > 0) Then
                                Call InicializarVotacion
                                'Call AbstenerVector(SEPARADOR_VECTOR, 0, 0, 0)
                                Call AbstenerVector(Join(EstadoActual.VectorAbstencion, SEPARADOR_VECTOR), 0, 0, 0)
                            End If
                            
                            EstadoActual.LimpiarResultados = 1
                            Call PintarTodasLasBancas
                            EstadoActual.ActaGrabada = 0
                            
                            Mensaje2Banca.sTipo = "mset"
                            Mensaje2Banca.sObjeto = "brc"
                            Mensaje2Banca.sComponente = "term.keyb"
                            Mensaje2Banca.sAtributo = "state"
                            Mensaje2Banca.sValor = "off" & EstadoActual.TipoDeOperacion
                            Mensaje2Banca.sComentario = EstadoActual.EstadoVotacion_y_PasList & "Inicializacion de votacion por el operador"
                            Call EnviarMensajesBancas(Mensaje2Banca)
                            'caso del presidente
                            
                            'Apaga leds teclado para todos (inc al presidente que queda prendido si hubo empate)
                            
                            Mensaje2Banca.sTipo = "mset"
                            Mensaje2Banca.sObjeto = "brc"
                            Mensaje2Banca.sComponente = "term.ledk1"
                            Mensaje2Banca.sAtributo = "state"
                            Mensaje2Banca.sValor = "off"
                            Mensaje2Banca.sComentario = EstadoActual.EstadoVotacion_y_PasList
                            Call EnviarMensajesBancas(Mensaje2Banca)
                                                    
                            'MsgBox "Usuario reinicia la votación"
                        End If
                        ' Usuario inicia votación de reconsideración
                    End If
                Case Is = "accion?votacionespecial"
                    EstadoActual.strError = "accion?votacionespecial"
                    If EstadoActual.TipoDeOperacion = "votnom" And EstadoActual.EstadoVotacion_y_PasList = "espera" Then
                        If Not IsNull(.sComponente) Then
                            xSesionReconsideracion = Int(.sComponente)
                            If Not IsNull(.sObjeto) Then
                                xActaReconsideracion = Int(.sObjeto)
                            Else
                                xActaReconsideracion = 0
                                EstadoActual.MensajeAlOperador = "Selección de votacion de reconsideracion: Acta invalida"
                                EstadoActual.strError = "**error"
                            End If
                        Else
                            xSesionReconsideracion = 0
                            EstadoActual.MensajeAlOperador = "Selección de votacion de reconsideracion: Sesion invalida"
                            EstadoActual.strError = "**error"
                        End If
                        If xSesionReconsideracion > 0 And xActaReconsideracion > 0 Then
                            'Call MostrarActaProyector(xSesionReconsideracion, xActaReconsideracion, 0)
                            If ArmarHabilitadosDeActa(xSesionReconsideracion, xActaReconsideracion) Then
                                'Armado de habilitados exitoso. Se cancelan a los no habilitados
                                
                            Else
                                EstadoActual.MensajeAlOperador = "Votacion de reconsideracion: No se pudo obtener los datos de la sesion " & Str(xSesionReconsideracion) & " acta Nro. " & Str(xActaReconsideracion) & " en el periodo legislativo actual " & EstadoActual.PeriodoLegislativo
                                EstadoActual.strError = "**error"
                            End If
                        End If
                    Else
                        EstadoActual.MensajeAlOperador = "La selección de votacion de reconsideracion debe realizarse antes de iniciar una votación"
                        EstadoActual.strError = "**error"
                    End If
                        ' Usuario cambia el voto de una banca
                Case Is = "cambiovoto"
                    EstadoActual.strError = "cambiovoto"
                    xBanca = -1
                    If Not IsNull(.sObjeto) Then
                        If Trim(LCase(.sObjeto)) = "presidente" Then
                            xBanca = 0
                        ElseIf .sObjeto <> "" Then
                            xBanca = Int(.sObjeto)
                        End If
                    End If
                    If Not IsNull(.sComponente) Then
                        xVotoOperador = Left(LCase(MensajeActual.sComponente), 1)
                        If xVotoOperador = "a" Then
                            xVotoOperador = ABSTENCION
                        End If
                    Else
                        xVotoOperador = "x"
                    End If
                    If (xBanca >= 0 And InStr("s n", xVotoOperador) > 0) And _
                        (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum") And _
                         (EstadoActual.EstadoVotacion_y_PasList = "votando" Or EstadoActual.EstadoVotacion_y_PasList = "larga" Or EstadoActual.EstadoVotacion_y_PasList = "empate") Then
                        'Filtro de banca 0 y mayor a leg
                        'If xBanca >= IIf(xPresidenteLegislador Or EstadoActual.EstadoVotacion_y_PasList = "empate", 0, 1) And xBanca <= xUltimaBanca Then
                        If xBanca >= IIf(EstadoActual.PresidenteHabilitadoParaVotar Or EstadoActual.EstadoVotacion_y_PasList = "empate", 0, 1) And xBanca <= xUltimaBanca Then
                            'Si es nominal, ver que este identificado, sino solo que esté presente
                            'Si es votacion larga, tambien se puede habiiltar para votar
                            'If (EstadoActual.VectorPresencia(xBanca) = PRESENTE Or (xBanca = 0 And (xPresidenteLegislador Or EstadoActual.EstadoVotacion_y_PasList = "empate"))) And _
                            '
                            
                            If (EstadoActual.VectorPresencia(xBanca) = PRESENTE Or (xBanca = 0 And (EstadoActual.PresidenteHabilitadoParaVotar Or EstadoActual.EstadoVotacion_y_PasList = "empate"))) And _
                                    (EstadoActual.TipoDeOperacion = "votnum" Or _
                                     (EstadoActual.TipoDeOperacion = "votnom" And EstadoActual.VectorIdentificacion(xBanca) <> NO_IDENTIFICADO) _
                                    ) Then
                                ' los abstencionistas no los cuento
                                If LCase(EstadoActual.VectorResultados(xBanca)) <> ABSTENCION_AUTORIZADA Then
                                    'objeto == term.keyb
                                    If (xVotoOperador = AFIRMATIVO Or xVotoOperador = NEGATIVO Or xVotoOperador = ABSTENCION) Then
                                        'Actualiza pendientes de votar si antes no habia votado
                                        If LCase(EstadoActual.VectorResultados(xBanca)) = ABSTENCION Then
                                            EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1
                                            Call AltaLogGeneral("SQV SERVER: consola cambiovoto 1", "If LCase(EstadoActual.VectorResultados(xBanca)) = ABSTENCION Then EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1: " & EstadoActual.PendientesEmitirVotos, 0, "0")
                                        End If
                                        
                                        'deshacer voto anterior
                                        'No apago las luces del voto anterior pues lo hace la banca automaticamente
                                        Select Case LCase(EstadoActual.VectorResultados(xBanca))
                                            Case AFIRMATIVO
                                                CartelActual.Afirmativos = CartelActual.Afirmativos - 1
                                            Case NEGATIVO
                                                CartelActual.Negativos = CartelActual.Negativos - 1
                                        End Select
                                        'Aplica voto nuevo, y prende la luz del acknowledge
                                        If xVotoOperador = ABSTENCION And xBanca = 0 And EstadoActual.TipoDeAbstencion = "absaut" Then
                                            EstadoActual.VectorResultados(0) = "AP"
                                        End If
                                        If xVotoOperador = AFIRMATIVO Then
                                            CartelActual.Afirmativos = CartelActual.Afirmativos + 1
                                            EstadoActual.VectorResultados(xBanca) = AFIRMATIVO
                                            Mensaje2Banca.sTipo = "mset"
                                            Mensaje2Banca.sObjeto = xBanca
                                            Mensaje2Banca.sComponente = "term.ledk1"
                                            Mensaje2Banca.sAtributo = "state"
                                            Mensaje2Banca.sValor = "on"
                                            Mensaje2Banca.sComentario = EstadoActual.EstadoVotacion_y_PasList
                                            Call EnviarMensajesBancas(Mensaje2Banca)
                                            If xBanca = 0 And EstadoActual.EstadoVotacion_y_PasList = "empate" Then
                                                xHuboDesempate = True
                                                CartelActual.Resultado = "AFIRMATIVO"
                                            End If
                                        ElseIf xVotoOperador = NEGATIVO Then
                                            CartelActual.Negativos = CartelActual.Negativos + 1
                                            EstadoActual.VectorResultados(xBanca) = NEGATIVO
                                            Mensaje2Banca.sTipo = "mset"
                                            Mensaje2Banca.sObjeto = xBanca
                                            Mensaje2Banca.sComponente = "term.ledk2"
                                            Mensaje2Banca.sAtributo = "state"
                                            Mensaje2Banca.sValor = "on"
                                            Mensaje2Banca.sComentario = EstadoActual.EstadoVotacion_y_PasList
                                            Call EnviarMensajesBancas(Mensaje2Banca)
                                            If xBanca = 0 And EstadoActual.EstadoVotacion_y_PasList = "empate" Then
                                                xHuboDesempate = True
                                                CartelActual.Resultado = "NEGATIVO"
                                            End If
                                            
                                        ElseIf xVotoOperador = ABSTENCION Then
                                            If xBanca > 0 Or (xBanca = 0 And Not EstadoActual.TipoDeAbstencion = "absaut") Then
                                                EstadoActual.VectorResultados(xBanca) = ABSTENCION
                                            End If
                                            Mensaje2Banca.sTipo = "mset"
                                            Mensaje2Banca.sObjeto = xBanca
                                            Mensaje2Banca.sComponente = "term.keyb"
                                            Mensaje2Banca.sAtributo = "state"
                                            If xBanca = 0 Then
                                                Mensaje2Banca.sValor = "offvotnom"
                                            Else
                                                Mensaje2Banca.sValor = "off" & EstadoActual.TipoDeOperacion
                                            End If
                                            Mensaje2Banca.sComentario = EstadoActual.EstadoVotacion_y_PasList
                                            Call EnviarMensajesBancas(Mensaje2Banca)
                                            EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos + 1
                                            Call AltaLogGeneral("SQV SERVER: consola cambiovoto 2", "EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos + 1: " & EstadoActual.PendientesEmitirVotos, 0, "0")
                                            If xBanca = 0 And EstadoActual.EstadoVotacion_y_PasList = "empate" Then
                                                xHuboDesempate = True
                                            End If
                                            'siiiiii
                                        End If 'MensajeActual.sComponente = "term.keyb.si o no
                                        PintarVectorColor (xBanca)
                                        If (xVotoOperador = AFIRMATIVO Or xVotoOperador = NEGATIVO Or xVotoOperador = ABSTENCION) And xBanca = 0 And EstadoActual.EstadoVotacion_y_PasList = "empate" Then
                                            EstadoActual.ActaGrabada = 0
                                            EstadoActual.SolicitudGrabarManual = 0

                                            EstadoActual.EstadoVotacion_y_PasList = "finalizada"
                                        End If
                                    End If 'MensajeActual.sComponente = "term.keyb.si" Or MensajeActual.sComponente = "term.keyb.no" And MensajeActual.sAtributo = "STATE" And MensajeActual.sComponente.sValor = "ON"
                                    Call AltaLogGeneral("SQVC", "Voto " & xVotoOperador & " por Consola Operacion en banca " & xBanca, Str(xBanca))
                                Else
                                    'Call AltaLogGeneral("SQVB", "Intento de voto por Operador con abstencion autorizada" & xBanca) 'aca3
                                    EstadoActual.MensajeAlOperador = "Intento de voto por Operador con abstencion autorizada"
                                    EstadoActual.strError = "**error"
                                End If 'LCase(EstadoActual.VectorResultados(xBanca)) <> absaut
                            Else
                                EstadoActual.MensajeAlOperador = "Intento de voto por Operador sin presencia."
                                EstadoActual.strError = "**error"
                            End If 'EstadoActual.VectorPresencia(xBanca) = PRESE...
                        Else
                            EstadoActual.MensajeAlOperador = "Intento de voto por Operador con banca y/o estado de votacion invalido."
                            EstadoActual.strError = "**error"
                        End If 'xBanca >= IIf(xPresidenteLegislador Or EstadoActual.EstadoVotacion_y_PasList = "empate", 0, 1) And xBanca <= xUltimaBanca Then
                    Else
                        EstadoActual.MensajeAlOperador = "Intento de voto por Operador con estado de votacion invalido."
                        EstadoActual.strError = "**error"
                    End If ' (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum") And
                    'MsgBox "Usuario cambia el voto de una banca"
                    'MsgBox "Usuario inicia votación de reconsideración"
                    ' Usuario cambia el identificador de una banca
                Case Is = "cambioid"
                    EstadoActual.strError = "cambioid"
                    xBanca = -1
                    If Not IsNull(.sObjeto) Then
                        If Trim(LCase(.sObjeto)) = "presidente" Then
                            xBanca = 0
                        ElseIf .sObjeto <> "" Then
                            xBanca = Int(.sObjeto)
                        End If
                    End If
                    If Not IsNull(.sComponente) Then
                        xNuevoID = LCase(MensajeActual.sComponente)
                    Else
                        xNuevoID = "x"
                    End If
                    If xBanca >= 1 And xBanca <= xUltimaBanca And xNuevoID <> "x" Then ' Solo procesa mensajes de identificacion, cuando la SB censo presencia.
                        If EstadoActual.VectorPresencia(xBanca) = PRESENTE Then
                            If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then ' y no estamos en modo Mantenimiento
                                ' Hay que verificar que el legislador se encuentre en al tabla de legisladores activos
                                strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                       & "Legisladores.departamento , Legisladores.cargo FROM Legisladores INNER JOIN legisladores_activos ON " _
                                       & "Legisladores.id = legisladores_activos.id WHERE (Legisladores.id = '" & xNuevoID & "' and Legisladores.tipo = 1)" '<AP 040115 Pide que sea tipo legislador
                            Else ' Si se encuentra en modo mantenimiento
                                ' IDs recibidos = TRIM (.sValor) & ';' & TRIM (IDs recibidos)
                                ' <AP 040115 Hay que verificar que el legislador se encuentre como personal de mantenimiento, sin join
                                strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                       & "Legisladores.departamento , Legisladores.cargo FROM Legisladores WHERE (Legisladores.id = '" & xNuevoID & "' and Legisladores.tipo = 0)" '<AP 040115 Pide que sea tipo personal de mantenimiento
                            End If
                            RsLocal.Open strSql, Cn, adOpenForwardOnly, adLockOptimistic
                            If RsLocal.RecordCount <= 0 Or RsLocal.EOF Or RsLocal.BOF Then ' Si NO es un legislador activo o personal de mant.
                                'xNuevoID = ""
                                EstadoActual.MensajeAlOperador = "Cambio ID por Operador invalido " & xNuevoID
                                EstadoActual.strError = "**error"
                            Else ' En cambio, SI ES un legislador activo o personal de mant.
                                For X = 1 To xUltimaBanca ' y me fijo si se identifico anteriormente en otra banca
                                    If Trim(LCase(EstadoActual.VectorIdentificacion(X))) = xNuevoID Then
                                        flIdDupOperador = True
                                        EstadoActual.MensajeAlOperador = "El diputado ya esta identificado en la banca " & Str(X)
                                        EstadoActual.strError = "**error"
                                        'xNuevoID = ""
                                    End If
                                Next X
                                ' Si no se i dentifico anteriormente en otra banca, pongo ID de legislador en Vector Identificacion
                                If flIdDupOperador = False Then ' identificar al legislador en vector identificacion
                                    'Verifica que no coincida con el presidente
                                    If Trim(LCase(EstadoActual.VectorIdentificacion(0))) = xNuevoID Then
                                        flIdDupOperador = True
                                        EstadoActual.MensajeAlOperador = "Cambio ID por Operador. Identificado como Presidente." & xNuevoID
                                        EstadoActual.strError = "**error"
                                    Else
                                        '>> Verifica que este habilitado para el caso de las votaciones de reconsideracion.
                                        ' En las otras situaciones siempre estan todos habilitados.
                                        If (True Or LegisladorHabilitado(xNuevoID)) Then
                                            'La identificacion ha sido exitosa!
                                            If "SAUTOD" = "SIN TACKNL" Then 'si no se cuenta con acknowledge desde la banca, se identifica directa...
                                                EstadoActual.VectorIdentificacion(xBanca) = xNuevoID
                                                EstadoActual.OcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados - 1
                                                
                                                Call EnviarMensajesFinAuth(Str(xBanca), "Autenticacion Operador")
                                            End If
                                            
                                            If EstadoActual.VectorIdentificacion(xBanca) = "0" Then
                                                Mensaje2Banca.sObjeto = xBanca '<AP 040115 faltaba indicar la banca>
                                                Mensaje2Banca.sTipo = "mset"
                                                Mensaje2Banca.sComponente = "term.led1"
                                                Mensaje2Banca.sAtributo = "state"
                                                Mensaje2Banca.sValor = "on_manual|" & Trim(xNuevoID)
                                                EstadoActual.VectorIdentificacion(xBanca) = xNuevoID
                                                Call ActualizarVector_enBD
                                                If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
                                                    Mensaje2Banca.sComentario = "Id aceptado Modo normal - led1 - 1"
                                                Else
                                                    Mensaje2Banca.sComentario = "Id aceptado Modo mantenimiento - led1"
                                                End If
                                                Call EnviarMensajesBancas(Mensaje2Banca)
                                            Else
                                                EstadoActual.MensajeAlOperador = "La banca se identificó mientras se buscaba el diputado en la asignación manual. Banca: " & Str(X) & " " & Now()
                                                EstadoActual.strError = "**error"
                                            End If
                                            If "SINFOR" = "HABILITADO" Then 'Si la banca acepta mostrar en el display los datos del legislador
                                                Mensaje2Banca.sObjeto = xBanca '<AP 040115 faltaba indicar la banca>
                                                Mensaje2Banca.sTipo = "mset"
                                                If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
                                                    Mensaje2Banca.sComponente = "term.display"
                                                    Mensaje2Banca.sAtributo = "text"
                                                    Mensaje2Banca.sValor = "Identificacion Aceptada"
                                                    Mensaje2Banca.sComentario = "Id aceptado Modo normal - display - manual"
                                                Else
                                                    Mensaje2Banca.sValor = "Identificacion de Prueba"
                                                    Mensaje2Banca.sComentario = "Id aceptado Modo mantenimiento - manual"
                                                End If
                                                Call EnviarMensajesBancas(Mensaje2Banca)
                                            End If
                                            If "SAUTOD" = "SIN TACKNL" Then 'si no se cuenta con acknowledge desde la banca, se identifica directa...
                                                Call PintarVectorColor(xBanca)
                                                'Hubo identificacion positiva
                                                'ver si hay que habilitarlo para votar
                                                If EstadoActual.TipoDeOperacion = "votnom" Then
                                                        If InStr(Join(EstadoActual.VectorAbstencion, SEPARADOR_VECTOR), Trim(xNuevoID)) > 0 Then
                                                            AbstenerBanca (xBanca)
                                                        ElseIf EstadoActual.EstadoVotacion_y_PasList = "votando" Or EstadoActual.EstadoVotacion_y_PasList = "larga" Then
                                                            Mensaje2Banca.sTipo = "mset"
                                                            Mensaje2Banca.sObjeto = xBanca
                                                            Mensaje2Banca.sComponente = "term.keyb"
                                                            Mensaje2Banca.sAtributo = "state"
                                                            Mensaje2Banca.sValor = "on" & IIf(xBanca > 0, EstadoActual.TipoDeOperacion, "votnum")
                                                            Mensaje2Banca.sComentario = EstadoActual.EstadoVotacion_y_PasList
                                                            Call EnviarMensajesBancas(Mensaje2Banca)
                                                    End If
                                                End If
                                            End If
                                        Else
                                            EstadoActual.MensajeAlOperador = "Cambio ID por Operador. El Legislador solicitado no esta habilitado para esta votacion de reconsideracion." & xNuevoID
                                            EstadoActual.strError = "**error"
                                            ' Ver con Alejandro que paso aca:
                                            ' Call AltaLogGeneral("Identificacion Operador", "Id no habilitado: " & strIdLegislador & ", Banca " & xBanca)
                                        End If
                                    End If
                                End If ' ya esta identificado, error ya se dio
                            End If
                            RsLocal.Close
                        Else 'no esta presente
                            EstadoActual.MensajeAlOperador = "Para asignar la identificacion del Legislador debe verificarse presencia en la banca. " & Chr(10) & "Nro. Identificacion solicitado " & xNuevoID
                            EstadoActual.strError = "**error"
                        End If
                    Else
                        EstadoActual.MensajeAlOperador = "Cambio ID por Operador. Banca Invalida." & xNuevoID & " Banca: " & Str(xBanca)
                        EstadoActual.strError = "**error"
                        'enviar error
                    End If 'banca invalida
                    ' Mensaje de error mostrado en la consola, limpiar
                Case Is = "mensajemostrado"
                    'Usado para limpiar el mensaje al operador EstadoActual.strError = "**error"
                    EstadoActual.strError = "mensajemostrado"
                    EstadoActual.MensajeAlOperador = ""
                    ' Usuario cambia el identificador de una banca
                Case Is = "abstener"
                    EstadoActual.strError = "abstener"
                    xBanca = -1
                    If Not IsNull(.sObjeto) Then
                        If Trim(LCase(.sObjeto)) = "presidente" Then
                            xBanca = 0
                        ElseIf .sObjeto <> "" Then
                            xBanca = Int(.sObjeto)
                        End If
                    End If
                    If xBanca >= 0 Then ' nro. de banca
                        ' If LCase(EstadoActual.TipoDeOperacion) = "votnom" And InStr("votando larga espera", EstadoActual.EstadoVotacion_y_PasList) > 0 Then
                        If (LCase(EstadoActual.TipoDeOperacion) = "votnom" Or LCase(EstadoActual.TipoDeOperacion) = "votnum") And InStr("votando larga espera", EstadoActual.EstadoVotacion_y_PasList) > 0 Then
                            'xBanca = Int(.sObjeto)
                            ' el presidente no puede abstenerse
                            If xBanca >= 1 And (xBanca) <= xUltimaBanca And EstadoActual.VectorPresencia(xBanca) = PRESENTE And Not (EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO) Then
                                If EstadoActual.VectorResultados(xBanca) = ABSTENCION Then
                                    Call AbstenerBanca(xBanca)
                                ElseIf EstadoActual.VectorResultados(xBanca) = ABSTENCION_AUTORIZADA Then
                                    Call CancelarAbstenerBanca(xBanca)
                                Else
                                    EstadoActual.strError = "**error"
                                    EstadoActual.MensajeAlOperador = "Solo puede abstenerse si no ha emitido voto aun. Banca: " & xBanca
                                End If
                            Else
                                EstadoActual.strError = "**error"
                                If Not EstadoActual.VectorPresencia(xBanca) = PRESENTE Then
                                    EstadoActual.MensajeAlOperador = "Banca sin presencia: " & xBanca
                                ElseIf (EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO) Then
                                    EstadoActual.MensajeAlOperador = "Legislador no identificado en la banca: " & xBanca
                                Else
                                    EstadoActual.MensajeAlOperador = "Banca invalida: " & xBanca
                                End If
                            End If
                        Else
                            EstadoActual.strError = "**error"
                            EstadoActual.MensajeAlOperador = "Solicitud de abstencion fuera de votacion nominal a iniciar o en curso"
                        End If
                        lblPendientesEmitirVotos = EstadoActual.PendientesEmitirVotos
                        lblAbsAut.Caption = EstadoActual.AbstencionistasAutorizados
                        lblOcupadosNoIdentificados(0) = EstadoActual.OcupadosNoIdentificados
                    End If
                    ' Usuario consulta DB utilizada
                Case Is = "configuradoaccion?consultasufijo"
                    Call AltaLogGeneral("Consola", "Usuario consulta DB utilizada")
                    'MsgBox "Usuario consulta DB utilizada"
                    ' Usuario prende luces
                Case Is = "accion?prenderluces"
                    Call AltaLogGeneral("Consola", "Usuario prende luces")
                    'MsgBox "Usuario prende luces"
                    ' Usuario Apaga luces
                Case Is = "accion?apagarluces"
                    Call AltaLogGeneral("Consola", "Usuario Apaga luces")
                    'MsgBox "Usuario Apaga luces"
                Case Is = "cambio?modovotapresidente" 'Habilitar identificación con Presencia con identificación: se cuenta como presente sólo si esta identificado (esto afecta al calculo de quorum).
                    'verificar
                    EstadoActual.strError = "cambio?modovotapresidente"
                    If EstadoActual.ModoVotaPresidente = 0 Then
                        EstadoActual.ModoVotaPresidente = 1
                            Call SolicitarHabilitarVotoPresidente
                        Else
                            Call DeshabilitarVotoPresidente
                    End If
            End Select
            ' si se acaba el tiempo de votacion, decirle a SB que cancele las bancas
            If EstadoActual.EstadoVotacion_y_PasList = "votando" Then
                If tFinVotacion < Now Then
                    'FinVotacionBrc ("fin tiempo")
                End If
            End If
        End With
    ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    ' Actualizar valores antes de buscar siguiente mensaje
    ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
    xUltimoMensajeCosola = rs.Fields("serial").Value
    .MoveNext
    Wend
End With
    
    ' ------------------------------------------------------------------------------------------
    ' Atender a todos los mensajes nuevos emitidos por el servidor de banca
    ' ------------------------------------------------------------------------------------------
    If ModoMant = False Then
        'Call EjecutarSQL("DELETE FROM sb_sqv_mensajes WHERE id < " & Str(xUltimoMensajeSB))
    End If
    strSql = "SELECT * FROM sb_sqv_mensajes WHERE id > " & Str(xUltimoMensajeSB)
    Call SetearRs(strSql)
    With rs
        While Not .EOF
            ' -----------------------------------------------------------------
            ' Leer mensaje de la banca: Se quitan espacios en blanco y se trabaja en minusculas
            ' -----------------------------------------------------------------
            MensajeActual.sTipo = GetCadena(LCase(Trim(.Fields("Tipo").Value)))
            MensajeActual.sComponente = GetCadena(LCase(Trim(.Fields("Componente").Value)))
            MensajeActual.sObjeto = GetCadena(LCase(Trim(.Fields("Objeto").Value)))
            MensajeActual.sAtributo = GetCadena(LCase(Trim(.Fields("Atributo").Value)))
            MensajeActual.sValor = GetCadena(LCase(Trim(.Fields("valor").Value)))
            MensajeActual.sComentario = GetCadena(LCase(Trim(.Fields("Comentario").Value)))
            xNroMensajeSB = .Fields("Id").Value
            If xPrimerMensajeSB = 0 Then
                xPrimerMensajeSB = xNroMensajeSB
            End If
            ' -------------------------------------------------------------------------------------
            ' Armo mensaje para log
            ' -------------------------------------------------------------------------------------
            strMensajeLog = "Tipo: " & MensajeActual.sTipo & "; Componente: " & MensajeActual.sComponente _
                        & "; Objeto: " & MensajeActual.sObjeto & "; Atributo: " & MensajeActual.sAtributo _
                        & "; Valor: " & MensajeActual.sValor
            nLogSQVPrueba = nLogSQVPrueba + 1
            strUltimoMensaje_SB_SQV = strMensajeLog

            'xLogSQVPrueba = xLogSQVPrueba & "    " & Format(nLogSQVPrueba, "00000000000") & "¦ " & Now & "¦ " & "Msj SB :" & Str(xNroMensajeSB) & "¦" & _
            " Msjs/seg: " & Format((xNroMensajeSB - xPrimerMensajeSB) / max(DateDiff("s", xFechaInicioProceso, Now), 0.001), "###.00") & _
            strMensajeLog & vbCrLf
            'Call AltaLogGeneral("sqv", " " & "  " & Format(nLogSQVPrueba, "0000000") & "¦ " & Now & "¦ " & _
            " Msjs/seg: " & Format((xNroMensajeSB - xPrimerMensajeSB) / max(DateDiff("s", xFechaInicioProceso, Now), 0.001), "###.00"))
            ' -------------------------------------------------------------------------------------
            ' Verificar si hay que dejar log de mensajes
            ' -------------------------------------------------------------------------------------
            If chkLog_Mensajes.Value = 1 Then
                Call AltaLogGeneral("Servidor de Bancas", strMensajeLog, MensajeActual.sObjeto)
            End If
            ' -----------------------------------------------------------------
            ' Mensaje switch: Alguien se sento o se paro
            ' -----------------------------------------------------------------
            With MensajeActual
                flSwitchExitoso = False
                xBanca = Int(.sObjeto)
                If xBanca = 256 Then
                    xBanca = Int(.sObjeto)
                End If
                If xBanca >= 1 And xBanca <= xUltimaBanca Then
                    If VL.modoExtendido Then
                        If .sAtributo = "switch" Then
                            If .sValor = "closed" Then 'Se sento!
                                VL.PresenciaReal(xBanca) = PRESENTE
                            Else
                                'Esta open
                                VL.PresenciaReal(xBanca) = AUSENTE
                                VL.PerdioIdentificacion(xBanca) = True
                            End If
                        End If
                    End If
                    If .sComponente = "term.seat" And VL.bancaValida(Int(.sObjeto)) = True Then
                        If .sAtributo = "switch" Then
                        If EstadoActual.VectorError(Val(.sObjeto)) = ERROR_IOC Then 'Si estaba en IOC lo restauro
                            EstadoActual.VectorError(Val(.sObjeto)) = ERROR_SIN_ERROR
                        End If
                            If .sValor = "closed" Then
                                ' Preguntar en el vector de presencia si la banca esta ocupada
                                If EstadoActual.VectorPresencia(xBanca) <> PRESENTE Then
                                    EstadoActual.VectorPresencia(xBanca) = PRESENTE
                                    
                                    If EstadoActual.VectorPresencia(0) <> PRESENTE Then
                                        EstadoActual.VectorPresencia(0) = IIf(xPresidenteLegislador, PRESENTE, AUSENTE)
                                        EstadoActual.VectorColor(0) = AsignarColor(0)

                                    End If
                                    
                                    ' para abstencion numerica
                                    If EstadoActual.TipoDeOperacion = "votnum" Then
                                      ' MODIFICACION AP 040921
                                      If InStr(SEPARADOR_VECTOR & Join(EstadoActual.VectorAbstencion, SEPARADOR_VECTOR), SEPARADOR_VECTOR & Trim(Str(xBanca)) & SEPARADOR_VECTOR) > 0 Then
                                          AbstenerBanca (xBanca)
                                      End If
                                      'If InStr(Join(EstadoActual.VectorAbstencion, SEPARADOR_VECTOR), Trim(Str(xBanca))) > 0 Then
                                      '    AbstenerBanca (xBanca)
                                      'End If
                                    End If ' fin para abstencion numerica
                                    
                                    If Not ((EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum") And EstadoActual.EstadoVotacion_y_PasList = "finalizada") Then
                                       EstadoActual.VectorColor(xBanca) = AsignarColor(xBanca)
                                    End If
                                    ' actualizo cantidad de legisladores en el recinto
                                    If PrimerRecuento = True Then
                                        EstadoActual.Presentes = EstadoActual.Presentes + 2
                                        EstadoActual.Ausentes = EstadoActual.Ausentes - 2
                                        PrimerRecuento = False
                                    Else
                                        EstadoActual.Presentes = EstadoActual.Presentes + 1
                                        EstadoActual.Ausentes = EstadoActual.Ausentes - 1
                                    End If
                                    'CartelActual.Presentes = CartelActual.Presentes + 1
                                    flSwitchExitoso = True
                                End If
                            ElseIf .sValor = "open" Then
                                If ModoMant Then
                                    VectorDesconectadas(xBanca) = False
                                End If
                                If EstadoActual.VectorPresencia(xBanca) = BANCA_INHABILITADA Then
                                    With Mensaje2Banca
                                        .sTipo = "mget"
                                        .sObjeto = Str(xBanca)
                                        .sComponente = "term"
                                        .sAtributo = "state"
                                        .sValor = ""
                                    End With
                                    Call EnviarMensajesBancas(Mensaje2Banca)
                                    Call AltaLogGeneral("BANCA " & Str(xBanca) & "INHABILITADA", "Banca inhabilitada por switch open. Esperando respuesta de SB", Str(xBanca), "2")
                                ElseIf EstadoActual.VectorPresencia(xBanca) = PRESENTE Then
                                    EstadoActual.VectorPresencia(xBanca) = AUSENTE
                                    If Not ((EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum") And EstadoActual.EstadoVotacion_y_PasList = "finalizada") Then
                                        EstadoActual.VectorColor(xBanca) = AsignarColor(xBanca)
                                    End If
                                    ' actualizo cantidad de legisladores en el recinto
                                    EstadoActual.Presentes = EstadoActual.Presentes - 1
                                    EstadoActual.Ausentes = EstadoActual.Ausentes + 1
                                    flSwitchExitoso = True
                                End If
                            End If
                        End If
                    End If
                End If
                If EstadoActual.VectorColor(xBanca) <> cMarronClaro Then
                    Call PintarVectorColor(xBanca)
                End If
                If xBanca = 1 Then
                    xBanca = 1
                End If
                Call CalcularMinimoParaQuorum
            End With
            ' */*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/
            ' Escuchar mensaje de votacion
            ' */*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/
            'If EstadoActual.EstadoVotacion_y_PasList = "votando" Or EstadoActual.EstadoVotacion_y_PasList = "larga" Then
            '    Call Votacion(MensajeActual)
            'End If
            ' */*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/
            ' Fin mensaje de votacion
            ' */*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/
            
            ' -----------------------------------------------------------------
            ' Inicializar un proceso de identificacion de banca
            ' -----------------------------------------------------------------
            With MensajeActual
                flBancaIdentifPosExitosa = False
                flExitoPierdeIdDup = False
                flExitoPierdeIdDupConPresdte = False
                flExitoPierdeID = False
'                If (EstadoActual.TipoDeOperacion = "votnom" And Not (.sComponente = "term.keyb.si" Or .sComponente = "term.keyb.no")) Or EstadoActual.TipoDeOperacion = "paslis" Or (EstadoActual.TipoDeOperacion = "quorum" And EstadoActual.Modo_Ident_Nom = 1) Or (EstadoActual.TipoDeOperacion = "votnum" And EstadoActual.Modo_Ident_Nom = 0) Then
                If Not (.sComponente = "term.keyb.si" Or .sComponente = "term.keyb.no") Then
                   'antes nominal  If (EstadoActual.TipoDeOperacion = "votnom" And Not (.sComponente = "term.keyb.si" Or .sComponente = "term.keyb.no")) Or EstadoActual.TipoDeOperacion = "paslis" Then
                   Call Identificacion(MensajeActual)
                End If
                If EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum" Then
                   Call Votacion(MensajeActual)
                End If
                If (.sComponente = "term.mon" Or .sComponente = "term" Or .sComponente = "term.ioc") And LCase(.sAtributo) = "state" Or .sAtributo = "einact" Or .sAtributo = "eidacp" Or .sAtributo = "eidrxh" Then
                    If VL.bancaValida(CInt(xBanca)) Then
                        Call ManejoDeFallas(MensajeActual)
                    End If
                End If
                If EstadoActual.ModoMantenimientoBancas > 0 And Not IsNull(.sObjeto) Then
                    Call MantenimientoBancas(MensajeActual)
                End If
                If xBancaPruebaScan > 0 And xBanca = xBancaPruebaScan And EstadoActual.TipoDeOperacion = "quorum" Then
                    Call BancaPruebaScan(MensajeActual)
                End If
            End With
            ' -----------------------------------------------------------------
            ' Actualizar valores antes de buscar siguiente mensaje
            ' -----------------------------------------------------------------
            xUltimoMensajeSB = .Fields("Id").Value
            .MoveNext
            CntInterno = CntInterno + 1
            If CntInterno > 5 Then
                Call ActualizarTiempoCartel
                CntInterno = 0
            End If
        Wend
        'ver si se valida el fin de tiempo de votacion.
        'VERIFICACION DE TIEMPO CUMPLIDO DE VOTACION solo en estado votando sea nominal o numerica
        Call ControlTiempoCumplidoVotacion
        'Operaciones en caso de cierre de votacion
        Call CierreVotacion
        'Actualizacion de cartel y quorum
        xTiempoVotacionTranscurrido = DateDiff("s", EstadoActual.FechaVotacion, Now)
        xTiempoRestanteVotacion = EstadoActual.TiempoParaVotacion - xTiempoVotacionTranscurrido
        CartelActual.LeyendaTiempo = _
            IIf(EstadoActual.EstadoVotacion_y_PasList = "espera", "", _
                IIf(EstadoActual.EstadoVotacion_y_PasList = "votando", _
                    IIf(xTiempoRestanteVotacion > EstadoActual.TiempoParaVotacion, "", _
                        IIf(xTiempoRestanteVotacion > 59, Str(xTiempoRestanteVotacion), _
                            IIf(xTiempoVotacionTranscurrido > EstadoActual.TiempoParaVotacion, " 0", Right(Str(xTiempoRestanteVotacion), 2)))), _
                IIf(EstadoActual.EstadoVotacion_y_PasList = "larga", " 0", _
                IIf(EstadoActual.EstadoVotacion_y_PasList = "cancelada", "VOTACION CANCELADA", " 0"))))
        If CartelActual.LeyendaTiempo = "VOTACION CANCELADA" Then
            lblGeneralTiempoDato.top = 1600
            lblGeneralTituloTiempo.Visible = False
        Else
            lblGeneralTiempoDato.top = lblGeneralTituloTiempo.top
        End If
        '>> Leyenda del tipo de operacion, presentes, ausentes, etc.  para el cartel mural. La formula permite obtener esta leyenda.
        CartelActual.LeyendaTipoOperacion = IIf(EstadoActual.TipoDeOperacion = "quorum", "QUORUM", IIf(EstadoActual.TipoDeOperacion = "votnum", "VOTACION NUMERICA", IIf(EstadoActual.TipoDeOperacion = "paslis", "PASE DE LISTA", IIf(EstadoActual.TipoDeOperacion = "votnom", "VOTACION NOMINAL", ""))))
        ' calculo de presentes

        CartelActual.Presentes = IIf(EstadoActual.TipoDeOperacion <> "quorum" And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate" Or EstadoActual.EstadoVotacion_y_PasList = "cierre"), EstadoActual.PresentesCongelados, Presentes())
        CartelActual.Ausentes = IIf(EstadoActual.TipoDeOperacion <> "quorum" And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate" Or EstadoActual.EstadoVotacion_y_PasList = "cierre"), EstadoActual.AusentesCongelados, Ausentes())
        'Tipos de quorum:
        'MAN: mantenimiento, con dos presentes hay quorum (1 mas el "presidente")
        '121 La mitad mas uno del cuerpo
        '120 La mitad del cuerpo
        
        '>> Leyenda de quorum segun sea el tipo de quorum seleccionado. Esta leyenda variara dinamicamente segun cambien los presentes y los parametros de tipo de quorum.
        CartelActual.LeyendaQuorum = CalculoQuorum()
        'CartelActual.LeyendaQuorum = IIf(CartelActual.Presentes >= Fix(xMinimoParaQuorumEntero), "QUORUM", "NO HAY QUORUM")
        
       ' Cierre de Pase de lista
        If EstadoActual.TipoDeOperacion = "paslis" And EstadoActual.EstadoVotacion_y_PasList = "finalizada" Then
            '|  134|En pases de lista, muestra de datos en carteles                          |
            '|  135|Fin de congelamiento de datos en carteles en pase de lista               |
            'NO Se debe grabar el pase de lista por pedido manual.
            '>> Luego de un tiempo de espera definido en la configuracion, sale de estado "finalizada"
            If DateDiff("s", EstadoActual.FechaVotacion, Now) >= xTiempoEsperaPaseLista Then
                EstadoActual.EstadoVotacion_y_PasList = "espera"
                CartelActual.Resultado = ""
                'EstadoActual.OcupadosNoIdentificados = 0
            End If
        End If
        'WATCH DOG
        'SIN USO
'        If DateDiff("s", xFechaUltimoReset, Now) > 30 Then
'            For i = 0 To xUltimaBanca
'                If EstadoActual.VectorPresencia(i) = BANCA_INHABILITADA Then
'                    With Mensaje2Banca ' Mensaje para SB
'                        .sTipo = "mget"
'                        .sComponente = "term.mon"
'                        .sObjeto = Str(i)
'                        .sAtributo = "action"
'                        .sValor = "reset"
'                    End With
'                    Call EnviarMensajesBancas(Mensaje2Banca)
'                End If
'            Next
'            xFechaUltimoReset = Now
'        End If
    End With
    ' Rs.Close
End Function
Private Sub CalcularMinimoParaQuorum()
    ' variable que permite calcular el minimo necesario para obtener quorum
    xMinimoParaQuorum = IIf(LCase(EstadoActual.TipoMayoriaQuorum) = "man", 1, Fix(xMiembrosDelCuerpo / 2) + IIf(EstadoActual.TipoMayoriaQuorum = "121", 1, 0)) 'NICO 1 x 2
    If xMiembrosDelCuerpo Mod 2 = 1 Then
        xMinimoParaQuorumEntero = xMinimoParaQuorum + 1
    Else
        xMinimoParaQuorumEntero = xMinimoParaQuorum
    End If
End Sub
Private Sub CalcularMinimoParaQuorumold()
    Dim xMinimoConDecimales As Double
    ' variable que permite calcular el minimo necesario para obtener quorum
    xMinimoConDecimales = IIf(LCase(EstadoActual.TipoMayoriaQuorum) = "man", 1, xMiembrosDelCuerpo / 2 + IIf(EstadoActual.TipoMayoriaQuorum = "121", 1, 0))
    xMinimoParaQuorum = Fix(xMinimoConDecimales)
' variable que permite calcular el minimo necesario para obtener quorum aplicando la funcion FIX que toma solo la parte entera para descartar los decimales si el cuerpo tiene legisladores pares.
    If EstadoActual.TipoMayoriaQuorum = "120" Or EstadoActual.TipoMayoriaQuorum = "121" And xMinimoConDecimales > Fix(xMinimoConDecimales) Then
        xMinimoParaQuorumEntero = xMinimoParaQuorum + 1
    End If
        Call AltaLogGeneral("SERVER SQV", CartelActual.LeyendaQuorum & " " & xMinimoParaQuorum & " " & xMinimoParaQuorumEntero & " " & CartelActual.Presentes)
        'MsgBox CartelActual.LeyendaQuorum & " " & xMinimoParaQuorum & " " & xMinimoParaQuorumEntero & " " & CartelActual.Presentes
End Sub
Private Sub BancaPruebaScan(MensajeActual As MensajeSistema)

    Dim xActualBanca  As Long
    Dim Mensaje2Banca As MensajeSistema
    Dim strSql        As String
            
            With MensajeActual
                    xActualBanca = Int(.sObjeto)
                    If .sComponente = "term.auth" And EstadoActual.VectorPresencia(xActualBanca) = PRESENTE Then
                        xActualBanca = Int(.sObjeto)
                        If LCase(.sAtributo) = "result" Then
                            If LCase(.sValor) = "negative" Then
                                EstadoActual.strError = "pruebascan"
                                EstadoActual.MensajeAlOperador = "Banca " & Trim(Str(xActualBanca)) & " en prueba de Scan = negative"
                            ElseIf LCase(.sValor) <> "negative" Then ' Busca mValor en base de datos legisladores
                                strSql = "SELECT * FROM legisladores WHERE id = " & Val("&H" & Trim(.sValor)) & " AND tipo = 1"
                                Call SetearOtroRs(strSql)
                                If RsOtro.RecordCount = 0 Or RsOtro.EOF = True Then ' Si no lo encuentra entre los legisladores, lo busca entre el personal de mantenimiento
                                    strSql = "SELECT * FROM legisladores WHERE id = " & Val("&H" & Trim(.sValor)) & " AND tipo = 0"
                                    Call SetearOtroRs(strSql)
                                    If RsOtro.RecordCount > 0 Then
                                        ' Lo encontro entre la gente de mantenimiento: ackowledge al operador
                                        EstadoActual.strError = "pruebascan"
                                        If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
                                            'modo normal
                                            EstadoActual.MensajeAlOperador = "Prueba Scan Invalida: " & Trim(.sValor) & " es el valor recibido, no se encuentra en base de datos de Legisladores (Man)"
                                        Else ' modo mantenimiento
                                            EstadoActual.MensajeAlOperador = RsOtro.Fields("id")
                                        End If
                                    Else ' no es ni Leg, ni Man
                                        EstadoActual.strError = "pruebascan"
                                        EstadoActual.MensajeAlOperador = "Prueba Scan Invalida: " & Trim(.sValor) & " es el valor recibido, no se encuentra en base de datos (Leg - Man)."
                                    End If
                                Else
                                    ' Lo encontro como legislador: ackowledge al operador
                                    EstadoActual.strError = "pruebascan"
                                    'EstadoActual.strError = "**error"
                                    'EstadoActual.MensajeAlOperador = RsOtro.Fields("Apellido").Value & " " & RsOtro.Fields("Nombre").Value & " identificado Ok (Leg.)"
                                    EstadoActual.MensajeAlOperador = RsOtro.Fields("id")
                                    ' acknowledge al Legislador
'                                    Mensaje2Banca.sObjeto = xActualBanca '<AP 040115 faltaba indicar la banca>
'                                    Mensaje2Banca.sTipo = "mset"
'                                    Mensaje2Banca.sComponente = "term.display"
'                                    Mensaje2Banca.sAtributo = "text"
'                                    Mensaje2Banca.sValor = "Prueba Identificacion Valida"
'                                    Mensaje2Banca.sComentario = "Prueba Scan Id Aceptado Modo normal"
'                                    Call EnviarMensajesBancas(Mensaje2Banca)
                                     Mensaje2Banca.sObjeto = xBanca '<AP 040115 faltaba indicar la banca>
                                     Mensaje2Banca.sTipo = "mset"
                                     Mensaje2Banca.sComponente = "term.led1"
                                     Mensaje2Banca.sAtributo = "state"
                                     Mensaje2Banca.sValor = "on"
                                     If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
                                         Mensaje2Banca.sComentario = "Id aceptado Operador Modo normal - led1 - B"
                                     Else
                                         Mensaje2Banca.sComentario = "Id aceptado Operador Modo mantenimiento"
                                     End If
                                'End With
                                    Call EnviarMensajesBancas(Mensaje2Banca)
                                End If ' Fin si no esta entre personal de mantenimiento
                                RsOtro.Close
                            End If ' Fin LCase(.sValor) <> "negative"
                        End If
                    End If
            End With
    'MsgBox "PRUEBA SCAN terminar segun documentacion "
End Sub
Private Sub MantenimientoBancas(MensajeBanca As MensajeSistema)
    Dim xBanca As Long
    Dim xAuxNroIdAsignado As String
    Dim xIdentificadorMant As Long
    Dim i As Long
    Dim strRes As String
    'Dim mensajebanca As MensajeSistema
    Dim bActualizar As Boolean
    Dim flBancaPresidente As Boolean
    
    flBancaPresidente = False
    bActualizar = False
    
    With MensajeBanca
        xBanca = Int(.sObjeto)
        If MensajeBanca.sComponente = "term.seat" And MensajeBanca.sAtributo = "switch" And flSwitchExitoso Then
            If LCase(MensajeBanca.sValor) = "closed" Then
                'Closed: inicia o reinicia mantenimiento de la banca
                Call MensajeDisplayTerminal(Str(xBanca), "MAN-0:Panel " & xIdentificadorMant & " - Identifiquese")
                EstadoActual.VMantEstado(xBanca) = ABSTENCION
                If xIdentificadorMant > 0 Then
                    EstadoActual.VMantInfo(xIdentificadorMant) = " "
                End If
                EstadoActual.MantPresencias = Trim(xBanca) & ";" & Trim(EstadoActual.MantPresencias)
                If xBanca = 0 Then
                    flBancaPresidente = True
                End If
                bActualizar = True
            End If
        End If
        If flBancaIdentifPosExitosa Or flBancaPresidente Then
            flBancaPresidente = False
            xAuxNroIdAsignado = IIf(xBanca = 0, "P", EstadoActual.VectorIdentificacion(xBanca))
            xIdentificadorMant = 0
            For i = 1 To cUltimoPanelMant
                If EstadoActual.VMantIdentificacion(i) = xAuxNroIdAsignado Then
                    xIdentificadorMant = i
                End If
            Next i
            'si no lo encontro en ninguna, busca un panel disponible
            If xIdentificadorMant = 0 Then
                For i = 1 To cUltimoPanelMant
                    If Left(EstadoActual.VMantInfo(i), 3) = "FIN" Or Trim(EstadoActual.VMantInfo(i)) = "" Then
                        xIdentificadorMant = i
                    End If
                Next i
            End If
            'si no hay ninguno disponible lo pone en el ultimo
            If xIdentificadorMant = 0 Then
                xIdentificadorMant = cUltimoPanelMant
            End If
            'Ahora el identificador mantenimiento tiene el numero correspondiente a su area de mantenimiento de 1 a 4
            'Lo inserta en la pantalla correspondiente
            
            EstadoActual.VMantInfo(xIdentificadorMant) = IIf(xBanca = 0, "PRESID", xNombreUltimoIdentificado)
            EstadoActual.VMantBanca(xIdentificadorMant) = xBanca
            EstadoActual.VMantIdentificacion(xIdentificadorMant) = xAuxNroIdAsignado
            'acaa
            Call MensajeDisplayTerminal(Str(xBanca), "MAN-1:Panel " & xIdentificadorMant & " - Presione NO")
            bActualizar = True
        Else
            'ubica la banca en el vector
            xIdentificadorMant = 0
            For i = 1 To cUltimoPanelMant
                If EstadoActual.VMantBanca(i) = xBanca Then
                    xIdentificadorMant = i
                End If
            Next i
        End If ' fin id positiva
        If xBanca = 0 Then
            If EstadoActual.VectorResultados(xBanca) = NEGATIVO Then
                PresidenteEstuvoMantenimiento = True
                Call MensajeDisplayTerminal(Str(xBanca), "MAN-2:Panel " & xIdentificadorMant & " - Presione SI")
                lblMantenimientostrPanel3.Caption = "BANCA PRESIDENTE..."
                lblOperador4.Caption = "PRESIDENTE V. NO"
                EstadoActual.VMantEstado(xBanca) = NEGATIVO
                bActualizar = True
            ElseIf EstadoActual.VMantEstado(xBanca) = NEGATIVO And EstadoActual.VectorResultados(xBanca) = AFIRMATIVO Then
                PresidenteEstuvoMantenimiento = True
                Call MensajeDisplayTerminal(Str(xBanca), "MAN-3:Panel " & xIdentificadorMant & " - FINALIZADO OK")
                EstadoActual.VMantEstado(xBanca) = AFIRMATIVO
                EstadoActual.VMantInfo(xIdentificadorMant) = "FINOK " & Trim(EstadoActual.VMantInfo(xIdentificadorMant))
                'lblMantenimientostrPanel3.Caption = "BANCA PRESIDENTE OK"
                lblOperador4.Caption = "PRESIDENTE OK"
                Call EnviarMensajesFinAuth("0", "Banca Presidente en Modo Mant")
                bActualizar = True
            End If
        Else
            If xIdentificadorMant > 0 Then
                If EstadoActual.VectorIdentificacion(xBanca) <> NO_IDENTIFICADO Or (xBanca = 0) Then
                    If EstadoActual.VectorResultados(xBanca) = NEGATIVO Then
                        Call MensajeDisplayTerminal(Str(xBanca), "MAN-2:Panel " & xIdentificadorMant & " - Presione SI")
                        EstadoActual.VMantEstado(xBanca) = NEGATIVO
                        bActualizar = True
                    ElseIf EstadoActual.VMantEstado(xBanca) = NEGATIVO And EstadoActual.VectorResultados(xBanca) = AFIRMATIVO Then
                        Call MensajeDisplayTerminal(Str(xBanca), "MAN-3:Panel " & xIdentificadorMant & " - FINALIZADO OK")
                        EstadoActual.VMantEstado(xBanca) = AFIRMATIVO
                        EstadoActual.VMantInfo(xIdentificadorMant) = "FINOK " & Trim(EstadoActual.VMantInfo(xIdentificadorMant))
                        VectorDesconectadas(Val(xBanca)) = False
                        bActualizar = True
                    End If
                End If
            End If
        End If
        If MensajeBanca.sComponente = "term.seat" And MensajeBanca.sAtributo = "switch" And flSwitchExitoso Then
            If LCase(MensajeBanca.sValor) = "open" And Not (EstadoActual.VMantEstado(xBanca) = AFIRMATIVO) Then
                'OPEN y no completó secuencia de verificacion mantenimiento
                Call MensajeDisplayTerminal(Str(xBanca), "MAN-FIN ERROR 3:Panel " & xIdentificadorMant & " - INCOMPLETO")
                If xIdentificadorMant > 0 Then
                    EstadoActual.VMantInfo(xIdentificadorMant) = "FIN??" & Trim(EstadoActual.VMantInfo(xIdentificadorMant))
                End If
                bActualizar = True
            End If
        End If
    End With
    Dim cConta As Integer
    cConta = 0
    If bActualizar Then
        EstadoActual.MantCantFallas = 0
        EstadoActual.MantCantPendientes = 0
        EstadoActual.MantListaFallas = " "
        EstadoActual.MantListaPendientes = " "
        For i = 0 To xUltimaBanca
            cConta = cConta + 1
            If cConta > 7 Then
                cConta = 0
            End If
            If EstadoActual.VMantEstado(i) = ABSTENCION Then
                EstadoActual.MantCantPendientes = EstadoActual.MantCantPendientes + 1
                EstadoActual.MantListaPendientes = Trim(EstadoActual.MantListaPendientes) & "," & Trim(Str(i))
                If cConta = 7 Then
                    EstadoActual.MantListaPendientes = EstadoActual.MantListaPendientes & vbCrLf
                End If
            ElseIf Not (EstadoActual.VMantEstado(i) = AFIRMATIVO) And Not (EstadoActual.VectorPresencia(i) = PRESENTE) Then
                    EstadoActual.MantCantFallas = EstadoActual.MantCantFallas + 1
                    EstadoActual.MantListaFallas = Trim(EstadoActual.MantListaFallas) & "," & Trim(Str(i))
            End If
        Next i
    End If
End Sub
Private Function CuentaOcupadosNoIdentificadosCongelados() As Long
    Dim i As Long
    CuentaOcupadosNoIdentificadosCongelados = 0
    With EstadoActual
        For i = 1 To xUltimaBanca
            If .VectorPresenciaCong(i) = PRESENTE And Val(.VectorIdentificacionCong(i)) = NO_IDENTIFICADO Then
                CuentaOcupadosNoIdentificadosCongelados = CuentaOcupadosNoIdentificadosCongelados + 1
            End If
        Next i
    End With
End Function
            

Private Function CuentaOcupadosNoIdentificadosCong() As Long
    Dim i As Long
    CuentaOcupadosNoIdentificadosCong = 0
    With EstadoActual
        For i = 1 To xUltimaBanca
            If .VectorPresencia(i) = PRESENTE And Val(.VectorIdentificacion(i)) = NO_IDENTIFICADO Then
                CuentaOcupadosNoIdentificadosCong = CuentaOcupadosNoIdentificadosCong + 1
            End If
        Next i
    End With
End Function

Private Sub ManejoDeFallas(MensajeBanca As MensajeSistema)
    Dim xBanca As Long
    Dim strRes As String
    Dim strSql                      As String
    
    With MensajeBanca
        xBanca = Int(.sObjeto)
        If xBanca = 209 Then
            xBanca = 209
        End If
        If LCase(.sAtributo) = "eidrxh" Then
            If EstadoActual.VectorPresencia(xBanca) = "0" Then
               EstadoActual.VectorPresencia(xBanca) = "1"
            End If
        End If
        If ModoMant = False And xBanca > 0 And .sAtributo = "einact" Then
            If .sValor = "p" Then
                If EstadoActual.VectorPresencia(xBanca) = "0" Then
                   EstadoActual.VectorPresencia(xBanca) = "1"
                End If
            ElseIf .sValor = "a" Then
                If EstadoActual.VectorPresencia(xBanca) = "1" Then
                    EstadoActual.VectorPresencia(xBanca) = "0"
                End If
            End If
        End If
        If ModoMant = False And xBanca = 0 And LCase(.sAtributo) = "einactp" Or LCase(.sAtributo) = "einact p" Or (LCase(.sAtributo) = "einact" And LCase(.sValor) = "p") Then
            If Not EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO And xBanca > 0 Then

                Call AltaLogGeneral("SQV SERVER: Manejo fallas", "BANCA EINACT " & .sValor & " IDENTIFICADA = " & Trim(EstadoActual.VectorIdentificacion(xBanca)), Str(xBanca), "0")
                EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO
                If .sValor = "p" Then
                    EstadoActual.VectorPresencia(xBanca) = PRESENTE
                    If EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis" Or EstadoActual.Modo_Ident_Nom Then
                        'Call EnviarMensajesComienzoAuth(Trim(Str(xBanca)), "BANCA EINACT " & .sValor & " IDENTIFICADA = " & Trim(EstadoActual.VectorIdentificacion(xBanca)))
                    Else
                        'Call EnviarMensajesFinAuth(Trim(Str(xBanca)), "BANCA EINACT " & .sValor & " IDENTIFICADA = " & Trim(EstadoActual.VectorIdentificacion(xBanca)))
                    End If
                Else
                    'EstadoActual.VectorPresencia(xBanca) = AUSENTE
                    'Call EnviarMensajesFinAuth(Trim(Str(xBanca)), "BANCA EINACT " & .sValor & " IDENTIFICADA = " & Trim(EstadoActual.VectorIdentificacion(xBanca)))
                End If
                PintarVectorColor (xBanca)
            Else
                If ModoMant = False And EstadoActual.TipoDeOperacion = "votnum" And (EstadoActual.EstadoVotacion_y_PasList = "votando" Or EstadoActual.EstadoVotacion_y_PasList = "larga") Then
                    If EstadoActual.VectorPresencia(xBanca) = PRESENTE Then
                        With Mensaje2Banca
                            .sTipo = "mset"
                            .sObjeto = xBanca
                            .sComponente = "term.keyb"
                            .sAtributo = "state"
                            .sValor = "onvotnum"
                            .sComentario = "Reincorporacion a votnum en modo einact | estado :" & EstadoActual.EstadoVotacion_y_PasList
                        End With
                        Call EnviarMensajesBancas(Mensaje2Banca)
                    End If
                ElseIf ModoMant = False And (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis" Or EstadoActual.Modo_Ident_Nom = 1) And _
                    ((((DateDiff("s", EstadoActual.FechaVotacion, Now)) < EstadoActual.TiempoParaVotacion + xSegundosFinOperacion Or EstadoActual.EstadoVotacion_y_PasList = "espera") _
                    And EstadoActual.ExtensionDeTiempoPorPresidente = False And Not EstadoActual.EstadoVotacion_y_PasList = "finalizada")) Then
                    If xBanca <> 0 And EstadoActual.EstadoVotacion_y_PasList <> "esperafin" Then
                        If EstadoActual.VectorPresencia(xBanca) = PRESENTE And EstadoActual.VectorIdentificacion(xBanca) = "0" Then
                        If GetTickCount - VectorControlDobleTick(xBanca) > 15000 Then
                            VectorControlDobleTick(xBanca) = GetTickCount
                            VectorControlDoble(xBanca) = 0
                        End If
                            If VectorControlDoble(xBanca) > 1 Then
                                Call EnviarMensajesComienzoAuth(Trim(Str(xBanca)), "BANCA EINACT " & .sValor & " A EIDRXH = " & Trim(EstadoActual.VectorIdentificacion(xBanca)))
                            Else
                                VectorControlDoble(xBanca) = VectorControlDoble(xBanca) + 1
                            End If
                        End If
                    End If
                End If
            End If
        ElseIf LCase(.sAtributo) = "eidrxh" And xBanca > 0 And ModoMant = False Then
            If EstadoActual.VectorIdentificacion(xBanca) <> NO_IDENTIFICADO Then
                If TiempoEidrxh(xBanca) = 0 Then
                    TiempoEidrxh(xBanca) = GetTickCount
                End If
                CantidadEidrxh(xBanca) = CantidadEidrxh(xBanca) + 1
                If GetTickCount - TiempoEidrxh(xBanca) > 15000 Then
                    CantidadEidrxh(xBanca) = 0
                    TiempoEidrxh(xBanca) = GetTickCount
                End If
                If CantidadEidrxh(xBanca) >= 2 Then
                    Call EnviarMensajesFinAuth(Trim(Str(xBanca)), "ESTADO EIDRXH ESTANDO IDENTIFICADA")
                    'Call EnviarMensajesComienzoAuth(Trim(Str(xBanca)), "ESTADO EIDRXH ESTANDO IDENTIFICADA 2")
                    EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO
                    PintarVectorColor (xBanca)
                    CantidadEidrxh(xBanca) = 0
                    TiempoEidrxh(xBanca) = GetTickCount
                End If
            End If
        ElseIf LCase(.sAtributo) = "eidacp" And xBanca > 0 And ModoMant = False Then
            If EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO Then
                Call EnviarMensajesFinAuth(Trim(Str(xBanca)), "Cancelacion de ID por estado EIDACP sin identificacion")
            ElseIf (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum") And (EstadoActual.EstadoVotacion_y_PasList = "votando" Or EstadoActual.EstadoVotacion_y_PasList = "larga") And EstadoActual.VectorResultados(xBanca) = ABSTENCION And xBanca > 0 And ModoMant = False Then
                Call AltaLogGeneral("SQV SERVER: Manejo fallas", "BANCA EIDACP " & .sValor & " IDENTIFICADA = " & Trim(EstadoActual.VectorIdentificacion(xBanca)), Str(xBanca), "0")
                Dim MensajeParaBanca As MensajeSistema
                If EstadoActual.VectorPresencia(xBanca) = PRESENTE Then
                    With MensajeParaBanca
                        .sTipo = "mset"
                        .sObjeto = xBanca
                        .sComponente = "term.keyb"
                        .sAtributo = "state"
                        .sValor = "onvot" & EstadoActual.TipoDeOperacion
                        .sComentario = "Votacion.Identificacion exitosa banca" & EstadoActual.EstadoVotacion_y_PasList
                    End With
                    Call EnviarMensajesBancas(MensajeParaBanca)
                End If
            End If
        ElseIf LCase(.sTipo) = "mevt" And (LCase(.sComponente) = "term" Or LCase(.sComponente) = "term.ioc") And LCase(.sAtributo) = "state" Then
            If .sValor = "ok" Then
                VectorDesconectadas(xBanca) = False
            End If
            If LCase(.sValor) = "ok" And EstadoActual.VectorError(Val(.sObjeto)) = ERROR_SIN_ERROR Then
                'If EstadoActual.VectorPresencia(xBanca) = BANCA_INHABILITADA Then
                    If xBanca > 0 Then
                        EstadoActual.VectorPresencia(xBanca) = AUSENTE
                        EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO 'Agregado 21/03/2010
                        'Esta limpieza se dispara con el evento CONNECT del Socket del SB
                        'Se limpia el vector de identificación para que no quede la condición
                        'AUSENTE + IDENTIFICADO
                    Else 'ES PRESIDENTE: ver si es Legislador, para ponerlo presente.
                        ' Buscar en la base si es legislador
                        strSql = "SELECT Es_Legislador FROM Legisladores WHERE id = '" & Trim(EstadoActual.VectorIdentificacion(xBanca)) & "'"
                        rsTemp.CursorLocation = adUseClient
                        rsTemp.Open strSql, Cn, adOpenForwardOnly, adLockReadOnly
                        If rsTemp.RecordCount > 0 And (rsTemp.EOF = False Or rsTemp.BOF = False) Then
                            If rsTemp("Es_Legislador").Value = 0 Then
                                xPresidenteLegislador = False
                            Else
                                xPresidenteLegislador = True
                            End If
                            If xPresidenteLegislador = True Then
                                'If EstadoActual.VectorPresencia(0) = AUSENTE Then
                                    EstadoActual.VectorPresencia(0) = PRESENTE
                                    'EstadoActual.Presentes = EstadoActual.Presentes + 1
                                    'EstadoActual.Ausentes = EstadoActual.Ausentes - 1
                                'End If
                            Else
                                'If EstadoActual.VectorPresencia(0) = PRESENTE Then
                                    EstadoActual.VectorPresencia(0) = AUSENTE
                                    'EstadoActual.Presentes = EstadoActual.Presentes - 1
                                    'EstadoActual.Ausentes = EstadoActual.Ausentes + 1
                                'End If
                            End If
                            Call PintarVectorColor(0)
                        Else
                            'No encontro al presidente
                            Call AltaLogGeneral("SQV SERVER: Manejo fallas", "Recuperacion banca 0 y no existe presidente: " & Trim(EstadoActual.VectorIdentificacion(xBanca)), Str(xBanca), "0")
                        End If
                        rsTemp.Close
                    End If
                    PintarVectorColor (xBanca)
                    Call AltaLogGeneral("sqv", "Habilitacion banca:" & Str(xBanca) & " Estado:" & .sValor, Str(xBanca))
                'End If
            End If
            If LCase(.sValor) = "off" Or LCase(.sValor) = "error" Then
                If ModoMant = True Then
                    VectorDesconectadas(Val(xBanca)) = True
                End If
                If .sValor = "off" Then
                    EstadoActual.VectorError(xBanca) = ERROR_SIN_ERROR
                End If
                If EstadoActual.VectorPresencia(xBanca) = PRESENTE Then
                    If xBanca > 0 Then
                        'EstadoActual.Presentes = max(EstadoActual.Presentes - 1, 1) 'Si la banca es 0 no se resta
                        'EstadoActual.Ausentes = Min(xMiembrosDelCuerpo - 1, EstadoActual.Ausentes + 1)
                        EstadoActual.Presentes = max(EstadoActual.Presentes - 1, IIf(xPresidenteLegislador, 1, 0)) 'Si la banca es 0 no se resta
                        EstadoActual.Ausentes = xMiembrosDelCuerpo - EstadoActual.Presentes
                        If (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis") Or EstadoActual.Modo_Ident_Nom = 1 Then  'ooooooiii
                            If (EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO) Then
                                EstadoActual.OcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados - 1
                            End If
                        End If
                        If (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And EstadoActual.EstadoVotacion_y_PasList <> "finalizada" And EstadoActual.EstadoVotacion_y_PasList <> "cierre" Then
                            strRes = LCase(EstadoActual.VectorResultados(xBanca))
                            If strRes = AFIRMATIVO Or strRes = NEGATIVO Then
                                Call AltaLogGeneral("SQV SERVER: ManejoDeFallas 1", "Afirmativo o negativo no cambia: " & EstadoActual.PendientesEmitirVotos, 0, "0")
                                'EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1
                                Select Case strRes
                                    Case AFIRMATIVO
                                        CartelActual.Afirmativos = CartelActual.Afirmativos - 1
                                    Case NEGATIVO
                                        CartelActual.Negativos = CartelActual.Negativos - 1
                                End Select
                            ElseIf strRes = ABSTENCION_AUTORIZADA Then
                                EstadoActual.AbstencionistasAutorizados = EstadoActual.AbstencionistasAutorizados - 1
                            Else
                                EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1
                                Call AltaLogGeneral("SQV SERVER: ManejoDeFallas 2", "Caso ELSE (ni afirmativo, ni negativo ni abstencion autorizada) EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1: " & EstadoActual.PendientesEmitirVotos, Str(xBanca), "0")
                            End If
                            EstadoActual.VectorResultados(xBanca) = ABSTENCION
                        End If
                    End If
                End If
                EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO
                EstadoActual.VectorPresencia(xBanca) = BANCA_INHABILITADA
                If LCase(.sValor) = "error" Then
                    EstadoActual.VectorError(xBanca) = ERROR_IOC
                Else
                    'EstadoActual.VectorError(xBanca) = ERROR_SIN_ERROR
                End If
                EstadoActual.VectorColor(xBanca) = "1"
                PintarVectorColor (xBanca)
                Call AltaLogGeneral("sqv", "Inhabilitacion banca:" & Str(xBanca) & " Estado:" & .sValor, Str(xBanca))
            End If
        End If
    End With
End Sub

Private Sub Votacion(MensajeBanca As MensajeSistema)

    Dim xBanca              As Long
    Dim xPendientesDeVotar  As Long
    Dim xOcupadosCongelados As Long
    Dim MensajeParaBanca    As MensajeSistema
   
    'Solo durante el tiempo de votación: manejo de botones de votacion
    xBanca = Int(MensajeBanca.sObjeto)
    If EstadoActual.EstadoVotacion_y_PasList = "votando" Or EstadoActual.EstadoVotacion_y_PasList = "larga" Then
        'Filtro de banca 0 y mayor a leg
        'If xBanca >= IIf(xPresidenteLegislador, 0, 1) And xBanca <= xUltimaBanca Then
        If xBanca >= IIf(EstadoActual.PresidenteHabilitadoParaVotar, 0, 1) And xBanca <= xUltimaBanca Then
            'Si es nominal, ver que este identificado, sino solo que esté presente
            'Si es votacion larga, tambien se puede habiiltar para votar
            If (EstadoActual.VectorPresencia(xBanca) = PRESENTE Or xBanca = 0) And _
                    (EstadoActual.TipoDeOperacion = "votnum" Or _
                     (EstadoActual.TipoDeOperacion = "votnom" And (EstadoActual.VectorIdentificacion(xBanca) <> NO_IDENTIFICADO Or (EstadoActual.ModoMantenimientoBancas And xBanca = 0))) _
                    ) Then
                ' los abstencionistas no los cuento
                If LCase(EstadoActual.VectorResultados(xBanca)) <> ABSTENCION_AUTORIZADA Then
                    'objeto == term.keyb
                    If (MensajeBanca.sComponente = "term.keyb.si" Or MensajeBanca.sComponente = "term.keyb.no") And LCase(MensajeBanca.sAtributo) = "state" And LCase(MensajeBanca.sValor) = "on" Then
                        If EstadoActual.EstadoVotacion_y_PasList = "larga" Then
                            Mensaje2Banca.sTipo = "mset"
                            Mensaje2Banca.sObjeto = xBanca
                            Mensaje2Banca.sComponente = "term.keyb"
                            Mensaje2Banca.sAtributo = "state"
                            Mensaje2Banca.sValor = "off" & EstadoActual.TipoDeOperacion
                            Mensaje2Banca.sComentario = EstadoActual.EstadoVotacion_y_PasList & "Bloqueo de teclado en votacion larga"
                            Call EnviarMensajesBancas(Mensaje2Banca)
                        End If
                        'Actualiza pendientes de votar si antes no habia votado
                        If LCase(EstadoActual.VectorResultados(xBanca)) = ABSTENCION Then
                            EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1
                            Call AltaLogGeneral("SQV SERVER: Votacion 1", "LCase(EstadoActual.VectorResultados(xBanca)) = ABSTENCION Then EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1: " & EstadoActual.PendientesEmitirVotos, Str(xBanca), "0")
                        End If
                        'deshacer voto anterior
                        'No apago las luces del voto anterior pues lo hace la banca automaticamente
                        Select Case LCase(EstadoActual.VectorResultados(xBanca))
                            Case AFIRMATIVO
                                CartelActual.Afirmativos = CartelActual.Afirmativos - 1
                            Case NEGATIVO
                                CartelActual.Negativos = CartelActual.Negativos - 1
                        End Select
                        'Aplica voto nuevo, y prende la luz del acknowledge
                        If MensajeBanca.sComponente = "term.keyb.si" Then
                            CartelActual.Afirmativos = CartelActual.Afirmativos + 1
                            EstadoActual.VectorResultados(xBanca) = AFIRMATIVO
                            With MensajeParaBanca
                                .sTipo = "mset"
                                .sObjeto = xBanca
                                .sComponente = "term.ledk1"
                                .sAtributo = "state"
                                .sValor = "on"
                                .sComentario = EstadoActual.EstadoVotacion_y_PasList
                            End With
                            Call EnviarMensajesBancas(MensajeParaBanca)
                        ElseIf MensajeBanca.sComponente = "term.keyb.no" Then
                            CartelActual.Negativos = CartelActual.Negativos + 1
                            EstadoActual.VectorResultados(xBanca) = NEGATIVO
                            With MensajeParaBanca
                                .sTipo = "mset"
                                .sObjeto = xBanca
                                .sComponente = "term.ledk2"
                                .sAtributo = "state"
                                .sValor = "on"
                                .sComentario = EstadoActual.EstadoVotacion_y_PasList
                            End With
                            Call EnviarMensajesBancas(MensajeParaBanca)
                            If EstadoActual.EstadoVotacion_y_PasList = "larga" Then
                                'Mato al tipo
                            End If
                        End If 'MensajeBanca.sComponente = "term.keyb.si o no
                        PintarVectorColor (xBanca)
                    End If 'MensajeBanca.sComponente = "term.keyb.si" Or MensajeBanca.sComponente = "term.keyb.no" And MensajeBanca.sAtributo = "STATE" And MensajeBanca.sComponente.sValor = "ON"
                Else
                    Call AltaLogGeneral("SQVB", "Intento de voto con abstencion autorizada" & xBanca, Str(xBanca)) 'aca3
                End If 'LCase(EstadoActual.VectorResultados(xBanca)) <> absaut
            End If 'EstadoActual.VectorPresencia(xBanca) = PRESE...
            'obj term.seat switch
            If MensajeBanca.sComponente = "term.seat" And MensajeBanca.sAtributo = "switch" And flSwitchExitoso Then
                'Si se levanta en medio de la votación >> PROCESO DE OPEN
                Select Case LCase(MensajeBanca.sValor)
                Case "open"
                    'queda uno menos pendiente de votar, si no habia votado antes en estado = votando o larga                                              |
                    If (EstadoActual.TipoDeOperacion = "votnum" Or flExitoPierdeID Or (EstadoActual.ModoMantenimientoBancas And xBanca = 0)) Then
                        'deshacer voto anterior
                        'No apago las luces del voto anterior pues lo hace la banca automaticamente
                        Select Case LCase(EstadoActual.VectorResultados(xBanca))
                            Case AFIRMATIVO
                                CartelActual.Afirmativos = CartelActual.Afirmativos - 1
                            Case NEGATIVO
                                CartelActual.Negativos = CartelActual.Negativos - 1
                            Case ABSTENCION_AUTORIZADA
                                EstadoActual.AbstencionistasAutorizados = EstadoActual.AbstencionistasAutorizados - 1
                            Case ABSTENCION
                                EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1
                                Call AltaLogGeneral("SQV SERVER: Votacion 2", "Case ABSTENCION EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1: " & EstadoActual.PendientesEmitirVotos, Str(xBanca), "0")
                        End Select
                        'limpiar vector resultado con " "
                        'If EstadoActual.VectorResultados(xBanca) <> ABSTENCION_AUTORIZADA Then
                        '    EstadoActual.VectorResultados(xBanca) = ABSTENCION
                        'End If
                        EstadoActual.VectorResultados(xBanca) = ABSTENCION
                        With MensajeParaBanca
                            .sTipo = "mset"
                            .sObjeto = xBanca
                            .sComponente = "term.keyb"
                            .sAtributo = "state"
                            .sValor = "off" & IIf(xBanca > 0, EstadoActual.TipoDeOperacion, "votnum")
                            .sComentario = EstadoActual.EstadoVotacion_y_PasList
                        End With
                        Call EnviarMensajesBancas(MensajeParaBanca)
                        ' Apagar las luces de votacion de la banca
                        With MensajeParaBanca
                            .sTipo = "mset"
                            .sObjeto = xBanca
                            .sComponente = "term.ledk1"
                            .sAtributo = "state"
                            .sValor = "off"
                            .sComentario = EstadoActual.EstadoVotacion_y_PasList
                        End With
                        Call EnviarMensajesBancas(MensajeParaBanca)
                    
                    End If
                    If (EstadoActual.TipoDeOperacion = "votnom" And Not flExitoPierdeID) Then
                        EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1
                        Call AltaLogGeneral("SQV SERVER: Votacion 3 ATT - no pierde id", "If (EstadoActual.TipoDeOperacion = votnom And Not flExitoPierdeID) Then EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1: " & EstadoActual.PendientesEmitirVotos, Str(xBanca), "0")
                    End If
                '>> PROCESO DE CLOSED se permite ingresar a habilitarse para votar
                Case "closed"
                    'Si es numérica y esta presente
                    If (EstadoActual.TipoDeOperacion = "votnum" And EstadoActual.VectorPresencia(xBanca) = PRESENTE Or (EstadoActual.ModoMantenimientoBancas = 1 And xBanca = 0)) And Not EstadoActual.EstadoVotacion_y_PasList = "larga" Then
                        'Habilitar para votar
                        With MensajeParaBanca
                            .sTipo = "mset"
                            .sObjeto = xBanca
                            .sComponente = "term.keyb"
                            .sAtributo = "state"
                            .sValor = "on" & IIf(xBanca > 0, EstadoActual.TipoDeOperacion, "votnum")
                            .sComentario = EstadoActual.EstadoVotacion_y_PasList
                        End With
                        Call EnviarMensajesBancas(MensajeParaBanca)
                    End If
                    'Suma a pendientes de votar tanto en numerica como nominal                |
                    EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos + 1
                    Call AltaLogGeneral("SQV SERVER: Votacion 4 ", "Suma a pendientes de votar tanto en numerica como nominal EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos + 1: " & EstadoActual.PendientesEmitirVotos, Str(xBanca), "0")
                End Select 'SWITCH
            End If 'MensajeBanca.sComponente = "term.seat"
            '>> PROCESA MENSAJES DE IDENTIFICACION PARA VER SI LO HABILITA PARA VOTAR tanto en modo de votacion normal = votando como larga y en tipo votacion nominal se permite ingresar a habilitarse para votar
            If EstadoActual.TipoDeOperacion = "votnom" And MensajeBanca.sComponente = "term.auth" And MensajeBanca.sAtributo = "result" And MensajeBanca.sValor <> "negative" Then
                If flBancaIdentifPosExitosa Then
                    'Hubo identificacion positiva en el mismo ciclo Habilitar para votar
                    With MensajeParaBanca
                        .sTipo = "mset"
                        .sObjeto = xBanca
                        .sComponente = "term.keyb"
                        .sAtributo = "state"
                        .sValor = "on" & IIf(xBanca > 0, EstadoActual.TipoDeOperacion, "votnum")
                        .sComentario = "Votacion.Identificacion exitosa banca" & EstadoActual.EstadoVotacion_y_PasList
                    End With
                    Call EnviarMensajesBancas(MensajeParaBanca)
                End If
                'Nota: voto presidente es habilitado directamente en procedimiento de identificacion.
                If flExitoPierdeIdDup And Not (flExitoPierdeIdDupConPresdte) Then
                    'Pierde el voto en la banca duplicada
                    'If EstadoActual.VectorResultados(xBancaDuplicada) <> ABSTENCION_AUTORIZADA Then
                        'deshacer voto anterior
                        'No apago las luces del voto anterior pues lo hace la banca automaticamente
                        'Si ya votó, tampoco queda pendiente de votar
                        'Si no votó, lo resta de pendiente de votar.
                        Select Case LCase(EstadoActual.VectorResultados(xBancaDuplicada))
                            Case AFIRMATIVO
                                CartelActual.Afirmativos = CartelActual.Afirmativos - 1
                                EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos + 1 'test091012
                                Call AltaLogGeneral("SQV SERVER: Votacion 5 ATT Pierde ID", "CartelActual.Afirmativos = CartelActual.Afirmativos - 1 SUMA 1 pend: " & EstadoActual.PendientesEmitirVotos, Str(xBanca), "0")
                            Case NEGATIVO
                                CartelActual.Negativos = CartelActual.Negativos - 1
                                EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos + 1 'test091012
                                Call AltaLogGeneral("SQV SERVER: Votacion 6 ATT Pierde ID", "CartelActual.Negativos = CartelActual.Negativos - 1 SUMA 1 pendi: " & EstadoActual.PendientesEmitirVotos, Str(xBanca), "0")
                            Case ABSTENCION
                                'EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1 ' ap 091021
                                Call AltaLogGeneral("SQV SERVER: Votacion 7 ATT Pierde ID", "Case ABSTENCION EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1: " & EstadoActual.PendientesEmitirVotos, Str(xBanca), "0")
                            Case ABSTENCION_AUTORIZADA
                                EstadoActual.AbstencionistasAutorizados = EstadoActual.AbstencionistasAutorizados - 1
                                Call AltaLogGeneral("SQV SERVER: Votacion 8 ATT Pierde ID", "Case ABSTENCION_AUTORIZADA EstadoActual.AbstencionistasAutorizados = EstadoActual.AbstencionistasAutorizados - 1 No hace nada: " & EstadoActual.PendientesEmitirVotos, Str(xBanca), "0")
                        End Select
                        EstadoActual.VectorResultados(xBancaDuplicada) = ABSTENCION
                        'apaga teclado (FINVT)
                        With MensajeParaBanca
                            .sTipo = "mset"
                            .sObjeto = xBancaDuplicada
                            .sComponente = "term.keyb"
                            .sAtributo = "state"
                            .sValor = "off" & IIf(xBancaDuplicada > 0, EstadoActual.TipoDeOperacion, "votnum")
                            .sComentario = EstadoActual.EstadoVotacion_y_PasList
                        End With
                        Call EnviarMensajesBancas(MensajeParaBanca)
                        'reinicia la identificacion
                        With MensajeParaBanca
                            .sObjeto = xBancaDuplicada
                            .sTipo = "mset"
                            .sComponente = "term.auth"
                            .sAtributo = "action"
                            .sComentario = EstadoActual.EstadoVotacion_y_PasList & " Duplicada restart "
                            If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
                                .sComentario = .sComentario & " Modo normal"
                            Else
                                .sComentario = .sComentario & " Modo mantenimiento"
                            End If
                            .sValor = Trim("auth_restart")  'comienzo normal
                        End With
                        Call EnviarMensajesBancas(MensajeParaBanca)
                    'End If
                End If 'EstadoActual.VectorResultados(xBancaDuplicada) <> ABSTENCION_AUTORIZADA Then
            End If '>> PROCESA MENSAJES DE IDENTIFICACION PARA VER SI LO HABILITA PARA VOTAR tanto en modo de votacion normal = votando como larga y en tipo votacion nominal se permite ingresar a habilitarse para votar
        End If 'Filtro de banca 0 y mayor a leg
        '>> OPERACIONES NO ASOCIADAS AL PROCESO DE MENSAJES.
        'VERIFICACION DE TIEMPO CUMPLIDO DE VOTACION solo en estado votando sea nominal o numerica
        Call ControlTiempoCumplidoVotacion
    End If 'EstadoActual.EstadoVotacion_y_PasList = "votando" Or EstadoActual.EstadoVotacion_y_PasList = "larga" Then
    ' Manejo de switch en el caso de abstencion autorizada
    'MODIFICACION AP 040921
    If LCase(EstadoActual.VectorResultados(xBanca)) = ABSTENCION_AUTORIZADA And (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum") And Not (EstadoActual.EstadoVotacion_y_PasList = "votando" Or EstadoActual.EstadoVotacion_y_PasList = "larga") Then
   'If LCase(EstadoActual.VectorResultados(xBanca)) = ABSTENCION_AUTORIZADA And  EstadoActual.TipoDeOperacion = "votnom"                                             And Not (EstadoActual.EstadoVotacion_y_PasList = "votando" Or EstadoActual.EstadoVotacion_y_PasList = "larga") Then
        If MensajeBanca.sComponente = "term.seat" And MensajeBanca.sAtributo = "switch" And flSwitchExitoso Then
            'Si se levanta
            If LCase(MensajeBanca.sValor) = "open" Then
                'MODIFICACION AP 040921
                If (flExitoPierdeID) Or EstadoActual.TipoDeOperacion = "votnum" Then
               'If (flExitoPierdeID) Then ' or EstadoActual.TipoDeOperacion = "votnum"
                    EstadoActual.AbstencionistasAutorizados = EstadoActual.AbstencionistasAutorizados - 1
                    EstadoActual.VectorResultados(xBanca) = ABSTENCION
                End If
            '>> PROCESO DE CLOSED se permite ingresar a habilitarse para votar
            End If 'SWITCH
        End If 'MensajeBanca.sComponente = "term.seat"
    End If
    If MensajeBanca.sComponente = "term.auth" And MensajeBanca.sAtributo = "result" And MensajeBanca.sValor <> "negative" Then
        If flExitoPierdeIdDup And Not (flExitoPierdeIdDupConPresdte) Then
           If LCase(EstadoActual.VectorResultados(xBancaDuplicada)) = ABSTENCION_AUTORIZADA And EstadoActual.TipoDeOperacion = "votnom" And Not (EstadoActual.EstadoVotacion_y_PasList = "votando" Or EstadoActual.EstadoVotacion_y_PasList = "larga") Then
                EstadoActual.AbstencionistasAutorizados = EstadoActual.AbstencionistasAutorizados - 1
                EstadoActual.VectorResultados(xBancaDuplicada) = ABSTENCION
            End If 'MensajeBanca.sComponente = "term.seat"
        End If
    End If
    '>> SITUACION DE EMPATE
    '>> TODOS LOS DEMAS LEGISLADORES QUEDAN "CONGELADOS" mientras vota el presidente como vice gobernador.
    'Si es empate, procesar el voto del presidente solamente
    If xBanca = 0 And EstadoActual.EstadoVotacion_y_PasList = "empate" And _
       (MensajeBanca.sComponente = "term.keyb.si" Or MensajeBanca.sComponente = "term.keyb.no") And MensajeBanca.sAtributo = "state" And MensajeBanca.sValor = "on" Then
        'Aplica voto nuevo, y prende la luz del acknowledge
        If MensajeBanca.sComponente = "term.keyb.si" Then
            CartelActual.Afirmativos = CartelActual.Afirmativos + 1
            EstadoActual.VectorResultados(xBanca) = AFIRMATIVO
            EstadoActual.VectorResultadosCong(xBanca) = AFIRMATIVO
            CartelActual.Resultado = "AFIRMATIVO"
            xHuboDesempate = True
            With MensajeParaBanca
                .sTipo = "mset"
                .sObjeto = xBanca
                .sComponente = "term.ledk1"
                .sAtributo = "state"
                .sValor = "on"
                .sComentario = EstadoActual.EstadoVotacion_y_PasList
            End With
            Call EnviarMensajesBancas(MensajeParaBanca)
        ElseIf MensajeBanca.sComponente = "term.keyb.no" Then
            CartelActual.Negativos = CartelActual.Negativos + 1
            EstadoActual.VectorResultados(xBanca) = NEGATIVO
            EstadoActual.VectorResultadosCong(xBanca) = NEGATIVO
            CartelActual.Resultado = "NEGATIVO"
            xHuboDesempate = True
            With MensajeParaBanca
                .sTipo = "mset"
                .sObjeto = xBanca
                .sComponente = "term.ledk2"
                .sAtributo = "state"
                .sValor = "on"
                .sComentario = EstadoActual.EstadoVotacion_y_PasList
            End With
            Call EnviarMensajesBancas(MensajeParaBanca)
        End If 'MensajeBanca.sComponente = "term.keyb.si o no
        If "hay" = "cartel serial" Then Call CartelSerial("resultado")
        PintarVectorColor (xBanca)
        'Presentar resultados en cartel
        EstadoActual.EstadoVotacion_y_PasList = "finalizada"
        'cancela teclado presidente: en este caso se trata como numerica, porque no se habilito por el operador, y la banca queda en modo no identificada

        With MensajeParaBanca
            .sTipo = "mset"
            .sObjeto = xBanca
            .sComponente = "term.keyb"
            .sAtributo = "state"
            .sValor = "off" & IIf(EstadoActual.PresidenteHabilitadoParaVotar, "votnom", "votnum")
            .sComentario = EstadoActual.EstadoVotacion_y_PasList
        End With
        Call EnviarMensajesBancas(MensajeParaBanca)
        EstadoActual.ActaGrabada = 0
        EstadoActual.SolicitudGrabarManual = 0
        'Vuelve a pintar la pantalla del operador con el resultado (puede haberse perdido al levantarse algun leg.)
        Call PintarTodasLasBancas
        Call AltaLogGeneral("Votacion", "EMPATE VOTA PRESIDENTE " & CartelActual.Resultado, Str(0))
    End If 'empate
    'ver si se debe cerrar la votacion si todos votaron
    'y hacer operaciones durante el cierre de una votacion
    Call CierreVotacion
End Sub
Private Sub ControlTiempoCumplidoVotacion()
        Dim Mensaje2Banca          As MensajeSistema
        Dim X As Integer
        Dim xStrVector As String
        Dim i As Integer
        If Tick_InicioPasLis > 0 Then 'Se inicio un pase de lista
            If EstadoActual.EstadoVotacion_y_PasList = "canpas" Then
                Tick_InicioPasLis = 0
                lblGeneralInformacion.Caption = "PASE DE LISTA CANCELADO"
                EstadoActual.ActaGrabada = 0
                ActualizarVector_enBD
            End If
            If GetTickCount - Tick_InicioPasLis > 5000 Then
                If EstadoActual.EstadoVotacion_y_PasList <> "canpas" Then
                    Call FinPasLis
                End If
                Tick_InicioPasLis = 0
            End If
        End If
        'VERIFICACION DE TIEMPO CUMPLIDO DE VOTACION solo en estado votando sea nominal o numerica
        If EstadoActual.EstadoVotacion_y_PasList = "votando" Or EstadoActual.EstadoVotacion_y_PasList = "larga" And Not (EstadoActual.ModoMantenimientoBancas) Then
            If EstadoActual.TipoDeAbstencion = "absaut" And EstadoActual.ExtensionDeTiempoPorPresidente = True Then
                If Trim(EstadoActual.VectorResultados(0)) <> "" Then
                    If EstadoActual.VectorResultados(0) = "AP" Then 'Si fue abstenido
                        EstadoActual.VectorResultados(0) = ABSTENCION 'Lo abstengo realmente
                    End If
                    VL.log "--------------------------- Llamada 4 ---------------------------"
                    EstadoActual.EstadoVotacion_y_PasList = "cierre"
                    EstadoActual.ExtensionDeTiempoPorPresidente = False
                End If
            End If
            If (DateDiff("s", EstadoActual.FechaVotacion, Now)) >= EstadoActual.TiempoParaVotacion + xSegundosFinOperacion Then
                'si es modalidad de abstencion automatica
                If EstadoActual.TipoDeAbstencion = "absaut" Then
                    If PrimeraVezControl = True And (EstadoActual.PresidenteHabilitadoParaVotar And Trim(EstadoActual.VectorResultados(0)) = "") Then 'Si el presidente puede votar y todavía no votó hay que esperar a q vote
                        Dim StrTempCadena As String
                        PrimeraVezControl = False
                        'EstadoActual.EstadoVotacion_y_PasList = "larga"
                        'Cancelar todos los teclados menos el presidente
                        StrTempCadena = "0" & SEPARADOR_VECTOR
                        For X = 1 To xUltimaBanca
                            If EstadoActual.VectorPresencia(X) = PRESENTE Then
                                StrTempCadena = StrTempCadena & "1" & SEPARADOR_VECTOR
                            Else
                                StrTempCadena = StrTempCadena & "0" & SEPARADOR_VECTOR
                            End If
                        Next X
                        With Mensaje2Banca ' Mensaje para SB
                            .sTipo = "mset"
                            .sComponente = "term.keyb"
                            .sObjeto = StrTempCadena
                            .sAtributo = "state"
                            .sComentario = "Cancelacion de teclado en abstencion automatica"
                            .sValor = "off" & EstadoActual.TipoDeOperacion
                        End With
                        Call EnviarMensajesBancas(Mensaje2Banca)
                        If EstadoActual.Modo_Ident_Nom = 1 Or EstadoActual.TipoDeOperacion <> "votnum" Then
                            xStrVector = "0" & SEPARADOR_VECTOR 'presidente
                            For i = 1 To UBound(EstadoActual.VectorPresencia)
                                xStrVector = xStrVector & IIf(EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO, "1", "0") & SEPARADOR_VECTOR
                            Next i
                            Call EnviarMensajesFinAuth(xStrVector, "Fin modo nominal desde " & EstadoActual.TipoDeOperacion & "a " & "")
                            EstadoActual.Modo_Ident_Nom = 0
                        End If
'                        While Trim(EstadoActual.ResultadoVotoPresidente) = ""
'                            DoEvents
'                        Wend
                        EstadoActual.ExtensionDeTiempoPorPresidente = True
                    ElseIf EstadoActual.PresidenteHabilitadoParaVotar = False Or (EstadoActual.PresidenteHabilitadoParaVotar = True And Trim(EstadoActual.VectorResultados(0)) <> "") Then
                        If EstadoActual.VectorResultados(0) = "AP" Then
                            EstadoActual.VectorResultados(0) = ABSTENCION
                        End If
                        VL.log "--------------------------- Llamada 5 ---------------------------"
                        EstadoActual.EstadoVotacion_y_PasList = "cierre"
                    End If
                ElseIf EstadoActual.TipoDeAbstencion = "votlar" Then 'si es modalidad de votacion larga
                    Dim Pendientes As Boolean
                    Pendientes = False
                    For i = 1 To 256
                        'If EstadoActual.VectorIdentificacion(i) <> "0" And EstadoActual.VectorResultados(i) = " " Then
                        If EstadoActual.VectorIdentificacion(i) <> "0" And (EstadoActual.VectorResultados(i) = " " Or EstadoActual.VectorColor(i) <> cGRIS) Then
                            Pendientes = True
                            i = 256
                            DoEvents
                        End If
                    Next i
                    If EstadoActual.PresidenteHabilitadoParaVotar = True And Trim(EstadoActual.VectorResultados(0)) = "" Then
                        Pendientes = True
                    End If
                    If Pendientes = False Then
                        If (Not VL.votosPendientes()) Then
                            PintarTodasLasBancas
                            VL.log "--------------------------- Llamada 6 ---------------------------"
                            EstadoActual.EstadoVotacion_y_PasList = "cierre"
                            PrimerControlLarga = False
                            Exit Sub
                        End If
                    Else
                        EstadoActual.EstadoVotacion_y_PasList = "larga"
                    End If
                    If Not (CartelActual.LeyendaQuorum = "QUORUM") Then 'Si no hay quorum, la cancela
                        EstadoActual.EstadoVotacion_y_PasList = "cancelada"
                        FinVotacionBrc ("cancelada")
                    Else
                        If PrimerControlLarga = False Then
                            VL.modoExtendido = True
                            Call VL.guardarEstado
                            PrimerControlLarga = True
                            EstadoActual.EstadoVotacion_y_PasList = "larga"
                            'CANCELAR TECLADOS DE TODOS MENOS LOS PENDIENTES DE VOTAR
                            If Not EstadoActual.PresidenteHabilitadoParaVotar Or (EstadoActual.PresidenteHabilitadoParaVotar And Trim(EstadoActual.ResultadoVotoPresidente) = "") Then
                                StrTempCadena = "0" & SEPARADOR_VECTOR
                            ElseIf EstadoActual.PresidenteHabilitadoParaVotar And EstadoActual.ResultadoVotoPresidente <> "" Then 'Si ya votó
                                StrTempCadena = "1" & SEPARADOR_VECTOR 'Le cancelo el teclado
                            End If
                            For X = 1 To xUltimaBanca
                                If EstadoActual.VectorPresencia(X) = PRESENTE And Trim(EstadoActual.VectorResultados(X)) <> "" Then
                                    StrTempCadena = StrTempCadena & "1" & SEPARADOR_VECTOR
                                Else 'Si esta ausente o si no votó
                                    StrTempCadena = StrTempCadena & "0" & SEPARADOR_VECTOR
                                End If
                            Next X
                            With Mensaje2Banca ' Mensaje para SB
                                .sTipo = "mset"
                                .sComponente = "term.keyb"
                                .sObjeto = StrTempCadena
                                .sAtributo = "state"
                                .sComentario = "Cancelacion de teclado en votlar"
                                .sValor = "off" & EstadoActual.TipoDeOperacion
                            End With
                            Call EnviarMensajesBancas(Mensaje2Banca)
                            If EstadoActual.Modo_Ident_Nom = 1 Or EstadoActual.TipoDeOperacion <> "votnum" Then
                                xStrVector = "0" & SEPARADOR_VECTOR 'presidente
                                For i = 1 To UBound(EstadoActual.VectorPresencia)
                                    Dim xIdent As String
                                    xStrVector = xStrVector & IIf(EstadoActual.VectorPresencia(i) = PRESENTE And EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO, "1", "0") & SEPARADOR_VECTOR
                                Next i
                                Call EnviarMensajesFinAuth(xStrVector, "Fin modo nominal desde " & EstadoActual.TipoDeOperacion & "a " & "")
                                EstadoActual.Modo_Ident_Nom = 0
                            End If
                        End If
                    End If
                End If
            End If
        End If 'fin chequeo tiempo cumplido
End Sub
Private Sub CierreVotacion()
    Dim xBanca As Long
    Dim xPendientesDeVotar As Long
    Dim xOcupadosCongelados As Long
    Dim MensajeParaBanca As MensajeSistema
    Dim Mensaje2Banca          As MensajeSistema
    'ver si se debe cerrar la votacion si todos votaron
    If EstadoActual.EstadoVotacion_y_PasList = "larga" And Not (EstadoActual.ModoMantenimientoBancas) And EstadoActual.PendientesEmitirVotos - (0 * EstadoActual.AbstencionistasAutorizados) <= 0 Then 'no hace falta restarlos, pues al abstener se restan de pendientes de emitir votos.
        If (Not VL.votosPendientes()) Then
            'si no tengo votos pendientes
            VL.log "--------------------------- Llamada 1 ---------------------------"
            EstadoActual.EstadoVotacion_y_PasList = "cierre"
        End If
    End If
    'En estado de cierre de votacion
    'Si no hay quorum, y todavia no voto el presidente (no empa) la cancela
    'CartelActual.LeyendaQuorum = IIf(CartelActual.Presentes >= Fix(xMinimoParaQuorumEntero), "QUORUM", "NO HAY QUORUM")
    
    If EstadoActual.EstadoVotacion_y_PasList = "cierre" Then
        Dim ci As Integer
        ci = 0
        VL.log "--------------------------- Cierre ---------------------------"
        VL.log "Votos Pendientes: " & VL.votosPendientes
        VL.log "EstadoVotacion_y_PasList: " & EstadoActual.EstadoVotacion_y_PasList
        VL.log ""
        VL.log "////// Votos //////"
        For ci = 0 To 256
            VL.log "BANCA:" & ci & ";ID:" & EstadoActual.VectorIdentificacion(ci) & ";VOTO:" & EstadoActual.VectorResultados(ci)
        Next ci
        'If Not (CalculoQuorum() = "QUORUM") And ((EstadoActual.VectorResultados(0) <> AFIRMATIVO And EstadoActual.VectorResultados(0) <> NEGATIVO) Or xPresidenteLegislador) Then
        If Not (CalculoQuorum() = "QUORUM") And ((EstadoActual.VectorResultados(0) <> AFIRMATIVO And EstadoActual.VectorResultados(0) <> NEGATIVO) Or EstadoActual.PresidenteHabilitadoParaVotar) Then
           ' antes nominal If Not (IIf(EstadoActual.Presentes >= Fix(xMinimoParaQuorumEntero), "QUORUM", "NO HAY QUORUM") = "QUORUM") And ((EstadoActual.VectorResultados(0) <> AFIRMATIVO And EstadoActual.VectorResultados(0) <> NEGATIVO) Or xPresidenteLegislador) Then
            EstadoActual.EstadoVotacion_y_PasList = "cancelada"
            FinVotacionBrc ("cancelada")
        Else
            'If (Not xCierreEmpateOperador Or Trim(CartelActual.Resultado) <> "EMPATE") Then CartelActual.Resultado = CalculoResultado(EstadoActual.BaseMayoria, EstadoActual.TipoMayoria, xMiembrosDelCuerpo, EstadoActual.Presentes, CartelActual.Afirmativos, CartelActual.Negativos, "", 0, 0, CartelActual.MinimoVotosParaAfirmativo, EstadoActual.VectorResultados(0), IIf(xPresidenteLegislador, 1, 0))
            If (Not xCierreEmpateOperador Or Trim(CartelActual.Resultado) <> "EMPATE") Then CartelActual.Resultado = CalculoResultado(EstadoActual.BaseMayoria, EstadoActual.TipoMayoria, xMiembrosDelCuerpo, EstadoActual.Presentes, CartelActual.Afirmativos, CartelActual.Negativos, "", 0, 0, CartelActual.MinimoVotosParaAfirmativo, EstadoActual.VectorResultados(0), IIf(EstadoActual.PresidenteHabilitadoParaVotar, 1, 0))
            'MsgBox CartelActual.Resultado
            ' datos que debe utilizar la funcion de calculo.
            '+-#--+Var+Exp-+--- Description ------------------------------------------------------+------ Table -------+------ Name ----+
            '|   1|BH |    |xBase Mayoria                                                         |Virtual             |                |
            '|   2|BI |    |xTipo Mayoria                                                         |Virtual             |                |
            '|   3|K  |    |Cantidad de Legisladores                                              |Configuracion       |                |
            '|   4|T  |    |xPresentes                                                            |Virtual             |                |
            '|   5|BA |    |xAfirmativos                                                          |Virtual             |                |
            '|   6|BB |    |xNegativos                                                            |Virtual             |                |
            '|   7|Z  |    |xResultado                                                            |Virtual             |                |
            '|   8|BE |    |xMinimo votos para afirm                                              |Virtual             |                |
            '+----+---+----+----------------------------------------------------------------------+--------------------+----------------+
            'datos para actualizar el vector (VER SI EXISTE DIFERENCIA ENTRE VOTLAR Y ABSAUT)
            If EstadoActual.TipoDeOperacion = "votnom" Then
               EstadoActual.OcupadosNoIdentificados = GetNoIdentificadosSobrePresentes 'CuentaOcupadosNoIdentificadosCong()
               EstadoActual.OcupadosNoIdentificadosCongelados = EstadoActual.OcupadosNoIdentificados
            End If
            CartelActual.Abstenciones = getPresentes - CartelActual.Afirmativos - CartelActual.Negativos - IIf(EstadoActual.TipoDeOperacion = "votnom", GetNoIdentificadosSobrePresentes, 0) - IIf(EstadoActual.ModoVotaPresidente = False, 1, 0)  '- IIf(InStr("sn", EstadoActual.VectorResultados(0)) = 0, 0, 1)
            If CartelActual.Abstenciones <> EstadoActual.AbstencionistasAutorizados Then
                Call AltaLogGeneral("Sistema", "Abst. calc. " & Str(CartelActual.Abstenciones) & " Abs. Autor. " & Str(EstadoActual.AbstencionistasAutorizados))
            End If
            Call CartelSerial("resultado")
            Call AltaLogGeneral("Sistema", "Cierre votacion P:" & Str(EstadoActual.Presentes) & "A:" & Str(EstadoActual.Ausentes) & "S:" & Str(CartelActual.Afirmativos) & "N:" & Str(CartelActual.Negativos) & "A" & Str(EstadoActual.AbstencionistasAutorizados) & "  o:" & Str(EstadoActual.OcupadosNoIdentificados))
            'Cancela teclados y luces de teclados en todas las bancas
            FinVotacionBrc (EstadoActual.EstadoVotacion_y_PasList)
            '>> Congela presentes
            EstadoActual.PresentesCongelados = Presentes() ' EstadoActual.Presentes
            EstadoActual.AusentesCongelados = Ausentes() 'EstadoActual.Ausentes
            EstadoActual.EstadoVotacion_y_PasList = "finalizada"
            'Congela  identificados siempre, para evitar perder o recibir nvos id entre empate y fin o entre fin y grabacion  manual
            EstadoActual.VectorPresenciaCong = EstadoActual.VectorPresencia
            EstadoActual.VectorIdentificacionCong = EstadoActual.VectorIdentificacion
            EstadoActual.VectorResultadosCong = EstadoActual.VectorResultados
            'Salva el resultado para que no se lo pise cuando desempate
            If Not (ultimoResultadoEvaluado = "EMPATE" And UCase(Trim(CartelActual.Resultado)) = "EMPATE") Then
                'Si no se cerro la votacion en empate
                If EstadoActual.PresidenteHabilitadoParaVotar Then
                    EstadoActual.ResultadoVotoPresidente = EstadoActual.VectorResultados(0)
                    EstadoActual.VectorResultados(0) = ABSTENCION 'Limpia el vector
                    'para que pueda volver a votar en caso de empate
                End If
            End If
            ultimoResultadoEvaluado = UCase(Trim(CartelActual.Resultado))
            'Call FinalizarVotacionPresidente
            'Call DeshabilitarVotoPresidente 'por si acaso estuviera en modo nominal la banca y habilitada para votar
            'Se pinta con los resultados en votacion numerica tambien
                'Solo si es votacion nominal Actualizar informacion en pantalla operador cambiando colores
                'If EstadoActual.TipoDeOperacion = "votnom" Then
                '    Call PintarTodasLasBancas
                'End If
            Call PintarTodasLasBancas
            
            'si es empate
            'If CartelActual.Resultado = "EMPATE" And EstadoActual.VectorResultados(0) <> AFIRMATIVO And EstadoActual.VectorResultados(0) <> NEGATIVO And Not (lAbstencionPresidente) Then
            If CartelActual.Resultado = "EMPATE" And Not xHuboDesempate And Not (lAbstencionPresidente) Then
                lAbstencionPresidente = True ' bandera que permite que se cierre la votacion en empate con abstencion del presidente, si vuelve a pasar por esta seccion, ya no ingresa a la opcion de empate
                EstadoActual.EstadoVotacion_y_PasList = "empate"
                'habilitar al presidente para votar EN MODO NUMERICA si no esta habilitado
                'Nominal si tiene sautod
                Call BorrarVotoPresidente
                With MensajeParaBanca
                    .sTipo = "mset"
                    .sObjeto = 0
                    .sComponente = "term.keyb"
                    .sAtributo = "state"
                    .sValor = "onvot" & IIf(EstadoActual.PresidenteHabilitadoParaVotar, "nom", "num")
                    .sComentario = EstadoActual.EstadoVotacion_y_PasList
                End With
                Call EnviarMensajesBancas(MensajeParaBanca)
            End If
        End If 'hay quorum o desempato el presidente,
        Call ActualizarVector_enBD
        Call MostrarCartel
    End If 'EstadoActual.EstadoVotacion_y_PasList = "cierre"
    If EstadoActual.EstadoVotacion_y_PasList = "finalizada" Then
        'Grabacion de acta / Grabar acta / guardar acta / guarda acta / graba acta
        If EstadoActual.ActaGrabada = 0 Then
            EstadoActual.ActaGrabada = EstadoActual.NroActa
            If EstadoActual.TipoDeOperacion = "votnom" And EstadoActual.EstadoVotacion_y_PasList = "finalizada" Then
               'MsgBox "GrabaVotacionNominal"
               Call AlmacenarActa
                '+-#--+Var+Exp-+--- Description ------------------------------------------------------+------ Table -------+------ Name ----+
                '|   1|EE |    |Vector Identificacion Congelado                                       |Virtual             |                |
                '|   2|BM |    |xVector resultados                                                    |Virtual             |                |
                '|   3|DR |    |cResultado                                                            |Virtual             |                |
                '|   4|BF |    |xSesion                                                               |Virtual             |                |
                '|   5|BG |    |xNro. de Acta                                                         |Virtual             |                |
                '|   6|DM |    |cPresentes cartel                                                     |Virtual             |                |
                '|   7|DN |    |cAusentes cartel                                                      |Virtual             |                |
                '|   8|?? |  10|cAfirmativos - IF (MID (xVector resultados,1,1) = "S",1,0)                                                                                                                                              |
                '|   9|?? |  11|cNegativos - IF (MID (xVector resultados,1,1) = "N",1,0)                                                                                                                                                |
                '|  10|DU |    |cAbstenciones                                                         |Virtual             |                |
                '|  11|BD |    |xOcupados no identificados                                            |Virtual             |                |
                '|  12|DW |    |cMinimo votos para afirm                                              |Virtual             |                |
                '|  13|BT |    |xTipo Mayoria Quorum                                                  |Virtual             |                |
                '|  14|BH |    |xBase Mayoria                                                         |Virtual             |                |
                '|  15|BI |    |xTipo Mayoria                                                         |Virtual             |                |
                '|  16|BQ |    |xTitulo del Acta                                                      |Virtual             |                |
                '|  17|?? |  15|DATE ()                                                               |DATE ()                                                                                                                                                                                                 |
                '|  18|?? |  14|TIME ()                                                               |TIME ()                                                                                                                                                                                                 |
                '|  19|K  |    |Cantidad de Legisladores                                              |Configuracion       |                |
                '|  20|?? |  12|IF (MID (xVector resultados,1,1) = "S",1,0)                                                                                                                                                             |
                '|  21|?? |  13|IF (MID (xVector resultados,1,1) = "N",1,0)                                                                                                                                                             |
                '|  22|BU |    |xPeríodo Legislativo                                                  |Virtual             |                |
                '+----+---+----+----------------------------------------------------------------------+--------------------+----------------+'
            End If
            
            If EstadoActual.TipoDeOperacion = "votnum" And EstadoActual.EstadoVotacion_y_PasList = "finalizada" Then
               Call AlmacenarActa
               'MsgBox "GrabaVotacionNumerica"
                '+-#--+Var+Exp-+--- Description ------------------------------------------------------+------ Table -------+------ Name ----+
                '|   1|X  |    |xVector Identificacion                                                |Virtual             |                |
                '|   2|BM |    |xVector resultados                                                    |Virtual             |                |
                '|   3|DR |    |cResultado                                                            |Virtual             |                |
                '|   4|BF |    |xSesion                                                               |Virtual             |                |
                '|   5|BG |    |xNro. de Acta                                                         |Virtual             |                |
                '|   6|DM |    |cPresentes cartel                                                     |Virtual             |                |
                '|   7|DN |    |cAusentes cartel                                                      |Virtual             |                |
                '|   8|?? |  10|cAfirmativos - IF (MID (xVector resultados,1,1) = "S",1,0)                                                                                                                                              |
                '|   9|?? |  11|cNegativos - IF (MID (xVector resultados,1,1) = "N",1,0)                                                                                                                                                |
                '|  10|DU |    |cAbstenciones                                                         |Virtual             |                |
                '|  11|?? |   9|Cantidad de Legisladores - 1                                                                                                                                                                            |
                '|  12|DW |    |cMinimo votos para afirm                                              |Virtual             |                |
                '|  13|BT |    |xTipo Mayoria Quorum                                                  |Virtual             |                |
                '|  14|BH |    |xBase Mayoria                                                         |Virtual             |                |
                '|  15|BI |    |xTipo Mayoria                                                         |Virtual             |                |
                '|  16|BQ |    |xTitulo del Acta                                                      |Virtual             |                |
                '|  17|?? |  15|DATE ()                                                               |DATE ()                                                                                                                                                                                                 |
                '|  18|?? |  14|TIME ()                                                               |TIME ()                                                                                                                                                                                                 |
                '|  19|K  |    |Cantidad de Legisladores                                              |Configuracion       |                |
                '|  20|?? |  12|IF (MID (xVector resultados,1,1) = "S",1,0)                                                                                                                                                             |
                '|  21|?? |  13|IF (MID (xVector resultados,1,1) = "N",1,0)                                                                                                                                                             |
                '|  22|BU |    |xPeríodo Legislativo                                                  |Virtual             |                |
                '+----+---+----+----------------------------------------------------------------------+--------------------+----------------+
            End If
            Call CartelSerial("resultado")
            'Dim MensajeParaBanca As MensajeSistema
            ' *** Probando si es aca la cosa...
            'With MensajeParaBanca
            '    .sTipo = "mset"
            '    .sObjeto = "brc"
            '    .sComponente = "term.ledk1"
            '    .sAtributo = "state"
            '    .sValor = "off"
            '    .sComentario = EstadoActual.EstadoVotacion_y_PasList
            'End With
            'Call EnviarMensajesBancas(MensajeParaBanca)
            lAbstencionPresidente = False 'se limpia esta bandera que permite que se cierre la votacion en empate con abstencion del presidente
            If Imprimio = False Then
                MandarImprimir
                Imprimio = True
            End If
        End If ' fin grabar
    End If ' EstadoActual.EstadoVotacion_y_PasList = "finalizada"
        
    'Marcos anulo esto porque se llama a cada rato
    'Dim MensajeParaBanca As MensajeSistema
    'With MensajeParaBanca
    '    .sTipo = "mset"
    '    .sObjeto = "brc"
    '    .sComponente = "term.ledk1"
    '    .sAtributo = "state"
    '    .sValor = "off"
    '    .sComentario = EstadoActual.EstadoVotacion_y_PasList
    'End With
    'Call EnviarMensajesBancas(MensajeParaBanca)
    
    'MsgBox "VOTAC terminar segun documentacion "
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As Integer
    Dim Mensaje2Banca As MensajeSistema
    With Mensaje2Banca ' Mensaje para SB
        .sTipo = "mset"
        .sComponente = "sb"
        .sObjeto = "sb"
        .sAtributo = "shutdown"
        .sValor = "now"
    End With
    Call EnviarMensajesBancas(Mensaje2Banca)
    
    Call AltaLogGeneral("Finalizando y saliendo de SQV Server", Now)
    'Call closeLogfile
    ShowCursor True
    ResetearVectores
    ActualizarVector_enBD
    End
End Sub

Private Sub HabilitarSeguimientoPizarraCartel_Click()
    If blMostrarEstadoCartel = True Then
        HabilitarSeguimientoPizarraCartel.Caption = "&Mostrar Estado de Cartel"
        blMostrarEstadoCartel = False
    Else
        HabilitarSeguimientoPizarraCartel.Caption = "&Ocultar Estado de Cartel"
        blMostrarEstadoCartel = True
    End If
End Sub

Private Sub HabilitarSeguimientoPizarraRecinto_Click()
    If blMostrarEstadoRecinto = True Then
        HabilitarSeguimientoPizarraRecinto.Caption = "&Mostrar Estado de Recinto"
        blMostrarEstadoRecinto = False
    Else
        HabilitarSeguimientoPizarraRecinto.Caption = "&Ocultar Estado de Recinto"
        blMostrarEstadoRecinto = True
    End If
End Sub

Private Sub PintarTodasLasBancas()
    ' -----------------------------------------------------------------------------
    ' Recorrer todas las bancas y asignarle el color correspondiente
    ' -----------------------------------------------------------------------------
    Dim X      As Long
    For X = 0 To xUltimaBanca
        Call PintarVectorColor(X)
    Next X
End Sub
Private Sub PintarVectorColor(X As Long)
    ' -----------------------------------------------------------------------------
    ' Asignarle color a la banca x
    ' -----------------------------------------------------------------------------
    If X > xUltimaBanca Then
        Exit Sub
    End If
    If X = 2 Then
        X = X
    End If
    If EstadoActual.VectorColor(X) <> cMarronClaro Then
        EstadoActual.VectorColor(X) = AsignarColor(X)
    Else
        'If EstadoActual.VectorIdentificacion(X) <> "0" Then
            EstadoActual.VectorColor(X) = AsignarColor(X)
        'End If
    End If

End Sub
Private Sub OpenLogFile()
    On Error GoTo TrapError
    Dim strArchivo    As String
    Dim strCadena     As String
    Dim strFecha      As String
    Dim strDirectorio As String
    
    xFileSqv = FreeFile()
    strFecha = Trim(Replace(Date, "/", "-"))
    strDirectorio = Trim(App.Path & "\" & "logsqv" & strFecha)
    'strArchivo = strDirectorio & "\logsqv" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & ".txt" ' & "-" & Trim(Format(Time, "HHMMSS")) & ".txt"
    strArchivo = strDirectorio & "\" & Day(Now) & ".txt" ' & "-" & Trim(Format(Time, "HHMMSS")) & ".txt"
    Open strArchivo For Append As #xFileSqv
Exit Sub
TrapError:
    Select Case err.Number
        Case 76
            MkDir strDirectorio
            Resume
        Case Else
            Resume Next
    End Select
End Sub
Private Sub closeLogfile()
    Dim strCadena As String
    Close #xFileSqv
End Sub

Private Sub Timer_Timer()
    Dim xRespuesta As Long
    Dim xx         As String
    
    Static xVueltas As Long
    
    If blBanderaTimer Then
        blBanderaTimer = False
        xCiclosTotales = xCiclosTotales + 1 ' Contador de ciclos
        
        If (xCiclosTotales Mod 80 = 0) Or (xCiclosTotales = 1) Then
            If xCiclosTotales = 80 Then
                lblVersionCartel.Visible = False
            End If
            If (DateDiff("s", xFechaArranque, Now)) > 0 Then
               lblCiclos.Caption = "SBº" & Str(xNroMensajeSB) & " ac " & Trim(Format(xCiclosTotales / (DateDiff("s", xFechaArranque, Now)), "##.00"))
               Call OpenLogFile
               Print #xFileSqv, , xLogSQVPrueba
               Call closeLogfile
               xLogSQVPrueba = ""
               xFechaArranque = Now
               xCiclosTotales = 0
            End If
        End If
        'If EstadoActual.TipoDeOperacion = "votnom" Or _
        '        EstadoActual.TipoDeOperacion = "paslis" Or _
        '        (EstadoActual.TipoDeOperacion = "quorum" And EstadoActual.Modo_Ident_Nom = 1) And _
        '        (xCiclosTotales Mod 40 = 0) Then
        '        lblCiclos.Caption = "RID-SBº" & Str(xNroMensajeSB)
        '   Call SolicitarIdentificacionPendientes("Identificacion preventiva en modo " & EstadoActual.TipoDeOperacion, "restart")
        'End If
        ' -------------------------------------------------------------------------------------
        ' Leer vectores y valores de la base de datos y mostrarlo en
        ' pantalla de SQV SERVER y mostrar hora
        ' -------------------------------------------------------------------------------------
        'Call LeerEstadoRecinto
        
        Call CalcularMinimoParaQuorum
        
        lblFechaInicioServer.Caption = Now
        DoEvents
        ' -------------------------------------------------------------------------------------
        ' PROCESOS DE MENSAJES DE QUORUM: Mensajes de Presencia (a su vez esta se ocupa de invocar a los modulos de
        '  indentificación, id prueba, votación, fallas y mantenimiento según corresponda
        ' -------------------------------------------------------------------------------------
         xRespuesta = ProcesoDeMensajesQuorum
        ' -------------------------------------------------------------------------------------
        ' ACTUALIZACION ESTADO PARA CONSOLA: Se llama al programa que actualiza la información que va a presentar la consola
        '  (Mensaje a la Consola).
        ' -------------------------------------------------------------------------------------
        
        ' delay
        
        ' Control de modo presentación de formularios
        
        
        ' -------------------------------------------------------------------------------------
        ' mostrar estado de valores de cartel
        ' -------------------------------------------------------------------------------------
        If blMostrarEstadoCartel = True Then
            Call MostrarCartel
            frmEstadoCartel.Visible = True
        Else
            frmEstadoCartel.Visible = False
        End If
        ' -------------------------------------------------------------------------------------
        ' mostrar estado de valores de recinto
        ' -------------------------------------------------------------------------------------
        If blMostrarEstadoRecinto = True Then
            frmEstadoRecinto.Visible = True
        Else
            frmEstadoRecinto.Visible = False
        End If
        ' -------------------------------------------------------------------------------------
        ' Leer vectores y valores de recinto en memoria y actualizar en vector de BD
        ' -------------------------------------------------------------------------------------
        Call ActualizarVector_enBD
        Call PublicarEstadoRecinto
        blBanderaTimer = True
    End If
End Sub

Private Sub txtVecesPorSegundo_Change()
    
    If txtVecesPorSegundo.text = "" Then
        txtVecesPorSegundo.text = 2
        Exit Sub
    End If
    ' blServerPrendido = False
    blServerPrendido = True
    Call ServerOnOff
    If Trim(txtVecesPorSegundo.text) <> Str(xIntervalo) Then
        xIntervalo = Int(txtVecesPorSegundo.text)
        Call ServerOnOff
    End If
    
End Sub

Private Sub txtVecesPorSegundo_GotFocus()
    With txtVecesPorSegundo
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub txtVecesPorSegundo_KeyPress(KeyAscii As Integer)

    ' MsgBox KeyAscii
    If KeyAscii = 13 Then
        cmdTerminar_Click
    ElseIf KeyAscii = 8 Then
        Exit Sub
    ElseIf (KeyAscii > 57 Or KeyAscii < 48) Then
        KeyAscii = 0
    End If
End Sub
Private Function GetCadena(Optional strCadena As Variant) As String
    GetCadena = Trim(strCadena) & ""
End Function
Private Function GetNumero(Optional xNumero As Variant) As Integer
    If IsNull(xNumero) Then
        GetNumero = 0
    Else
        GetNumero = Val(xNumero)
    End If
End Function
Private Sub PublicarEstadoRecinto()
    
    With CartelActual
        lblcrt_Presentes.Caption = .Presentes
        'lblcrt_Afirmativos.Caption = .Afirmativos
        lblcrt_Ausentes.Caption = .Ausentes
        'lblcrt_Negativos.Caption = .Negativos
        lblcrt_Resultado.Caption = .Resultado
        lblcrt_Abstenciones.Caption = .Abstenciones
        lblcrt_MinimoParaAfirmativo.Caption = .MinimoVotosParaAfirmativo
        lblcrt_LeyendaQuorum.Caption = .LeyendaQuorum
        lblcrt_LeyendaTiempo.Caption = .LeyendaTiempo
    End With
    
    With EstadoActual
        lblVectorColor.Caption = ""
        llVectorPresencia.Caption = ""
        lblVectorIdentificacion.Caption = ""
        lblVectorResultado.Caption = ""
        lblOcupadosNoIdentificados(0).Caption = Str(.OcupadosNoIdentificados)
        lblPendientesEmitirVotos.Caption = .PendientesEmitirVotos
        lblSesion.Caption = .Sesion
        lblPeriodoLegislativo.Caption = .PeriodoLegislativo
        lblNumeroActa.Caption = .NroActa
        lblTituloActa.Caption = .TituloDelActa
        lblIdentificadorFormulario.Caption = .IdentificadorDeFormulario
        lblIpConsola.Caption = .IP_Consola
        lblError.Caption = .strError
        lblEstadoVotacionyPaseDeLista.Caption = .EstadoVotacion_y_PasList
        lblModalidadVotacion.Caption = .TipoDeAbstencion
        lblMensajeAlOperador.Caption = .MensajeAlOperador
        lblModoMantenimientoBancas.Caption = .ModoMantenimientoBancas
        lblModoMantenimientosistema.Caption = .ModoNormalMantSistema
        lblCartelEncendido.Caption = .CartelEncendido
        lblBaseMayoria.Caption = .BaseMayoria
        lblTipoMayoria.Caption = .TipoMayoria
        lbltipoMayoriaQuorum.Caption = .TipoMayoriaQuorum
        lblTipoOperacion.Caption = .TipoDeOperacion
        lblTiempoParaVotacion.Caption = .TiempoParaVotacion
        lblGrabarAutomaticamente.Caption = .GrabarAutomaticamente
        lblListarAutomaticamente.Caption = .ListarAutomaticamente
        lblActaGrabada.Caption = .ActaGrabada
        lblSolicituGrabarManual.Caption = .SolicitudGrabarManual
        lblEstadoSesion.Caption = .EstadoSesion
    End With
End Sub


Private Sub ActualizarVector_enBD()
    Dim X                 As Long
    Dim xMax              As Long
    Dim strColor          As String
    Dim strPresencia      As String
    Dim strIdentificacion As String
    Dim strResultados     As String
    Dim strAbstencion     As String
    Dim strCadenaVector   As String
    Dim strSql            As String
    ' ------------------------------------------------------------------------------------
    ' Armar string vector COLOR
    ' ------------------------------------------------------------------------------------
    strCadenaVector = ""
    xMax = UBound(EstadoActual.VectorColor)
    For X = 0 To xMax
        strCadenaVector = strCadenaVector & EstadoActual.VectorColor(X) & SEPARADOR_VECTOR
    Next X
    strCadenaVector = Trim(strCadenaVector)
    If Right(strCadenaVector, 1) = SEPARADOR_VECTOR Then
        strCadenaVector = Left(strCadenaVector, Len(strCadenaVector) - 1)
    End If
    strColor = strCadenaVector
    ' ------------------------------------------------------------------------------------
    ' Armar string vector PRESENCIA
    ' ------------------------------------------------------------------------------------
    strCadenaVector = ""
    xMax = UBound(EstadoActual.VectorPresencia)
    For X = 0 To xMax
        strCadenaVector = strCadenaVector & EstadoActual.VectorPresencia(X) & SEPARADOR_VECTOR
    Next X
    strCadenaVector = Trim(strCadenaVector)
    If Right(strCadenaVector, 1) = SEPARADOR_VECTOR Then
        strCadenaVector = Left(strCadenaVector, Len(strCadenaVector) - 1)
    End If
    strPresencia = strCadenaVector
    ' ------------------------------------------------------------------------------------
    ' Armar string vector IDENTIFICACION
    ' ------------------------------------------------------------------------------------
    strCadenaVector = ""
    xMax = UBound(EstadoActual.VectorIdentificacion)
    For X = 0 To xMax
        strCadenaVector = strCadenaVector & EstadoActual.VectorIdentificacion(X) & SEPARADOR_VECTOR
    Next X
    strCadenaVector = Trim(strCadenaVector)
    If Right(strCadenaVector, 1) = SEPARADOR_VECTOR Then
        strCadenaVector = Left(strCadenaVector, Len(strCadenaVector) - 1)
    End If
    strIdentificacion = strCadenaVector
    ' ------------------------------------------------------------------------------------
    ' Armar string vector RESULTADOS
    ' ------------------------------------------------------------------------------------
    strCadenaVector = ""
    xMax = UBound(EstadoActual.VectorResultados)
    For X = 0 To xMax
        strCadenaVector = strCadenaVector & EstadoActual.VectorResultados(X) & SEPARADOR_VECTOR
    Next X
    strCadenaVector = Trim(strCadenaVector)
    If Right(strCadenaVector, 1) = SEPARADOR_VECTOR Then
        strCadenaVector = Left(strCadenaVector, Len(strCadenaVector) - 1)
    End If
    strResultados = strCadenaVector
    
    ' ------------------------------------------------------------------------------------
    ' Armar string vector Abstencion
    ' ------------------------------------------------------------------------------------
    strCadenaVector = ""
    xMax = UBound(EstadoActual.VectorAbstencion)
    For X = 0 To xMax
        strCadenaVector = strCadenaVector & EstadoActual.VectorAbstencion(X) & SEPARADOR_VECTOR
    Next X
    strCadenaVector = Trim(strCadenaVector)
    If Right(strCadenaVector, 1) = SEPARADOR_VECTOR Then
        strCadenaVector = Left(strCadenaVector, Len(strCadenaVector) - 1)
    End If
    strAbstencion = strCadenaVector
        
    ' ------------------------------------------------------------------------------------
    ' Armar string SQL para actualizacion de registro
    ' ------------------------------------------------------------------------------------
    With EstadoActual
    
   ' strSql = "UPDATE vector SET " _
               & "Presentes = " & .Presentes & ", Ausentes = " & .Ausentes & ", vector_colores = '" & strColor & "', " _
               & "vector_presencia = '" & strPresencia & "', vector_identificacion = '" & strIdentificacion & "', " _
               & "Identificador_tipo_de_operacion = '" & .TipoDeOperacion & "', Afirmativos = " & CartelActual.Afirmativos & ", " _
               & "Resultado = '" & CartelActual.Resultado & "', Negativos = " & CartelActual.Negativos & ", Abstenciones = " & CartelActual.Abstenciones & ", " _
               & "Ocupadas_no_identificadas = " & .OcupadosNoIdentificados & ", Minimo_de_votos_para_afirmativa = " & CartelActual.MinimoVotosParaAfirmativo & ", " _
               & "Sesión = " & .Sesion & ", Nro_de_Acta = " & .NroActa & ", Titulo_del_Acta = '" & .TituloDelActa & "', " _
               & "Base_de_Mayoría = '" & .BaseMayoria & "', Tipo_de_Mayoría = '" & .TipoMayoria & "', Modo_identifica_nom_Obsoleto = " & .Modo_Ident_Nom & ", " _
               & "strError = '" & .strError & "', Estado_de_votacion = '" & .EstadoVotacion_y_PasList & "', Vector_resultado = '" & strResultados & "', " _
               & "Tipo_de_Abstención = '" & .TipoDeAbstencion & "', Mensaje_al_operador = '" & .MensajeAlOperador & "', Pendientes_Emitir_Voto = " & .PendientesEmitirVotos & ", " _
               & "Grabar_automaticamente = " & .GrabarAutomaticamente & ", Listar_automaticamente = " & .ListarAutomaticamente & ", " _
               & "Tipo_Mayoria_Quorum = '" & .TipoMayoriaQuorum & "', Leyenda_Quorum = '" & CartelActual.LeyendaQuorum & "', Período_Legislativo = '" & .PeriodoLegislativo & "', " _
               & "Fecha = '" & Format(Now, FORMATOFECHA&" hh:mm:ss") & "', Hora = '" & Time & "', Acta_Grabada = " & .ActaGrabada & ", Solicitud_Grabacion_Manual = " & .SolicitudGrabarManual & ", " _
               & "Tiempo_de_votación = " & .TiempoParaVotacion & ", IP_Consola_Habilitada = '" & .IP_Consola & "', Modo_Mantenimiento_Bancas = " & .ModoMantenimientoBancas & ", " _
               & "Modo_Normal_Mant_Sistema = " & .ModoNormalMantSistema & ", Identificador_de_Formulario = '" & .IdentificadorDeFormulario & "', Encender_Carteles = " & .CartelEncendido & ", Estado_Sesion = '" & .EstadoSesion & "', FechaVotacion = '" & Format(EstadoActual.FechaVotacion, FORMATOFECHA & " hh:mm:ss") & "', HoraVotacion = '" & EstadoActual.HoraVotacion & "'"
    ' Misma sentencia SQL, pero con SP
    'MsgBox .PeriodoLegislativos
    .IdentificadorDeFormulario = IIf(.Modo_Ident_Nom = 1, "1", "0") & IIf(.Modo_Presencia_Nom = 1, "1", "0")
    'Se emplea identificador formulario para guardar variables adicionales al vector. en caso de requerirse esta funcionalidad de formularios, se debera crear otra columna y mover esta actualización.
    strSql = "update_vector " & getPresentes & ", " & GetAusentes & ", '" & strColor & "', " & _
             "'" & strPresencia & "','" & strIdentificacion & "','" & .TipoDeOperacion & "', " & _
             "'" & CartelActual.Resultado & "', " & CartelActual.Afirmativos & ", " & CartelActual.Negativos & ", " & _
             " " & CartelActual.Abstenciones & ", " & .OcupadosNoIdentificados & ", " & CartelActual.MinimoVotosParaAfirmativo & ", " & _
             " " & .Sesion & ", " & .NroActa & ", '" & .TituloDelActa & "', " & _
             " '" & .EstadoSesion & "', '" & .BaseMayoria & "', '" & .TipoMayoria & "', " & _
             " " & .Modo_Ident_Nom & ", '" & .strError & "', '" & .EstadoVotacion_y_PasList & "', " & _
             " '" & strResultados & "', '" & .TipoDeAbstencion & "', " & .PendientesEmitirVotos & "," & _
             " '" & .MensajeAlOperador & "', " & .GrabarAutomaticamente & ", " & .ListarAutomaticamente & ", " & _
             " '" & .TipoMayoriaQuorum & "', '" & CartelActual.LeyendaQuorum & "', '" & .PeriodoLegislativo & "', " & _
             " '" & Format(Now, FORMATOFECHA & " hh:mm:ss") & "', '" & Format(Now(), "H:mm:ss") & "', " & .ActaGrabada & ", " & _
             " " & .SolicitudGrabarManual & ", " & .TiempoParaVotacion & ", '" & .IP_Consola & "', " & _
             " " & .ModoMantenimientoBancas & ", " & .ModoNormalMantSistema & " , '" & .IdentificadorDeFormulario & "', " & _
             " " & .CartelEncendido & ", '" & Format(EstadoActual.FechaVotacion, FORMATOFECHA & " hh:mm:ss") & "', '" & EstadoActual.HoraVotacion & "', '" & strAbstencion & "'" & ", " & .Reunion & ", " & "'" & .Orador & "', " & _
             " " & IIf(.ModoVotaPresidente, 1, 0) & _
             " " & "," & IIf(.Expresiones_Minoria, 1, 0)
    '")"
    End With
    Cn.Execute (strSql)
    DoEvents


End Sub
Private Sub AltasActas()
    On Error GoTo TrapError
    Dim strSql              As String
    Dim xMax                As Long
    Dim X                   As Long
    Dim xPresIdentificado   As Long
    Dim xPresNoIdentificado As Long
    Dim xAusentesTotales    As Long
    Dim xPresentesTotales   As Long
    strSql = "SELECT Tipo_de_operación, Período_Legislativo, Sesión, Número_de_Acta, Versión_Acta, Ultima_Versión_Acta, " _
           & "Nombre_del_Acta, Fecha, Hora, Tipo_de_Quorum, Base_de_Mayoria, Tipo_de_Mayoria, Miembros_del_cuerpo, " _
           & "Desempate, Votacion, Presidente, Presentes_Identificables, Presentes_No_Identificables, Presentes_Total, " _
           & "Ausentes_Total, Votos_Afirm_Identificables, Votos_Afirm_No_Identificables, Votos_Afirm_Desempate, " _
           & "Votos_Afirm_Total, Votos_Neg_Identificables, Votos_Neg_No_Identificables, Votos_Neg_Desempate, " _
           & "Votos_Neg_Total, Abstenciones_Identificables, Abstenciones_No_Identificables, Abstenciones_Total, " _
           & "Fecha_Modificacion, Usuario_Modificacion, Hora_Modificacion , IP_Modificacion, Observaciones " _
           & "From Actas"
    ' ---------------------------------------------------------------------------------------------------------
    ' Determinar Presentes identificables, no identificables, presentes totales, asusentes, etc
    ' ---------------------------------------------------------------------------------------------------------
    xMax = xUltimaBanca
    For X = 0 To xMax
        ' Contar presentes y ausentes
        If EstadoActual.VectorPresencia(X) = PRESENTE Then
            xPresentesTotales = xPresentesTotales + 1
        Else
            xAusentesTotales = xAusentesTotales - 1
        End If
        ' Contar identificados o no identificados
        If EstadoActual.VectorIdentificacion(X) = NO_IDENTIFICADO Then
            xPresNoIdentificado = xPresNoIdentificado + 1
        Else
            xPresIdentificado = xPresIdentificado + 1
        End If
    Next X
    Call SetearRs(strSql)
    With rs
        .AddNew
        .Fields("Tipo_de_operación").Value = EstadoActual.TipoDeOperacion
        .Fields("Período_Legislativo").Value = EstadoActual.PeriodoLegislativo
        .Fields("Sesión").Value = EstadoActual.Sesion
        .Fields("Número_de_Acta").Value = EstadoActual.NroActa
        .Fields("Versión_Acta").Value = ""
        .Fields("Ultima_Versión_Acta").Value = ""
        .Fields("Reunion").Value = EstadoActual.Reunion
        .Fields("Nombre_del_Acta").Value = ""
        .Fields("Fecha").Value = Date
        .Fields("Hora").Value = Time
        .Fields("Tipo_de_Quorum").Value = EstadoActual.TipoMayoriaQuorum
        .Fields("Base_de_Mayoria").Value = EstadoActual.BaseMayoria
        .Fields("Tipo_de_Mayoria").Value = EstadoActual.TipoMayoria
        .Fields("Miembros_del_cuerpo").Value = xMax
        .Fields("Desempate").Value = ""
        .Fields("Votacion").Value = ""
        .Fields("Presidente").Value = ""
        .Fields("Presentes_Identificables").Value = xPresIdentificado
        .Fields("Presentes_No_Identificables").Value = xPresNoIdentificado
        .Fields("Presentes_Total").Value = xPresentesTotales
        .Fields("Ausentes_Total").Value = xAusentesTotales
        .Fields("Votos_Afirm_Identificables").Value = ""
        .Fields("Votos_Afirm_No_Identificables").Value = ""
        .Fields("Votos_Afirm_Desempate").Value = ""
        .Fields("Votos_Afirm_Total").Value = ""
        .Fields("Votos_Neg_Identificables").Value = ""
        .Fields("Votos_Neg_No_Identificables").Value = ""
        .Fields("Votos_Neg_Desempate").Value = ""
        .Fields("Votos_Neg_Total").Value = ""
        .Fields("Abstenciones_Identificables").Value = ""
        .Fields("Abstenciones_No_Identificables").Value = ""
        .Fields("Abstenciones_Total").Value = ""
        .Fields("Fecha_Modificacion").Value = ""
        .Fields("Usuario_Modificacion").Value = ""
        .Fields("Hora_Modificacion").Value = ""
        .Fields("IP_Modificacion").Value = ""
        .Fields("Observaciones").Value = ""
        .Update
        .Close
    End With
Exit Sub
TrapError:
    Select Case err.Number
        Case Else
            Call AltaLogGeneral("SERVER SQV", "Alta actas Error Nº" & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source)
            'MsgBox "Error Nº" & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            End
    End Select
End Sub
Private Sub GuardarActas()  'OBSOLETO NO SE USA
    On Error GoTo TrapError
    Dim strSql              As String
    Dim xMax                As Long
    Dim X                   As Long
    Dim xPresIdentificado   As Long
    Dim xPresNoIdentificado As Long
    Dim xAusentesTotales    As Long
    Dim xPresentesTotales   As Long
    strSql = "SELECT Tipo_de_operación, Período_Legislativo, Sesión, Número_de_Acta, Versión_Acta, Ultima_Versión_Acta, " _
           & "Nombre_del_Acta, Fecha, Hora, Tipo_de_Quorum, Base_de_Mayoria, Tipo_de_Mayoria, Miembros_del_cuerpo, " _
           & "Desempate, Votacion, Presidente, Presentes_Identificables, Presentes_No_Identificables, Presentes_Total, " _
           & "Ausentes_Total, Votos_Afirm_Identificables, Votos_Afirm_No_Identificables, Votos_Afirm_Desempate, " _
           & "Votos_Afirm_Total, Votos_Neg_Identificables, Votos_Neg_No_Identificables, Votos_Neg_Desempate, " _
           & "Votos_Neg_Total, Abstenciones_Identificables, Abstenciones_No_Identificables, Abstenciones_Total, " _
           & "Fecha_Modificacion, Usuario_Modificacion, Hora_Modificacion , IP_Modificacion, Observaciones, Reunion " _
           & "From Actas"
    ' ---------------------------------------------------------------------------------------------------------
    ' Determinar Presentes identificables, no identificables, presentes totales, asusentes, etc
    ' ---------------------------------------------------------------------------------------------------------
    xMax = xUltimaBanca
    For X = 0 To xMax
        ' Contar presentes y ausentes
        If EstadoActual.VectorPresencia(X) = PRESENTE Then
            xPresentesTotales = xPresentesTotales + 1
        Else
            xAusentesTotales = xAusentesTotales - 1
        End If
        ' Contar identificados o no identificados
        If EstadoActual.VectorIdentificacion(X) = NO_IDENTIFICADO Then
            xPresNoIdentificado = xPresNoIdentificado + 1
        Else
            xPresIdentificado = xPresIdentificado + 1
        End If
    Next X
    Call SetearRs(strSql)
    With rs
        .AddNew
        .Fields("Tipo_de_operación").Value = EstadoActual.TipoDeOperacion
        .Fields("Período_Legislativo").Value = EstadoActual.PeriodoLegislativo
        .Fields("Sesión").Value = EstadoActual.Sesion
        .Fields("Número_de_Acta").Value = EstadoActual.NroActa
        .Fields("Versión_Acta").Value = ""
        .Fields("Ultima_Versión_Acta").Value = ""
        .Fields("Reunion").Value = EstadoActual.Reunion
        .Fields("Nombre_del_Acta").Value = ""
        .Fields("Fecha").Value = Date
        .Fields("Hora").Value = Time
        .Fields("Tipo_de_Quorum").Value = EstadoActual.TipoMayoriaQuorum
        .Fields("Base_de_Mayoria").Value = EstadoActual.BaseMayoria
        .Fields("Tipo_de_Mayoria").Value = EstadoActual.TipoMayoria
        .Fields("Miembros_del_cuerpo").Value = xMax
        .Fields("Desempate").Value = ""
        .Fields("Votacion").Value = ""
        .Fields("Presidente").Value = ""
        .Fields("Presentes_Identificables").Value = xPresIdentificado
        .Fields("Presentes_No_Identificables").Value = xPresNoIdentificado
        .Fields("Presentes_Total").Value = xPresentesTotales
        .Fields("Ausentes_Total").Value = xAusentesTotales
        .Fields("Votos_Afirm_Identificables").Value = ""
        .Fields("Votos_Afirm_No_Identificables").Value = ""
        .Fields("Votos_Afirm_Desempate").Value = ""
        .Fields("Votos_Afirm_Total").Value = ""
        .Fields("Votos_Neg_Identificables").Value = ""
        .Fields("Votos_Neg_No_Identificables").Value = ""
        .Fields("Votos_Neg_Desempate").Value = ""
        .Fields("Votos_Neg_Total").Value = ""
        .Fields("Abstenciones_Identificables").Value = ""
        .Fields("Abstenciones_No_Identificables").Value = ""
        .Fields("Abstenciones_Total").Value = ""
        .Fields("Fecha_Modificacion").Value = ""
        .Fields("Usuario_Modificacion").Value = ""
        .Fields("Hora_Modificacion").Value = ""
        .Fields("IP_Modificacion").Value = ""
        .Fields("Observaciones").Value = ""
        .Update
        .Close
    End With
Exit Sub
TrapError:
    Select Case err.Number
        Case Else
            Call AltaLogGeneral("SERVER SQV", "Guardar Actas Error Nº" & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source)
            'MsgBox "Error Nº" & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            End
    End Select
End Sub
Private Function LegisladorHabilitado(strIdLegislador As String) As Boolean
    Dim X As Long
    LegisladorHabilitado = False
    For X = 0 To xUltimaBanca '
       If Trim(LCase(EstadoActual.VectorIdentificacionHabilitados(X))) = strIdLegislador Then
          LegisladorHabilitado = True
       End If
    Next X
End Function


Private Sub EnviarMensajesComienzoAuth(xBanca As String, xComentario As String, Optional xModo As String) '< unifica llamadas de encencido del scanner
    Dim MsgSistema As MensajeSistema
    Dim nI As Long
    Dim xStrVectorHuella As String
    Dim xStrVectorTeclado As String
    Dim VectorBancas() As String
    ReDim VectorBancas(0 To xUltimaBanca)
    ' Prender scan nuevamente y solicitar identificacion, dependiendo
    ' Segun xModo, se trata de una reconexion preventiva o un pedido de inicio real.
    ' En modo start, se envia primero un scancl a la banca.
    
    xModo = "start" 'siempre el modo es start
    'Prepara mensaje en general, y luego revisa a que bancas se lo manda
    With MsgSistema
        .sObjeto = xBanca '<AP 040115 faltaba indicar la banca>
        .sTipo = "mset"
        .sComponente = "term.auth"
        .sAtributo = "action"
        If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
            .sComentario = xComentario & " Modo normal"
        Else
            .sComentario = xComentario & " Modo mantenimiento"
        End If
    End With
    'Caso general: es una banca individual o no hay identificaciones por teclado. Atiendo rapidamente todo el vector porque ninguno se identificac por teclado
    If (Len(Trim(xBanca)) <= 4 And Not (Trim(xBanca) = "brc")) Or Not (InStr(Join(EstadoActual.VTipoIdentificacion, SEPARADOR_VECTOR), TIPO_IDENTIFICACION_TECLADO) > 0) Then
        With MsgSistema
            .sObjeto = xBanca '<AP 040115 faltaba indicar la banca>
            If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
                .sValor = Trim("auth_" & xModo) 'comienzo normal
                If EstadoActual.VTipoIdentificacion(Val(xBanca)) = TIPO_IDENTIFICACION_TECLADO Then
                    .sValor = "auth_key_start"
                Else
                    .sValor = "auth_start"
                End If
            Else
                If Len(Trim(xBanca)) <= 4 Then 'es una banca sola hay que ver si es id. por teclado
                    If EstadoActual.VTipoIdentificacion(Val(xBanca)) = TIPO_IDENTIFICACION_TECLADO Then
                        .sValor = "auth_test_key_start"
                    Else
                        .sValor = "auth_test"
                    End If
                Else 'sino, si no es una banca sola, como por el "if" anterior no hay identificaciones por teclado, manda identificacion por huella, para cualquier sea el valor de xbanca
                        .sValor = "auth_test"
                End If
            End If
        End With
        Call EnviarMensajesBancas(MsgSistema)
    Else
        'No es una banca sola y ademas, hay al menos una identificacion por teclado, entonces hago dos comandos
        'uno para teclado y el otro para huella
        
        If Trim(xBanca) = "brc" Then ' si es brc, compone un vector con todos habilitados para recibir el comando, menos el presidente.
            VectorBancas(0) = "0"
            For nI = 1 To xUltimaBanca
                VectorBancas(nI) = "1"
            Next nI
        Else ' es un vector, entonces toma el vector directamente
            VectorBancas = Split(xBanca, SEPARADOR_VECTOR)
        End If

        xStrVectorHuella = "0" & SEPARADOR_VECTOR 'presidente no se identifica nunca
        xStrVectorTeclado = "0" & SEPARADOR_VECTOR 'presidente no se identifica nunca
        For nI = 1 To UBound(EstadoActual.VectorPresencia)
                xStrVectorHuella = xStrVectorHuella & IIf(VectorBancas(nI) = "0" Or EstadoActual.VTipoIdentificacion(nI) = TIPO_IDENTIFICACION_TECLADO, "0", "1") & SEPARADOR_VECTOR
                xStrVectorTeclado = xStrVectorHuella & IIf(VectorBancas(nI) = "1" And EstadoActual.VTipoIdentificacion(nI) = TIPO_IDENTIFICACION_TECLADO, "1", "0") & SEPARADOR_VECTOR
        Next nI
        MsgSistema.sObjeto = xStrVectorHuella
        MsgSistema.sComentario = "IDXH " & MsgSistema.sComentario
        Call EnviarMensajesBancas(MsgSistema)
        MsgSistema.sObjeto = xStrVectorTeclado
        MsgSistema.sComentario = "IDXT " & MsgSistema.sComentario
        Call EnviarMensajesBancas(MsgSistema)
    End If
End Sub

Private Sub EnviarMensajesFinAuth(xBanca As String, xComentario As String) '< unifica llamadas de encencido del scanner
    Dim MsgSistema As MensajeSistema
                            
    ' detener escaner y perder identificacion
    ' si se esta en modo normal o modo mantenimiento.-
    With MsgSistema
        .sObjeto = xBanca '<AP 040115 faltaba indicar la banca>
        .sTipo = "mset"
        .sComponente = "term.auth"
        .sAtributo = "action"
        If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
            .sValor = "auth_cancel"
            .sComentario = xComentario & " Modo normal"
        Else
            .sValor = "auth_cancel"
            '.sValor = "auth_test_cancel"
            .sComentario = xComentario & " Modo mantenimiento"
        End If
    End With
    Call EnviarMensajesBancas(MsgSistema)
End Sub

Private Sub EnviarMensajesActualizarAuth(xBanca As String, xComentario As String) '< unifica llamadas de actualizacion de datos legisladores
    Dim MsgSistema As MensajeSistema
                            
    With MsgSistema
        .sObjeto = xBanca
        .sTipo = "mset"
        .sComponente = "term.auth"
        .sAtributo = "action"
        If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
            .sValor = "auth_data_refresh"
            .sComentario = xComentario & " Modo normal"
        Else
            .sValor = "auth_test_refresh"
            .sComentario = xComentario & " Modo mantenimiento"
        End If
    End With
    Call EnviarMensajesBancas(MsgSistema)
End Sub

Private Function DigitoHexADec(charHex) As Long
    If charHex >= "0" And charHex <= "9" Then
        DigitoHexADec = Asc(charHex) - 48
    Else
        DigitoHexADec = Asc(charHex) - 55
    End If
End Function

Private Sub Identificacion(xMsj As MensajeSistema)

    Dim xBanca                As Long
    Dim MsgSistema            As MensajeSistema
    Dim xLecturasNegativas    As Long
    Dim xIntentosRealizados   As Long
    Dim strSql                As String
    Dim strIdLegislador       As String
    Dim strInforme            As String
    Dim X                     As Long
    Dim vTemporal             As Variant
    
    With xMsj
        xBanca = Int(.sObjeto)
        .sAtributo = Trim(LCase(.sAtributo))
        .sComponente = LCase(.sComponente)
        
        If xBanca >= 1 And xBanca <= xUltimaBanca Then ' Solo procesa mensajes de identificacion, cuando la SB censo presencia.
                If xBanca = 56 Then
                    xBanca = 56
                End If
            If EstadoActual.VectorPresencia(xBanca) = PRESENTE And .sComponente = "term.auth" Then
                'If Trim(LCase(.sAtributo)) = "sautod" And Trim(LCase(.sValor)) = "ok" Then
                    'el sautod viene por auth result
                If LCase(.sAtributo) = "result" Then
                    .sValor = Left(Trim(.sValor), 16)
                    If InStr(.sValor, "|") > 0 Then ' Puede recibirse en el Valor una serie de parametros de reintentos del usuario los cuales toma aqui.
                        'MsgBox "hablar con marcos para determinar como me van a llegar los mensajes de banca con
                        ' problemas para identificarse... y para ver este asunto de cuantos intentos de identificacion
                        ' tiene esa banca..."
                    End If '<AP 040115 abro el endif porque sino no procesaria negative con parametros>
                    If Trim(.sValor) = "negative" Then ' Se recibio un resultado negativo, es decir no lo identifico.
                        ' Ver si no fue identificado manualmente
                        EstadoActual.VectorColor(xBanca) = cMarronClaro
                        If EstadoActual.VectorIdentificacion(xBanca) = 0 Then
                            'Call EnviarMensajesComienzoAuth(Str(xBanca), "Negative")
                            ' En funcion de los parametros de sensibilidad de reintentos, se pone de color alarma
                            ' la banca, para llamar la atencion al operador respecto de que el Leg. esta teniendo problemas para identificarse.
                            If xLecturasNegativas >= xSensibilidadReintentos Then
                                EstadoActual.VectorColor(xBanca) = cVERDE
                                Call AltaLogGeneral("REINTENTO ID", Trim(.sValor) & "Banca Nro. " & Trim(Str(xBanca)), Str(xBanca), "1")    '<AP 040115 agrego log>
                            End If
                        End If
                    End If
                    If Trim(.sValor) = "timeout" Then
                        EstadoActual.VectorColor(xBanca) = cAMARILLO
                    End If
                    If xIntentosRealizados >= 1 And xLecturasNegativas = 0 And EstadoActual.VectorColor(xBanca) = cVERDE Then
                        'Desactivar indicador de reintentos anterior si dejo de intentar cuando tras al menos dos ciclos de scan no se observó ningún intento de identificacion negativa.
                        Call PintarVectorColor(xBanca)
                        Call AltaLogGeneral("Normalizacion desde REINTENTO ID", Trim(.sValor) & "Banca Nro. " & Trim(Str(xBanca)), Str(xBanca), "1")    '<AP 040115 agrego log>
                    End If
                    If Not (.sValor = "negative") And (Not .sValor = "timeout") Then ' Si se identificó correctamente al legislador
                        EstadoActual.EnIdentificacion(xBanca) = True
                        If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then ' y no estamos en modo Mantenimiento
                            ' Hay que verificar que el legislador se encuentre en al tabla de legisladores activos
                            ' strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                   & "Legisladores.departamento , Legisladores.cargo FROM Legisladores INNER JOIN legisladores_activos ON " _
                                   & "Legisladores.id = legisladores_activos.id WHERE (Legisladores.id = '" & CerosIzquierda(.sValor, 8) & "' and Legisladores.tipo = 1)" '<AP 040115 Pide que sea tipo legislador
                            
                            ' TECLADO: si se trata de una identificacion por teclado
                            If True Then ' implementacion cordoba 03
                                vTemporal = Val("&H" & Trim(.sValor))
                            Else
                                vTemporal = Val(Left(Trim(.sValor), 16))
                                
                            End If
                            If False And vTemporal > 99999 Then ' es identificacion por teclado y vino el numero de PIN (revisar)
                                .sValor = Trim(Str(vTemporal))
                                .sValor = CerosIzquierda(.sValor, 8)
                                .sValor = Encripta.EncryptString(.sValor)
                                strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                       & "Legisladores.departamento , Legisladores.cargo FROM Legisladores INNER JOIN legisladores_activos ON " _
                                       & "Legisladores.id = legisladores_activos.id WHERE (Legisladores.Pin = '" & Trim(.sValor) & "') AND (Legisladores.tipo = 1)"  '<AP 040115 Pide que sea tipo legislador
                            Else
                                If Not IsNumeric(.sValor) Then
                                    If vTemporal > 0 Then
                                        .sValor = vTemporal
                                    Else
                                        .sValor = "999999"
                                    End If
                                Else
                                    .sValor = vTemporal
                                End If
                                strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                       & "Legisladores.departamento , Legisladores.cargo FROM Legisladores INNER JOIN legisladores_activos ON " _
                                       & "Legisladores.id = legisladores_activos.id WHERE Cast(Legisladores.id AS Int) = " & Str(Int(.sValor)) & " and Legisladores.tipo = 1"  '<AP 040115 Pide que sea tipo legislador
                            End If
                        Else ' Si se encuentra en modo mantenimiento
                            '< AP 041007a manejo id por teclado en mantenimiento
                            If False Then ' implementacion cordoba 03
                                vTemporal = Val("&H" & .sValor)
                                If vTemporal > 99999 Then ' es identificacion por teclado y vino el numero de PIN (revisar)
                                    .sValor = Trim(Str(vTemporal))
                                    .sValor = CerosIzquierda(.sValor, 8)
                                    .sValor = Encripta.EncryptString(.sValor)
                                    strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                           & "Legisladores.departamento , Legisladores.cargo FROM Legisladores " _
                                           & "WHERE (Legisladores.Pin = '" & Trim(.sValor) & "') AND (Legisladores.tipo = 0)"  '<AP 040115 Pide que sea tipo personal mantenimiento
                                Else
                                    If Not IsNumeric(.sValor) Then
                                        .sValor = "999999"
                                    End If
                                    ' IDs recibidos = TRIM (.sValor) & ';' & TRIM (IDs recibidos)
                                    ' <AP 040115 Hay que verificar que el legislador se encuentre como personal de mantenimiento, sin join
                                    ' strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                           & "Legisladores.departamento , Legisladores.cargo FROM Legisladores WHERE Legisladores.id  = '" & CerosIzquierda(.sValor, 8) & "' and Legisladores.tipo = 0)" '<AP 040115 Pide que sea tipo personal de mantenimiento
                                    strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                           & "Legisladores.departamento , Legisladores.cargo FROM Legisladores WHERE Cast(Legisladores.id AS Int) = " & Str(Int(.sValor)) & " and Legisladores.tipo = 0" '<AP 040115 Pide que sea tipo personal de mantenimiento
                                End If
                            Else
                                vTemporal = Val("&H" & (Trim(.sValor)))
                                If Not IsNumeric(.sValor) Then
                                    If vTemporal > 0 Then
                                        .sValor = vTemporal
                                    Else
                                        .sValor = "999999"
                                    End If
                                Else
                                    .sValor = vTemporal
                                End If
                                ' con banca asignada
                                'strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                '       & "Legisladores.departamento , Legisladores.cargo FROM Legisladores INNER JOIN legisladores_activos ON " _
                                '       & "Legisladores.id = legisladores_activos.id WHERE Cast(Legisladores.id AS Int) = " & Str(Int(.sValor)) & " and Legisladores.tipo = 0"  '<AP 040115 Pide que sea tipo legislador
                                'sin banca asignada
                                strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                       & "Legisladores.departamento , Legisladores.cargo FROM Legisladores " _
                                       & "WHERE Cast(Legisladores.id AS Int) = " & Str(Int(.sValor)) & " and Legisladores.tipo = 0"  '<AP 040115 Pide que sea tipo legislador
                            End If
                            
                            'Arma la lista de mantenimientos
                            EstadoActual.MantIdentificaciones = Trim(.sValor) & ";" & Trim(EstadoActual.MantIdentificaciones)
                        
                        End If
                            ' Call SetearRs(strSql)
                            'If Rs.State = 1 Then
                            '    RsLocal.Close
                            'End If
                            'RsLocal.CursorLocation = adUseClient
                            
                            RsLocal.Open strSql, Cn, adOpenForwardOnly, adLockOptimistic
                            If RsLocal.RecordCount <= 0 Or RsLocal.EOF Or RsLocal.BOF Then ' Si NO es un legislador activo o personal de mant.
                                ' RsLocal.Close
                                strIdLegislador = ""
                                'Reenciende el scanner
                                'Call EnviarMensajesComienzoAuth(Str(xBanca), "ID Invalido") 'Unifica envio mensaje comienzo autorizacion
                                'Call MensajeDisplayTerminal(Str(xBanca), "Id. invalida:" & strIdLegislador & " Por favor reintente.")
                                'MsgBox "id invalido" & .sValor & xBanca
                            Else ' En cambio, SI ES un legislador activo o personal de mant.
                                strIdLegislador = Trim(LCase(RsLocal.Fields("id").Value)) ' Levanto Id de Legisl.
                                xNombreUltimoIdentificado = Trim(RsLocal.Fields("apellido").Value)
                                For X = 1 To xUltimaBanca ' y me fijo si se identifico anteriormente en otra banca
                                    If Trim(LCase(EstadoActual.VectorIdentificacion(X))) = strIdLegislador Then
                                        flExitoPierdeIdDup = True
                                        xBancaDuplicada = X ' Tomo el numero de la banca duplicada
                                    End If
                                Next X
                                ' Si no se identifico anteriormente en otra banca, pongo ID de legislador en Vector Identificacion
                                If Not (flExitoPierdeIdDup) Then ' identificar al legislador en vector identificacion
                                    'Verifica que no coincida con el presidente
                                    If Trim(LCase(EstadoActual.VectorIdentificacion(0))) = strIdLegislador Then
                                        flExitoPierdeIdDupConPresdte = True
                                        xBancaDuplicada = 0
                                        ' Guardo log completo de lo sucedido
                                        strInforme = "LEGISLADOR: " & RsLocal.Fields("Apellido").Value & ", " & RsLocal.Fields("Nombre").Value & ", " & RsLocal.Fields("bloque_politico").Value & ", " & RsLocal.Fields("departamento").Value & ", " & RsLocal.Fields("Cargo").Value & ", se intento identificar en la banca " & Str(xBanca) & ", ya esta identificado como Presidente"
                                        Call AltaLogGeneral("BANCA DUPLICADA PRESIDENTE", strInforme, Str(xBanca), "5")
                                        EstadoActual.VectorColor(xBanca) = cROJO ' Avisar al operador lo que esta pasando
                                        EstadoActual.VectorColor(xBancaDuplicada) = cROJO ' <AP 040115 Ambas bancas en rojo
                                        'Solo al ultimo le reenciende el scanner
                                        Call EnviarMensajesComienzoAuth(Str(xBanca), "Banca Duplicada - ultimo intento") 'Unifica envio mensaje comienzo autorizacion
                                    Else
                                        '>> Verifica que este habilitado para el caso de las votaciones de reconsideracion.
                                        ' En las otras situaciones siempre estan todos habilitados.
                                        If EstadoActual.VectorIdentificacion(xBanca) = 0 And ((EstadoActual.ModoMantenimientoBancas = 1 Or EstadoActual.ModoNormalMantSistema = 1) Or (True Or (LegisladorHabilitado(strIdLegislador)))) Then ' revisar el indice para caso de reconsideracion ATENCION REVISAR 090518
                                            'La identificacion ha sido exitosa!
                                            EstadoActual.VectorIdentificacion(xBanca) = strIdLegislador
                                            EstadoActual.EnIdentificacion(xBanca) = False
                                            EstadoActual.OcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados - 1
                                                If Not LCase(.sComentario) = "sautod" Then
                                                     MsgSistema.sObjeto = xBanca '<AP 040115 faltaba indicar la banca>
                                                     MsgSistema.sTipo = "mset"
                                                     MsgSistema.sComponente = "term.led1"
                                                     MsgSistema.sAtributo = "state"
                                                     MsgSistema.sValor = "on"
                                                     If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
                                                         MsgSistema.sComentario = "Id aceptado Modo normal - led1 - B"
                                                     Else
                                                         MsgSistema.sComentario = "Id aceptado Modo mantenimiento"
                                                     End If
                                                     Call EnviarMensajesBancas(MsgSistema)
                                                 End If
                                                                                                  
                                              If EstadoActual.TipoDeOperacion = "votnom" Then
                                                If InStr(Join(EstadoActual.VectorAbstencion, SEPARADOR_VECTOR), Trim(strIdLegislador)) > 0 Then
                                                    AbstenerBanca (xBanca)
                                                End If
                                              End If

                                              ' envia mensaje por display... opcional
                                              With MsgSistema
                                                   MsgSistema.sObjeto = xBanca '<AP 040115 faltaba indicar la banca>
                                                   MsgSistema.sTipo = "mset"
                                                   MsgSistema.sComponente = "term.display"
                                                   MsgSistema.sAtributo = "text"
                                                   If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
                                                       MsgSistema.sValor = "Identificacion Aceptada"
                                                       MsgSistema.sComentario = "Id aceptado Modo normal"
                                                   Else
                                                       MsgSistema.sValor = "Identificacion de Prueba"
                                                       MsgSistema.sComentario = "Id aceptado Modo mantenimiento"
                                                   End If
                                              End With
                                              Call EnviarMensajesBancas(MsgSistema)
                                              
                                              flBancaIdentifPosExitosa = True
                                              Call PintarVectorColor(xBanca)
                                        Else
                                            Call MensajeDisplayTerminal(Str(xBanca), "Reconsideracion: No habilitado.")
                                            Call AltaLogGeneral("Identificacion", "Id no habilitado: " & strIdLegislador & ", Banca " & xBanca, Str(xBanca), "2")
                                        End If
                                    End If
                                Else ' Si el legislador ya esta identificado en otra banca: BANCA DUPLICADA!
                                    If xBancaDuplicada <> xBanca Then ' Cancelo indentificacion de las dos bancas
                                        'Cancela identificacion banca duplicada anterior. La banca que se intenta identificar ya queda sin identificar
                                        If EstadoActual.VectorIdentificacion(xBancaDuplicada) = strIdLegislador Then
                                             EstadoActual.VectorIdentificacion(xBancaDuplicada) = NO_IDENTIFICADO
                                             EstadoActual.OcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados + 1
                                        Else
                                            flExitoPierdeIdDup = False
                                        End If
                                        If flExitoPierdeIdDup = True Then
                                            'Elimine el fornext y deje la instruccion anterior porque ya tengo xbancaduplicada
                                            'Next x
                                            'For x = 1 To xUltimaBanca
                                            '    If EstadoActual.VectorIdentificacion(x) = strIdLegislador Then
                                            '        EstadoActual.VectorIdentificacion(x) = NO_IDENTIFICADO
                                            '    End If
                                            'Next x
                                            ' Guardo log completo de lo sucedido
                                            strInforme = "LEGISLADOR: " & RsLocal.Fields("Apellido").Value & ", " & RsLocal.Fields("Nombre").Value & ", " & RsLocal.Fields("bloque_politico").Value & ", " & RsLocal.Fields("departamento").Value & ", " & RsLocal.Fields("Cargo").Value & ", se intento identificar en la banca " & Str(xBanca) & ", ya estando identificado en la banca " & Str(xBancaDuplicada)
                                            Call AltaLogDuplicada("BANCA DUPLICADA Ultima", strInforme, Str(xBanca), "5")
                                            Call AltaLogDuplicada("BANCA DUPLICADA Primera", strInforme, Str(xBancaDuplicada), "5")
                                            EstadoActual.VectorColor(xBanca) = cROJO ' Avisar al operador lo que esta pasando
                                            EstadoActual.VectorColor(xBancaDuplicada) = cROJO ' <AP 040115 Ambas bancas en rojo
                                            'A ambos les reenciende el scanner
                                            Call EnviarMensajesComienzoAuth(Str(xBanca), "Banca Duplicada - ultimo intento") 'Unifica envio mensaje comienzo autorizacion
                                            Call EnviarMensajesComienzoAuth(Str(xBancaDuplicada), "Banca Duplicada - anterior id") 'Unifica envio mensaje comienzo autorizacion
                                            Dim identifConsulta As String
                                            identifConsulta = "INSERT INTO LogIdentificaciones(banca,id_diputado,fecha,hora,duplicidad) VALUES (" & Str(xBanca) & _
                                            ",'" & RsLocal.Fields("id").Value & "','" & Format(Now(), "YYYYMMDD") & "','" & Format(Now(), "HH:mm") & "'," & Str(xBancaDuplicada) & ")"
                                            Call EjecutarSQL(identifConsulta)
                                        End If
                                    End If
                                    PintarVectorColor (1)
                                End If
                            End If
                            RsLocal.Close
                    End If ' .sValor <> "negative" (si fue negative, se trato en if anterior.)
                End If ' FIN .sAtributo = "result"
            End If 'term.auth
            If LCase(.sComponente) = "term.seat" Then
                If ModoMant Then
                    VectorDesconectadas(xBanca) = False
                End If
                If LCase(.sAtributo) = "switch" And flSwitchExitoso = True Then
                    If LCase(.sValor) = "closed" Then
                        'If EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO And EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis" Or EstadoActual.Modo_Ident_Nom Then
                        If EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO Then
                                If ((EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis" Or EstadoActual.Modo_Ident_Nom = 1)) Then
                                    If ((((DateDiff("s", EstadoActual.FechaVotacion, Now)) < EstadoActual.TiempoParaVotacion + xSegundosFinOperacion Or EstadoActual.EstadoVotacion_y_PasList = "espera") _
                                        And EstadoActual.ExtensionDeTiempoPorPresidente = False And Not EstadoActual.EstadoVotacion_y_PasList = "finalizada")) Then
                                        EstadoActual.OcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados + 1
                                    End If
                                End If
                                If ModoMant = True Then
                                    Call EnviarMensajesComienzoAuth(Str(xBanca), "SW Closed")
                                ElseIf ModoMant = False And (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis" Or EstadoActual.Modo_Ident_Nom = 1) And _
                                    ((((DateDiff("s", EstadoActual.FechaVotacion, Now)) < EstadoActual.TiempoParaVotacion + xSegundosFinOperacion Or EstadoActual.EstadoVotacion_y_PasList = "espera") _
                                    And EstadoActual.ExtensionDeTiempoPorPresidente = False And Not EstadoActual.EstadoVotacion_y_PasList = "finalizada")) Then
                                    If xBanca <> 0 And EstadoActual.EstadoVotacion_y_PasList <> "esperafin" Then
                                        If EstadoActual.VectorPresencia(xBanca) = PRESENTE And EstadoActual.VectorIdentificacion(xBanca) = "0" Then
                                            Call EnviarMensajesComienzoAuth(Str(xBanca), "SW Closed")
                                        End If
                                    End If
                                End If
                        End If 'no id
                    End If 'closed
                    If LCase(.sValor) = "open" Then '>> El legislador se levanta, pierde la identificacion.
                        If Not (EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO) Then
                            flExitoPierdeID = True
                        Else
                            EstadoActual.OcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados - 1
                        End If 'no id
                        EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO 'aca
                        Call EnviarMensajesFinAuth(Str(xBanca), "SW Open")  'apagar scanner
                        ' el led se debe apagar solo
                    End If 'open
                End If 'switch
            End If 'seat
        ElseIf xBanca = 0 Then 'ACK VOTO PRESIDENTE habilitado
            'Identificacion del presidente. Solo con el objetivo de procesar la habilitación del presidente para votar, prendiendole el led1.
            If EstadoActual.ModoVotaPresidente Then
                If InStr(UCase(xMsj.sComentario), "SAUTOD") > 0 Then 'cuando se recibe un tidval con comentario sautod en el SB, corresponde al caso de la habilitacion del presidente para votar, porque la consola no deja identificar por operador al presidente (solo lo hace por la funcion de cambio de presidente)
                    
                    EstadoActual.PresidenteHabilitadoParaVotar = True
                    Call ComenzarVotacionPresidente
                End If
            End If
        End If 'banca valida
    End With 'xMsj
    
    
End Sub
Private Sub InicializarVotacion()
Dim X As Long

    CartelActual.Resultado = ""
    CartelActual.Afirmativos = 0
    CartelActual.Negativos = 0
    CartelActual.Abstenciones = 0
    EstadoActual.OcupadosNoIdentificados = CuentaOcupadosNoIdentificadosCong
    EstadoActual.EstadoVotacion_y_PasList = "espera"
    EstadoActual.PendientesEmitirVotos = 0
    EstadoActual.VectorIdentificacionHabilitados = xVectorIdentificacionHabilitados
    CartelActual.LeyendaTiempo = ""
    EstadoActual.VectorIdentificacionCong = EstadoActual.VectorIdentificacion
    EstadoActual.VectorResultadosCong = EstadoActual.VectorResultados
    Call CalcularMinimoAfirmativaCartel
    EstadoActual.AbstencionistasAutorizados = 0
    EstadoActual.ActaGrabada = 0
    For X = 0 To (xUltimaBanca)
        EstadoActual.VectorResultados(X) = ABSTENCION
    Next X
    EstadoActual.ModoVotaPresidente = False 'hcdn110218
    EstadoActual.ResultadoVotoPresidente = " "
    EstadoActual.EsperarVotoPresidente = False
    EstadoActual.PresidenteHabilitadoParaVotar = False
    
End Sub
Private Sub CalcularMinimoAfirmativaCartel()
    CartelActual.MinimoVotosParaAfirmativo = 0
    If (EstadoActual.BaseMayoria = "legpre" Or EstadoActual.BaseMayoria = "votemi") And Not (CartelActual.LeyendaQuorum = "QUORUM") Then
        CartelActual.LeyendaMinimoVotosParaAfirmativo = "N/D"
    Else
        'Call CalculoResultado(IIf(EstadoActual.BaseMayoria = "votemi", "legpre", EstadoActual.BaseMayoria), EstadoActual.TipoMayoria, xMiembrosDelCuerpo, CartelActual.Presentes, 0, 0, "", 0, 0, CartelActual.MinimoVotosParaAfirmativo, " ", IIf(xPresidenteLegislador, 1, 0))
        Call CalculoResultado(IIf(EstadoActual.BaseMayoria = "votemi", "legpre", EstadoActual.BaseMayoria), EstadoActual.TipoMayoria, xMiembrosDelCuerpo, CartelActual.Presentes, 0, 0, "", 0, 0, CartelActual.MinimoVotosParaAfirmativo, " ", IIf(EstadoActual.PresidenteHabilitadoParaVotar, 1, 0))
        CartelActual.LeyendaMinimoVotosParaAfirmativo = Str(CartelActual.MinimoVotosParaAfirmativo)
    End If
End Sub

Private Sub ReinicioSistema()
    
    Dim Mensaje2Banca As MensajeSistema
    
    Call InicializarVotacion
    Call ResetearVectores
    With EstadoActual
        .LimpiarResultados = 0
        .TipoDeOperacion = "quorum"
        .EstadoVotacion_y_PasList = "espera"
        
        .Presentes = 0 '1
        .Ausentes = xMiembrosDelCuerpo - .Presentes
        .PresentesCongelados = 0
        .AusentesCongelados = xMiembrosDelCuerpo
                
        .ActaGrabada = 0
        
        .OcupadosNoIdentificados = 0
        .PendientesEmitirVotos = 0
        .AbstencionistasAutorizados = 0
        .ModoVotaPresidente = False
        .ResultadoVotoPresidente = " "
        .EsperarVotoPresidente = False
        .PresidenteHabilitadoParaVotar = False
    End With
    With CartelActual
        .Abstenciones = 0
        .Afirmativos = 0
        .LeyendaQuorum = "NO HAY QUORUM"
        .MinimoVotosParaAfirmativo = 0
        .Negativos = 0
        .Presentes = 1
        .Ausentes = xMiembrosDelCuerpo - .Presentes
        .Resultado = 0
    End With
    Call CalcularMinimoParaQuorum
    Call PintarTodasLasBancas
    Call ActualizarVector_enBD
'    With Mensaje2Banca ' Mensaje para SB
'        .sTipo = "mget"
'        .sComponente = "term.mon"
'        .sObjeto = "brc"
'        .sComentario = "Reinicio Sistema"
'        .sAtributo = "action"
'        .sValor = "reset"
'    End With
'    Call EnviarMensajesBancas(Mensaje2Banca)
    With Mensaje2Banca ' Mensaje para SB
        .sTipo = "mget"
        .sComponente = "term"
        .sObjeto = "brc"
        .sAtributo = "state"
        .sValor = ""
    End With
    Call EnviarMensajesBancas(Mensaje2Banca)
End Sub



Private Function ArmarHabilitadosDeActa(pSesion As Long, pActa As Long) As Boolean
    ' Actualizar vector habilitados con los id de los legisladores que NO figuren como ausentes en la sesion indicada en el periodo legislativo actual.
    ' Devolver true si da ok.

    On Error GoTo TrapError
    Dim pPeriodo                As String
    Dim pVersion                As Long
    Dim rstAux                  As New ADODB.Recordset
    Dim strSql                  As String
    Dim i                       As Long
    Dim j                       As Long
    Dim xHacerLogHabilitados    As Boolean
    Dim xLogHabilitados         As String
    Dim xCantidadHabilitados    As Long
    Dim xStrVector              As String
    
    
    
    
    ArmarHabilitadosDeActa = False 'Error
    xHacerLogHabilitados = True
    xLogHabilitados = ";"
    xCantidadHabilitados = 0
    pPeriodo = EstadoActual.PeriodoLegislativo
    pVersion = 0
    
    strSql = "SELECT *, rtrim(Legisladores.nombre) + ', '+ rtrim(Legisladores.nombre) as legislador " _
            & " FROM detalleactas LEFT OUTER JOIN Legisladores ON detalleactas.Legislador_asignado = Legisladores.id " _
            & " WHERE (Período_Legislativo='" & pPeriodo & "') AND (Sesión=" & pSesion & ") AND (Nro_de_Acta=" & pActa & ") AND (Versión_Acta= " & pVersion & " ) " _
            & " ORDER BY Numero_de_banca"
    SetearRsAux strSql, rstAux
    If rstAux.EOF = False Then
        i = 0
        Do While Not (rstAux.EOF)
            If Trim(rstAux!Resultado) <> "AUSENTE" Then
                EstadoActual.VectorIdentificacionHabilitados(i) = rstAux!legislador_asignado
                xCantidadHabilitados = xCantidadHabilitados + 1
            Else
                EstadoActual.VectorIdentificacionHabilitados(i) = NO_IDENTIFICADO
            End If
            xLogHabilitados = xLogHabilitados & Trim(EstadoActual.VectorIdentificacionHabilitados(i)) & ","
            
            i = i + 1
            rstAux.MoveNext
        Loop
        If i <= xUltimaBanca Then
            For j = i To xUltimaBanca
                    EstadoActual.VectorIdentificacionHabilitados(j) = NO_IDENTIFICADO
                    If xHacerLogHabilitados Then
                        xLogHabilitados = xLogHabilitados & EstadoActual.VectorIdentificacionHabilitados(j) & ","
                    End If
            Next j
        End If
        'Se deben cancelar las identificaciones de los que no estan habilitados
        xStrVector = ""
        For i = 0 To xUltimaBanca
            If EstadoActual.VectorPresencia(i) = PRESENTE And Not (EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO) And Not (InStr(xLogHabilitados, ";" & Trim(EstadoActual.VectorIdentificacion(i)) & ";") > 0) Then
                EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO
                EstadoActual.VectorColor(i) = AsignarColor(i)
                EstadoActual.OcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados + 1
                xStrVector = xStrVector & "1" & SEPARADOR_VECTOR
            Else
                xStrVector = xStrVector & "0" & SEPARADOR_VECTOR
            End If
        Next i
        Call EnviarMensajesComienzoAuth(xStrVector, "Reidentificarse por no permitir identificacion por Votacion de Reconsideracion")
        If xHacerLogHabilitados Then
            Call AltaLogGeneral("SQV Server Reconsideracion", "Habilitados: " & xCantidadHabilitados & "Lista: " & xLogHabilitados)
        End If
        ArmarHabilitadosDeActa = True
    Else
        ArmarHabilitadosDeActa = False
        'MsgBox "Ha ocurrido un error al recuperar el detalle del acta.", vbInformation + vbOKOnly
        Call AltaLogGeneral("SQV Server Reconsideracion", "Ha ocurrido un error al recuperar el detalle del acta." & vbInformation & vbOKOnly)
    End If
Exit Function
TrapError:
    Select Case err.Number
        Case Else
            Call AltaLogGeneral("SQV Server Reconsideracion", "Ha ocurrido un error al recuperar el detalle del acta." & vbInformation & vbOKOnly)
            ArmarHabilitadosDeActa = False 'MsgBox "Ha ocurrido un error al recuperar el detalle del acta.", vbInformation + vbOKOnly
            Resume Next
    End Select
Return
End Function
Private Sub AltaLogGeneral(strOrigen As String, strDetalle As String, Optional nObjeto As String = "", Optional strSeveridad As String = "0")
    ' Alta de Log se hace a traves de Stored Procedure en Base de datos
    Dim strSP As String
'    If strSeveridad >= cSEVERIDAD_MINIMA Then
'        strSP = "insert_log_general('" & strOrigen & "','" & Trim(strDetalle) & " SB " & Trim(strUltimoMensaje_SB_SQV) & " SQV " & Trim(strUltimoMensaje_SQV_SB) & "','" & nObjeto & "','" & strSeveridad & "')"
'        Cn.Execute strSP
'    End If
End Sub
Private Sub AltaLogDuplicada(strOrigen As String, strDetalle As String, Optional nObjeto As String = "", Optional strSeveridad As String = "0")
    Dim strSP As String
    strSP = "insert_log_general2('" & strOrigen & "','" & Trim(strDetalle) & " SB " & Trim(strUltimoMensaje_SB_SQV) & " SQV " & Trim(strUltimoMensaje_SQV_SB) & "','" & nObjeto & "','" & strSeveridad & "')"
    Cn.Execute strSP
End Sub
Private Sub PintarBancasCartel()
    On Error Resume Next
    Dim i     As Integer
    Dim clave As String
    'busco los datos del presi sólo si cambia
    For i = 0 To UBound(EstadoActual.VectorColor)
        'clave = i
        ctrBanca(i).FillColor = mColores(Val(EstadoActual.VectorColor(i)))
        lblBanca(i).ForeColor = mColoresFuente(Val(EstadoActual.VectorColor(i)))
'        If ctrBanca(i).BackColor = &H0 Then
'            ctrBanca(i).ForeColor = &HE0E0E0
'        End If
    Next i
End Sub
Private Sub cargarColores()
    Dim Color As String
    Dim clave As String
    Dim i As Long
    ReDim mColores(0 To 7)
    i = 0
    'cargo el diccionario de manera estática
    'GRIS
    Color = "&HC0C0C0"
    clave = "0"
    mColores(i) = Color
    i = i + 1
    'BLANCO
    Color = "&HFFFFFF"
    clave = "1"
    mColores(i) = Color
    i = i + 1
    'AMARILLO
    Color = "&HFFFF"
    clave = "2"
    mColores(i) = Color
    i = i + 1
    'ROJO
    Color = "&HFF"
    clave = "3"
    mColores(i) = Color
    i = i + 1
    'CELESTE
    Color = "&HFFFF00"
    clave = "4"
    mColores(i) = Color
    i = i + 1
    'NARANJA
    Color = "&H80FF"
    clave = "5"
    mColores(i) = Color
    i = i + 1
    'VERDE
    Color = "&HFF00"
    clave = "6"
    mColores(i) = Color
    i = i + 1
    'NEGRO
    Color = "&H0"
    clave = "7"
    mColores(i) = Color
    i = i + 1
End Sub

Private Sub ArmarBancasCartel()
    'Dim xBanca As Long
    'For xBanca = 0 To xUltimaBanca
    '    ctrBanca(xBanca).Caption = xBanca
    'Next xBanca
    
    
    
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
    Dim xUltimaBanca As Integer
    xUltimaBanca = 46
    Call AsignarFuentes
    xCentroIzquierdo = (imgC(0).Width / 2) - 400
    xCentroDerecho = (imgC(0).Width / 2) + 400
    yCentro = imgC(0).top + imgC(0).Height - 1000
    pi = 4 * Atn(1)   ' Calculo el valor de pi
    picC(1).Cls
    picC(1).Refresh
    picC(0).Cls
    picC(0).Refresh
    For xBanca = 0 To xUltimaBanca
        ctrBanca(xBanca).Visible = True
        lblBanca(xBanca).Visible = True
        ctrBanca(xBanca).Width = 800
        ctrBanca(xBanca).Height = 800
        ctrBanca(xBanca).FillStyle = 0
        If xBanca >= 1 And xBanca <= 22 Then
            radio = (imgC(0).Width / 2) * 0.85
            Step = pi / 21
            xCentro = IIf(xBanca < 12, xCentroIzquierdo, xCentroDerecho)
            offset = 1
        ElseIf xBanca >= 23 And xBanca <= 42 Then
            radio = (imgC(0).Width / 2) * 0.7
            Step = pi / 19
            xCentro = IIf(xBanca < 33, xCentroIzquierdo, xCentroDerecho)
            offset = 23
        ElseIf xBanca >= 43 And xBanca <= 58 Then
            radio = (imgC(0).Width / 2) * 0.55
            Step = pi / 15
            xCentro = IIf(xBanca < 51, xCentroIzquierdo, xCentroDerecho)
            offset = 43
        ElseIf xBanca >= 59 And xBanca <= 70 Then
            radio = (imgC(0).Width / 2) * 0.4
            Step = pi / 11
            xCentro = IIf(xBanca < 65, xCentroIzquierdo, xCentroDerecho)
            offset = 59
        End If
        
        If xBanca = 0 Then
            xObjeto = imgC(0).Width / 2
            yObjeto = yCentro
        Else
            xObjeto = xCentro - Cos(Step * (xBanca - offset)) * radio
            yObjeto = yCentro - Sin(Step * (xBanca - offset)) * radio
        End If
        ctrBanca(xBanca).Left = xObjeto - (ctrBanca(xBanca).Width / 2)
        ctrBanca(xBanca).top = yObjeto - (ctrBanca(xBanca).Height / 2)
        ctrBanca(xBanca).FillColor = MiBlanco
        lblBanca(xBanca).Caption = xBanca
        lblBanca(xBanca).AutoSize = True
        lblBanca(xBanca).FontSize = 28
        lblBanca(xBanca).Left = xObjeto - (lblBanca(xBanca).Width / 2)
        lblBanca(xBanca).top = yObjeto - (lblBanca(xBanca).Height / 2)
        ctrBanca(xBanca).ZOrder 0
    Next
    For xBanca = IIf(cFORMULARIO_MOSTRAR_BANCAS, xUltimaBanca + 1, 0) To 70
        lblBanca(xBanca).Visible = False
        ctrBanca(xBanca).Visible = False
    Next
    imgC(0).ZOrder 1
End Sub


Private Function LeyendaSesion(nLinea As Integer) As String
    Const ORDINAL_MASCULINO = "°"
    Const ORDINAL_FEMENINO = "ª"
    Dim strEtiqueta      As String
    Dim strEtiquetaLinea1     As String
    Dim strEtiquetaLinea2     As String
    Dim strEtiquetaLinea3     As String

    strEtiquetaLinea1 = Left(EstadoActual.PeriodoLegislativo, 3) & ORDINAL_MASCULINO & " - Periodo "
    strEtiqueta = strEtiquetaLinea1
    
    Select Case UCase(Mid(EstadoActual.PeriodoLegislativo, 4, 1))
        Case "O"
            'strEtiqueta = strEtiqueta & "Ordinario "
            'strEtiqueta = strEtiqueta & "Legislativo "
            strEtiquetaLinea1 = strEtiquetaLinea1 & "Ordinario "
            'strEtiquetaLinea1 = strEtiquetaLinea1 & "Legislativo "
        Case "E"
            strEtiqueta = strEtiqueta & "Extraord. "
            strEtiquetaLinea1 = strEtiquetaLinea1 & "Extraordinario "
        Case "P"
            strEtiqueta = strEtiqueta & "Prepar. "
            strEtiquetaLinea1 = strEtiquetaLinea1 & "Preparatorio "
    End Select

    strEtiqueta = ""

    Select Case UCase(Mid(EstadoActual.PeriodoLegislativo, 5, 1))
        Case "T"
            strEtiqueta = strEtiqueta & "    " & Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & "  " & IIf(Len(Trim(EstadoActual.Sesion)) = 2, "", "  ") & "Sesión "
            strEtiqueta = strEtiqueta & "- Tablas" ' no aplicable en SBA09
            'strEtiquetaLinea2 = ""
            strEtiquetaLinea2 = strEtiqueta
        Case "E"
            'strEtiqueta = strEtiqueta & Str(EstadoActual.Sesion) & ORDINAL_FEMENINO & " Sesión "
            strEtiqueta = strEtiqueta & "    " & Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & "  " & IIf(Len(Trim(EstadoActual.Sesion)) = 2, "", "  ") & "Sesión "
            strEtiqueta = strEtiqueta & "- Especial"
            'strEtiquetaLinea2 = Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & " - Sesión "
            'strEtiquetaLinea2 = strEtiquetaLinea2 & "Especial"
            strEtiquetaLinea2 = strEtiqueta
        Case "A"
            strEtiqueta = strEtiqueta & "    " & Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & "  " & IIf(Len(Trim(EstadoActual.Sesion)) = 2, "", "  ") & "Sesión "
            'strEtiqueta = strEtiqueta & Str(EstadoActual.Sesion) & ORDINAL_FEMENINO & " "
            'strEtiqueta = strEtiqueta & "Asamblea Legislativa"
            strEtiqueta = strEtiqueta & "- Asamblea Prep."
            'strEtiquetaLinea2 = Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & " - "
            'strEtiquetaLinea2 = strEtiquetaLinea2 & "Asamblea Legislativa"
            strEtiquetaLinea2 = strEtiqueta
        Case "O"
            'strEtiqueta = strEtiqueta & Str(EstadoActual.Sesion) & ORDINAL_FEMENINO & " Sesión "
            strEtiqueta = strEtiqueta & "    " & Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & "  " & IIf(Len(Trim(EstadoActual.Sesion)) = 2, "", "  ") & "Sesión "
            strEtiqueta = strEtiqueta & "- Ordinaria"
            'strEtiquetaLinea2 = Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & " - Sesión "
            'strEtiquetaLinea2 = strEtiquetaLinea2 & "Ordinaria"
            strEtiquetaLinea2 = strEtiqueta
        Case "X"
            'strEtiqueta = strEtiqueta & Str(EstadoActual.Sesion) & ORDINAL_FEMENINO & " Sesión "
            strEtiqueta = strEtiqueta & "    " & Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & "  " & IIf(Len(Trim(EstadoActual.Sesion)) = 2, "", "  ") & "Sesión "
            strEtiqueta = strEtiqueta & "- Extraordinaria"
            'strEtiquetaLinea2 = Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & " - Sesión "
            'strEtiquetaLinea2 = strEtiquetaLinea2 & "Extraordinaria"
            strEtiquetaLinea2 = strEtiqueta
        Case "P"
            'strEtiqueta = strEtiqueta & Str(EstadoActual.Sesion) & ORDINAL_FEMENINO & " Sesión "
            strEtiqueta = strEtiqueta & "    " & Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & "  " & IIf(Len(Trim(EstadoActual.Sesion)) = 2, "", "  ") & "Sesión "
            strEtiqueta = strEtiqueta & "- Preparatoria"
            'strEtiquetaLinea2 = Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & " - Sesión "
            'strEtiquetaLinea2 = strEtiquetaLinea2 & "Preparatoria"
            strEtiquetaLinea2 = strEtiqueta
        Case "I"
            'strEtiqueta = strEtiqueta & Str(EstadoActual.Sesion) & ORDINAL_FEMENINO & " Sesión "
            strEtiqueta = strEtiqueta & "    " & Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & "  " & IIf(Len(Trim(EstadoActual.Sesion)) = 2, "", "  ") & "Sesión "
            strEtiqueta = strEtiqueta & "- Informativa"
            'strEtiquetaLinea2 = Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & " - Sesión "
            'strEtiquetaLinea2 = strEtiquetaLinea2 & "Preparatoria"
            strEtiquetaLinea2 = strEtiqueta
        Case "H"
            'strEtiqueta = strEtiqueta & Str(EstadoActual.Sesion) & ORDINAL_FEMENINO & " Sesión "
            strEtiqueta = strEtiqueta & "    " & Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & "  " & IIf(Len(Trim(EstadoActual.Sesion)) = 2, "", "  ") & "Sesión "
            strEtiqueta = strEtiqueta & "- Homenajes"
            'strEtiquetaLinea2 = Trim(Str(EstadoActual.Sesion)) & ORDINAL_FEMENINO & " - Sesión "
            'strEtiquetaLinea2 = strEtiquetaLinea2 & "Preparatoria"
            strEtiquetaLinea2 = strEtiqueta
    End Select
    
    'strEtiqueta = strEtiqueta & " - Próximo Nº de Acta: " & Str(EstadoActual.NroActa)
    If EstadoActual.Reunion > 0 Then
        strEtiqueta = strEtiqueta & " - " & "Reunión " & Str(EstadoActual.Reunion)
        Dim espacios As String
        Dim i As Integer
        For i = 1 To Len(Trim(Str(EstadoActual.Reunion)))
            espacios = espacios & " "
        Next i
        If Len(Trim(Str(EstadoActual.Reunion))) > 1 Then
            strEtiquetaLinea3 = "  " & Trim(Str(EstadoActual.Reunion) & "ª    Reunión ") 'Alineado
        Else
            strEtiquetaLinea3 = "    " & Trim(Str(EstadoActual.Reunion) & "ª    Reunión ") 'Alineado
        End If
    Else
        strEtiquetaLinea3 = ""
    End If
    LeyendaSesion = strEtiqueta
    Select Case nLinea
        Case 0
            LeyendaSesion = strEtiqueta
        Case 1
            LeyendaSesion = strEtiquetaLinea1
        Case 2
            LeyendaSesion = strEtiquetaLinea2
        Case 3
            LeyendaSesion = strEtiquetaLinea3
    End Select
End Function

Private Function Ordinal(xNumero As Integer, xGenero As String)
    Select Case xGenero
        Case "M"
            Select Case xNumero
                Case 1
                    Ordinal = "er."
                Case 2
                    Ordinal = "do."
                Case 3
                    Ordinal = "er."
                Case 4
                    Ordinal = "to."
                Case 5
                    Ordinal = "to."
                Case 6
                    Ordinal = "to."
                Case 7
                    Ordinal = "mo."
                Case 8
                    Ordinal = "vo."
                Case 9
                    Ordinal = "no."
                Case 0
                    Ordinal = "o."
            End Select
        Case "F"
            Select Case xNumero
                Case 1
                    Ordinal = "era."
                Case 2
                    Ordinal = "da."
                Case 3
                    Ordinal = "era."
                Case 4
                    Ordinal = "ta."
                Case 5
                    Ordinal = "ta."
                Case 6
                    Ordinal = "ta."
                Case 7
                    Ordinal = "ma."
                Case 8
                    Ordinal = "va."
                Case 9
                    Ordinal = "na."
                Case 0
                    Ordinal = "a."
            End Select
    End Select
    
End Function
Private Function LeyendaSesionCartelSerial() As String
    Dim strEtiqueta      As String

    strEtiqueta = Mid(EstadoActual.PeriodoLegislativo, 1, 3) & Ordinal(Val(Mid(EstadoActual.PeriodoLegislativo, 3, 1)), "M") & " P.Leg."
    
    Select Case UCase(Mid(EstadoActual.PeriodoLegislativo, 4, 1))
        Case "O"
            strEtiqueta = strEtiqueta & "Ordinario"
        Case "E"
            strEtiqueta = strEtiqueta & "Extraordinario"
        Case "P"
            strEtiqueta = strEtiqueta & "Pre."
    End Select
    Select Case UCase(Mid(EstadoActual.PeriodoLegislativo, 5, 1))
        Case "T"
            strEtiqueta = strEtiqueta & " " & Trim(Str(EstadoActual.Sesion)) & Ordinal(EstadoActual.Sesion Mod 10, "F") & " Sesion "
            strEtiqueta = strEtiqueta & "Tablas"
        Case "E"
            strEtiqueta = strEtiqueta & " Sesion "
            strEtiqueta = strEtiqueta & "Especial"
        Case "P"
            strEtiqueta = strEtiqueta & " " & Trim(Str(EstadoActual.Sesion)) & Ordinal(EstadoActual.Sesion Mod 10, "F") & " Sesion "
            strEtiqueta = strEtiqueta & "Preparatoria"
    End Select
    
    'strEtiqueta = strEtiqueta & " A." & Trim(Str(EstadoActual.NroActa))
    
    LeyendaSesionCartelSerial = strEtiqueta
End Function

Private Function LeyendaTipoOperacion() As String
    Select Case EstadoActual.TipoDeOperacion
        Case "paslis"
            LeyendaTipoOperacion = "Pase de Lista"
        Case "quorum"
            LeyendaTipoOperacion = "Censando Quórum"
        Case "votnom"
            If xTipoVotacion = "votnum" Then
                LeyendaTipoOperacion = "Votación Numérica"
            Else
                LeyendaTipoOperacion = "Votación Nominal"
            End If
        Case "votnum"
            LeyendaTipoOperacion = "Votación Numérica"
    End Select
End Function
Private Function LeyendaTipoOperacionCartelSerial() As String
    Select Case EstadoActual.TipoDeOperacion
        Case "paslis"
            LeyendaTipoOperacionCartelSerial = "        Pase de Lista"
        Case "quorum"
            LeyendaTipoOperacionCartelSerial = "Quórum"
        Case "votnom"
            If xTipoVotacion = "votnum" Then
                LeyendaTipoOperacionCartelSerial = "Numerica"
            Else
                LeyendaTipoOperacionCartelSerial = "Nominal "
            End If
        Case "votnum"
            LeyendaTipoOperacionCartelSerial = "Numerica"
    End Select
End Function

Private Function LeyendaTipoMayoria() As String
If True Then
    LeyendaTipoMayoria = EtiquetasCartel.strTipo
Else
    Select Case EstadoActual.TipoMayoria
        Case "120"
                   LeyendaTipoMayoria = "Más de la mitad"
        Case "121"
                   LeyendaTipoMayoria = "Mitad más uno"
        Case "14"
                   LeyendaTipoMayoria = "Un cuarto"
        Case "15"
                   LeyendaTipoMayoria = "Un quinto"
        Case "110"
                   LeyendaTipoMayoria = "Un decimo"
        Case "23"
                   LeyendaTipoMayoria = "Dos tercios"
        Case "34"
                   LeyendaTipoMayoria = "Tres cuartos"
        Case "100"
                   LeyendaTipoMayoria = "Unanimidad"
    End Select
End If
End Function


Private Function LeyendaBaseMayoria() As String
If True Then
    LeyendaBaseMayoria = EtiquetasCartel.strBase
Else
    Select Case EstadoActual.BaseMayoria
        Case "legpre"
                   LeyendaBaseMayoria = "Senadores Presentes"
        Case "miecue"
                   LeyendaBaseMayoria = "Miembros del Cuerpo"
        Case "votemi"
                   LeyendaBaseMayoria = "Votos Emitidos"
    End Select
End If
End Function
    
            
Private Function LeyendaBaseMayoriaCartelSerial() As String
    Select Case EstadoActual.BaseMayoria
        Case "legpre"
                   LeyendaBaseMayoriaCartelSerial = "L.Presentes"
        Case "miecue"
                   LeyendaBaseMayoriaCartelSerial = "M.Cuerpo"
        Case "votemi"
                   LeyendaBaseMayoriaCartelSerial = "V.Emitidos"
    End Select
End Function
    
            
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        'If MsgBox("¿Está seguro que desea cerrar el SQV Server?", vbQuestion + vbOKCancel, "Cerrar SQV Server") = vbOK Then
            Unload Me
        'End If
    End If
End Sub




Private Sub MostrarDatosSesion(mSesion As Long, mActa As Long, mVersion As Long)
On Error GoTo TrapError

Dim blEsLegislador  As Boolean
Dim strTipoOperacion  As String
    
    Dim strSql As String
    strSql = "SELECT actas.*, " _
        & " tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, basemay.Descripcion AS descBaseMay, " _
        & " tipmay.Descripcion AS descTipoMay, rtrim(Legisladores.apellido) + ', ' + rtrim(legisladores.nombre) AS Legislador, Actas.Tipo_de_Quorum " _
        & " FROM Legisladores RIGHT OUTER JOIN actas ON Legisladores.id = actas.Presidente LEFT OUTER JOIN " _
        & " tipmay ON actas.Tipo_de_Mayoria = tipmay.identificador_en_mensajes LEFT OUTER JOIN " _
        & " basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON " _
        & " actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT Outer Join tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes " _
        & " WHERE (Período_Legislativo='" & Trim(EstadoActual.PeriodoLegislativo) & "') AND (Sesión=" & mSesion & ") AND (Número_de_Acta=" & mActa & ") AND (Versión_Acta=" & mVersion & ") "
    SetearRsAux strSql, rstActa
    If rstActa.EOF = False Then
        With rstActa
            'xUltimaVersionActa = !Ultima_Versión_Acta
            'strResultadoEsperado = Trim(UCase(Trim(!Votacion)))
            'strSesion = Trim(!Sesión)
            'xNumeroActa = !Número_de_Acta
            'xVersionActa = !versión_acta
            'strTipoQuorum = Trim(!Tipo_de_Quorum)
            'strTipoMayoria = Trim(!Tipo_de_Mayoria)
            'strBaseMayoria = Trim(!Base_de_mayoria)
            'If IsNull(!NroOrdenDia) Then
            '    xNroOrdenDia = 0
            'Else
            '    xNroOrdenDia = !NroOrdenDia
            'End If
            strTipoOperacion = Trim(!Tipo_de_Operación)
            'strPeriodo_Legislativo = Trim(!Período_Legislativo)
            'strNombreActa = Trim(!Nombre_del_Acta)
            'xMiembrosDelCuerpo = !Miembros_del_cuerpo
            'xPresentesTotal = !Presentes_Total
            'xVotosAfirmativosTotal = !Votos_Afirm_Total
            'xVotosNegativosTotal = !Votos_Neg_Total
            'xVotosAfirmativosIdentificables = !Votos_Afirm_Identificables
            'xVotosAfirmativosTotal = !Votos_Afirm_Total
            'xVotosAfirmativosNoIdentificables = !Votos_Afirm_No_Identificables
            'xVotosAfirmativosDesempate = !Votos_Afirm_Desempate
            'xVotosNegativosIdentificables = !Votos_Neg_Identificables
            'xVotosNegativosTotal = !Votos_Neg_Total
            'xVotosNegativosNoIdentificables = !Votos_Neg_No_Identificables
            'xVotosNegativosDesempate = !Votos_Neg_Desempate
            'xAbstencionesIdentificables = !Abstenciones_Identificables
            'xAbstencionesNOIdentificables = !Abstenciones_No_Identificables
            'xAbstencionesTotal = !Abstenciones_Total
            'xPresentesIdentificables = !Presentes_Identificables
            'xPresentesNOIdentificables = !Presentes_No_Identificables
            'xPresentesTotal = !Presentes_Total
            If (!vota_presidente = 1) Then
                blEsLegislador = True
            Else
                blEsLegislador = False
            End If
            If IsNull(!descTipoOp) = False Then
                txtTipoOperacion.text = !descTipoOp
            End If
            If IsNull(!Sesión) = False Then
                txtSesion.text = !Sesión
            End If
            If IsNull(!Número_de_Acta) = False Then
                txtNroActa.text = !Número_de_Acta
            End If
            If IsNull(!Versión_Acta) = False Then
                txtVersion.text = !Versión_Acta
                If !Ultima_Versión_Acta = 0 Then
                    txtVersion.text = "Original"
                Else
                   If !Versión_Acta = 0 Then
                        txtVersion.text = "Ult.Mod.Ver. " & Val(!Ultima_Versión_Acta) + 1
                   Else
                        txtVersion.text = "Ver. " & Val(!Ultima_Versión_Acta) + 1
                   End If
                End If
                txtVersion.Tag = !Versión_Acta
            End If
            If IsNull(!Reunion) = False Then
                'txtReunion.Text = !Reunion 'Habilitar en cartel
            End If
            If IsNull(!Nombre_del_Acta) = False Then
                txtNombre.text = Trim(!Nombre_del_Acta)
            End If
            If IsNull(!Fecha) = False Then
                txtFecha.text = !Fecha
            End If
            If IsNull(!Hora) = False Then
                txtHora.text = !Hora
            End If
            If IsNull(!descTipoMayQuo) = False Then
                txtTipoQuorum.text = !descTipoMayQuo
            End If
            If IsNull(!Miembros_del_cuerpo) = False Then
                txtMiembros.text = !Miembros_del_cuerpo
            End If
            If IsNull(!Desempate) = False Then
                txtDesempate.text = !Desempate
            End If
            If IsNull(!descTipoMay) = False Then
                txtTipoMayoria.text = !descTipoMay
            End If
            If IsNull(!descBaseMay) = False Then
                txtBase.text = !descBaseMay
            End If
            If IsNull(!Votacion) = False Then
                txtVotacion.text = !Votacion
            End If
            If IsNull(!presidente) = False Then
                txtCodigoPresidente.text = !presidente
            End If
            If IsNull(!legislador) = False Then
                txtNombrePresidente.text = !legislador
            End If
            If IsNull(!Observaciones) = False Then
                txtObservaciones.text = Trim(!Observaciones)
            End If
            If IsNull(!Presentes_Identificables) = False Then
                txtPresentesId.text = !Presentes_Identificables
            Else
                txtPresentesId.text = "0"
            End If
            If IsNull(!Presentes_No_Identificables) = False Then
                txtPresentesNoId.text = !Presentes_No_Identificables
            Else
                txtPresentesNoId.text = "0"
            End If
            If IsNull(!Presentes_Total) = False Then
                txtPresentesTotal.text = !Presentes_Total
            Else
                txtPresentesTotal.text = "0"
            End If
            If IsNull(!Ausentes_Total) = False Then
                txtAusentesTotal.text = !Ausentes_Total
            Else
                txtAusentesTotal.text = "0"
            End If
            If IsNull(!Votos_Afirm_Identificables) = False Then
                txtAfirmativosId.text = !Votos_Afirm_Identificables
            Else
                txtAfirmativosId.text = "0"
            End If
            If IsNull(!Votos_Afirm_No_Identificables) = False Then
                txtAfirmativosNoId.text = !Votos_Afirm_No_Identificables
            Else
                txtAfirmativosNoId.text = "0"
            End If
            If IsNull(!Votos_Afirm_Desempate) = False Then
                txtAfirmativosDesempate.text = !Votos_Afirm_Desempate
            Else
                txtAfirmativosDesempate.text = "0"
            End If
            If IsNull(!Votos_Afirm_Total) = False Then
                txtAfirmativosTotal.text = !Votos_Afirm_Total
            Else
                txtAfirmativosTotal.text = "0"
            End If
            If IsNull(!Votos_Neg_Identificables) = False Then
                txtNegativoID.text = !Votos_Neg_Identificables
            Else
                txtNegativoID.text = "0"
            End If
            If IsNull(!Votos_Neg_No_Identificables) = False Then
                txtNegativoNoId.text = !Votos_Neg_No_Identificables
            Else
                txtNegativoNoId.text = "0"
            End If
            If IsNull(!Votos_Neg_Desempate) = False Then
                txtNegativoDesempate.text = !Votos_Neg_Desempate
            Else
                txtNegativoDesempate.text = "0"
            End If
            If IsNull(!Votos_Neg_Total) = False Then
                txtNegativoTotales.text = !Votos_Neg_Total
            Else
                txtNegativoTotales.text = "0"
            End If
            If IsNull(!Abstenciones_Identificables) = False Then
                txtAbstencionesId.text = !Abstenciones_Identificables
            Else
                txtAbstencionesId.text = "0"
            End If
            If IsNull(!Abstenciones_No_Identificables) = False Then
                txtAbstencionesNoId.text = !Abstenciones_No_Identificables
            Else
                txtAbstencionesNoId.text = "0"
            End If
            If IsNull(!Abstenciones_Total) = False Then
                txtAbstencionesTotales.text = !Abstenciones_Total
            Else
                txtAbstencionesTotales.text = "0"
            End If
        End With
        
    End If
    If strTipoOperacion <> "votnum" Then
        'Call MostrarDetalleActa(rstActa!Período_Legislativo, txtSesion.Text, txtNroActa.Text, txtVersion.Tag)
        'Call CargarComboResultados
    Else
        'vsGrilla.Visible = False
        'lblBuscarLegislador.Visible = False
        'txtBuscar.Visible = False
    End If
    'ControlesHabilitados = False
rstActa.Close

Exit Sub
TrapError:
    Select Case err.Number
        Case Else
            Call AltaLogGeneral("SERVER SQV", "MostrarDatosSesion Error Nº" & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source)
            'MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            End
    End Select
End Sub



Private Sub MostrarActaProyector(mSesion As Long, mActa As Long, mVersion As Long)
FrameSQVActa.ZOrder 0
Call MostrarDatosSesion(mSesion, mActa, mVersion)
'MsgBox " dd"

End Sub

Private Function max(a As Long, b As Long) As Long
    max = IIf(a > b, a, b)
End Function
Private Function Min(a As Long, b As Long) As Long
    Min = IIf(a < b, a, b)
End Function

Private Function CerosIzquierda(strText As String, nLong As Long) As String
    If nLong > Len(strText) Then
        CerosIzquierda = Left(String(nLong - Len(strText), "0") & strText, nLong)
    Else
        CerosIzquierda = Right(strText, nLong)
    End If
End Function

Function CalculoQuorum() As String
    CalculoQuorum = IIf(CartelActual.Presentes >= Fix(xMinimoParaQuorumEntero), "QUORUM", "NO HAY QUORUM") 'COMPAQ
End Function
Private Function PresentesIdentificados() As Long
    PresentesIdentificados = GetIdentificados 'EstadoActual.Presentes - EstadoActual.OcupadosNoIdentificados
End Function
Private Function Presentes() As Long
    If EstadoActual.Modo_Ident_Nom = 1 And EstadoActual.Modo_Presencia_Nom = 1 Then
        Presentes = PresentesIdentificados()
    Else
        Presentes = getPresentes 'EstadoActual.Presentes
    End If
End Function

Private Function Ausentes() As Long
    If EstadoActual.Modo_Ident_Nom = 1 And EstadoActual.Modo_Presencia_Nom = 1 Then
        Ausentes = xMiembrosDelCuerpo - PresentesIdentificados()
    Else
        Ausentes = GetAusentes 'EstadoActual.Ausentes
    End If
End Function

Private Sub SolicitarIdentificacionPendientes(xStrMensaje As String, xModo As String)
    Dim xStrVector As String
    Dim i As Long
    
    xStrVector = "0" & SEPARADOR_VECTOR 'presidente
    For i = 1 To UBound(EstadoActual.VectorPresencia)
        xStrVector = xStrVector & IIf(EstadoActual.VectorPresencia(i) = PRESENTE And EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO, "1", "0") & SEPARADOR_VECTOR
    Next i
    Call EnviarMensajesComienzoAuth(xStrVector, xStrMensaje, xModo)
End Sub

Private Sub AbstenerBanca(xBanca As Long)
    EstadoActual.VectorResultados(xBanca) = ABSTENCION_AUTORIZADA
    PintarVectorColor (xBanca)
    EstadoActual.AbstencionistasAutorizados = EstadoActual.AbstencionistasAutorizados + 1
    ' If (EstadoActual.TipoDeOperacion = "votnom" ) And InStr("votando larga", EstadoActual.EstadoVotacion_y_PasList) > 0 Then
    If (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum") And InStr("votando larga", EstadoActual.EstadoVotacion_y_PasList) > 0 Then
        EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1
        With Mensaje2Banca ' Mensaje para SB
            .sTipo = "mset"
            .sComponente = "term.ledk1"
            .sObjeto = Str(xBanca)
            .sAtributo = "state"
            .sValor = "off"
        End With
        Call EnviarMensajesBancas(Mensaje2Banca)
        Call AltaLogGeneral("SQV SERVER: AbstenerBanca 2", "If (EstadoActual.TipoDeOperacion = votnom Or EstadoActual.TipoDeOperacion = votnum) And InStr(votando larga, EstadoActual.EstadoVotacion_y_PasList) > 0 Then EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1: " & EstadoActual.PendientesEmitirVotos, Str(xBanca), "0")
        'apaga teclados
        With Mensaje2Banca ' Mensaje para SB
            .sTipo = "mset"
            .sComponente = "term.keyb"
            .sObjeto = Str(xBanca)
            .sAtributo = "state"
            .sValor = "off" & IIf(xBanca > 0, EstadoActual.TipoDeOperacion, "votnum")
        End With
        Call EnviarMensajesBancas(Mensaje2Banca)
    End If
    Call AltaLogGeneral("SQV SERVER: Abstencion", "Abstencion Autorizada Banca: " & Str(xBanca) & " Leg. ID " & EstadoActual.VectorIdentificacion(xBanca))
End Sub

Private Sub CancelarAbstenerBanca(xBanca As Long)
    EstadoActual.VectorResultados(xBanca) = ABSTENCION
    PintarVectorColor (xBanca)
    EstadoActual.AbstencionistasAutorizados = EstadoActual.AbstencionistasAutorizados - 1
    ' If EstadoActual.TipoDeOperacion = "votnom" And InStr("votando larga", EstadoActual.EstadoVotacion_y_PasList) > 0 Then
    If (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum") And InStr("votando larga", EstadoActual.EstadoVotacion_y_PasList) > 0 Then
        EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos + 1
        Call AltaLogGeneral("SQV SERVER: CancelarAbstenerBanca 1", "EstadoActual.PendientesEmitirVotos + 1: " & EstadoActual.PendientesEmitirVotos & " Banca " & Str(xBanca), Str(xBanca), "0")
        'apaga teclados
        With Mensaje2Banca ' Mensaje para SB
            .sTipo = "mset"
            .sComponente = "term.keyb"
            .sObjeto = Str(xBanca)
            .sAtributo = "state"
            .sValor = "on" & IIf(xBanca > 0, EstadoActual.TipoDeOperacion, "votnum")
        End With
        Call EnviarMensajesBancas(Mensaje2Banca)
    End If
    Call AltaLogGeneral("SQV SERVER: Abstencion", "Cancelacion Abstencion Autorizada Banca: " & Str(xBanca) & " Leg. ID " & EstadoActual.VectorIdentificacion(xBanca))
End Sub
Private Sub AbstenerVector(strVector As String, Optional nNuevosAbstenidos As Long, Optional nNuevosCancelados As Long, Optional nTotalAbstenciones As Long)
    Dim i As Long
    nNuevosAbstenidos = 0
    nNuevosCancelados = 0
    nTotalAbstenciones = 0
    'En hcdn no hay abstencion autorizada
    If False Then ' funcion deshabilitada
        Call AltaLogGeneral("SQV SERVER: Abstencion vector", strVector, , "0")
    End If
    strVector = SEPARADOR_VECTOR & Trim(strVector)
    For i = 1 To (xUltimaBanca) 'el presidente no puede abstenerse
        'Si ya esta autorizado a abstenerse, y tambien esta identificado
        ' If EstadoActual.VectorResultados(i) = ABSTENCION_AUTORIZADA And Not (EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO) Then
        ' If i = 70 Then Stop
        If EstadoActual.VectorResultados(i) = ABSTENCION_AUTORIZADA _
                   And EstadoActual.VectorPresencia(i) = PRESENTE _
                   And (Not (EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO) _
                   Or EstadoActual.TipoDeOperacion = "votnum") Then
            'y solo si no esta en la lista de autorizados recibida
            If Not (InStr(strVector, IIf(EstadoActual.TipoDeOperacion = "votnum", SEPARADOR_VECTOR & Trim(Str(i)) & SEPARADOR_VECTOR, SEPARADOR_VECTOR & Trim(EstadoActual.VectorIdentificacion(i)) & SEPARADOR_VECTOR)) > 0) Then
                CancelarAbstenerBanca (i)
                nNuevosCancelados = nNuevosCancelados + 1
            Else
                nTotalAbstenciones = nTotalAbstenciones + 1
            End If
        'Si no esta aun autorizado a abstenerse, y esta identificado alguien en la banca
        ' ElseIf EstadoActual.VectorResultados(i) = ABSTENCION And Not (EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO) Then
        ElseIf EstadoActual.VectorResultados(i) = ABSTENCION And EstadoActual.VectorPresencia(i) = PRESENTE And (Not (EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO) Or EstadoActual.TipoDeOperacion = "votnum") Then
            'y si esta en la lista de autorizados recibida
            'cambiox
            If (InStr(strVector, IIf(EstadoActual.TipoDeOperacion = "votnum", SEPARADOR_VECTOR & Trim(Str(i)) & SEPARADOR_VECTOR, SEPARADOR_VECTOR & Trim(EstadoActual.VectorIdentificacion(i)) & SEPARADOR_VECTOR)) > 0) Then
                AbstenerBanca (i)
                nNuevosAbstenidos = nNuevosAbstenidos + 1
                nTotalAbstenciones = nTotalAbstenciones + 1
            End If
        End If
    Next i
    lblPendientesEmitirVotos = EstadoActual.PendientesEmitirVotos
    lblAbsAut.Caption = EstadoActual.AbstencionistasAutorizados
    lblOcupadosNoIdentificados(0) = EstadoActual.OcupadosNoIdentificados
End Sub

Private Sub FinVotacionBrc(xComentario As String)
Dim xStrVector As String
Dim i As Integer
Dim xMax As Integer
Dim X As Integer
Dim StrTempCadena As String
xMax = UBound(EstadoActual.VectorPresencia)
Dim Mensaje2Banca As MensajeSistema
    If EstadoActual.TipoDeOperacion = "votnum" Then
    
        StrTempCadena = ("0") & SEPARADOR_VECTOR ' el presidente aunque vote, se lo trata como nominal mas abajo
        For X = 1 To xMax
            If EstadoActual.VectorPresencia(X) = PRESENTE Then
                If (EstadoActual.TipoDeOperacion = "votnum" _
                    And (EstadoActual.VectorIdentificacion(X) = NO_IDENTIFICADO)) _
                    And EstadoActual.VectorResultados(X) <> ABSTENCION_AUTORIZADA Then 'Solo los identif
                    StrTempCadena = StrTempCadena & PRESENTE & SEPARADOR_VECTOR
                Else
                    StrTempCadena = StrTempCadena & AUSENTE & SEPARADOR_VECTOR
                End If
            Else
                StrTempCadena = StrTempCadena & AUSENTE & SEPARADOR_VECTOR
            End If
        Next X
        With Mensaje2Banca ' Mensaje para SB
            .sTipo = "mset"
            .sComponente = "term.keyb"
            .sObjeto = StrTempCadena
            .sAtributo = "state"
            .sValor = "offvotnum"
        End With
        Call EnviarMensajesBancas(Mensaje2Banca)
        
        'Ahora cancelar los teclados de los identificados, no importa el tipo de votacion
        StrTempCadena = ("0") & SEPARADOR_VECTOR
        For X = 1 To xMax
            If EstadoActual.VectorPresencia(X) = PRESENTE Then
                If (EstadoActual.TipoDeOperacion = "votnum" _
                    And (EstadoActual.VectorIdentificacion(X) <> NO_IDENTIFICADO)) _
                    And EstadoActual.VectorResultados(X) <> ABSTENCION_AUTORIZADA Then 'Solo los identif
                    StrTempCadena = StrTempCadena & PRESENTE & SEPARADOR_VECTOR
                Else
                    StrTempCadena = StrTempCadena & AUSENTE & SEPARADOR_VECTOR
                End If
            Else
                StrTempCadena = StrTempCadena & AUSENTE & SEPARADOR_VECTOR
            End If



        Next X
        With Mensaje2Banca ' Mensaje para SB
            .sTipo = "mset"
            .sComponente = "term.keyb"
            .sObjeto = StrTempCadena
            .sAtributo = "state"
            .sValor = "offvotnom"
        End With
        Call EnviarMensajesBancas(Mensaje2Banca)
    Else 'Si esta en votacion nominal
        Mensaje2Banca.sTipo = "mset"

        Dim strX As String
        strX = Space(2)
        strX = "1;"
        Mensaje2Banca.sObjeto = "0;" & Replicar(xUltimaBanca, strX)
        Mensaje2Banca.sComponente = "term.keyb"
        Mensaje2Banca.sAtributo = "state"
        Mensaje2Banca.sValor = "off" & EstadoActual.TipoDeOperacion
        Mensaje2Banca.sComentario = xComentario
        Call EnviarMensajesBancas(Mensaje2Banca)
        'MsgBox "verificar cierre de votacion"
'        Mensaje2Banca.sObjeto = "0"
'        Mensaje2Banca.sValor = "offvotnum"
'        Call EnviarMensajesBancas(Mensaje2Banca)
    End If
    Call FinalizarVotacionPresidente 'en caso de que estuviera habilitado...
    Call AltaLogGeneral("SQV SERVER: Fin Votacion", xComentario, , "1")
    '***********Entre cierre e inicializacion hay que detener el proceso de identificacion***********
    'Copiar este proceso a "larga"
    If EstadoActual.Modo_Ident_Nom = 1 Or EstadoActual.TipoDeOperacion = "votnom" Then
        xStrVector = "0" & SEPARADOR_VECTOR 'presidente
        For i = 1 To UBound(EstadoActual.VectorPresencia)
            'xStrVector = xStrVector & IIf(EstadoActual.VectorPresencia(i) = "1", "1", "0") & SEPARADOR_VECTOR
            xStrVector = xStrVector & IIf(EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO, "1", "0") & SEPARADOR_VECTOR
        Next i
        Call EnviarMensajesFinAuth(xStrVector, "Fin modo nominal desde " & EstadoActual.TipoDeOperacion & "a " & "")
        EstadoActual.Modo_Ident_Nom = 0
    End If
End Sub

Private Function Replicar(xCant As Long, strCadena As String) As String
    Dim X As Long
    For X = 1 To xCant
        Replicar = Replicar & strCadena
    Next X
End Function




'******************************************************************************************************************************
'PROCEDIMIENTOS AGREGADOS POR A02
'******************************************************************************************************************************
Private Sub ConfigurarFrames()
    Dim i As Integer
    'Frame A: Superior: Fecha Hora presentes ausentes, etc.
    '
    For i = 0 To 2
        picA(i).Height = imgA(i).Height
        picA(i).Width = imgA(i).Width
        picA(i).top = 0
        picA(i).Left = 0
        imgA(i).top = 0
        imgA(i).Left = 0
    Next
    For i = 0 To 4
        picB(i).top = picA(0).Height
        picB(i).Left = 0
        picB(i).Width = 15360
        picB(i).Height = imgB(i).Height
        imgB(i).Left = 0
        imgB(i).top = 0
    Next
    For i = 0 To 1
        picC(i).top = picA(0).Height + picB(0).Height
        picC(i).Left = 0
        picC(i).Width = 15360
        picC(i).Height = imgC(i).Height
        picC(i).BackColor = &H80000008
        picC(i).ForeColor = &H80000008
        imgC(i).top = 0
        imgC(i).Left = 0
    Next
    'lsv.View = lvwSmallIcon
    'lsv.Arrange = lvwAutoTop
    'lsv.Top = 250
    'lsv.Left = 280
    'lsv.Height = 7700
    'lsv.Width = 14760
    lsv.Visible = False
    InicializarImagenes
End Sub



Private Sub CargarColoresFuente()
    Dim Color As String
    Dim clave As String
    Dim i As Long
    ReDim mColoresFuente(0 To 7)
    i = 0
    'cargo el diccionario de manera estática
    'GRIS
    'CONTRASTE NEGRO
    Color = "&H0"
    mColoresFuente(i) = Color
    
    i = i + 1
    'BLANCO
    'CONTRASTE NEGRO
    Color = "&H0"
    mColoresFuente(i) = Color
    
    i = i + 1
    'AMARILLO
    'CONTRASTE NEGRO
    Color = "&H0"
    mColoresFuente(i) = Color
    
    i = i + 1
    'ROJO
    'CONTRASTE BLANCO
    Color = "&HFFFFFF"
    mColoresFuente(i) = Color
    
    
    i = i + 1
    'CELESTE
    'CONTRASTE NEGRO
    Color = "&H0"
    mColoresFuente(i) = Color
        
    i = i + 1
    'NARANJA
    'CONTRASTE NEGRO
    Color = "&H0"
    mColoresFuente(i) = Color
    
    i = i + 1
    'VERDE
    'CONTRASTE NEGRO
    Color = "&H0"
    mColoresFuente(i) = Color
    
    i = i + 1
    'NEGRO
    'CONTRASTE BLANCO
    Color = "&HFFFFFF"
    mColoresFuente(i) = Color
End Sub

Private Sub AsignarFuentes()
'On Error Resume Next 'manzana
    Const MAX_OBJETOS = 37
    Dim i As Integer
    Dim Fuente As String
    Dim FuenteNombre As String
    Dim FuenteTamano As Single
    Dim FuenteBold As Boolean
    Dim ExisteFuentePrincipal As Boolean
    Dim ExisteFuenteSecundaria As Boolean
    Dim ExisteFuenteTerciaria As Boolean
    Dim f As String
    Dim obj(MAX_OBJETOS) As Object
    For i = 0 To Screen.FontCount - 1
        f = Screen.Fonts(i)
        If f = "Accidental Presidency" Then
            ExisteFuentePrincipal = True
        ElseIf f = "Haettenschweiler" Then
            ExisteFuenteSecundaria = True
        ElseIf f = "Arial" Then
            ExisteFuenteTerciaria = True
        End If
    Next
    
    If ExisteFuentePrincipal Then
        FuenteNombre = "Accidental Presidency"
        FuenteBold = False
        FuenteTamano = 24
    ElseIf ExisteFuenteSecundaria Then
        FuenteNombre = "Haettenschweiler"
        FuenteBold = False
        FuenteTamano = 22
    ElseIf ExisteFuenteTerciaria Then 'SBA 2009
        'FuenteNombre = "Arial"
        FuenteNombre = "Times New Roman"
        FuenteBold = True
        FuenteTamano = 20
    Else
        FuenteNombre = "MS Sans Serif"
        FuenteBold = True
        FuenteTamano = 20
    End If
    
    For i = 0 To xUltimaBanca
        'Load lblBanca(xUltimaBanca)'manzana
        If i >= 71 Then
            Load lblBanca(i)
        End If
        lblBanca(i).Font = FuenteNombre
        lblBanca(i).FontBold = FuenteBold
        lblBanca(i).FontSize = FuenteTamano
    Next
    
    Set obj(1) = lblGeneralAbstencionesDato
    Set obj(2) = lblGeneralAfirmativosDato
    Set obj(3) = lblGeneralAusentesDato
    Set obj(4) = lblGeneralFechaDato
    Set obj(5) = lblGeneralHoraDato
    Set obj(6) = lblGeneralInformacion
    Set obj(7) = lblGeneralLeyendaQuorumDato
    Set obj(8) = lblGeneralMayoriaDato(0)
    Set obj(9) = lblGeneralMayoriaDato(1)
    Set obj(10) = lblGeneralNegativosDato
    Set obj(11) = lblGeneralPresentesDato
    Set obj(12) = lblGeneralResultadoDato
    Set obj(13) = lblGeneralSesionDato(0)
    Set obj(14) = lblGeneralSesionDato(1)
    Set obj(15) = lblGeneralTipoOperacionDato
    Set obj(16) = lblGeneralTituloDato(0)
    Set obj(17) = lblGeneralTituloDato(1)
    Set obj(18) = lblGeneralTituloDato(2)
    Set obj(19) = lblGeneralTituloDato(3)
    Set obj(20) = lblGeneralTituloDato(4)
    Set obj(21) = lblGeneralTiempoDato
    Set obj(22) = lblGeneralMayoriaDato(2)
    Set obj(23) = lblGeneralTituloTiempo
    Set obj(24) = lblTituloPresentesIdentificados
    Set obj(25) = lblTituloOcupadosNoIdentificados(0)
    Set obj(26) = lblPresentesIdentificados
    Set obj(27) = lblOcupadosNoIdentificados(0)
    Set obj(28) = lblOcupadosNoIdentificados(1)
    Set obj(29) = lblOrador01
    Set obj(30) = lblOrador02
    Set obj(31) = lblOrador03
    Set obj(32) = lblOrador04
    Set obj(33) = lblGeneralSesionDato(3)
    Set obj(34) = lblGeneralSesionDato(4)
    Set obj(35) = lblGeneralSesionDato(5)
    Set obj(36) = lblTituloOcupadosNoIdentificados(1)
    Set obj(37) = lblOcupadosNoIdentificados(2)
    For i = 1 To MAX_OBJETOS
        obj(i).Font = FuenteNombre
        obj(i).FontBold = FuenteBold
        obj(i).FontSize = FuenteTamano
        obj(i).ForeColor = &HFFFF&    ' &HC0FFFF 'Color para todas las fuentes
        Set obj(i) = Nothing
    Next
    lblGeneralMayoriaDato(2).ForeColor = MiRojo
    lblTituloBaseYTipoDeMayoria.ForeColor = MiBlanco
    lblGeneralSesionDato(3).ForeColor = MiBlanco
    lblGeneralSesionDato(4).ForeColor = MiBlanco
    lblGeneralSesionDato(5).ForeColor = MiBlanco
    lblGeneralTituloDato(4).ForeColor = MiRojo
    lblGeneralFechaDato.ForeColor = MiBlanco
    lblGeneralHoraDato.ForeColor = MiBlanco
    lblGeneralNegativosDato.Alignment = vbRightJustify
    lblGeneralAfirmativosDato.Alignment = vbRightJustify
    lblGeneralAbstencionesDato.Alignment = vbRightJustify
    lblLeyendaVotoAfirmativo.FontSize = lblLeyendaVotoAfirmativo.FontSize - 4
    lblLeyendaVotoNegativo.FontSize = lblLeyendaVotoNegativo.FontSize - 4
    lblLeyendaVotoAbstencion.FontSize = lblLeyendaVotoAbstencion.FontSize - 4
    lblGeneralFechaDato.FontSize = 40
    lblGeneralHoraDato.FontSize = 40
    lblGeneralResultadoDato.FontSize = lblLeyendaVotoNegativo.FontSize + 8
    lblGeneralAbstencionesDato.FontSize = 48
    lblGeneralAfirmativosDato.FontSize = 48
    lblGeneralNegativosDato.FontSize = 48
    lblGeneralPresentesDato.FontSize = 48
    lblGeneralAusentesDato.FontSize = 48
    lblGeneralTipoOperacionDato.FontSize = 36
    lblGeneralTipoOperacionDato.FontUnderline = True
    
    lblGeneralInformacion.FontSize = 40
    'lblGeneralInformacion.Alignment = 2
    lblGeneralInformacion.Visible = True
    lblGeneralInformacion.ForeColor = MiRojo
    lblGeneralInformacion.FontUnderline = False
    lblGeneralTiempoDato.FontSize = 56
    lblGeneralTiempoDato.ForeColor = &H80000005
    lblGeneralTituloTiempo.FontSize = 56
    lblGeneralTituloTiempo.ForeColor = &H80000005
    
'    lblGeneralSesionDato(0).FontSize = 26
'    lblGeneralSesionDato(1).FontSize = 26
'    lblGeneralSesionDato(3).FontSize = 32
'    lblGeneralSesionDato(4).FontSize = 32
'    lblGeneralSesionDato(5).FontSize = 32
    lblGeneralSesionDato(0).FontSize = 26
    lblGeneralSesionDato(1).FontSize = 26
    lblGeneralSesionDato(3).FontSize = 28
    lblGeneralSesionDato(4).FontSize = 28
    lblGeneralSesionDato(5).FontSize = 28
    lblTituloPresentesIdentificados.FontSize = 36
    lblTituloOcupadosNoIdentificados(1).FontSize = 32
    'lblTituloOcupadosNoIdentificados(1).FontBold = False
    lblTituloPresentesIdentificados.ForeColor = &H80000005
    lblTituloOcupadosNoIdentificados(1).ForeColor = &H80000005
    lblPresentesIdentificados.FontSize = 36
    lblOcupadosNoIdentificados(0).FontSize = 36
    lblOcupadosNoIdentificados(1).FontSize = 36
    lblOcupadosNoIdentificados(2).FontSize = 36
    
    lblOrador01.FontSize = 32
    lblOrador01.ForeColor = &H80000005
    lblOrador02.FontSize = 32
    lblOrador03.FontSize = 32
    lblOrador04.FontSize = 32
    
    lblTituloBaseYTipoDeMayoria.FontSize = 26
    lblGeneralMayoriaDato(2).FontSize = 26
    
    'TAMAÑO del titulo
    lblGeneralTituloDato(4).FontSize = 28
    lblTituloOcupadosNoIdentificados(1).FontSize = lblGeneralTituloDato(4).FontSize
    lblLeyendaVotoNegativo.FontName = "Times New Roman"
    lblLeyendaVotoAfirmativo.FontName = "Times New Roman"
    lblLeyendaVotoAbstencion.FontName = "Times New Roman"
    'lsv.Font = FuenteNombre
    'lsv.Font.Size = 18
End Sub

Private Sub MostrarPIC(Nivel As String, ElementoVisible As Integer)
    Dim i As Integer
    Select Case Nivel
        Case "A"
            For i = 0 To 2
                picA(i).Visible = False
            Next
            picA(ElementoVisible).Visible = True
        Case "B"
            For i = 0 To 4
                picB(i).Visible = False
            Next
            If ElementoVisible < 2 Then
                TimerPic.Enabled = True
                picB(ElementoVisible).Visible = True
                TimerCounter = 1
            Else
                TimerPic.Enabled = False
                TimerCounter = 0
                picB(ElementoVisible).Visible = True
            End If
        Case "C"
            For i = 0 To 1
                picC(i).Visible = False
            Next
            picC(ElementoVisible).Visible = True
    End Select
End Sub

Private Sub CrearColeccionLegisladores()
    On Error GoTo TrapError
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido FROM Legisladores order by Legisladores.apellido, Legisladores.nombre"
    
    With rs
        If .State = adStateOpen Then
            .Close
            .Source = strSql
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open
        Else
            .Source = strSql
            .ActiveConnection = Cn
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .CursorLocation = adUseClient
            .Open
        End If
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        While Not .EOF
            colLeg.Add .Fields("apellido").Value, .Fields("nombre").Value, UCase(.Fields("apellido").Value) & ", " & .Fields("nombre").Value, .Fields("ID").Value
            .MoveNext
        Wend
    End With
    Set rs = Nothing
Exit Sub
TrapError:
    Call AltaLogGeneral("SERVER SQV", "CrearColeccionLegisladores Error Nº" & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source)
    'MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
    Resume
End Sub

Private Sub PanelResultadosInicializar()
    'lsv.SortKey = 0
    'lsv.ListItems.Clear
End Sub

Private Sub PanelResultadosCargar(strOperacion As String, strIdLegislador As String, strResultado As String)
    Dim Icon As Integer
    If (strOperacion = "paslis" Or strOperacion = "votnom") And strIdLegislador <> "" Then
        Select Case strResultado
            Case "AFIRMATIVO"
                Icon = 1
            Case "NEGATIVO"
                Icon = 2
            Case "ABSTENCION"
                Icon = 3
            Case "PRESENTE"
                Icon = 4
            Case "AUSENTE"
                Icon = 5
        End Select
        'lsv.ListItems.Add , , colLeg(strIdLegislador).ApellidoNombre, Icon, Icon
    End If
    'lsv.Arrange = lvwAutoTop
    'lsv.View = lvwIcon
    'lsv.View = lvwSmallIcon
    'lsv.Arrange = lvwAutoLeft
End Sub


Private Sub TimerPic_Timer()
    'Muestra cinco segundos la Sesión y 10 segundos el asunto en tratamiento siempre
    'y cuando no este vacio
    Select Case TimerCounter
        Case Is < 1
            picB(0).Visible = True
            picB(1).Visible = False
            TimerCounter = TimerCounter + 1
        Case 1
            If EstadoActual.TituloDelActa <> "" Then
                picB(0).Visible = True ' False para alternar
                picB(1).Visible = False ' True para alternar...
                TimerCounter = 2
            Else
                picB(0).Visible = True
                picB(1).Visible = False
                TimerCounter = 1
            End If
        Case 2
            TimerCounter = 0
    End Select
End Sub


Private Sub CargarImagenes()
    On Error GoTo TrapError
    'Fondo del frame general
    FrameSQVGeneral.BackColor = (cFORMULARIO_COLOR_FONDO)
    FrameSQVGeneral.Left = 0
    imgA(0).Picture = LoadPicture(imagePath & "AStandard.jpg")
    imgA(1).Picture = LoadPicture(imagePath & "AAsunto.jpg")
'    Set imgA(0).Picture = Nothing
'    Set imgA(1).Picture = Nothing
    imgA(2).Picture = LoadPicture(imagePath & "ASesion.jpg")
    
    imgB(0).Picture = LoadPicture(imagePath & "BSesion.jpg")
    imgB(1).Picture = LoadPicture(imagePath & "BAsunto.jpg")
    imgB(2).Picture = LoadPicture(imagePath & "BVotOn.jpg")
    imgB(3).Picture = LoadPicture(imagePath & "BResultado.jpg")
    imgB(4).Picture = LoadPicture(imagePath & "BInfo.jpg")
    'Set imgB(4).Picture = Nothing
    
    imgC(0).Picture = LoadPicture(imagePath & "CMapa.jpg")
    imgC(1).Picture = LoadPicture(imagePath & "CListado.jpg")
    picB(4).Picture = LoadPicture(App.Path & "\Imagenes\todonegro.jpg")
    Exit Sub
TrapError:
    Call AltaLogGeneral("SERVER SQV", "Problemas al abrir archivos de piel de la aplicación. " & vbCrLf & "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source)
    'MsgBox "Problemas al abrir archivos de piel de la aplicación. " & vbCrLf & "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
    End
End Sub
    
Private Sub InicializarImagenes()
    MostrarPIC "A", 0 'Muestro Standard
    MostrarPIC "B", 0 'Muestro Datos de Sesion alternando con Asunto en Tratamiento
    MostrarPIC "C", 1 '091011 Muestra sin resultados 'Muestro Mapa de Bancas
End Sub

Private Sub SincronizarBancas(xBanca As String)
    Dim MensajeBanca As MensajeSistema

    With MensajeBanca
        .sTipo = "mset"
        .sObjeto = xBanca
        .sComponente = "term.mon"
        .sAtributo = "action"
        .sValor = "sync"
        .sComentario = "Sincronizar datos bancas"
    End With

    Call EnviarMensajesBancas(MensajeBanca)
End Sub

Private Sub SolicitarHabilitarVotoPresidente()
'si vota presidente
'1. Prende led
'2. si esta en estado votando, habilitar el teclado del presidente S='votando' AND CS (si vota presidente)
'    pend votar + 1

'HCDN 2011: Este mensaje solicita la identificacion del presidente para que prenda el led1, (mensaje sautod), entonces, cuando ser recibe respuesta de identificacion de la banca, se habilita el teclado. Como queda en modo identificado, los mensajes se tratan como votacion nominal siempre.
    EstadoActual.ModoVotaPresidente = True
    EstadoActual.PresidenteHabilitadoParaVotar = False
    'prende led mediante una identificación manual
    Mensaje2Banca.sObjeto = "0" '<AP 040115 faltaba indicar la banca>
    Mensaje2Banca.sTipo = "mset"
    Mensaje2Banca.sComponente = "term.led1"
    Mensaje2Banca.sAtributo = "state"
    Mensaje2Banca.sValor = "on_manual|" & Trim(EstadoActual.VectorIdentificacion(0))
    If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
        Mensaje2Banca.sComentario = "Presidente habilitado para votar Modo normal - led1 - 1"
    Else
        Mensaje2Banca.sComentario = "Presidente habilitado para votar Modo mantenimiento - led1"
    End If
    Call EnviarMensajesBancas(Mensaje2Banca)
    EstadoActual.PresidenteHabilitadoParaVotar = True
End Sub
Private Sub DeshabilitarVotoPresidente()
'apagar led1
'HCDN 2011: Este mensaje solicita la identificacion del presidente para que prenda el led1, (mensaje sautod), entonces, cuando ser recibe respuesta de identificacion de la banca, se habilita el teclado. Como queda en modo identificado, los mensajes se tratan como votacion nominal siempre.
    If EstadoActual.PresidenteHabilitadoParaVotar Then
        Call FinalizarVotacionPresidente
        Call EnviarMensajesFinAuth("0", "DeshabilitarVotoPresidente")  'apagar scanner / led
    End If
    EstadoActual.PresidenteHabilitadoParaVotar = False
    EstadoActual.ModoVotaPresidente = False
End Sub
Private Sub ComenzarVotacionPresidente()
    If EstadoActual.ModoVotaPresidente And EstadoActual.PresidenteHabilitadoParaVotar Then
        'habilitar teclado
        If (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And EstadoActual.EstadoVotacion_y_PasList = "votando" Then
            'se debe habilitar el teclado para votar
            With Mensaje2Banca ' Mensaje para SB
                .sTipo = "mset"
                .sComponente = "term.keyb"
                .sObjeto = 0
                .sAtributo = "state"
                .sValor = "onvotnom" 'SE TRATA COMO NOMINAL (SVOTAR)
                .sComentario = EstadoActual.EstadoVotacion_y_PasList & "Inicializacion de voto presidente"
            End With
            Call EnviarMensajesBancas(Mensaje2Banca)
        End If
    End If
End Sub
Private Sub FinalizarVotacionPresidente()
    'apaga teclado (pero no cancela modo habilitado)
    If EstadoActual.ModoVotaPresidente And EstadoActual.PresidenteHabilitadoParaVotar Then
        'deshabilitar teclado
        With Mensaje2Banca ' Mensaje para SB
            .sTipo = "mset"
            .sComponente = "term.keyb"
            .sObjeto = 0
            .sAtributo = "state"
            .sValor = "offvotnom"
            .sComentario = EstadoActual.EstadoVotacion_y_PasList & "Fin de voto presidente"
        End With
        Call EnviarMensajesBancas(Mensaje2Banca)
    End If
End Sub
Private Sub BorrarVotoPresidente()
    Mensaje2Banca.sTipo = "mset"
    Mensaje2Banca.sObjeto = "0"
    Mensaje2Banca.sComponente = "term.ledk1"
    Mensaje2Banca.sAtributo = "state"
    Mensaje2Banca.sValor = "off"
    Mensaje2Banca.sComentario = "Borro teclado presidente"
    Call EnviarMensajesBancas(Mensaje2Banca)
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
Private Function GetIdentificados() As Integer
Dim i As Integer
Dim total As Integer
total = 0
For i = 0 To 256
    If (EstadoActual.VectorIdentificacion(i) <> 0) Then
    total = total + 1
    End If
Next i
GetIdentificados = total
End Function
Private Function getPresentes() As Integer
Dim i As Integer
Dim total As Integer
total = 0
For i = 0 To 256
    If (EstadoActual.VectorPresencia(i) = PRESENTE Or i = 0) Then
        total = total + 1
    End If
Next i
'VotoRemoto
total = total + VotoRemoto.getPresentes()
getPresentes = total
End Function
Private Function GetAusentes() As Integer
Dim i As Integer
Dim total As Integer
total = 0
For i = 1 To 256
    If (EstadoActual.VectorPresencia(i) = AUSENTE Or EstadoActual.VectorPresencia(i) = "X") Then
        total = total + 1
    End If
Next i
GetAusentes = total
End Function
Public Sub FinPasLis()
Dim StrTempCadena As String
Dim Mensaje2Banca As MensajeSistema
Dim X As Integer
EstadoActual.EstadoVotacion_y_PasList = "esperafin"
EstadoActual.Modo_Ident_Nom = 0
EstadoActual.OcupadosNoIdentificados = GetNoIdentificadosSobrePresentes
EstadoActual.ActaGrabada = EstadoActual.NroActa
EstadoActual.SolicitudGrabarManual = 0
EstadoActual.PresentesCongelados = Presentes()
EstadoActual.AusentesCongelados = Ausentes()
EstadoActual.OcupadosNoIdentificadosCongelados = EstadoActual.OcupadosNoIdentificados
CartelActual.Resultado = "FINALIZADO"
EstadoActual.FechaVotacion = Now
StrTempCadena = "0" & SEPARADOR_VECTOR
For X = 1 To 256
    If EstadoActual.VectorPresencia(X) = PRESENTE And EstadoActual.VectorIdentificacion(X) = "0" Then ' AND EstadoActual.EnIdentificacion(X) = False And EstadoActual.VectorIdentificacion(X) = "0" Then
        StrTempCadena = StrTempCadena & "1" & SEPARADOR_VECTOR
    Else
        StrTempCadena = StrTempCadena & "0" & SEPARADOR_VECTOR
    End If
Next X
Call EnviarMensajesFinAuth(StrTempCadena, "Fin Pase de Lista")
lblGeneralInformacion.Caption = "PASE DE LISTA FINALIZADO"
Call MostrarCartel
Call AlmacenarActa
For X = 1 To 256
    EstadoActual.EnIdentificacion(X) = False
Next X
MandarImprimir
End Sub
Public Function GetDesconectadas() As String
Dim i As Integer
Dim Buff As String
Dim cConta As Integer
Dim cTotal As Integer
cConta = 0
Buff = ""
For i = 0 To 256
    If VectorDesconectadas(i) = True Then
        Buff = Buff & "," & Trim(Str(i))
        cConta = cConta + 1
        If cConta = 7 Then
            Buff = Buff & vbCrLf
            cConta = 0
        End If
    End If
Next i
GetDesconectadas = Buff
End Function
