VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Poder Legislativo - Servidor de Consolas"
   ClientHeight    =   12000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12000
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameSQVGeneral 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "SQV General"
      ForeColor       =   &H80000008&
      Height          =   12000
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15360
      Begin VB.PictureBox Picture3 
         Height          =   11600
         Left            =   90
         Picture         =   "frmMain.frx":0442
         ScaleHeight     =   11535
         ScaleWidth      =   15300
         TabIndex        =   104
         Top             =   -60
         Width           =   15360
         Begin MSWinsockLib.Winsock Ws 
            Left            =   1635
            Top             =   6285
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin VB.CommandButton cmdGeneralSalir 
            Caption         =   "&Salir"
            Height          =   315
            Left            =   14400
            TabIndex        =   197
            Top             =   10320
            Width           =   735
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   0
            Left            =   7470
            TabIndex        =   105
            Top             =   10125
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            Caption         =   "0"
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   1
            Left            =   2185
            TabIndex        =   106
            Top             =   9995
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            Caption         =   "1"
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   2
            Left            =   2245
            TabIndex        =   107
            Top             =   9220
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            Caption         =   "2"
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   3
            Left            =   2405
            TabIndex        =   108
            Top             =   8485
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            Caption         =   "3"
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   4
            Left            =   2680
            TabIndex        =   109
            Top             =   7770
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            Caption         =   "4"
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   5
            Left            =   3060
            TabIndex        =   110
            Top             =   7110
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            Caption         =   "5"
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   6
            Left            =   3540
            TabIndex        =   111
            Top             =   6510
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            Caption         =   "6"
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   7
            Left            =   4090
            TabIndex        =   112
            Top             =   6000
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            Caption         =   "7"
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   8
            Left            =   4710
            TabIndex        =   113
            Top             =   5560
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            Caption         =   "8"
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   9
            Left            =   5400
            TabIndex        =   114
            Top             =   5230
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   10
            Left            =   6120
            TabIndex        =   115
            Top             =   5010
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   11
            Left            =   6870
            TabIndex        =   116
            Top             =   4890
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   12
            Left            =   8040
            TabIndex        =   117
            Top             =   4890
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   13
            Left            =   8790
            TabIndex        =   118
            Top             =   5010
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   14
            Left            =   9510
            TabIndex        =   119
            Top             =   5220
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   15
            Left            =   10200
            TabIndex        =   120
            Top             =   5550
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   16
            Left            =   10830
            TabIndex        =   121
            Top             =   5970
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   17
            Left            =   11400
            TabIndex        =   122
            Top             =   6510
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   18
            Left            =   11880
            TabIndex        =   123
            Top             =   7080
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   19
            Left            =   12270
            TabIndex        =   124
            Top             =   7770
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   20
            Left            =   12540
            TabIndex        =   125
            Top             =   8460
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   21
            Left            =   12690
            TabIndex        =   126
            Top             =   9210
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   22
            Left            =   12750
            TabIndex        =   127
            Top             =   9960
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   23
            Left            =   3070
            TabIndex        =   128
            Top             =   9990
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   24
            Left            =   3130
            TabIndex        =   129
            Top             =   9295
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   25
            Left            =   3300
            TabIndex        =   130
            Top             =   8620
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   26
            Left            =   3570
            TabIndex        =   131
            Top             =   7980
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   27
            Left            =   3950
            TabIndex        =   132
            Top             =   7405
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   28
            Left            =   4420
            TabIndex        =   133
            Top             =   6890
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   29
            Left            =   4970
            TabIndex        =   134
            Top             =   6460
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   30
            Left            =   5580
            TabIndex        =   135
            Top             =   6120
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   31
            Left            =   6230
            TabIndex        =   136
            Top             =   5890
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   32
            Left            =   6920
            TabIndex        =   137
            Top             =   5770
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   33
            Left            =   8010
            TabIndex        =   138
            Top             =   5780
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   34
            Left            =   8680
            TabIndex        =   139
            Top             =   5910
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   35
            Left            =   9340
            TabIndex        =   140
            Top             =   6120
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   36
            Left            =   9960
            TabIndex        =   141
            Top             =   6450
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   37
            Left            =   10500
            TabIndex        =   142
            Top             =   6870
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   38
            Left            =   10980
            TabIndex        =   143
            Top             =   7380
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   39
            Left            =   11360
            TabIndex        =   144
            Top             =   7950
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   40
            Left            =   11640
            TabIndex        =   145
            Top             =   8580
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   41
            Left            =   11820
            TabIndex        =   146
            Top             =   9270
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   42
            Left            =   11880
            TabIndex        =   147
            Top             =   9960
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   43
            Left            =   3960
            TabIndex        =   148
            Top             =   9990
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   44
            Left            =   4030
            TabIndex        =   149
            Top             =   9300
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   45
            Left            =   4250
            TabIndex        =   150
            Top             =   8640
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   46
            Left            =   4590
            TabIndex        =   151
            Top             =   8040
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   47
            Left            =   5050
            TabIndex        =   152
            Top             =   7520
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   48
            Left            =   5610
            TabIndex        =   153
            Top             =   7110
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   49
            Left            =   6240
            TabIndex        =   154
            Top             =   6820
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   50
            Left            =   6920
            TabIndex        =   155
            Top             =   6680
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   51
            Left            =   8010
            TabIndex        =   156
            Top             =   6670
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   52
            Left            =   8690
            TabIndex        =   157
            Top             =   6820
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   53
            Left            =   9320
            TabIndex        =   158
            Top             =   7100
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   54
            Left            =   9870
            TabIndex        =   159
            Top             =   7500
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   55
            Left            =   10350
            TabIndex        =   160
            Top             =   8010
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   56
            Left            =   10690
            TabIndex        =   161
            Top             =   8610
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   57
            Left            =   10910
            TabIndex        =   162
            Top             =   9270
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   58
            Left            =   10980
            TabIndex        =   163
            Top             =   9960
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   59
            Left            =   4860
            TabIndex        =   164
            Top             =   9990
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   60
            Left            =   4950
            TabIndex        =   165
            Top             =   9300
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   61
            Left            =   5230
            TabIndex        =   166
            Top             =   8675
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   62
            Left            =   5680
            TabIndex        =   167
            Top             =   8152
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   63
            Left            =   6250
            TabIndex        =   168
            Top             =   7775
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   64
            Left            =   6920
            TabIndex        =   169
            Top             =   7570
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   65
            Left            =   8000
            TabIndex        =   170
            Top             =   7570
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   66
            Left            =   8670
            TabIndex        =   171
            Top             =   7770
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   67
            Left            =   9240
            TabIndex        =   172
            Top             =   8130
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   68
            Left            =   9690
            TabIndex        =   173
            Top             =   8650
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   69
            Left            =   9990
            TabIndex        =   174
            Top             =   9270
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin Project1.ctrBanca ctrBanca 
            Height          =   420
            Index           =   70
            Left            =   10080
            TabIndex        =   175
            Top             =   9960
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   741
            BackColor       =   8438015
            ForeColor       =   3
         End
         Begin VB.Label lblPruebas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base de Pruebas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   360
            Left            =   6480
            TabIndex        =   341
            Top             =   8520
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.Label lblGeneralMayoriaDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Mayoria"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   465
            Left            =   240
            TabIndex        =   196
            Top             =   3120
            Width           =   14895
         End
         Begin VB.Label lblGeneralAbstenciones 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "ABSTENCIONES:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   585
            Left            =   10200
            TabIndex        =   195
            Top             =   10800
            Width           =   4095
         End
         Begin VB.Label lblGeneralAfirmativos 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "AFIRMATIVOS:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   585
            Left            =   360
            TabIndex        =   194
            Top             =   10800
            Width           =   3615
         End
         Begin VB.Label lblGeneralNegativos 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "NEGATIVOS:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   585
            Left            =   5520
            TabIndex        =   193
            Top             =   10800
            Width           =   3255
         End
         Begin VB.Label lblGeneralAbstencionesDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   14280
            TabIndex        =   192
            Top             =   10725
            Width           =   855
         End
         Begin VB.Label lblGeneralAfirmativosDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   3840
            TabIndex        =   191
            Top             =   10730
            Width           =   1455
         End
         Begin VB.Label lblGeneralNegativosDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   8640
            TabIndex        =   190
            Top             =   10725
            Width           =   1455
         End
         Begin VB.Label lblGeneralResultadoDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "RESULTADO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   5760
            TabIndex        =   189
            Top             =   8920
            Width           =   3855
         End
         Begin VB.Label lblGeneralTiempoDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "CUMPLIDO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   11160
            TabIndex        =   188
            Top             =   3760
            Width           =   3975
         End
         Begin VB.Label lblGeneralTiempo 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "TIEMPO:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   585
            Left            =   9120
            TabIndex        =   187
            Top             =   3840
            Width           =   2415
         End
         Begin VB.Label lblGeneralTipoOperacionDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Votacion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   585
            Left            =   240
            TabIndex        =   186
            Top             =   3840
            Width           =   6495
         End
         Begin VB.Label lblGeneralOrdenDiaDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Orden del Dia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   465
            Left            =   240
            TabIndex        =   185
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label lblGeneralTituloDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Titulo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   465
            Left            =   240
            TabIndex        =   184
            Top             =   2640
            Width           =   14895
         End
         Begin VB.Label lblGeneralLeyendaQuorumDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "NO HAY QUORUM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   465
            Left            =   10440
            TabIndex        =   183
            Top             =   1080
            Width           =   4575
         End
         Begin VB.Label lblGeneralHoraDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "08:50"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   465
            Left            =   5880
            TabIndex        =   182
            Top             =   1080
            Width           =   3615
         End
         Begin VB.Label lblGeneralFechaDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "29/01/2004"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   465
            Left            =   720
            TabIndex        =   181
            Top             =   1080
            Width           =   3615
         End
         Begin VB.Label lblGeneralAusentesDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   11760
            TabIndex        =   180
            Top             =   135
            Width           =   1335
         End
         Begin VB.Label lblGeneralPresentesDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   4440
            TabIndex        =   179
            Top             =   135
            Width           =   1335
         End
         Begin VB.Label lblGeneralAusentes 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "AUSENTES:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   465
            Left            =   8160
            TabIndex        =   178
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label lblGeneralPresentes 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "PRESENTES:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   465
            Left            =   840
            TabIndex        =   177
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label lblGeneralSesionDato 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Período Legislativo y Sesión "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   585
            Left            =   360
            TabIndex        =   176
            Top             =   1830
            Width           =   14895
         End
      End
   End
   Begin VB.Frame FrameControl 
      Caption         =   "Control de SQV"
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9855
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
         TabIndex        =   79
         Top             =   1680
         Width           =   9675
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Presentes : "
            Height          =   195
            Left            =   1920
            TabIndex        =   97
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label xx 
            Alignment       =   1  'Right Justify
            Caption         =   "Ausentes : "
            Height          =   195
            Left            =   4320
            TabIndex        =   96
            Top             =   255
            Width           =   1005
         End
         Begin VB.Label rr 
            Alignment       =   1  'Right Justify
            Caption         =   "Resultado : "
            Height          =   195
            Left            =   6720
            TabIndex        =   95
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label cc 
            Alignment       =   1  'Right Justify
            Caption         =   "Afirmativos : "
            Height          =   195
            Left            =   1920
            TabIndex        =   94
            Top             =   525
            Width           =   1005
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Negativos : "
            Height          =   195
            Left            =   4320
            TabIndex        =   93
            Top             =   540
            Width           =   1005
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Abstenciones : "
            Height          =   195
            Left            =   6720
            TabIndex        =   92
            Top             =   555
            Width           =   1005
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Minimo de Votos Para Afirmativo : "
            Height          =   255
            Left            =   1800
            TabIndex        =   91
            Top             =   930
            Width           =   2595
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Leyenda Quorum : "
            Height          =   255
            Left            =   5280
            TabIndex        =   90
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
            TabIndex        =   89
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
            TabIndex        =   88
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
            TabIndex        =   87
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
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   84
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
            TabIndex        =   83
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
            TabIndex        =   82
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
            TabIndex        =   81
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
            TabIndex        =   80
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
         TabIndex        =   15
         Top             =   3240
         Width           =   9675
         Begin VB.Label lblVersionSQV 
            Alignment       =   1  'Right Justify
            Caption         =   "Merge040225a:"
            Height          =   195
            Left            =   5160
            TabIndex        =   340
            Top             =   4100
            Width           =   2175
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Vector Color() :"
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Vector Presencia() :"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   735
            Width           =   2655
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Vector Identificacion() :"
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   990
            Width           =   2655
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Vector Resultados() :"
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   1245
            Width           =   2655
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Ocupados No Identificados :"
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   1500
            Width           =   2655
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Pendientes de Emitir Votos :"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   1755
            Width           =   2655
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Sesión :"
            Height          =   255
            Left            =   240
            TabIndex        =   72
            Top             =   2010
            Width           =   2655
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Periodo Legislativo :"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   2265
            Width           =   2655
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Nº de Acta :"
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   2520
            Width           =   2655
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Título Del Acta :"
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   2775
            Width           =   2655
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Identificador De Formulario :"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   3030
            Width           =   2655
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "IP Consola :"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   3285
            Width           =   2655
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Base Mayoria :"
            Height          =   195
            Left            =   5160
            TabIndex        =   66
            Top             =   1515
            Width           =   2205
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Mayoria :"
            Height          =   195
            Left            =   5160
            TabIndex        =   65
            Top             =   1785
            Width           =   2205
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Mayoria Quórum :"
            Height          =   195
            Left            =   5160
            TabIndex        =   64
            Top             =   2040
            Width           =   2205
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo De Operación :"
            Height          =   195
            Left            =   5160
            TabIndex        =   63
            Top             =   2295
            Width           =   2205
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "Tiempo Para Votación :"
            Height          =   195
            Left            =   5160
            TabIndex        =   62
            Top             =   2565
            Width           =   2205
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Error :"
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   3540
            Width           =   2655
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Estado de Votación y Pase de Lista :"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   3795
            Width           =   2655
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Modalidad de Votación :"
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   4050
            Width           =   2655
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Mensaje Al Operador :"
            Height          =   195
            Left            =   5160
            TabIndex        =   58
            Top             =   480
            Width           =   2205
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Modo Mantenimiento de Bancas :"
            Height          =   195
            Left            =   4920
            TabIndex        =   57
            Top             =   735
            Width           =   2445
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Modo Normal Mant Sistema :"
            Height          =   195
            Left            =   5160
            TabIndex        =   56
            Top             =   1005
            Width           =   2205
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Cartel Encendido :"
            Height          =   195
            Left            =   5160
            TabIndex        =   55
            Top             =   1260
            Width           =   2205
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "Grabar Automáticamente :"
            Height          =   195
            Left            =   5160
            TabIndex        =   54
            Top             =   2820
            Width           =   2205
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "Listar Automáticamente :"
            Height          =   195
            Left            =   5160
            TabIndex        =   53
            Top             =   3075
            Width           =   2205
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   "Acta Grabada :"
            Height          =   195
            Left            =   5160
            TabIndex        =   52
            Top             =   3345
            Width           =   2205
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "Solicitud Grabar Manual :"
            Height          =   195
            Left            =   5160
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            Left            =   2880
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
            Top             =   3795
            Width           =   1485
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            Caption         =   "Estado de Sesión :"
            Height          =   195
            Left            =   5160
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label LblAbstencionesAut 
            Alignment       =   1  'Right Justify
            Caption         =   "Abs.Aut"
            Height          =   255
            Left            =   120
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   4080
            Width           =   525
         End
      End
      Begin VB.Timer Timer 
         Interval        =   100
         Left            =   6000
         Top             =   945
      End
      Begin VB.TextBox txtVecesPorSegundo 
         Alignment       =   2  'Center
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
         Left            =   2280
         TabIndex        =   14
         Text            =   "1"
         Top             =   945
         Width           =   615
      End
      Begin VB.PictureBox Picture1 
         Height          =   1050
         Left            =   7920
         ScaleHeight     =   990
         ScaleWidth      =   1695
         TabIndex        =   11
         Top             =   720
         Width           =   1750
         Begin VB.CommandButton cmdSalir 
            Caption         =   "&Salir"
            Height          =   495
            Left            =   0
            TabIndex        =   13
            Top             =   510
            Width           =   1695
         End
         Begin VB.CommandButton cmdTerminar 
            Caption         =   "&Iniciar Server"
            Height          =   495
            Left            =   0
            TabIndex        =   12
            Top             =   15
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   810
         Left            =   0
         ScaleHeight     =   750
         ScaleWidth      =   2055
         TabIndex        =   8
         Top             =   480
         Width           =   2120
         Begin VB.CommandButton HabilitarSeguimientoPizarraRecinto 
            Caption         =   "&Ocultar Estado de Recinto"
            Height          =   375
            Left            =   0
            TabIndex        =   10
            Top             =   375
            Width           =   2055
         End
         Begin VB.CommandButton HabilitarSeguimientoPizarraCartel 
            Caption         =   "&Ocultar Estado de Cartel"
            Height          =   375
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   2055
         End
      End
      Begin VB.CheckBox chkLog_Mensajes 
         Caption         =   "Guardar Copia de Mensajes"
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   1365
         Width           =   2535
      End
      Begin VB.CommandButton cmdResetarVectores 
         Caption         =   "Reset"
         Height          =   315
         Left            =   7080
         TabIndex        =   6
         ToolTipText     =   "Resetear vectores de estado"
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Config"
         Height          =   315
         Left            =   6360
         TabIndex        =   5
         ToolTipText     =   "Resetear vectores de estado"
         Top             =   1185
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Dar Quorum"
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   1065
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "SB"
         Height          =   495
         Left            =   4560
         TabIndex        =   3
         Top             =   1065
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cartel"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   1185
         Width           =   975
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
         TabIndex        =   103
         Top             =   585
         Width           =   1215
      End
      Begin VB.Label lblFechaInicioServer 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   102
         Top             =   585
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ciclo por segundo"
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
         Left            =   3000
         TabIndex        =   101
         Top             =   990
         Width           =   1545
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
         Left            =   5520
         TabIndex        =   100
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label lblCiclos 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   6960
         TabIndex        =   99
         Top             =   360
         Width           =   2655
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
         TabIndex        =   98
         Top             =   885
         Width           =   3135
      End
   End
   Begin VB.Frame FrameSQVApagado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12000
      Left            =   0
      TabIndex        =   198
      Top             =   0
      Width           =   15360
   End
   Begin VB.Frame FrameMantenimiento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12000
      Left            =   0
      TabIndex        =   199
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
         TabIndex        =   217
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
         TabIndex        =   216
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
         TabIndex        =   215
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
         TabIndex        =   214
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
         TabIndex        =   213
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
         TabIndex        =   212
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
         TabIndex        =   211
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
         TabIndex        =   210
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
         TabIndex        =   209
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
         TabIndex        =   208
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
         TabIndex        =   207
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
         TabIndex        =   206
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
         TabIndex        =   205
         Top             =   3540
         Width           =   11505
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
         TabIndex        =   204
         Top             =   2565
         Width           =   11505
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
         TabIndex        =   203
         Top             =   1860
         Width           =   11505
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
         TabIndex        =   202
         Top             =   1095
         Width           =   11505
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
         TabIndex        =   201
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
         TabIndex        =   200
         Top             =   -15
         Width           =   1215
      End
   End
   Begin VB.Frame FrameSQVActa 
      BorderStyle     =   0  'None
      Height          =   12000
      Left            =   0
      TabIndex        =   218
      Top             =   0
      Width           =   15360
      Begin VB.TextBox txtPagina 
         Height          =   285
         Left            =   5685
         TabIndex        =   339
         Text            =   "Text1"
         Top             =   11640
         Width           =   1770
      End
      Begin VB.Frame frameActaDatos 
         Height          =   11595
         Left            =   45
         TabIndex        =   219
         Top             =   0
         Width           =   15285
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   27
            Left            =   13110
            TabIndex        =   338
            Top             =   11100
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   27
            Left            =   7905
            TabIndex        =   337
            Top             =   11100
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   26
            Left            =   13110
            TabIndex        =   336
            Top             =   10620
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   26
            Left            =   7890
            TabIndex        =   335
            Top             =   10620
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   25
            Left            =   13110
            TabIndex        =   334
            Top             =   10170
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   25
            Left            =   7890
            TabIndex        =   333
            Top             =   10170
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   24
            Left            =   13095
            TabIndex        =   332
            Top             =   9690
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   24
            Left            =   7875
            TabIndex        =   331
            Top             =   9690
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   23
            Left            =   13095
            TabIndex        =   330
            Top             =   9195
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   23
            Left            =   7875
            TabIndex        =   329
            Top             =   9195
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   22
            Left            =   13080
            TabIndex        =   328
            Top             =   8685
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   22
            Left            =   7860
            TabIndex        =   327
            Top             =   8715
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   21
            Left            =   13065
            TabIndex        =   326
            Top             =   8220
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   21
            Left            =   7845
            TabIndex        =   325
            Top             =   8220
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   20
            Left            =   13050
            TabIndex        =   324
            Top             =   7740
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   20
            Left            =   7830
            TabIndex        =   323
            Top             =   7740
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   19
            Left            =   13050
            TabIndex        =   322
            Top             =   7245
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   19
            Left            =   7830
            TabIndex        =   321
            Top             =   7245
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   18
            Left            =   13035
            TabIndex        =   320
            Top             =   6750
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   18
            Left            =   7815
            TabIndex        =   319
            Top             =   6765
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   17
            Left            =   13035
            TabIndex        =   318
            Top             =   6270
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   17
            Left            =   7815
            TabIndex        =   317
            Top             =   6270
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   16
            Left            =   13020
            TabIndex        =   316
            Top             =   5790
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   16
            Left            =   7800
            TabIndex        =   315
            Top             =   5790
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   15
            Left            =   13020
            TabIndex        =   314
            Top             =   5295
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   15
            Left            =   7800
            TabIndex        =   313
            Top             =   5295
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   14
            Left            =   13005
            TabIndex        =   312
            Top             =   4800
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   14
            Left            =   7770
            TabIndex        =   311
            Top             =   4815
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   13
            Left            =   5475
            TabIndex        =   310
            Top             =   11085
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   13
            Left            =   270
            TabIndex        =   309
            Top             =   11085
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   12
            Left            =   5475
            TabIndex        =   308
            Top             =   10605
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   12
            Left            =   255
            TabIndex        =   307
            Top             =   10605
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   11
            Left            =   5475
            TabIndex        =   306
            Top             =   10155
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   11
            Left            =   255
            TabIndex        =   305
            Top             =   10155
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   10
            Left            =   5460
            TabIndex        =   304
            Top             =   9675
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   10
            Left            =   240
            TabIndex        =   303
            Top             =   9675
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   9
            Left            =   5460
            TabIndex        =   302
            Top             =   9180
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   9
            Left            =   240
            TabIndex        =   301
            Top             =   9180
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   8
            Left            =   5445
            TabIndex        =   300
            Top             =   8670
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   8
            Left            =   225
            TabIndex        =   299
            Top             =   8700
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   7
            Left            =   5430
            TabIndex        =   298
            Top             =   8205
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   7
            Left            =   210
            TabIndex        =   297
            Top             =   8205
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   6
            Left            =   5430
            TabIndex        =   296
            Top             =   7725
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   6
            Left            =   195
            TabIndex        =   295
            Top             =   7725
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   5
            Left            =   5415
            TabIndex        =   294
            Top             =   7230
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   5
            Left            =   195
            TabIndex        =   293
            Top             =   7230
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   4
            Left            =   5400
            TabIndex        =   292
            Top             =   6735
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   4
            Left            =   180
            TabIndex        =   291
            Top             =   6750
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   3
            Left            =   5400
            TabIndex        =   290
            Top             =   6255
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   3
            Left            =   180
            TabIndex        =   289
            Top             =   6255
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   2
            Left            =   5385
            TabIndex        =   288
            Top             =   5775
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   2
            Left            =   165
            TabIndex        =   287
            Top             =   5775
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   1
            Left            =   5385
            TabIndex        =   286
            Top             =   5280
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   1
            Left            =   165
            TabIndex        =   285
            Top             =   5280
            Width           =   5055
         End
         Begin VB.TextBox txtActaResultado 
            Height          =   420
            Index           =   0
            Left            =   5370
            TabIndex        =   284
            Top             =   4785
            Width           =   2190
         End
         Begin VB.TextBox txtActaLegislador 
            Height          =   420
            Index           =   0
            Left            =   135
            TabIndex        =   283
            Top             =   4800
            Width           =   5055
         End
         Begin VB.CommandButton cmdPresidente 
            Caption         =   "Cam&biar presidente"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9150
            TabIndex        =   254
            Top             =   1950
            Width           =   1515
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   315
            Left            =   1830
            TabIndex        =   253
            Top             =   2370
            Width           =   7185
         End
         Begin VB.Frame Frame3 
            Height          =   45
            Left            =   510
            TabIndex        =   252
            Top             =   4410
            Width           =   11235
         End
         Begin VB.TextBox txtAbstencionesTotales 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   11100
            Locked          =   -1  'True
            TabIndex        =   251
            Top             =   4020
            Width           =   615
         End
         Begin VB.TextBox txtAbstencionesNoId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   11100
            Locked          =   -1  'True
            TabIndex        =   250
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox txtAbstencionesId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   11100
            Locked          =   -1  'True
            TabIndex        =   249
            Top             =   2850
            Width           =   615
         End
         Begin VB.TextBox txtNegativoTotales 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   8100
            Locked          =   -1  'True
            TabIndex        =   248
            Top             =   4020
            Width           =   615
         End
         Begin VB.TextBox txtNegativoDesempate 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   8100
            Locked          =   -1  'True
            TabIndex        =   247
            Top             =   3630
            Width           =   615
         End
         Begin VB.TextBox txtNegativoNoId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   8100
            Locked          =   -1  'True
            TabIndex        =   246
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox txtNegativoID 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   8100
            Locked          =   -1  'True
            TabIndex        =   245
            Top             =   2850
            Width           =   615
         End
         Begin VB.TextBox txtAfirmativosTotal 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   244
            Top             =   4020
            Width           =   615
         End
         Begin VB.TextBox txtAfirmativosDesempate 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   243
            Top             =   3630
            Width           =   615
         End
         Begin VB.TextBox txtAfirmativosNoId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   242
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox txtAfirmativosId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   241
            Top             =   2850
            Width           =   615
         End
         Begin VB.TextBox txtAusentesTotal 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   240
            Top             =   4020
            Width           =   615
         End
         Begin VB.TextBox txtPresentesTotal 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   239
            Top             =   3630
            Width           =   615
         End
         Begin VB.TextBox txtPresentesNoId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   238
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox txtPresentesId 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   237
            Top             =   2850
            Width           =   615
         End
         Begin VB.Frame Frame2 
            Height          =   45
            Left            =   510
            TabIndex        =   236
            Top             =   2730
            Width           =   11235
         End
         Begin VB.TextBox txtNombrePresidente 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   2850
            Locked          =   -1  'True
            TabIndex        =   235
            Top             =   1980
            Width           =   6165
         End
         Begin VB.TextBox txtCodigoPresidente 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   234
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
            TabIndex        =   233
            Top             =   1590
            Width           =   1365
         End
         Begin VB.TextBox txtBase 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   232
            Top             =   1590
            Width           =   2115
         End
         Begin VB.TextBox txtDesempate 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   10350
            Locked          =   -1  'True
            TabIndex        =   231
            Top             =   1200
            Width           =   1365
         End
         Begin VB.TextBox txtMiembros 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   230
            Top             =   1200
            Width           =   2115
         End
         Begin VB.TextBox txtTipoMayoria 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   229
            Top             =   1590
            Width           =   3105
         End
         Begin VB.TextBox txtTipoQuorum 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   228
            Top             =   1200
            Width           =   3105
         End
         Begin VB.Frame Frame1 
            Height          =   45
            Left            =   510
            TabIndex        =   227
            Top             =   1110
            Width           =   11235
         End
         Begin VB.TextBox txtHora 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   10350
            Locked          =   -1  'True
            TabIndex        =   226
            Top             =   750
            Width           =   1365
         End
         Begin VB.TextBox txtFecha 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   9180
            Locked          =   -1  'True
            TabIndex        =   225
            Top             =   750
            Width           =   1035
         End
         Begin VB.TextBox txtNombre 
            Height          =   315
            Left            =   1830
            TabIndex        =   224
            Top             =   750
            Width           =   7155
         End
         Begin VB.TextBox txtVersion 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   10350
            Locked          =   -1  'True
            TabIndex        =   223
            Top             =   360
            Width           =   1365
         End
         Begin VB.TextBox txtNroActa 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   7500
            Locked          =   -1  'True
            TabIndex        =   222
            Top             =   360
            Width           =   1485
         End
         Begin VB.TextBox txtSesion 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5730
            Locked          =   -1  'True
            TabIndex        =   221
            Top             =   360
            Width           =   1035
         End
         Begin VB.TextBox txtTipoOperacion 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   220
            Top             =   360
            Width           =   3105
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Left            =   480
            TabIndex        =   282
            Top             =   2430
            Width           =   1065
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            Caption         =   "Abstenciones Totales"
            Height          =   195
            Left            =   8880
            TabIndex        =   281
            Top             =   4080
            Width           =   1530
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Abstenciones No Identificables"
            Height          =   195
            Left            =   8880
            TabIndex        =   280
            Top             =   3300
            Width           =   2190
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            Caption         =   "Abstenciones identificables"
            Height          =   195
            Left            =   8880
            TabIndex        =   279
            Top             =   2910
            Width           =   1920
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            Caption         =   "Votos Negativos Totales"
            Height          =   195
            Left            =   6030
            TabIndex        =   278
            Top             =   4080
            Width           =   1740
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            Caption         =   "Votos Neg. Desempate"
            Height          =   195
            Left            =   6030
            TabIndex        =   277
            Top             =   3690
            Width           =   1650
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            Caption         =   "Votos Neg. No Identificables"
            Height          =   195
            Left            =   6030
            TabIndex        =   276
            Top             =   3300
            Width           =   2025
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "Votos Neg. Identificables"
            Height          =   195
            Left            =   6030
            TabIndex        =   275
            Top             =   2910
            Width           =   1770
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            Caption         =   "Votos Afirmativos Total"
            Height          =   195
            Left            =   3180
            TabIndex        =   274
            Top             =   4080
            Width           =   1620
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "Votos Afirm. Desempate"
            Height          =   195
            Left            =   3180
            TabIndex        =   273
            Top             =   3690
            Width           =   1695
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            Caption         =   "Votos Afirm. No Identificables"
            Height          =   195
            Left            =   3180
            TabIndex        =   272
            Top             =   3300
            Width           =   2070
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "Votos Afirm. Identificables"
            Height          =   195
            Left            =   3180
            TabIndex        =   271
            Top             =   2910
            Width           =   1815
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            Caption         =   "Ausentes Total"
            Height          =   195
            Left            =   480
            TabIndex        =   270
            Top             =   4080
            Width           =   1065
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "Presentes Total"
            Height          =   195
            Left            =   480
            TabIndex        =   269
            Top             =   3690
            Width           =   1110
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "Presentes no identificables"
            Height          =   195
            Left            =   480
            TabIndex        =   268
            Top             =   3300
            Width           =   1890
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Presentes identificables"
            Height          =   195
            Left            =   480
            TabIndex        =   267
            Top             =   2910
            Width           =   1665
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "Presidente"
            Height          =   195
            Left            =   480
            TabIndex        =   266
            Top             =   2040
            Width           =   750
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Votación"
            Height          =   195
            Left            =   9180
            TabIndex        =   265
            Top             =   1650
            Width           =   630
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Base"
            Height          =   195
            Left            =   5190
            TabIndex        =   264
            Top             =   1650
            Width           =   360
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Desempate"
            Height          =   195
            Left            =   9180
            TabIndex        =   263
            Top             =   1260
            Width           =   810
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Miembros del cuerpo"
            Height          =   195
            Left            =   5190
            TabIndex        =   262
            Top             =   1260
            Width           =   1470
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de mayoría"
            Height          =   195
            Left            =   480
            TabIndex        =   261
            Top             =   1650
            Width           =   1155
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de quorum"
            Height          =   195
            Left            =   480
            TabIndex        =   260
            Top             =   1260
            Width           =   1110
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Nombre del acta"
            Height          =   195
            Left            =   480
            TabIndex        =   259
            Top             =   810
            Width           =   1170
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Versión"
            Height          =   195
            Left            =   9180
            TabIndex        =   258
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Nº acta"
            Height          =   195
            Left            =   6870
            TabIndex        =   257
            Top             =   420
            Width           =   540
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Sesión"
            Height          =   195
            Left            =   5190
            TabIndex        =   256
            Top             =   420
            Width           =   480
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de operación"
            Height          =   195
            Left            =   480
            TabIndex        =   255
            Top             =   420
            Width           =   1290
         End
      End
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
Private WithEvents Rs          As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1
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
Private xFechaInicioProceso As Date
Private blBanderaTimer         As Boolean
Private xFileSqv As Long
Private xNroMensajeSB           As Long
Private xPrimerMensajeSB As Long

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
    Set Rs = New ADODB.Recordset
    Set RsWrite = New ADODB.Recordset
Exit Sub
TrapError:
    MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
    Resume
End Sub
Private Sub SetearRs(strCadena As String)
    On Error GoTo TrapError
    'Set Rs = New ADODB.Recordset
    With Rs
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
    Resume
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
    Resume
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
    Resume
End Sub
Public Function SetearRsAux(pCadena As String, ByRef pRst As ADODB.Recordset) As Boolean
    SetearRsAux = False
    pRst.CursorLocation = adUseClient
    pRst.Open pCadena, Cn, adOpenForwardOnly, adLockReadOnly
    If Not pRst.BOF And Not pRst.EOF Then
         SetearRsAux = True
    End If
End Function
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
Private Sub ResetearVectores()
    Dim X      As Long
    With EstadoActual
        For X = 0 To (xUltimaBanca)
            .VectorPresencia(X) = BANCA_INHABILITADA
            .VectorIdentificacion(X) = NO_IDENTIFICADO
            .VectorColor(X) = AsignarColor(X)
            .VectorResultados(X) = ABSTENCION
            .VMantEstado(X) = ABSTENCION
        Next X
        
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
    With Rs
        While Not .EOF
            X = Int(.Fields("deskid").Value)
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
        .ActaGrabada = 0
        .Ausentes = xMiembrosDelCuerpo
        .BaseMayoria = "legpre"
        
        'cartel control
        .CartelEncendido = 1
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
    Call ReinicioSistema
    EstadoActual.TipoDeOperacion = "votnom"
    Call InicializarVotacion
    EstadoActual.EstadoVotacion_y_PasList = "votando"
    EstadoActual.TiempoParaVotacion = 9999
    EstadoActual.FechaVotacion = DateAdd("s", xtiempoInicioVotac, Now)
    EstadoActual.TituloDelActa = "MANTENIMIENTO DEL SISTEMA SQV"
    xPresidenteLegislador = True
End Sub
Private Sub Fin_Mantenimiento_SQV()
    With EstadoActual
        .CartelEncendido = 0
        FrameSQVApagado.ZOrder 0
        .ModoMantenimientoBancas = 0
        .ModoNormalMantSistema = 0
        .TituloDelActa = " "
    End With
    Call ReinicioSistema
    EstadoActual.TiempoParaVotacion = 15
    xPresidenteLegislador = False
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
    MsgBox "error"
End Sub
Private Sub Levanta_Banca()
    Ws.Close
    Ws.RemoteHost = Trim(strIpServer)
    Ws.RemotePort = strPuerto
    Ws.Connect
    strPath = strExeSb
    DoEvents
    While Ws.State = 6
        DoEvents
    Wend
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

Private Sub AlmacenarActa()

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
               & "Zona_asignada , Bloque_político, Departamento " _
               & "From DetalleActas WHERE 1 = 0"
        Call SetearRsW(strSql)
        ' Determinar si se debe considerar al presidente como legislador o como vicegobernador
    End If
    xInicio = IIf(xPresidenteLegislador, 0, 1)
    For X = xInicio To xUltimaBanca
        ' VECTOR IDENTIFICACION
        strIdLegislador = ""
        If EstadoActual.VectorPresencia(X) = PRESENTE _
            And (((EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis") And EstadoActual.VectorIdentificacion(X) <> NO_IDENTIFICADO) _
                Or (EstadoActual.TipoDeOperacion = "votnum")) Then
            If (((EstadoActual.TipoDeOperacion = "votnom" And xTipoVotacion <> "votnum") Or EstadoActual.TipoDeOperacion = "paslis") And EstadoActual.VectorIdentificacion(X) <> NO_IDENTIFICADO) Then
                ' Identificar al legislador
                strIdLegislador = Trim(LCase(EstadoActual.VectorIdentificacion(X)))
                strBuscarLegislador = "SELECT * FROM Legisladores WHERE id = '" & strIdLegislador & "'"
                rsTemp.Open strBuscarLegislador, Cn, adOpenForwardOnly, adLockReadOnly
                If Not rsTemp.EOF Then
                    strBloquePolitico = GetCadena(rsTemp.Fields("Bloque_Politico").Value)
                    strDepartamento = GetCadena(rsTemp.Fields("Departamento").Value)
                    xZonaAsignada = GetNumero(rsTemp.Fields("Zona").Value)
                Else
                    strBloquePolitico = "BLOQUE NO IDENTIFICADO"
                    strDepartamento = "ND"
                    xZonaAsignada = -1
                End If
                rsTemp.Close
            End If
            Select Case EstadoActual.TipoDeOperacion
            Case "votnom", "votnum"
                Select Case EstadoActual.VectorResultados(X)
                    Case AFIRMATIVO
                        strResultado = "AFIRMATIVO"
                    Case NEGATIVO
                        strResultado = "NEGATIVO"
                    Case ABSTENCION, ABSTENCION_AUTORIZADA
                        strResultado = "ABSTENCION"
                End Select
            Case "paslis"
                strResultado = "PRESENTE"
            End Select
        Else
            strResultado = "AUSENTE"
        End If
        ' Escribir tabla solo en votnom, paslis
        If (EstadoActual.TipoDeOperacion = "votnom" And xTipoVotacion <> "votnum") Or EstadoActual.TipoDeOperacion = "paslis" Then
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
            RsWrite.Fields("Departamento").Value = Trim(strDepartamento)
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
                    If EstadoActual.VectorResultados(X) = AFIRMATIVO Then
                        xVotosAfirmIdentificables = xVotosAfirmIdentificables + 1
                    ElseIf EstadoActual.VectorResultados(X) = NEGATIVO Then
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
                If EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum" Then
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
    Next X
    If xPresidenteLegislador Then
        xAusentesTotales = xAusentesTotales - 1
    End If
    If (EstadoActual.TipoDeOperacion = "votnom" And xTipoVotacion <> "votnum") Or EstadoActual.TipoDeOperacion = "paslis" Then
        RsWrite.Close
    End If
    'Totales generales
    xVotosAfirmTotales = xVotosAfirmIdentificables + xVotosAfirmNOIdentificables
    xVotosNegatTotales = xVotosNegatNOIdentificables + xVotosNegatIdentificables
    xAbstencionesTotales = xAbstencionesIdentificables + xAbstencionesNOIdentificables
    xPresentesTotales = xPresentesIdentificables + xPresentesNOIdentificables
    
    
    strDesempate = IIf((xVotosAfirmTotales = xVotosNegatTotales) And xVotosAfirmTotales > 0, "Si", "No")
    
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
           & "NroOrdenDia , Tipo, Origen, Destino, vota_presidente From Actas " _
           & " WHERE 1=0 "
    Call SetearRsW(strSql)
    
    RsWrite.AddNew
    RsWrite.Fields("Tipo_de_operación").Value = IIf(EstadoActual.TipoDeOperacion = "votnom" And xTipoVotacion = "votnum", "votnum", EstadoActual.TipoDeOperacion)
    RsWrite.Fields("Período_Legislativo").Value = EstadoActual.PeriodoLegislativo
    RsWrite.Fields("Sesión").Value = EstadoActual.Sesion
    RsWrite.Fields("Número_de_Acta").Value = EstadoActual.NroActa
    RsWrite.Fields("Versión_Acta").Value = 0
    RsWrite.Fields("Ultima_Versión_Acta").Value = 0
    RsWrite.Fields("Nombre_del_Acta").Value = EstadoActual.TituloDelActa
    RsWrite.Fields("Fecha").Value = EstadoActual.FechaVotacion
    RsWrite.Fields("Hora").Value = EstadoActual.HoraVotacion
    RsWrite.Fields("Tipo_de_Quorum").Value = EstadoActual.TipoMayoriaQuorum
    RsWrite.Fields("Base_de_Mayoria").Value = EstadoActual.BaseMayoria
    RsWrite.Fields("Tipo_de_Mayoria").Value = EstadoActual.TipoMayoria
    RsWrite.Fields("Miembros_del_cuerpo").Value = xMiembrosDelCuerpo
    RsWrite.Fields("Desempate").Value = strDesempate
    RsWrite.Fields("Votacion").Value = IIf(EstadoActual.TipoDeOperacion = "paslis", CartelActual.LeyendaQuorum, CartelActual.Resultado)
    RsWrite.Fields("Presidente").Value = EstadoActual.VectorIdentificacion(0)
    RsWrite.Fields("Presentes_Identificables").Value = xPresentesIdentificables
    RsWrite.Fields("Presentes_No_Identificables").Value = xPresentesNOIdentificables
    RsWrite.Fields("Presentes_Total").Value = xPresentesTotales
    RsWrite.Fields("Ausentes_Total").Value = xAusentesTotales
    RsWrite.Fields("Votos_Afirm_Identificables").Value = xVotosAfirmIdentificables
    RsWrite.Fields("Votos_Afirm_No_Identificables").Value = xVotosAfirmNOIdentificables
    RsWrite.Fields("Votos_Afirm_Desempate").Value = IIf(blHayDesempate And LCase(CartelActual.Resultado) = "afirmativo", 1, 0)
    RsWrite.Fields("Votos_Afirm_Total").Value = xVotosAfirmTotales
    RsWrite.Fields("Votos_Neg_Identificables").Value = xVotosNegatIdentificables
    RsWrite.Fields("Votos_Neg_No_Identificables").Value = xVotosNegatNOIdentificables
    RsWrite.Fields("Votos_Neg_Desempate").Value = IIf(blHayDesempate And LCase(CartelActual.Resultado) = "negativo", 1, 0)
    RsWrite.Fields("Votos_Neg_Total").Value = xVotosNegatTotales
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
    'cmdConfig.Enabled = Not blServerPrendido
    'If blServerPrendido = False Then
    '    cmdTerminar.Caption = "&Iniciar Server"
    '    lblFechaInicioServer.Caption = "SERVER DETENIDO"
    'Else
    '    cmdTerminar.Caption = "&Detener Server"
    '    lblFechaInicioServer.Caption = Now
    'End If
    Timer.Interval = 1000 / xIntervalo
    Timer.Enabled = True ' blServerPrendido ' Prendo o apago el server.. segun se requiera
Exit Sub
TrapError:
    Select Case err.Number
        Case 11
            xIntervalo = 2
             txtVecesPorSegundo.Text = ""
            Resume
        Case Else
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            Resume
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
    MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
    Resume
End Sub
Private Sub LeerEstadoRecinto()

    Dim strSql As String
    
    strSql = "SELECT * From vector"
    Call SetearRs(strSql)
    With CartelActual
        .Abstenciones = Rs.Fields("Abstenciones").Value
        .Afirmativos = Rs.Fields("Afirmativos").Value
        .Ausentes = GetNumero(Rs.Fields("Ausentes").Value)
        .LeyendaQuorum = Rs.Fields("Leyenda_Quorum").Value
        .MinimoVotosParaAfirmativo = Rs.Fields("Minimo_de_votos_para_afirmativa").Value
        .Negativos = Rs.Fields("Negativos").Value
        .Presentes = Rs.Fields("presentes").Value
        .Resultado = Rs.Fields("Resultado").Value
    End With
    
    With EstadoActual
        .ActaGrabada = GetNumero(Rs.Fields("Acta_Grabada"))
        .Ausentes = GetNumero(Rs.Fields("Ausentes").Value)
        .BaseMayoria = GetCadena(Rs.Fields("Base_de_Mayoría").Value)
        .CartelEncendido = GetNumero(Rs.Fields("Encender_Carteles").Value)
        .EstadoVotacion_y_PasList = GetCadena(Rs.Fields("Estado_de_votacion").Value)
        .GrabarAutomaticamente = GetNumero(Rs.Fields("Grabar_automaticamente").Value)
        .IdentificadorDeFormulario = GetCadena(Rs.Fields("Identificador_de_Formulario").Value)
        .IP_Consola = GetCadena(Rs.Fields("IP_Consola_Habilitada").Value)
        .ListarAutomaticamente = GetNumero(Rs.Fields("Listar_automaticamente").Value)
        .MensajeAlOperador = GetCadena(Rs.Fields("Mensaje_al_operador").Value)
        .TipoDeAbstencion = GetCadena(Rs.Fields("tipo_de_abstención").Value)
        .ModoMantenimientoBancas = GetNumero(Rs.Fields("Modo_Mantenimiento_Bancas").Value)
        .ModoNormalMantSistema = GetNumero(Rs.Fields("Modo_Normal_Mant_Sistema").Value)
        .NroActa = GetNumero(Rs.Fields("Nro_de_Acta").Value)
        .OcupadosNoIdentificados = GetNumero(Rs.Fields("Ocupadas_no_identificadas").Value)
        .PendientesEmitirVotos = GetNumero(Rs.Fields("Pendientes_Emitir_Voto").Value)
        .PeriodoLegislativo = GetCadena(Rs.Fields("Período_Legislativo").Value)
        .Presentes = GetNumero(Rs.Fields("Presentes").Value)
        .Sesion = GetNumero(Rs.Fields("Sesión").Value)
        .SolicitudGrabarManual = GetNumero(Rs.Fields("Solicitud_Grabacion_Manual").Value)
        .TiempoParaVotacion = GetNumero(Rs.Fields("Tiempo_de_votación").Value)
        .TipoDeOperacion = GetCadena(Rs.Fields("Identificador_tipo_de_operacion").Value)
        .TipoMayoria = GetCadena(Rs.Fields("Tipo_de_Mayoría").Value)
        .TipoMayoriaQuorum = GetCadena(Rs.Fields("Tipo_Mayoria_Quorum").Value)
        .TituloDelActa = GetCadena(Rs.Fields("Titulo_del_Acta").Value)
        .strError = GetCadena(Rs.Fields("strError").Value)
        .EstadoSesion = GetCadena(Rs.Fields("Estado_sesion").Value)
        .FechaVotacion = GetCadena(Rs.Fields("FechaVotacion").Value)
        .HoraVotacion = GetCadena(Rs.Fields("HoraVotacion").Value)
        If .TipoDeAbstencion = "" Then
                .TipoDeAbstencion = "votlar"
        End If
        If .BaseMayoria = "" Then
            .BaseMayoria = "legpre"
        End If
        If .TipoMayoria = "" Then
            .TipoMayoria = "120"
        End If
        If .TipoMayoriaQuorum = "" Then
            .TipoMayoriaQuorum = "120"
        End If
    End With
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
    & "Tipo_de_período_sesión, Fecha_de_comienzo, Tipo_de_Sesión , Nro_de_Sesion_actual, Histórico " _
    & "FROM perparl WHERE Rtrim(Período_Legislativo) = '" & Trim(EstadoActual.PeriodoLegislativo) & "' ORDER BY Orden"
    Call SetearRs(strSql)
    ' Si no existe, se selecciona el ultimo disponible
    If Rs.RecordCount <= 0 Or Rs.EOF = True Or Rs.BOF = True Then
        ' si no esta definida, selecciono la ultima disponible
        Rs.Close
        DoEvents
        strSql = "SELECT * FROM perparl ORDER BY Orden DESC"
        Call SetearRs(strSql)
        Rs.MoveFirst
    End If
    EstadoActual.PeriodoLegislativo = Rs.Fields("Período_Legislativo").Value
    'sesiones
    Rs.Close
    DoEvents
    strSql = "SELECT * FROM Sesion WHERE Rtrim(Período_Legislativo) = '" & Trim(EstadoActual.PeriodoLegislativo) & "' AND Sesión = " & EstadoActual.Sesion
    Call SetearRs(strSql)
    If Rs.RecordCount <= 0 Or Rs.EOF = True Or Rs.BOF = True Then
        strSql = "SELECT * FROM Sesion WHERE Rtrim(Período_Legislativo) = '" & Trim(EstadoActual.PeriodoLegislativo) & "' ORDER BY Sesión DESC"
        Call SetearRs(strSql)
        If Rs.RecordCount <= 0 Or Rs.EOF = True Or Rs.BOF = True Then
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
            Rs.MoveFirst
        End If
    End If
    EstadoActual.Sesion = Rs.Fields("Sesión").Value
    EstadoActual.NroActa = Rs.Fields("Próximo_Acta").Value
    Rs.Close
    ' En ambos casos: tomo el numero de "proximo acta" y lo pongo en estadoactual.nroacta
End Sub

Private Sub InicializarServer()
    Dim strSql As String
    lblVersion.Caption = strVersion  ' Mostrar versión de sqv
    Call AbrirDB                     ' establece conexion con la base de datos
    Call BorrarMensajesTotales       ' Borrar mensajes relacionados con la consola y el servidor de bancas
    Call DeterminarValoresInicioServer
    Call Levanta_Banca
    
    ' ------------------------------------------------------------------
    ' Interfaz de usuario del servidor
    ' ------------------------------------------------------------------
    blServerPrendido = True
    Call ServerOnOff
    ' ------------------------------------------------------------------
    ' Setear estado inicial del recinto
    ' ------------------------------------------------------------------
    Call LeerEstadoRecinto
    Call cargarColores
    Call ArmarBancasCartel
    With EstadoActual
        .Presentes = 0
        .Ausentes = xUltimaBanca + 1
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
        .Ausentes = xUltimaBanca + 1
        .Negativos = 0
        .Presentes = 0
        .LeyendaQuorum = "NO HAY QUORUM"
    End With
    Screen.MousePointer = vbIbeam
    'ShowCursor False
    ShowCursor True
    Call CargarVectorIdentificacionHabilitados
    Call CalcularMinimoParaQuorum
    Call AltaLogGeneral("SQV SERVER", "Inicializando SQV Server " & Now)
    Call SetearSesionActiva
    Call ReinicioSistema

End Sub


Private Sub Command2_Click()
    Dim MensajePrueba As MensajeSistema
    
    With MensajePrueba
        .sAtributo = "text"
        .sComentario = ""
        .sComponente = "term.display"
        .sObjeto = "brc"
        .sTipo = "xx"
        .sValor = "Inicio votacion"
    End With
    
    Call EnviarMensajesBancas(MensajePrueba)
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
           & "Negativos, Presidente, VotoPresidente, NuevoRes, NuevoMinAfirmativo From prueba_resultados_d"
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
            
            .Fields("NuevoRes").Value = CalculoResultado(.Fields("base_mayoria").Value, Str(.Fields("tipo_mayoria").Value), .Fields("miecue").Value, .Fields("presentes").Value, .Fields("Afirmativos").Value, .Fields("Negativos").Value, "", 0, 0, xpMin_p_af_Calc, .Fields("VotoPresidente").Value, .Fields("Presidente").Value)
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
        pBase_de_Mayoria = "legpre"
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
    
    xBase_para_Mayoria = IIf(pBase_de_Mayoria = "legpre", pPresentes, IIf(pBase_de_Mayoria = "miecue", pMiembros_del_cuerpo, IIf(pBase_de_Mayoria = "votemi", xVotosEmitidos, 0)))
    xResto = xBase_para_Mayoria * xNumerador Mod xDenominador
    pAuxMinParaAfirmativa = Fix(xBase_para_Mayoria * xNumerador / xDenominador)
    pMin_p_afirmativa_Calculo = IIf(xResto > 0, pAuxMinParaAfirmativa + 1, pAuxMinParaAfirmativa)
    pMin_p_afirmativa_Calculo = pMin_p_afirmativa_Calculo + IIf(LCase(RsOtro.Fields("Rdo_si_Af_igual_Min_y_Resto_mayor_0").Value) = "n", 1, 0)
    pMin_p_afirmativa_Calculo = pMin_p_afirmativa_Calculo + IIf(xResto = 0, LCase(RsOtro.Fields("SumarMinAfSiRestoIgual0").Value), 0)

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
End Function

Private Sub Form_Load()
    
    If App.PrevInstance = True Then ' Si se esta ejecutando una instancia previa del server, se apaga!
        End
    End If
    
    Set RsLocal = New ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rstActa = New ADODB.Recordset
    Set RsOtro = New ADODB.Recordset
    
    ' Indicar si se esta utilizando la base de pruebas o la de produccion
    If blBanderaPruebas = True Then
        Frame1.Caption = Frame1.Caption & " BASE PRUEBAS"
        Frame1.Caption = Frame2.Caption & " BASE PRUEBAS"
        Frame1.Caption = Frame3.Caption & " BASE PRUEBAS"
        lblPruebas.Visible = True
    Else
        lblPruebas.Visible = False
    End If
    
    
    xFechaArranque = Now
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
    Servidor.Show
    lblAppMayor.Caption = App.Major
    lblAppRevision.Caption = App.Revision
    lblAppMinor.Caption = App.Minor
    lblVersionSQV.Caption = "Merge 040225b:"
    'Call ProbarCalculoResultado
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
                lblOcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados
        End If
        If EstadoActual.CartelEncendido = 2 Then  'Frame General separar apagado Separar luego el 0 acaa acaa accaca
                'control de visualizacion, solo si hubo cambios de tipo o estado de operacion
                If xControlCartelTipoOperacion <> EstadoActual.TipoDeOperacion Or xControlCartelEstadoOperacion <> EstadoActual.EstadoVotacion_y_PasList Then
                    xControlCartelTipoOperacion = EstadoActual.TipoDeOperacion
                    xControlCartelEstadoOperacion = EstadoActual.EstadoVotacion_y_PasList
                    lblGeneralTituloDato.Visible = (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom")
                    lblGeneralOrdenDiaDato.Visible = False
                    lblGeneralTipoOperacionDato.Visible = (EstadoActual.TipoDeOperacion <> "quorum")
                    lblGeneralTiempo.Visible = (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom")
                    lblGeneralTiempoDato.Visible = (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom")
                    lblGeneralNegativos.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
                    lblGeneralNegativosDato.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
                    lblGeneralAfirmativos.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
                    lblGeneralAfirmativosDato.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
                    lblGeneralAbstenciones.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
                    lblGeneralAbstencionesDato.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
                    lblGeneralResultadoDato.Visible = ((EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate"))
                    lblGeneralMayoriaDato.Visible = (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom")
                End If
                    
                lblGeneralPresentesDato.Caption = Min(.Presentes, xMiembrosDelCuerpo)
                lblGeneralAusentesDato.Caption = max(.Ausentes, 0)
                lblGeneralHoraDato.Caption = Format(Now, "HH:MM")
                lblGeneralFechaDato.Caption = Format(Now, "DD/MM/YYYY")
                lblGeneralLeyendaQuorumDato.ForeColor = IIf(.LeyendaQuorum = "QUORUM", &HFFFFFF, &HC0C0FF)
                lblGeneralLeyendaQuorumDato.Caption = .LeyendaQuorum
                
                lblGeneralSesionDato.Caption = LeyendaSesion()
                lblGeneralTituloDato.Caption = EstadoActual.TituloDelActa
                lblGeneralOrdenDiaDato.Caption = ""
                lblGeneralTipoOperacionDato.Caption = LeyendaTipoOperacion
                lblGeneralMayoriaDato.Caption = "Base y tipo de mayoria: " & LeyendaTipoMayoria & " de los " & LeyendaBaseMayoria
                
                If (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") Then
                    lblGeneralTiempoDato.Caption = .LeyendaTiempo
                    lblGeneralNegativosDato.Caption = .Negativos
                    lblGeneralAfirmativosDato.Caption = .Afirmativos
                    lblGeneralAbstencionesDato.Caption = .Abstenciones
                    lblGeneralResultadoDato.Caption = .Resultado
                End If
                
                If xCiclosTotales Mod 10 = 0 Then
                    Call PintarBancasCartel
                End If
        End If
        If EstadoActual.CartelEncendido = 3 Then  'Actualizacion cartel mantenimiento
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
             lblMantenimientostrPanel3.Caption = Trim(strInfoMant(i)) & Space(14 - Len(Trim(strInfoMant(i)))) & Left(Trim(strInfoMant(i + 1)), 13)
             
             lblMantenimientostrPresencias.Caption = Trim(EstadoActual.MantPresencias)
             lblMantenimientostrId.Caption = Trim(EstadoActual.MantPresencias)
             lblMantenimientostrFallas.Caption = EstadoActual.MantCantFallas
        
             lblMantenimientostrPendientes.Caption = EstadoActual.MantCantPendientes
        
             lblMantenimientostrMantListaPendientes.Caption = EstadoActual.MantListaPendientes
             lblMantenimientostrMantListaFallas.Caption = EstadoActual.MantListaFallas
        End If
        
        If EstadoActual.CartelEncendido >= 1 Then  'Actualizacion cartel Serial. Separar luego el 0 acaa acaa accaca
                
                sCartel.strPresentes = Str(Min(.Presentes, xMiembrosDelCuerpo))
                sCartel.strAusentes = Str(max(.Ausentes, 0))
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
        If EstadoActual.CartelEncendido = 0 Then ' renombrar por caso 0 aca acaaa caccac
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
    Dim strTempCadena               As String
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
    Dim nNuevosAbstenidos As Long
    Dim nNuevosCancelados As Long
    Dim nTotalAbstenciones As Long
    Dim vTemp As Variant
    
     
    ' Atender a todos los mensajes nuevos Emitidos por las consolas
   strSql = "SELECT * FROM consola_sqv_mensajes WHERE serial > " & Str(xUltimoMensajeCosola)
   Call SetearRs(strSql)
   'Call SetearRsCadena(xUltimoMensajeCosola)
    With Rs
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
                            Call EnviarMensajesComienzoAuth(MensajeActual.sObjeto, "", "key_start")
                        Else
                            If EstadoActual.VectorPresencia(MensajeActual.sObjeto) = PRESENTE Then
                                Call EnviarMensajesComienzoAuth(MensajeActual.sObjeto, "", "key_start")
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
                            FrameSQVGeneral.ZOrder 0
                        ElseIf .sComponente = "1" Then
                            EstadoActual.CartelEncendido = 1
                            FrameControl.ZOrder 0
                        ElseIf .sComponente = "3" Then 'Mantenimiento
                            EstadoActual.CartelEncendido = 3
                            FrameMantenimiento.ZOrder 0
                        ElseIf .sComponente = "4" Then
                            EstadoActual.CartelEncendido = 4
                            FrameSQVActa.ZOrder 0
                        Else '.sComponente = "n" Or .sComponente = "0"
                            EstadoActual.CartelEncendido = 0
                            FrameSQVApagado.ZOrder 0
                            txtTipoOperacion = "d"
                        End If
                        ' ---------------------------------------------------------------------------------
                        ' Usuario inicia prueba scan de una banca
                        ' ---------------------------------------------------------------------------------
                    Case Is = "pruebascan"
                        If EstadoActual.TipoDeOperacion = "quorum" Then
                            If xBancaPruebaScan > 0 Then 'Ya estaba otra banca en prueba
                                Call EnviarMensajesFinAuth(Str(xBancaPruebaScan), "Prueba Scan Fin por nuevo pedido de prueba scan")
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
                            End If
                        End If
                        ' ---------------------------------------------------------------------------------
                        ' Usuario cambia el tipo de operacion
                        ' ---------------------------------------------------------------------------------
                    Case Is = "cambio?tipoop"
                        EstadoActual.strError = "cambio?tipoop"
                        If Not IsNull(.sComponente) Then
                            If (InStr("votnom;votnum", Trim(EstadoActual.TipoDeOperacion)) > 0) Then
                                For i = 0 To Min(xUltimaBanca, UBound(EstadoActual.VectorAbstencion))
                                    EstadoActual.VectorAbstencion(i) = 0
                                Next
                                Call AbstenerVector(Join(EstadoActual.VectorAbstencion, SEPARADOR_VECTOR), 0, 0, 0)
                            End If
                            '>> IMPORTANTE: si pasa de un modo de no identifiacion a uno que permite la identificacion, debe habilitar a identificarse a todos los presentes
                            'cambio de no identificacion a identificacion  de (quorum o votacion numerica ) a (vnominal o pase de lista)
                            If (EstadoActual.Modo_Ident_Nom_Obsoleto = 1) Then
                                Call SolicitarIdentificacionPendientes("Comienzo modo nominal desde " & EstadoActual.TipoDeOperacion & "a " & .sComponente, "start")
                                EstadoActual.OcupadosNoIdentificados = CuentaOcupadosNoIdentificadosCong  'EstadoActual.Presentes - IIf(xPresidenteLegislador, 1, 0) 'aca5
                            ElseIf (Trim(EstadoActual.TipoDeOperacion) = "votnum" Or (Trim(EstadoActual.TipoDeOperacion) = "quorum" And Not (EstadoActual.Modo_Ident_Nom_Obsoleto = 1))) And InStr("votnom;paslis", Trim(.sComponente)) > 0 Then
                                'antes nominal If (InStr("quorum;votnum", Trim(EstadoActual.TipoDeOperacion)) > 0) And InStr("votnom;paslis", Trim(.sComponente)) > 0 Then
                                'enviar broadcast identificarse Msj mset/term.auth?ACTION=AUTH_START
                                '>> el siguiente programa envia a TODOS los que en el vector presencia tengan un 1 el comando de identificarse.
                                xStrVector = "0" & SEPARADOR_VECTOR 'presidente
                                For i = 1 To UBound(EstadoActual.VectorPresencia)
                                    xStrVector = xStrVector & IIf(EstadoActual.VectorPresencia(i) = "1", "1", "0") & SEPARADOR_VECTOR
                                Next i
                                Call EnviarMensajesComienzoAuth(xStrVector, "Comienzo modo nominal desde " & EstadoActual.TipoDeOperacion & "a " & .sComponente)
                                For i = 1 To (xUltimaBanca)
                                    EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO
                                    EstadoActual.VectorColor(i) = AsignarColor(i)
                                Next i
                                EstadoActual.OcupadosNoIdentificados = EstadoActual.Presentes - IIf(xPresidenteLegislador, 1, 0) 'aca5
                            '>> A la inversa, cancela las identificaciones y la posibilidad de identificarse a todos los presentes.
                            ElseIf (InStr("votnom;paslis", Trim(EstadoActual.TipoDeOperacion)) > 0) And _
                                 (Trim(.sComponente) = "votnum" Or (Trim(.sComponente) = "quorum" And Not (EstadoActual.Modo_Ident_Nom_Obsoleto = 1))) Then
                                'antes nominal If (InStr("votnom;paslis", Trim(EstadoActual.TipoDeOperacion)) > 0) And InStr("quorum;votnum", Trim(.sComponente)) > 0 Then
                                xStrVector = "0" & SEPARADOR_VECTOR 'presidente
                                For i = 1 To UBound(EstadoActual.VectorPresencia)
                                    xStrVector = xStrVector & IIf(EstadoActual.VectorPresencia(i) = "1", "1", "0") & SEPARADOR_VECTOR
                                Next i
                                Call EnviarMensajesFinAuth(xStrVector, "Fin modo nominal desde " & EstadoActual.TipoDeOperacion & "a " & .sComponente)
                                For i = 1 To (xUltimaBanca)
                                    EstadoActual.VectorIdentificacion(i) = NO_IDENTIFICADO
                                    EstadoActual.VectorColor(i) = AsignarColor(i)
                                Next i
                                EstadoActual.OcupadosNoIdentificados = 0
                            End If
                            If EstadoActual.Modo_Ident_Nom_Obsoleto = 1 And (Trim(.sComponente) = "votnum") Then
                                'Modo nominal, votacion numerica. Se trata como nominal pero se guarda el tipo de operacion en el auxiliar
                                EstadoActual.TipoDeOperacion = "votnom"
                                xTipoVotacion = "votnum"
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
                        ' Buscar en la base si es legislador
                        strSql = "SELECT Es_Legislador FROM Legisladores WHERE id = '" & Trim(.sComponente) & "'"
                        rsTemp.CursorLocation = adUseClient
                        rsTemp.Open strSql, Cn, adOpenForwardOnly, adLockReadOnly
                        If rsTemp.RecordCount > 0 Or (rsTemp.EOF = True Or rsTemp.BOF = True) Then
                            If rsTemp("Es_Legislador").Value = 0 Then
                                xPresidenteLegislador = False
                            Else
                                xPresidenteLegislador = True
                            End If
                            If xPresidenteLegislador = True Then
                                If EstadoActual.VectorPresencia(0) = AUSENTE Then
                                    EstadoActual.VectorPresencia(0) = PRESENTE
                                    EstadoActual.Presentes = EstadoActual.Presentes + 1
                                    EstadoActual.Ausentes = EstadoActual.Ausentes - 1
                                End If
                            Else
                                If EstadoActual.VectorPresencia(0) = PRESENTE Then
                                    EstadoActual.VectorPresencia(0) = AUSENTE
                                    EstadoActual.Presentes = EstadoActual.Presentes - 1
                                    EstadoActual.Ausentes = EstadoActual.Ausentes + 1
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
                        EstadoActual.strError = "cambio?basemayoria"
                    End If
                    ' ---------------------------------------------------------------------------------
                    ' Usuario cambia el tipo de la mayoría para la votación
                    ' ---------------------------------------------------------------------------------
                Case Is = "cambio?tipomayoriavotacion"
                    If Not IsNull(.sComponente) Then
                        EstadoActual.TipoMayoria = .sComponente
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
                    ' Usuario inicia una votación
                    ' ---------------------------------------------------------------------------------
                Case Is = "accion?iniciovotacion"
                    If EstadoActual.EstadoVotacion_y_PasList = "espera" And (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum") Then
                        EstadoActual.PendientesEmitirVotos = EstadoActual.Presentes - EstadoActual.AbstencionistasAutorizados '- IIf(xPresidenteLegislador, 0, 1)
                        ' Pasar el vector presencia a un string
                        xMax = UBound(EstadoActual.VectorPresencia)
                        'strTempCadena = IIf(xPresidenteLegislador, "1", "0") & SEPARADOR_VECTOR
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
                        
                        strTempCadena = ("0") & SEPARADOR_VECTOR
                        For X = 1 To xMax
                            If EstadoActual.VectorPresencia(X) = PRESENTE Then
                                If (EstadoActual.TipoDeOperacion = "votnum" _
                                    Or Not (EstadoActual.VectorIdentificacion(X) = NO_IDENTIFICADO)) _
                                    And EstadoActual.VectorResultados(X) <> ABSTENCION_AUTORIZADA Then
                                    strTempCadena = strTempCadena & PRESENTE & SEPARADOR_VECTOR
                                Else
                                    strTempCadena = strTempCadena & AUSENTE & SEPARADOR_VECTOR
                                End If
                            Else
                                    strTempCadena = strTempCadena & AUSENTE & SEPARADOR_VECTOR
                            End If
                        Next X
                        With Mensaje2Banca ' Mensaje para SB
                            .sTipo = "mset"
                            .sComponente = "term.keyb"
                            .sObjeto = strTempCadena
                            .sAtributo = "state"
                            .sValor = "on" & EstadoActual.TipoDeOperacion
                        End With
                        Call EnviarMensajesBancas(Mensaje2Banca)
                        EstadoActual.EstadoVotacion_y_PasList = "votando"
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
                    If EstadoActual.EstadoVotacion_y_PasList = "larga" Then
                        EstadoActual.EstadoVotacion_y_PasList = "cierre"
                        Call FinVotacionBrc("cierre operador")
                    End If
                    
                    ' ------------------------------------------------------------------------------------
                    ' Usuario cancela una votaciónUsuario cancela una votación
                    ' ------------------------------------------------------------------------------------
                Case Is = "accion?cancelavotacion"
                    With EstadoActual
                        If .EstadoVotacion_y_PasList = "larga" Or .EstadoVotacion_y_PasList = "empate" Or .EstadoVotacion_y_PasList = "votando" Then
                            ' Apagar todos los teclados de las bancas
                            Mensaje2Banca.sTipo = "mset"
                            If EstadoActual.EstadoVotacion_y_PasList = "larga" Then
                                EstadoActual.EstadoVotacion_y_PasList = "cierre"
                                Call FinVotacionBrc("cancela operador")
                            End If
                            .EstadoVotacion_y_PasList = "cancelada"
                            Call AltaLogGeneral("Operador del sistema", "Cancelacion votacion. PL: " & EstadoActual.PeriodoLegislativo & " S " & EstadoActual.Sesion & " A " & EstadoActual.NroActa & " Estado " & EstadoActual.EstadoVotacion_y_PasList, , "1")
                        End If
                    End With
                    ' ------------------------------------------------------------------------
                    ' Usuario cambia el titulo de un acta
                    ' ------------------------------------------------------------------------
                Case Is = "cambio?tacta"
                    If Not IsNull(.sObjeto) Then ' .sComponente tiene el Id del titulo del acta
                        EstadoActual.strError = "cambio?tacta"
                        EstadoActual.TituloDelActa = Rs.Fields("Parametro2").Value ' titulo del acta: Lo toma directamente de record set para mantener mayusculas.
                    End If
                    ' -------------------------------------------------------------------------------------
                    ' Usuario cambia el periodo legislativo
                    ' -------------------------------------------------------------------------------------
                Case Is = "cambioperiodo"
                    If Not IsNull(.sComponente) Then
                        EstadoActual.strError = "cambioperiodo"
                        EstadoActual.PeriodoLegislativo = .sComponente
                        ' Buscar en Periodo Legislativo la ultima sesion
                        strSql = "SELECT Nro_de_Sesion_actual From dbo.perparl Where Período_Legislativo = '" & EstadoActual.PeriodoLegislativo & "'"
                        Call SetearOtroRs(strSql)
                        EstadoActual.Sesion = RsOtro.Fields(0).Value
                        RsOtro.Close
                        ' Buscar en Sesion ultima acta
                        strSql = "SELECT Próximo_Acta, Estado_sesión From Sesion Where Sesión = '" & EstadoActual.Sesion & "' And Período_Legislativo = '" & EstadoActual.PeriodoLegislativo & "'"
                        Call SetearOtroRs(strSql)
                        If RsOtro.RecordCount = 0 Or Rs.EOF = True Or Rs.BOF = True Then
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
                    If EstadoActual.TipoDeOperacion = "paslis" And EstadoActual.EstadoVotacion_y_PasList = "espera" Then 'EstadoActual.EstadoVotacion_y_PasList <> "votando" And EstadoActual.EstadoVotacion_y_PasList <> "larga" And EstadoActual.EstadoVotacion_y_PasList <> "cierre" Then
                        Call AltaLogGeneral("Operador del sistema", "Pase de lista iniciado: " & Trim(EstadoActual.PeriodoLegislativo) & " sesion" & Trim(Str(EstadoActual.Sesion)) & " acta" & Trim(Str(EstadoActual.NroActa)))
                        EstadoActual.EstadoVotacion_y_PasList = "finalizada"
                        EstadoActual.OcupadosNoIdentificados = CuentaOcupadosNoIdentificadosCong
                        EstadoActual.ActaGrabada = EstadoActual.NroActa
                        EstadoActual.SolicitudGrabarManual = 0
                        EstadoActual.PresentesCongelados = Presentes() ' EstadoActual.Presentes
                        EstadoActual.AusentesCongelados = Ausentes() 'EstadoActual.Ausentes
                        CartelActual.Resultado = "FINALIZADO"
                        EstadoActual.FechaVotacion = Now
                        Call AlmacenarActa
                        'MsgBox "Falta llamar a Graba Pase de Lista" & EstadoActual.PresentesCongelados & EstadoActual.AusentesCongelados
                    End If
                    ' Usuario actualiza los datos de las bancas
                Case Is = "accion?recargardatossb"
                    EstadoActual.strError = "accion?recargardatossb"
                    MsgBox "actualizar datos de banca"
                    ' Usuario habilita una consola
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
                            .sObjeto = Str(xBanca)
                            .sAtributo = "action"
                            .sValor = "reset"
                        End With
                        Call EnviarMensajesBancas(Mensaje2Banca)
                    Else
                        EstadoActual.MensajeAlOperador = "Banca invalida"
                    End If
                    'MsgBox "Resetear SO de la banca " & Str(xActualBanca)
                    ' Usuario solicita al Sb se apague una banca y que no se le envíen más mensajes
                
                Case Is = "resethard"
                    EstadoActual.strError = "resethard"
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
                            .sObjeto = Str(xBanca)
                            .sAtributo = "action"
                            .sValor = "resethard"
                        End With
                        Call EnviarMensajesBancas(Mensaje2Banca)
                    Else
                        EstadoActual.MensajeAlOperador = "Banca invalida"
                    End If
                Case Is = "apagarbanca"
                    EstadoActual.strError = "apagarbanca"
                    If Not IsNull(.sComponente) Then
                        xActualBanca = Int(.sComponente)
                        'MsgBox "Apagar banca " & Str(xActualBanca)
                        MsgBox xActualBanca
                        Call PintarVectorColor(xActualBanca)
                    End If
                    ' Usuario solicita se limpie la cola de mensajes sqv->consola
                Case Is = "limpia"
                        MsgBox "Definir regla de negocio"
                    ' Usuario decide realizar una grabación manual No implementado
                Case Is = "accion?grabarmanual"
                        MsgBox "Usuario decide hacer grabacion manual"
                    ' Usuario decide pasar al modo de mantenimiento
                Case Is = "cambio?mantenimiento"
                    EstadoActual.strError = "cambio?mantenimiento"
                    If EstadoActual.ModoMantenimientoBancas = 0 Then
                        Call Mantenimiento_SQV
                    Else
                        Call Fin_Mantenimiento_SQV
                    End If
                    'MsgBox "Pasar a modo mantenimiento"
                    ' Usuario decide pasar al modo normal con mantenimiento
                    ' En este modo se puede operar el sistema en forma normal pero se valida la identificacion con registros de personas definidas como Personal de Mantenimiento (tipo = 0)
                Case Is = "cambio?nominal"
                    EstadoActual.strError = "cambio?nominal"
                    If EstadoActual.Modo_Ident_Nom_Obsoleto = 0 Then
                        EstadoActual.Modo_Ident_Nom_Obsoleto = 1
                        Call Fin_Mantenimiento_SQV
                    Else
                        EstadoActual.Modo_Ident_Nom_Obsoleto = 0
                    End If
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
                    MsgBox "Usuario cambia el Modo de presentacion de informacion fija"
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
                    MsgBox "Usuario reinicia SQV en modo mantenimiento"
                    ' Usuario sale del SQV
                Case Is = "accion?salirserver"
                    EstadoActual.strError = "accion?salirserver"
                    Call AltaLogGeneral("Saliendo de SQV Server por solicitud Consola", Now)
                    Call Salir_SQV
                    ' Usuario reinicia el server de quorum inmediatamente
                Case Is = "accion?inicioconsola"
                    MsgBox "Usuario reinicia el server de quorum inmediatamente"
                    ' Usuario consulta un acta grabada
                Case Is = "mostrar?periodo"
                    EstadoActual.strError = "mostrar?periodo"
                    MsgBox "Usuario consulta un acta grabada"
                    ' Usuario cambia la modalidad de grabación automática
                Case Is = "cambio?grabar"
                    MsgBox "Usuario cambia la modalidad de grabación automática"
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
                    If (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "cancelada") Then
                        '>> REINCIO DE VOTACION. CUANDO EL OPERADOR PRESIONA EL BOTON INICIALIZAR, SE PREPARA EL SISTEMA PARA UNA NUEVA VOTACION.
                        '>> en el caso de que la ultima haya sido una votacion de reconsideracion (nominal), a todos los presentes no identificados les permitira identificars enuevamente
                        If EstadoActual.TipoDeOperacion = "votnom" Then
                            xVotacionReconsideracion = False
                            For i = 0 To xUltimaBanca
                                If EstadoActual.VectorIdentificacionHabilitados(i) = NO_IDENTIFICADO Then
                                    xVotacionReconsideracion = True
                                End If
                            Next i
                            If xVotacionReconsideracion Then
                                xStrVector = ""
                                For i = 0 To xUltimaBanca
                                    If EstadoActual.VectorPresencia(i) = PRESENTE And Val(EstadoActual.VectorIdentificacion(i)) = NO_IDENTIFICADO Then
                                        xStrVector = xStrVector & "1" & SEPARADOR_VECTOR
                                    Else
                                        xStrVector = xStrVector & "0" & SEPARADOR_VECTOR
                                    End If
                                Next i
                                Call EnviarMensajesComienzoAuth(xStrVector, "Permitir identificacion tras reconsideracion")
                            End If
                        End If 'fin si es reconsideracion
                        If (InStr("votnom;votnum", Trim(EstadoActual.TipoDeOperacion)) > 0) Then
                            Call InicializarVotacion
                            'Call AbstenerVector(SEPARADOR_VECTOR, 0, 0, 0)
                            Call AbstenerVector(Join(EstadoActual.VectorAbstencion, SEPARADOR_VECTOR), 0, 0, 0)
                        End If
                        
                        EstadoActual.LimpiarResultados = 1
                        Call PintarTodasLasBancas
                        EstadoActual.ActaGrabada = 0
                        
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
                    Else
                        xVotoOperador = "x"
                    End If
                    If (xBanca >= 0 And InStr("s n", xVotoOperador) > 0) And _
                        (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum") And _
                         (EstadoActual.EstadoVotacion_y_PasList = "votando" Or EstadoActual.EstadoVotacion_y_PasList = "larga" Or EstadoActual.EstadoVotacion_y_PasList = "empate") Then
                        'Filtro de banca 0 y mayor a leg
                        If xBanca >= IIf(xPresidenteLegislador Or EstadoActual.EstadoVotacion_y_PasList = "empate", 0, 1) And xBanca <= xUltimaBanca Then
                            'Si es nominal, ver que este identificado, sino solo que esté presente
                            'Si es votacion larga, tambien se puede habiiltar para votar
                            If EstadoActual.VectorPresencia(xBanca) = PRESENTE And _
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
                                        ElseIf xVotoOperador = ABSTENCION Then
                                            EstadoActual.VectorResultados(xBanca) = ABSTENCION
                                            Mensaje2Banca.sTipo = "mset"
                                            Mensaje2Banca.sObjeto = xBanca
                                            Mensaje2Banca.sComponente = "term.ledk1"
                                            Mensaje2Banca.sAtributo = "state"
                                            Mensaje2Banca.sValor = "off"
                                            Mensaje2Banca.sComentario = EstadoActual.EstadoVotacion_y_PasList
                                            Call EnviarMensajesBancas(Mensaje2Banca)
                                            EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos + 1
                                        End If 'MensajeActual.sComponente = "term.keyb.si o no
                                        PintarVectorColor (xBanca)
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
                                        EstadoActual.MensajeAlOperador = "Cambio ID por Operador. Ya esta identificado." & xNuevoID & " Banca: " & Str(X)
                                        EstadoActual.strError = "**error"
                                        'xNuevoID = ""
                                    End If
                                Next X
                                ' Si no se identifico anteriormente en otra banca, pongo ID de legislador en Vector Identificacion
                                If flIdDupOperador = False Then ' identificar al legislador en vector identificacion
                                    'Verifica que no coincida con el presidente
                                    If Trim(LCase(EstadoActual.VectorIdentificacion(0))) = xNuevoID Then
                                        flIdDupOperador = True
                                        EstadoActual.MensajeAlOperador = "Cambio ID por Operador. Identificado como Presidente." & xNuevoID
                                        EstadoActual.strError = "**error"
                                    Else
                                        '>> Verifica que este habilitado para el caso de las votaciones de reconsideracion.
                                        ' En las otras situaciones siempre estan todos habilitados.
                                        If (LegisladorHabilitado(xNuevoID)) Then
                                            'La identificacion ha sido exitosa!
                                            EstadoActual.VectorIdentificacion(xBanca) = xNuevoID
                                            EstadoActual.OcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados - 1
                                            
                                            Call EnviarMensajesFinAuth(Str(xBanca), "Autenticacion Operador")
                                            
                                            Mensaje2Banca.sObjeto = xBanca '<AP 040115 faltaba indicar la banca>
                                            Mensaje2Banca.sTipo = "mset"
                                            If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
                                                Mensaje2Banca.sComponente = "term.led1"
                                                Mensaje2Banca.sAtributo = "state"
                                                Mensaje2Banca.sValor = "on"
                                                Mensaje2Banca.sComentario = "Id aceptado Modo normal"
                                            Else
                                                Mensaje2Banca.sValor = "Identificacion de Prueba"
                                                Mensaje2Banca.sComentario = "Id aceptado Modo mantenimiento"
                                            End If
                                            Call EnviarMensajesBancas(Mensaje2Banca)
                                            Mensaje2Banca.sObjeto = xBanca '<AP 040115 faltaba indicar la banca>
                                            Mensaje2Banca.sTipo = "mset"
                                            If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
                                                Mensaje2Banca.sComponente = "term.display"
                                                Mensaje2Banca.sAtributo = "text"
                                                Mensaje2Banca.sValor = "Identificacion Aceptada"
                                                Mensaje2Banca.sComentario = "Id aceptado Modo normal"
                                            Else
                                                Mensaje2Banca.sValor = "Identificacion de Prueba"
                                                Mensaje2Banca.sComentario = "Id aceptado Modo mantenimiento"
                                            End If
                                            Call EnviarMensajesBancas(Mensaje2Banca)
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
                                EstadoActual.MensajeAlOperador = "Banca invalida: " & xBanca
                            End If
                        Else
                            EstadoActual.strError = "**error"
                            EstadoActual.MensajeAlOperador = "Solicitud de abstencion fuera de votacion nominal a iniciar o en curso"
                        End If
                        lblPendientesEmitirVotos = EstadoActual.PendientesEmitirVotos
                        lblAbsAut.Caption = EstadoActual.AbstencionistasAutorizados
                        lblOcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados
                    End If
                    ' Usuario consulta DB utilizada
                Case Is = "configuradoaccion?consultasufijo"
                    MsgBox "Usuario consulta DB utilizada"
                    ' Usuario prende luces
                Case Is = "accion?prenderluces"
                    MsgBox "Usuario prende luces"
                    ' Usuario Apaga luces
                Case Is = "accion?apagarluces"
                    MsgBox "Usuario Apaga luces"
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
    xUltimoMensajeCosola = Rs.Fields("serial").Value
    .MoveNext
    Wend
End With
    
    ' ------------------------------------------------------------------------------------------
    ' Atender a todos los mensajes nuevos emitidos por el servidor de banca
    ' ------------------------------------------------------------------------------------------
    strSql = "SELECT * FROM sb_sqv_mensajes WHERE id > " & Str(xUltimoMensajeSB)
    Call SetearRs(strSql)
    With Rs
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
                If xBanca >= 1 Or xBanca <= xUltimaBanca Then
                    If .sComponente = "term.seat" Then
                        If .sAtributo = "switch" Then
                            If .sValor = "closed" Then
                                ' Preguntar en el vector de presencia si la banca esta ocupada
                                If EstadoActual.VectorPresencia(xBanca) <> PRESENTE Then
                                    EstadoActual.VectorPresencia(xBanca) = PRESENTE
                                    
                                    ' para abstencion numerica
                                    If EstadoActual.TipoDeOperacion = "votnum" Then
                                      'MODIFICACION AP 040921
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
                                    EstadoActual.Presentes = EstadoActual.Presentes + 1
                                    EstadoActual.Ausentes = EstadoActual.Ausentes - 1
                                    'CartelActual.Presentes = CartelActual.Presentes + 1
                                    flSwitchExitoso = True
                                End If
                            ElseIf .sValor = "open" Then
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
                Call PintarVectorColor(xBanca)
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
                If (EstadoActual.TipoDeOperacion = "votnom" And Not (.sComponente = "term.keyb.si" Or .sComponente = "term.keyb.no")) Or EstadoActual.TipoDeOperacion = "paslis" Or (EstadoActual.TipoDeOperacion = "quorum" And EstadoActual.Modo_Ident_Nom_Obsoleto = 1) Then
                   'antes nominal  If (EstadoActual.TipoDeOperacion = "votnom" And Not (.sComponente = "term.keyb.si" Or .sComponente = "term.keyb.no")) Or EstadoActual.TipoDeOperacion = "paslis" Then
                   Call Identificacion(MensajeActual)
                End If
                If EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum" Then
                   Call Votacion(MensajeActual)
                End If
                If (.sComponente = "term.mon" Or .sComponente = "term" Or .sComponente = "term.ioc") And LCase(.sAtributo) = "state" Then
                    Call ManejoDeFallas(MensajeActual)
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
        Wend
        'ver si se valida el fin de tiempo de votacion.
        'VERIFICACION DE TIEMPO CUMPLIDO DE VOTACION solo en estado votando sea nominal o numerica
        Call ControlTiempoCumplidoVotacion
        'Operaciones en caso de cierre de votacion
        Call CierreVotacion
        'Actualizacion de cartel y quorum
        xTiempoVotacionTranscurrido = DateDiff("s", EstadoActual.FechaVotacion, Now)
        xTiempoRestanteVotacion = EstadoActual.TiempoParaVotacion - xTiempoVotacionTranscurrido
        CartelActual.LeyendaTiempo = IIf(EstadoActual.EstadoVotacion_y_PasList = "espera", "", IIf(EstadoActual.EstadoVotacion_y_PasList = "votando", IIf(xTiempoRestanteVotacion > EstadoActual.TiempoParaVotacion, "", IIf(xTiempoRestanteVotacion > 59, Str(xTiempoRestanteVotacion), IIf(xTiempoVotacionTranscurrido > EstadoActual.TiempoParaVotacion, " FIN", Right(Str(xTiempoRestanteVotacion), 2)))), IIf(EstadoActual.EstadoVotacion_y_PasList = "larga", " 0", IIf(EstadoActual.EstadoVotacion_y_PasList = "cancelada", "CANCELADA", "FIN"))))
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
    End With
    ' Rs.Close
End Function
Private Sub CalcularMinimoParaQuorum()
    ' variable que permite calcular el minimo necesario para obtener quorum
    xMinimoParaQuorum = IIf(LCase(EstadoActual.TipoMayoriaQuorum) = "man", 1, Fix(xMiembrosDelCuerpo / 2) + IIf(EstadoActual.TipoMayoriaQuorum = "121", 1, 1))
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
        MsgBox CartelActual.LeyendaQuorum & " " & xMinimoParaQuorum & " " & xMinimoParaQuorumEntero & " " & CartelActual.Presentes
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
                                EstadoActual.MensajeAlOperador = "Banca " & Trim(Str(xActualBanca)) & " en prueba de Scan = 'negative'"
                            ElseIf LCase(.sValor) <> "negative" Then ' Busca mValor en base de datos legisladores
                                strSql = "SELECT * FROM legisladores WHERE id = " & Trim(.sValor) & " AND Tipo = 1"
                                Call SetearOtroRs(strSql)
                                If RsOtro.RecordCount = 0 Or RsOtro.EOF = True Then ' Si no lo encuentra entre los legisladores, lo busca entre el personal de mantenimiento
                                    strSql = "SELECT * FROM legisladores WHERE id = " & Trim(.sValor) & " AND Tipo = 0"
                                    Call SetearOtroRs(strSql)
                                    If RsOtro.RecordCount > 0 Then
                                        ' Lo encontro entre la gente de mantenimiento: ackowledge al operador
                                        EstadoActual.strError = "pruebascan"
                                        EstadoActual.MensajeAlOperador = "Prueba Scan Invalida: " & Trim(.sValor) & " es el valor recibido, no se encuentra en base de datos"
                                        RsOtro.Close
                                    Else
                                        EstadoActual.strError = "pruebascan"
                                        EstadoActual.MensajeAlOperador = RsOtro.Fields("Apellido").Value & " " & RsOtro.Fields("Nombre").Value & " identificado Ok"
                                    End If
                                Else
                                    ' Lo encontro como legislador: ackowledge al operador
                                    EstadoActual.strError = "pruebascan"
                                    'EstadoActual.strError = "**error"
                                    EstadoActual.MensajeAlOperador = RsOtro.Fields("Apellido").Value & " " & RsOtro.Fields("Nombre").Value & " identificado Ok"
                                    ' acknowledge al Legislador
                                    Mensaje2Banca.sObjeto = xActualBanca '<AP 040115 faltaba indicar la banca>
                                    Mensaje2Banca.sTipo = "mset"
                                    Mensaje2Banca.sComponente = "term.display"
                                    Mensaje2Banca.sAtributo = "text"
                                    Mensaje2Banca.sValor = "Prueba Identificacion Valida"
                                    Mensaje2Banca.sComentario = "Prueba Scan Id Aceptado Modo normal"
                                    Call EnviarMensajesBancas(Mensaje2Banca)
                                End If ' Fin si no esta entre personal de mantenimiento
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
                    bActualizar = True
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
    If bActualizar Then
        EstadoActual.MantCantFallas = 0
        EstadoActual.MantCantPendientes = 0
        EstadoActual.MantListaFallas = " "
        EstadoActual.MantListaPendientes = " "
        For i = 0 To xUltimaBanca
            If EstadoActual.VMantEstado(i) = ABSTENCION Then
                EstadoActual.MantCantPendientes = EstadoActual.MantCantPendientes + 1
                EstadoActual.MantListaPendientes = Trim(EstadoActual.MantListaPendientes) & "," & Trim(Str(i))
            ElseIf Not (EstadoActual.VMantEstado(i) = AFIRMATIVO) And Not (EstadoActual.VectorPresencia(i) = PRESENTE) Then
                    EstadoActual.MantCantFallas = EstadoActual.MantCantFallas + 1
                    EstadoActual.MantListaFallas = Trim(EstadoActual.MantListaFallas) & "," & Trim(Str(i))
            End If
        Next i
    End If
End Sub
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
    With MensajeBanca
        xBanca = Int(.sObjeto)
        If LCase(.sTipo) = "mevt" And LCase(.sComponente) = "term" And LCase(.sAtributo) = "state" Then
            If LCase(.sValor) = "ok" Then
                If EstadoActual.VectorPresencia(xBanca) = BANCA_INHABILITADA Then
                    EstadoActual.VectorPresencia(xBanca) = AUSENTE
                    PintarVectorColor (xBanca)
                    Call AltaLogGeneral("sqv", "Habilitacion banca:" & Str(xBanca) & " Estado:" & .sValor, Str(xBanca))
                End If
            End If
            If LCase(.sValor) = "off" Then
                If EstadoActual.VectorPresencia(xBanca) = PRESENTE Then
                    EstadoActual.Presentes = EstadoActual.Presentes - 1
                    EstadoActual.Ausentes = EstadoActual.Ausentes + 1
                    If (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "paslis") Or EstadoActual.Modo_Ident_Nom_Obsoleto = 1 Then  'ooooooiii
                        If (EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO) Then
                            EstadoActual.OcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados - 1
                        End If
                    End If
                    EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO
                    If (EstadoActual.TipoDeOperacion = "votnum" Or EstadoActual.TipoDeOperacion = "votnom") And EstadoActual.EstadoVotacion_y_PasList <> "finalizada" And EstadoActual.EstadoVotacion_y_PasList <> "cierre" Then
                        strRes = LCase(EstadoActual.VectorResultados(xBanca))
                        If strRes = AFIRMATIVO Or strRes = NEGATIVO Then
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
                        End If
                        EstadoActual.VectorResultados(xBanca) = ABSTENCION
                    End If
                    
                End If
                EstadoActual.VectorPresencia(xBanca) = BANCA_INHABILITADA
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
        If xBanca >= IIf(xPresidenteLegislador, 0, 1) And xBanca <= xUltimaBanca Then
            'Si es nominal, ver que este identificado, sino solo que esté presente
            'Si es votacion larga, tambien se puede habiiltar para votar
            If EstadoActual.VectorPresencia(xBanca) = PRESENTE And _
                    (EstadoActual.TipoDeOperacion = "votnum" Or _
                     (EstadoActual.TipoDeOperacion = "votnom" And (EstadoActual.VectorIdentificacion(xBanca) <> NO_IDENTIFICADO Or (EstadoActual.ModoMantenimientoBancas And xBanca = 0))) _
                    ) Then
                ' los abstencionistas no los cuento
                If LCase(EstadoActual.VectorResultados(xBanca)) <> ABSTENCION_AUTORIZADA Then
                    'objeto == term.keyb
                    If (MensajeBanca.sComponente = "term.keyb.si" Or MensajeBanca.sComponente = "term.keyb.no") And LCase(MensajeBanca.sAtributo) = "state" And LCase(MensajeBanca.sValor) = "on" Then
                        'Actualiza pendientes de votar si antes no habia votado
                        If LCase(EstadoActual.VectorResultados(xBanca)) = ABSTENCION Then
                            EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1
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
                    End If
                '>> PROCESO DE CLOSED se permite ingresar a habilitarse para votar
                Case "closed"
                    'Si es numérica y esta presente
                    If EstadoActual.TipoDeOperacion = "votnum" And EstadoActual.VectorPresencia(xBanca) = PRESENTE Or (EstadoActual.ModoMantenimientoBancas = 1 And xBanca = 0) Then
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
                        .sComentario = EstadoActual.EstadoVotacion_y_PasList
                    End With
                    Call EnviarMensajesBancas(MensajeParaBanca)
                End If
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
                            Case NEGATIVO
                                CartelActual.Negativos = CartelActual.Negativos - 1
                            Case ABSTENCION
                                EstadoActual.PendientesEmitirVotos = EstadoActual.PendientesEmitirVotos - 1
                            Case ABSTENCION_AUTORIZADA
                                EstadoActual.AbstencionistasAutorizados = EstadoActual.AbstencionistasAutorizados - 1
                        End Select
                        EstadoActual.VectorResultados(xBancaDuplicada) = ABSTENCION
                        With MensajeParaBanca
                            .sTipo = "mset"
                            .sObjeto = xBancaDuplicada
                            .sComponente = "term.keyb"
                            .sAtributo = "state"
                            .sValor = "off" & IIf(xBancaDuplicada > 0, EstadoActual.TipoDeOperacion, "votnum")
                            .sComentario = EstadoActual.EstadoVotacion_y_PasList
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
        Call CartelSerial("resultado")
        PintarVectorColor (xBanca)
        'Presentar resultados en cartel
        EstadoActual.EstadoVotacion_y_PasList = "finalizada"
        'cancela teclado presidente
        With MensajeParaBanca
            .sTipo = "mset"
            .sObjeto = xBanca
            .sComponente = "term.keyb"
            .sAtributo = "state"
            .sValor = "off" & IIf(xBanca > 0, EstadoActual.TipoDeOperacion, "votnum")
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
        'VERIFICACION DE TIEMPO CUMPLIDO DE VOTACION solo en estado votando sea nominal o numerica
        If EstadoActual.EstadoVotacion_y_PasList = "votando" And Not (EstadoActual.ModoMantenimientoBancas) Then
            If (DateDiff("s", EstadoActual.FechaVotacion, Now)) >= EstadoActual.TiempoParaVotacion + xSegundosFinOperacion Then
                'si es modalidad de abstencion automatica
                If EstadoActual.TipoDeAbstencion = "absaut" Then
                    EstadoActual.EstadoVotacion_y_PasList = "cierre"
                ElseIf EstadoActual.TipoDeAbstencion = "votlar" Then 'si es modalidad de votacion larga
                    If Not (CartelActual.LeyendaQuorum = "QUORUM") Then 'Si no hay quorum, la cancela
                        EstadoActual.EstadoVotacion_y_PasList = "cancelada"
                        FinVotacionBrc ("cancelada")
                    Else
                        EstadoActual.EstadoVotacion_y_PasList = "larga"
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
        EstadoActual.EstadoVotacion_y_PasList = "cierre"
    End If
    'En estado de cierre de votacion
    'Si no hay quorum, y todavia no voto el presidente (no empa) la cancela
    'CartelActual.LeyendaQuorum = IIf(CartelActual.Presentes >= Fix(xMinimoParaQuorumEntero), "QUORUM", "NO HAY QUORUM")
    If EstadoActual.EstadoVotacion_y_PasList = "cierre" Then
        If Not (CalculoQuorum() = "QUORUM") And ((EstadoActual.VectorResultados(0) <> AFIRMATIVO And EstadoActual.VectorResultados(0) <> NEGATIVO) Or xPresidenteLegislador) Then
           ' antes nominal If Not (IIf(EstadoActual.Presentes >= Fix(xMinimoParaQuorumEntero), "QUORUM", "NO HAY QUORUM") = "QUORUM") And ((EstadoActual.VectorResultados(0) <> AFIRMATIVO And EstadoActual.VectorResultados(0) <> NEGATIVO) Or xPresidenteLegislador) Then
            EstadoActual.EstadoVotacion_y_PasList = "cancelada"
            FinVotacionBrc ("cancelada")
        Else
            CartelActual.Resultado = CalculoResultado(EstadoActual.BaseMayoria, EstadoActual.TipoMayoria, xMiembrosDelCuerpo, EstadoActual.Presentes, CartelActual.Afirmativos, CartelActual.Negativos, "", 0, 0, CartelActual.MinimoVotosParaAfirmativo, EstadoActual.VectorResultados(0), IIf(xPresidenteLegislador, 1, 0))
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
               EstadoActual.OcupadosNoIdentificados = CuentaOcupadosNoIdentificadosCong()
            End If
            CartelActual.Abstenciones = EstadoActual.Presentes - CartelActual.Afirmativos - CartelActual.Negativos - IIf(EstadoActual.EstadoVotacion_y_PasList = "votnom", EstadoActual.OcupadosNoIdentificados, 0) '- IIf(InStr("sn", EstadoActual.VectorResultados(0)) = 0, 0, 1)
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
            
            'Se pinta con los resultados en votacion numerica tambien
                'Solo si es votacion nominal Actualizar informacion en pantalla operador cambiando colores
                'If EstadoActual.TipoDeOperacion = "votnom" Then
                '    Call PintarTodasLasBancas
                'End If
            Call PintarTodasLasBancas
            
            'si es empate
            If CartelActual.Resultado = "EMPATE" And EstadoActual.VectorResultados(0) <> AFIRMATIVO And EstadoActual.VectorResultados(0) <> NEGATIVO Then
                EstadoActual.EstadoVotacion_y_PasList = "empate"
                'habilitar al presidente para votar
                With MensajeParaBanca
                    .sTipo = "mset"
                    .sObjeto = 0
                    .sComponente = "term.keyb"
                    .sAtributo = "state"
                    .sValor = "onvotnum"
                    .sComentario = EstadoActual.EstadoVotacion_y_PasList
                End With
                Call EnviarMensajesBancas(MensajeParaBanca)
            End If
        End If 'hay quorum o desempato el presidente,
    End If 'EstadoActual.EstadoVotacion_y_PasList = "cierre"
    If EstadoActual.EstadoVotacion_y_PasList = "finalizada" Then
        'Grabacion de acta
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
    EstadoActual.VectorColor(X) = AsignarColor(X)
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
    strArchivo = strDirectorio & "\logsqv" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "-" & Trim(Format(Time, "HHMMSS")) & ".txt"
    Open strArchivo For Binary As #xFileSqv
        strCadena = "COMIENZO: " & Now & " - AC: " & Trim(Format(xCiclosTotales / (DateDiff("s", xFechaArranque, Now)), "##0.00")) & vbCrLf
        Put #xFileSqv, , strCadena
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
    
    strCadena = "FIN: " & Now
    Put #xFileSqv, , strCadena
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
            If (DateDiff("s", xFechaArranque, Now)) > 0 Then
               lblCiclos.Caption = "SBº" & Str(xNroMensajeSB) & " ac " & Trim(Format(xCiclosTotales / (DateDiff("s", xFechaArranque, Now)), "##.00"))
               Call OpenLogFile
               Put #xFileSqv, , xLogSQVPrueba
               Call closeLogfile
               xLogSQVPrueba = ""
               xFechaArranque = Now
               xCiclosTotales = 0
            End If
        End If
        'If EstadoActual.TipoDeOperacion = "votnom" Or _
        '        EstadoActual.TipoDeOperacion = "paslis" Or _
        '        (EstadoActual.TipoDeOperacion = "quorum" And EstadoActual.Modo_Ident_Nom_Obsoleto = 1) And _
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
    
    If txtVecesPorSegundo.Text = "" Then
        txtVecesPorSegundo.Text = 2
        Exit Sub
    End If
    ' blServerPrendido = False
    blServerPrendido = True
    Call ServerOnOff
    If Trim(txtVecesPorSegundo.Text) <> Str(xIntervalo) Then
        xIntervalo = Int(txtVecesPorSegundo.Text)
        Call ServerOnOff
    End If
    
End Sub

Private Sub txtVecesPorSegundo_GotFocus()
    With txtVecesPorSegundo
        .SelStart = 0
        .SelLength = Len(.Text)
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
        lblOcupadosNoIdentificados.Caption = Str(.OcupadosNoIdentificados)
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
               & "Base_de_Mayoría = '" & .BaseMayoria & "', Tipo_de_Mayoría = '" & .TipoMayoria & "', Modo_identifica_nom_Obsoleto = " & .Modo_Ident_Nom_Obsoleto & ", " _
               & "strError = '" & .strError & "', Estado_de_votacion = '" & .EstadoVotacion_y_PasList & "', Vector_resultado = '" & strResultados & "', " _
               & "Tipo_de_Abstención = '" & .TipoDeAbstencion & "', Mensaje_al_operador = '" & .MensajeAlOperador & "', Pendientes_Emitir_Voto = " & .PendientesEmitirVotos & ", " _
               & "Grabar_automaticamente = " & .GrabarAutomaticamente & ", Listar_automaticamente = " & .ListarAutomaticamente & ", " _
               & "Tipo_Mayoria_Quorum = '" & .TipoMayoriaQuorum & "', Leyenda_Quorum = '" & CartelActual.LeyendaQuorum & "', Período_Legislativo = '" & .PeriodoLegislativo & "', " _
               & "Fecha = '" & Format(Now, "dd/mm/yyyy hh:mm:ss") & "', Hora = '" & Time & "', Acta_Grabada = " & .ActaGrabada & ", Solicitud_Grabacion_Manual = " & .SolicitudGrabarManual & ", " _
               & "Tiempo_de_votación = " & .TiempoParaVotacion & ", IP_Consola_Habilitada = '" & .IP_Consola & "', Modo_Mantenimiento_Bancas = " & .ModoMantenimientoBancas & ", " _
               & "Modo_Normal_Mant_Sistema = " & .ModoNormalMantSistema & ", Identificador_de_Formulario = '" & .IdentificadorDeFormulario & "', Encender_Carteles = " & .CartelEncendido & ", Estado_Sesion = '" & .EstadoSesion & "', FechaVotacion = '" & Format(EstadoActual.FechaVotacion, "dd/mm/yyyy hh:mm:ss") & "', HoraVotacion = '" & EstadoActual.HoraVotacion & "'"
    ' Misma sentencia SQL, pero con SP
    'MsgBox .PeriodoLegislativo
    
    strSql = "update_vector(" & .Presentes & ", " & .Ausentes & ", '" & strColor & "', " & _
             "'" & strPresencia & "','" & strIdentificacion & "','" & .TipoDeOperacion & "', " & _
             "'" & CartelActual.Resultado & "', " & CartelActual.Afirmativos & ", " & CartelActual.Negativos & ", " & _
             " " & CartelActual.Abstenciones & ", " & .OcupadosNoIdentificados & ", " & CartelActual.MinimoVotosParaAfirmativo & ", " & _
             " " & .Sesion & ", " & .NroActa & ", '" & .TituloDelActa & "', " & _
             " '" & .EstadoSesion & "', '" & .BaseMayoria & "', '" & .TipoMayoria & "', " & _
             " " & .Modo_Ident_Nom_Obsoleto & ", '" & .strError & "', '" & .EstadoVotacion_y_PasList & "', " & _
             " '" & strResultados & "', '" & .TipoDeAbstencion & "', " & .PendientesEmitirVotos & "," & _
             " '" & .MensajeAlOperador & "', " & .GrabarAutomaticamente & ", " & .ListarAutomaticamente & ", " & _
             " '" & .TipoMayoriaQuorum & "', '" & CartelActual.LeyendaQuorum & "', '" & .PeriodoLegislativo & "', " & _
             " '" & Format(Now, "dd/mm/yyyy hh:mm:ss") & "', '" & Time & "', " & .ActaGrabada & ", " & _
             " " & .SolicitudGrabarManual & ", " & .TiempoParaVotacion & ", '" & .IP_Consola & "', " & _
             " " & .ModoMantenimientoBancas & ", " & .ModoNormalMantSistema & " , '" & .IdentificadorDeFormulario & "', " & _
             " " & .CartelEncendido & ", '" & Format(EstadoActual.FechaVotacion, "dd/mm/yyyy hh:mm:ss") & "', '" & EstadoActual.HoraVotacion & "', '" & strAbstencion & "'" & _
             ")"
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
    With Rs
        .AddNew
        .Fields("Tipo_de_operación").Value = EstadoActual.TipoDeOperacion
        .Fields("Período_Legislativo").Value = EstadoActual.PeriodoLegislativo
        .Fields("Sesión").Value = EstadoActual.Sesion
        .Fields("Número_de_Acta").Value = EstadoActual.NroActa
        .Fields("Versión_Acta").Value = ""
        .Fields("Ultima_Versión_Acta").Value = ""
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
            MsgBox "Error Nº" & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            Resume
    End Select
End Sub
Private Sub GuardarActas()
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
    With Rs
        .AddNew
        .Fields("Tipo_de_operación").Value = EstadoActual.TipoDeOperacion
        .Fields("Período_Legislativo").Value = EstadoActual.PeriodoLegislativo
        .Fields("Sesión").Value = EstadoActual.Sesion
        .Fields("Número_de_Acta").Value = EstadoActual.NroActa
        .Fields("Versión_Acta").Value = ""
        .Fields("Ultima_Versión_Acta").Value = ""
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
            MsgBox "Error Nº" & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            Resume
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
    ' Prender scan nuevamente y solicitar identificacion, dependiendo
    ' Segun xModo, se trata de una reconexion preventiva o un pedido de inicio real.
    ' En modo start, se envia primero un scancl a la banca.
    
    If Trim(LCase(xModo)) = "key_start" Then
        xModo = "key_start"
    Else
        If xModo <> "restart" Then
            xModo = "start"
        End If
    End If
    With MsgSistema
        .sObjeto = xBanca '<AP 040115 faltaba indicar la banca>
        .sTipo = "mset"
        .sComponente = "term.auth"
        .sAtributo = "action"
        If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
            .sValor = Trim("auth_" & xModo)
            .sComentario = xComentario & " Modo normal"
        Else
            .sValor = "auth_" & xModo
            '.sValor = "auth_test"
            .sComentario = xComentario & " Modo mantenimiento"
        End If
    End With
    Call EnviarMensajesBancas(MsgSistema)
End Sub

Private Sub EnviarMensajesFinAuth(xBanca As String, xComentario As String) '< unifica llamadas de encencido del scanner
    Dim MsgSistema As MensajeSistema
                            
    ' Prender scan nuevamente y solicitar identificacion, dependiendo
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
            If EstadoActual.VectorPresencia(xBanca) = PRESENTE And .sComponente = "term.auth" Then
                If LCase(.sAtributo) = "result" Then
                    If InStr(.sValor, "|") > 0 Then ' Puede recibirse en el Valor una serie de parametros de reintentos del usuario los cuales toma aqui.
                        'MsgBox "hablar con marcos para determinar como me van a llegar los mensajes de banca con
                        ' problemas para identificarse... y para ver este asunto de cuantos intentos de identificacion
                        ' tiene esa banca..."
                    End If '<AP 040115 abro el endif porque sino no procesaria negative con parametros>
                    If .sValor = "negative" Then ' Se recibio un resultado negativo, es decir no lo identifico.
                        ' Ver si no fue identificado manualmente
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
                    If xIntentosRealizados >= 1 And xLecturasNegativas = 0 And EstadoActual.VectorColor(xBanca) = cVERDE Then
                        'Desactivar indicador de reintentos anterior si dejo de intentar cuando tras al menos dos ciclos de scan no se observó ningún intento de identificacion negativa.
                        Call PintarVectorColor(xBanca)
                        Call AltaLogGeneral("Normalizacion desde REINTENTO ID", Trim(.sValor) & "Banca Nro. " & Trim(Str(xBanca)), Str(xBanca), "1")    '<AP 040115 agrego log>
                    End If
                    If Not (.sValor = "negative") Then  ' Si se identificó correctamente al legislador
                        If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then ' y no estamos en modo Mantenimiento
                            ' Hay que verificar que el legislador se encuentre en al tabla de legisladores activos
                            ' strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                   & "Legisladores.departamento , Legisladores.cargo FROM Legisladores INNER JOIN legisladores_activos ON " _
                                   & "Legisladores.id = legisladores_activos.id WHERE (Legisladores.id = '" & CerosIzquierda(.sValor, 8) & "' and Legisladores.tipo = 1)" '<AP 040115 Pide que sea tipo legislador
                            
                            ' TECLADO: si se trata de una identificacion por teclado
                            vTemporal = Val("&H" & .sValor)
                            If vTemporal > 99999 Then
                                .sValor = Trim(Str(vTemporal))
                                .sValor = Encripta.EncryptString(.sValor)
                                strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                       & "Legisladores.departamento , Legisladores.cargo FROM Legisladores INNER JOIN legisladores_activos ON " _
                                       & "Legisladores.id = legisladores_activos.id WHERE (Legisladores.Pin = '" & Trim(.sValor) & "') AND (Legisladores.tipo = 1)"  '<AP 040115 Pide que sea tipo legislador
                            Else
                                If Not IsNumeric(.sValor) Then
                                    .sValor = "999999"
                                End If
                                strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                       & "Legisladores.departamento , Legisladores.cargo FROM Legisladores INNER JOIN legisladores_activos ON " _
                                       & "Legisladores.id = legisladores_activos.id WHERE Cast(Legisladores.id AS Int) = " & Str(Int(.sValor)) & " and Legisladores.tipo = 1"  '<AP 040115 Pide que sea tipo legislador
                            End If
                        Else ' Si se encuentra en modo mantenimiento
                            'Arma la lista de mantenimientos
                            EstadoActual.MantIdentificaciones = Trim(.sValor) & ";" & Trim(EstadoActual.MantIdentificaciones)
                            ' IDs recibidos = TRIM (.sValor) & ';' & TRIM (IDs recibidos)
                            ' <AP 040115 Hay que verificar que el legislador se encuentre como personal de mantenimiento, sin join
                            ' strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                   & "Legisladores.departamento , Legisladores.cargo FROM Legisladores WHERE Legisladores.id  = '" & CerosIzquierda(.sValor, 8) & "' and Legisladores.tipo = 0)" '<AP 040115 Pide que sea tipo personal de mantenimiento
                            strSql = "SELECT Legisladores.id, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, " _
                                   & "Legisladores.departamento , Legisladores.cargo FROM Legisladores WHERE Cast(Legisladores.id AS Int) = " & Str(Int(.sValor)) & " and Legisladores.tipo = 0" '<AP 040115 Pide que sea tipo personal de mantenimiento
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
                                Call EnviarMensajesComienzoAuth(Str(xBanca), "ID Invalido") 'Unifica envio mensaje comienzo autorizacion
                                Call MensajeDisplayTerminal(Str(xBanca), "Id. invalida:" & strIdLegislador & " Por favor reintente.")
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
                                If flExitoPierdeIdDup = False Then ' identificar al legislador en vector identificacion
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
                                        If EstadoActual.VectorIdentificacion(xBanca) = 0 And (EstadoActual.ModoMantenimientoBancas = 1 Or (LegisladorHabilitado(strIdLegislador))) Then
                                            'La identificacion ha sido exitosa!
                                            EstadoActual.VectorIdentificacion(xBanca) = strIdLegislador
                                            EstadoActual.OcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados - 1
                                              'With MsgSistema
                                                   MsgSistema.sObjeto = xBanca '<AP 040115 faltaba indicar la banca>
                                                   MsgSistema.sTipo = "mset"
                                                   MsgSistema.sComponente = "term.led1"
                                                   MsgSistema.sAtributo = "state"
                                                   MsgSistema.sValor = "on"
                                                   If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
                                                       MsgSistema.sComentario = "Id aceptado Modo normal"
                                                   Else
                                                       MsgSistema.sComentario = "Id aceptado Modo mantenimiento"
                                                   End If
                                              'End With
                                              Call EnviarMensajesBancas(MsgSistema)
                                                                                                  
                                              If EstadoActual.TipoDeOperacion = "votnom" Then
                                                If InStr(Join(EstadoActual.VectorAbstencion, SEPARADOR_VECTOR), Trim(strIdLegislador)) > 0 Then
                                                    AbstenerBanca (xBanca)
                                                End If
                                              End If

                                              ' envia mensaje por display... opcional
                                              'With MsgSistema
                                                   'MsgSistema.sObjeto = xBanca '<AP 040115 faltaba indicar la banca>
                                                   'MsgSistema.sTipo = "mset"
                                                   'MsgSistema.sComponente = "term.display"
                                                   'MsgSistema.sAtributo = "text"
                                                   'If EstadoActual.ModoMantenimientoBancas = 0 And EstadoActual.ModoNormalMantSistema = 0 Then
                                                   '    MsgSistema.sValor = "Identificacion Aceptada"
                                                   '    MsgSistema.sComentario = "Id aceptado Modo normal"
                                                   'Else
                                                   '    MsgSistema.sValor = "Identificacion de Prueba"
                                                   '    MsgSistema.sComentario = "Id aceptado Modo mantenimiento"
                                                   'End If
                                              'End With
                                              Call EnviarMensajesBancas(MsgSistema)
                                              flBancaIdentifPosExitosa = True
                                              Call PintarVectorColor(xBanca)
                                        Else
                                            Call MensajeDisplayTerminal(Str(xBanca), "Reconsideracion: No habilitado.")
                                            Call AltaLogGeneral("Identificacion", "Id no habilitado: " & strIdLegislador & ", Banca " & xBanca, Str(xBanca), "2")
                                        End If
                                    End If
                                Else ' Si el legislador ya esta identificado en otra banca
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
                                            Call AltaLogGeneral("BANCA DUPLICADA Ultima", strInforme, Str(xBanca), "5")
                                            Call AltaLogGeneral("BANCA DUPLICADA Primera", strInforme, Str(xBancaDuplicada), "5")
                                            EstadoActual.VectorColor(xBanca) = cROJO ' Avisar al operador lo que esta pasando
                                            EstadoActual.VectorColor(xBancaDuplicada) = cROJO ' <AP 040115 Ambas bancas en rojo
                                            'A ambos les reenciende el scanner
                                            Call EnviarMensajesComienzoAuth(Str(xBanca), "Banca Duplicada - ultimo intento") 'Unifica envio mensaje comienzo autorizacion
                                            Call EnviarMensajesComienzoAuth(Str(xBancaDuplicada), "Banca Duplicada - anterior id") 'Unifica envio mensaje comienzo autorizacion
                                        End If
                                    End If
                                End If
                            End If
                            RsLocal.Close
                    End If ' .sValor <> "negative" (si fue negative, se trato en if anterior.)
                End If ' FIN .sAtributo = "result"
            End If 'term.auth
            If LCase(.sComponente) = "term.seat" Then
                If LCase(.sAtributo) = "switch" And flSwitchExitoso = True Then
                    If LCase(.sValor) = "closed" Then
                        If EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO Then
                            EstadoActual.OcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados + 1
                            Call EnviarMensajesComienzoAuth(Str(xBanca), "SW Closed")
                        End If 'no id
                    End If 'closed
                    If LCase(.sValor) = "open" Then '>> El legislador se levanta, pierde la identificacion.
                        If Not (EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO) Then
                            flExitoPierdeID = True
                        Else
                            EstadoActual.OcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados - 1
                        End If 'no id
                        EstadoActual.VectorIdentificacion(xBanca) = NO_IDENTIFICADO
                        Call EnviarMensajesFinAuth(Str(xBanca), "SW Open")  'apagar scanner
                        ' el led se debe apagar solo
                    End If 'open
                End If 'switch
            End If 'seat
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
End Sub
Private Sub CalcularMinimoAfirmativaCartel()
    CartelActual.MinimoVotosParaAfirmativo = 0
    If (EstadoActual.BaseMayoria = "legpre" Or EstadoActual.BaseMayoria = "votemi") And Not (CartelActual.LeyendaQuorum = "QUORUM") Then
        CartelActual.LeyendaMinimoVotosParaAfirmativo = "N/D"
    Else
        Call CalculoResultado(IIf(EstadoActual.BaseMayoria = "votemi", "legpre", EstadoActual.BaseMayoria), EstadoActual.TipoMayoria, xMiembrosDelCuerpo, CartelActual.Presentes, 0, 0, "", 0, 0, CartelActual.MinimoVotosParaAfirmativo, " ", IIf(xPresidenteLegislador, 1, 0))
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
        
        .Presentes = 0
        .Ausentes = xMiembrosDelCuerpo
        .PresentesCongelados = 0
        .AusentesCongelados = xMiembrosDelCuerpo
                
        .ActaGrabada = 0
        
        .OcupadosNoIdentificados = 0
        .PendientesEmitirVotos = 0
        .AbstencionistasAutorizados = 0
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
    Call PintarTodasLasBancas
    Call ActualizarVector_enBD
    With Mensaje2Banca ' Mensaje para SB
        .sTipo = "mget"
        .sComponente = "term.mon"
        .sObjeto = "brc"
        .sAtributo = "action"
        .sValor = "reset"
    End With
    Call EnviarMensajesBancas(Mensaje2Banca)
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
    strSP = "insert_log_general('" & strOrigen & "','" & strDetalle & "','" & nObjeto & "','" & strSeveridad & "')"
    Cn.Execute strSP
End Sub

Private Sub PintarBancasCartel()
    On Error Resume Next
    Dim i     As Integer
    Dim clave As String
    'busco los datos del presi sólo si cambia
    For i = 0 To UBound(EstadoActual.VectorColor)
        'clave = i
        ctrBanca(i).BackColor = mColores(Val(EstadoActual.VectorColor(i)))
        If ctrBanca(i).BackColor = &H0 Then
            ctrBanca(i).ForeColor = &HE0E0E0
        End If
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
    Dim xBanca As Long
    
    For xBanca = 0 To xUltimaBanca
        ctrBanca(xBanca).Caption = xBanca
    Next xBanca
End Sub


Private Function LeyendaSesion() As String
    Dim strEtiqueta      As String

    strEtiqueta = Left(EstadoActual.PeriodoLegislativo, 3) & "º Período Legislativo: "
    
    Select Case UCase(Mid(EstadoActual.PeriodoLegislativo, 4, 1))
        Case "O"
            strEtiqueta = strEtiqueta & "Ordinario "
        Case "E"
            strEtiqueta = strEtiqueta & "Especial "
        Case "P"
            strEtiqueta = strEtiqueta & "Preparatoria "
    End Select
    strEtiqueta = strEtiqueta & " - " & Str(EstadoActual.Sesion) & "º Sesión "
    Select Case UCase(Mid(EstadoActual.PeriodoLegislativo, 5, 1))
        Case "T"
            strEtiqueta = strEtiqueta & "Tablas"
        Case "E"
            strEtiqueta = strEtiqueta & "Especial"
        Case "P"
            strEtiqueta = strEtiqueta & "Preparatoria"
    End Select
    
    'strEtiqueta = strEtiqueta & " - Próximo Nº de Acta: " & Str(EstadoActual.NroActa)
    
    LeyendaSesion = strEtiqueta
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
            LeyendaTipoOperacion = "Quórum"
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
    Select Case EstadoActual.TipoMayoria
        Case "120"
                   LeyendaTipoMayoria = "Mas de la mitad"
        Case "121"
                   LeyendaTipoMayoria = "Mitad mas uno"
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
End Function


Private Function LeyendaBaseMayoria() As String
    Select Case EstadoActual.BaseMayoria
        Case "legpre"
                   LeyendaBaseMayoria = "Legisladores Presentes"
        Case "miecue"
                   LeyendaBaseMayoria = "Miembros del Cuerpo"
        Case "votemi"
                   LeyendaBaseMayoria = "Votos Emitidos"
    End Select
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
        Unload Me
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
                txtTipoOperacion.Text = !descTipoOp
            End If
            If IsNull(!Sesión) = False Then
                txtSesion.Text = !Sesión
            End If
            If IsNull(!Número_de_Acta) = False Then
                txtNroActa.Text = !Número_de_Acta
            End If
            If IsNull(!versión_acta) = False Then
                txtVersion.Text = !versión_acta
                If !Ultima_Versión_Acta = 0 Then
                    txtVersion.Text = "Original"
                Else
                   If !versión_acta = 0 Then
                        txtVersion.Text = "Ult.Mod.Ver. " & Val(!Ultima_Versión_Acta) + 1
                   Else
                        txtVersion.Text = "Ver. " & Val(!Ultima_Versión_Acta) + 1
                   End If
                End If
                txtVersion.Tag = !versión_acta
            End If
            If IsNull(!Nombre_del_Acta) = False Then
                txtNombre.Text = Trim(!Nombre_del_Acta)
            End If
            If IsNull(!Fecha) = False Then
                txtFecha.Text = !Fecha
            End If
            If IsNull(!Hora) = False Then
                txtHora.Text = !Hora
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
                txtAbstencionesTotales.Text = !Abstenciones_Total
            Else
                txtAbstencionesTotales.Text = "0"
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
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            Resume
    End Select
End Sub



Private Sub MostrarActaProyector(mSesion As Long, mActa As Long, mVersion As Long)
FrameSQVActa.ZOrder 0
Call MostrarDatosSesion(mSesion, mActa, mVersion)
MsgBox " dd"

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
    CalculoQuorum = IIf(CartelActual.Presentes >= Fix(xMinimoParaQuorumEntero), "QUORUM", "NO HAY QUORUM")
End Function
Private Function PresentesIdentificados() As Long
    PresentesIdentificados = EstadoActual.Presentes - EstadoActual.OcupadosNoIdentificados
End Function
Private Function Presentes() As Long
    If EstadoActual.Modo_Ident_Nom_Obsoleto = 1 Then
        Presentes = PresentesIdentificados()
    Else
        Presentes = EstadoActual.Presentes
    End If
End Function

Private Function Ausentes() As Long
    If EstadoActual.Modo_Ident_Nom_Obsoleto = 1 Then
        Ausentes = xMiembrosDelCuerpo - PresentesIdentificados()
    Else
        Ausentes = EstadoActual.Ausentes
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
    Call AltaLogGeneral("SQV SERVER: Abstencion vector", strVector, , "0")
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
    lblOcupadosNoIdentificados = EstadoActual.OcupadosNoIdentificados
End Sub

Private Sub FinVotacionBrc(xComentario As String)

    Dim Mensaje2Banca As MensajeSistema
        
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
        
        Mensaje2Banca.sObjeto = "0"
        Mensaje2Banca.sValor = "offvotnum"
        
        Call AltaLogGeneral("SQV SERVER: Fin Votacion", xComentario, , "1")
        
        Call EnviarMensajesBancas(Mensaje2Banca)
        'MsgBox "verificar cierre de votacion"
    
End Sub

Private Function Replicar(xCant As Long, strCadena As String) As String
    Dim X As Long
    For X = 1 To xCant
        Replicar = Replicar & strCadena
    Next X
End Function



