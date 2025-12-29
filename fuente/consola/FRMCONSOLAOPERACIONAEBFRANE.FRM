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
   ScaleHeight     =   11520
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox MR2 
      Align           =   4  'Align Right
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6210
      Left            =   735
      Picture         =   "frmConsolaOperacionAEBFranes.frx":0000
      ScaleHeight     =   6210
      ScaleWidth      =   14370
      TabIndex        =   66
      Top             =   780
      Visible         =   0   'False
      Width           =   14370
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   329
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   328
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   327
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   326
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   325
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   324
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   323
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   322
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   321
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   320
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   319
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   318
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   317
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   316
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   315
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   314
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   313
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   312
         Top             =   1020
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   311
         Top             =   900
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   310
         Top             =   540
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblBloqueInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pres: 04  Aus:15  Ident: 02"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   309
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   233
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   232
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   231
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   230
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   229
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   228
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   227
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   226
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   225
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   224
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   223
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   222
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   221
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   220
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   219
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   218
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   217
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   216
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   215
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   214
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   213
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   212
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   211
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   210
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   209
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   208
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   207
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   206
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   205
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   204
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   203
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   202
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   201
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   200
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   199
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   198
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   197
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   196
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   195
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   194
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   193
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   192
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   191
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   190
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   189
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   188
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   187
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   186
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   185
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   184
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   183
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   182
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   181
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   180
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   179
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   178
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   177
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   176
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   175
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   174
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
            Name            =   "Accidental Presidency"
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
            Name            =   "Accidental Presidency"
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
            Name            =   "Accidental Presidency"
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
            Name            =   "Accidental Presidency"
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
            Name            =   "Accidental Presidency"
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
            Name            =   "Accidental Presidency"
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
            Name            =   "Accidental Presidency"
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
            Name            =   "Accidental Presidency"
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
            Name            =   "Accidental Presidency"
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
            Name            =   "Accidental Presidency"
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   163
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   162
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
            Name            =   "Accidental Presidency"
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
         TabIndex        =   161
         Top             =   2580
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   160
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   159
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   158
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   157
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   156
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   155
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   154
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   153
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   152
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   151
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   150
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   149
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   148
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   147
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   146
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   145
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   144
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   143
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   142
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   141
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   140
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   139
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   138
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   137
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   136
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   135
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   134
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   133
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   132
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   131
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   130
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   129
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   128
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   127
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   126
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   125
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   124
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   123
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   122
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   121
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   120
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   119
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   118
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   117
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   116
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   115
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   114
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   113
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   112
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   111
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   110
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   109
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   108
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   107
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   106
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   105
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   104
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   103
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   102
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   101
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   100
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   99
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   98
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   97
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   96
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   95
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   94
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   93
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   92
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   91
         Top             =   1950
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   90
         Top             =   1500
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   89
         Top             =   1080
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   88
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   87
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   86
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   85
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   84
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   83
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   82
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   81
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   80
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   79
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   78
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   77
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   76
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   75
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   74
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   73
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   72
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   71
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   70
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   69
         Top             =   150
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblLegis 
         BackStyle       =   0  'Transparent
         Caption         =   "Legislador Mambrusco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   68
         Top             =   570
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label lblBloque 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Bloque"
         BeginProperty Font 
            Name            =   "Accidental Presidency"
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
         TabIndex        =   67
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
   Begin VB.PictureBox MR1 
      Align           =   4  'Align Right
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6210
      Left            =   -13635
      Picture         =   "frmConsolaOperacionAEBFranes.frx":8C6F
      ScaleHeight     =   6210
      ScaleWidth      =   14370
      TabIndex        =   236
      Top             =   780
      Width           =   14370
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
         Left            =   1260
         TabIndex        =   308
         Top             =   0
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
         TabIndex        =   307
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
         TabIndex        =   306
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
         TabIndex        =   305
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
         TabIndex        =   304
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
         TabIndex        =   303
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
         TabIndex        =   302
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
         TabIndex        =   301
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
         TabIndex        =   300
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
         TabIndex        =   299
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
         TabIndex        =   298
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
         TabIndex        =   297
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
         TabIndex        =   296
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
         TabIndex        =   295
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
         TabIndex        =   294
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
         TabIndex        =   293
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
         TabIndex        =   292
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
         TabIndex        =   291
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
         TabIndex        =   290
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
         TabIndex        =   289
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
         TabIndex        =   288
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
         TabIndex        =   287
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
         TabIndex        =   286
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
         TabIndex        =   285
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
         TabIndex        =   284
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
         TabIndex        =   283
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
         TabIndex        =   282
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
         TabIndex        =   281
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
         TabIndex        =   280
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
         TabIndex        =   279
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
         TabIndex        =   278
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
         TabIndex        =   277
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
         TabIndex        =   276
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
         TabIndex        =   275
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
         TabIndex        =   274
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
         TabIndex        =   273
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
         TabIndex        =   272
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
         TabIndex        =   271
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
         TabIndex        =   270
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
         TabIndex        =   269
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
         TabIndex        =   268
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
         TabIndex        =   267
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
         TabIndex        =   266
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
         TabIndex        =   265
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
         TabIndex        =   264
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
         TabIndex        =   263
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
         TabIndex        =   262
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
         TabIndex        =   261
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
         TabIndex        =   260
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
         TabIndex        =   259
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
         TabIndex        =   258
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
         TabIndex        =   257
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
         TabIndex        =   256
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
         TabIndex        =   255
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
         TabIndex        =   254
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
         TabIndex        =   253
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
         TabIndex        =   252
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
         TabIndex        =   251
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
         TabIndex        =   250
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
         TabIndex        =   249
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
         TabIndex        =   248
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
         TabIndex        =   247
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
         TabIndex        =   246
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
         TabIndex        =   245
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
         TabIndex        =   244
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
         TabIndex        =   243
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
         TabIndex        =   242
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
         TabIndex        =   241
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
         TabIndex        =   240
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
         TabIndex        =   239
         Top             =   480
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
         Left            =   900
         TabIndex        =   238
         Top             =   480
         Width           =   480
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
         TabIndex        =   237
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
         Left            =   3090
         Shape           =   3  'Circle
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.PictureBox ML 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   6210
      Index           =   0
      Left            =   0
      Picture         =   "frmConsolaOperacionAEBFranes.frx":1287C
      ScaleHeight     =   6210
      ScaleWidth      =   990
      TabIndex        =   235
      Top             =   780
      Width           =   990
      Begin VB.Image imgA 
         Height          =   360
         Index           =   0
         Left            =   450
         Picture         =   "frmConsolaOperacionAEBFranes.frx":13C99
         ToolTipText     =   "Ver mapa de Bancas"
         Top             =   150
         Width           =   360
      End
      Begin VB.Image imgB 
         Height          =   360
         Index           =   0
         Left            =   450
         Picture         =   "frmConsolaOperacionAEBFranes.frx":14383
         ToolTipText     =   "Ver Legisladores Presentes"
         Top             =   750
         Width           =   360
      End
   End
   Begin VB.PictureBox ML 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   6210
      Index           =   1
      Left            =   990
      Picture         =   "frmConsolaOperacionAEBFranes.frx":14A6D
      ScaleHeight     =   6210
      ScaleWidth      =   990
      TabIndex        =   234
      Top             =   780
      Visible         =   0   'False
      Width           =   990
      Begin VB.Image imgA 
         Height          =   360
         Index           =   1
         Left            =   450
         Picture         =   "frmConsolaOperacionAEBFranes.frx":15ED4
         ToolTipText     =   "Ver mapa de Bancas"
         Top             =   150
         Width           =   360
      End
      Begin VB.Image imgB 
         Height          =   360
         Index           =   1
         Left            =   450
         Picture         =   "frmConsolaOperacionAEBFranes.frx":165BE
         ToolTipText     =   "Ver Legisladores Presentes"
         Top             =   750
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   0
      Picture         =   "frmConsolaOperacionAEBFranes.frx":16CA8
      ScaleHeight     =   4530
      ScaleWidth      =   15105
      TabIndex        =   8
      Top             =   6990
      Width           =   15105
      Begin VB.TextBox txtTitulo 
         Height          =   675
         Left            =   1500
         MaxLength       =   80
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   2520
         Width           =   4395
      End
      Begin VB.CommandButton cmdControlSistema 
         Caption         =   "&Control de Sistema"
         Height          =   330
         Left            =   4530
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Período legislativo y Tipo de Sesión"
         Top             =   3900
         Width           =   1755
      End
      Begin VB.CommandButton cmdPeriodoLegislativo 
         Caption         =   "Período Legislativo y Tipo de Sesión"
         Height          =   345
         Left            =   540
         TabIndex        =   37
         ToolTipText     =   "Seleccionar Período legislativo y Tipo de Sesión"
         Top             =   1800
         Width           =   2760
      End
      Begin VB.CommandButton cmdTiempo 
         Caption         =   "Tiemp&o de Votación"
         Height          =   330
         Left            =   2640
         TabIndex        =   36
         ToolTipText     =   "Período legislativo y Tipo de Sesión"
         Top             =   3900
         Width           =   1815
      End
      Begin VB.CommandButton cmdTitulo 
         Caption         =   "Orden del Día"
         Height          =   345
         Left            =   5970
         TabIndex        =   35
         Top             =   1800
         Width           =   1245
      End
      Begin VB.CommandButton cmdPresidente 
         Caption         =   "&Presidente"
         Height          =   315
         Left            =   5970
         TabIndex        =   34
         ToolTipText     =   "Seleccionar presidente del total de legisladores"
         Top             =   2190
         Width           =   1245
      End
      Begin VB.CommandButton cmdTituloBlanco 
         Caption         =   "&Limpiar"
         Height          =   315
         Left            =   5970
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2910
         Width           =   1245
      End
      Begin VB.CommandButton cmdNuevaSesion 
         Caption         =   "&Nueva Sesión"
         Height          =   345
         Left            =   3390
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdCambiarSesion 
         Caption         =   "&Cambiar Sesión"
         Height          =   345
         Left            =   4620
         TabIndex        =   31
         Top             =   1800
         Width           =   1275
      End
      Begin VB.CommandButton cmdTituloRapido 
         Caption         =   "&Título"
         Height          =   315
         Left            =   5970
         TabIndex        =   30
         Top             =   2550
         Width           =   1245
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
         Left            =   630
         Picture         =   "frmConsolaOperacionAEBFranes.frx":22748
         TabIndex        =   29
         Top             =   3930
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CommandButton cmdCarteles 
         Caption         =   "E&ncender Carteles"
         Height          =   330
         Left            =   9810
         TabIndex        =   28
         ToolTipText     =   "Período legislativo y Tipo de Sesión"
         Top             =   3900
         Width           =   1620
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Tomar Control"
         Height          =   330
         Left            =   6630
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3900
         Width           =   1590
      End
      Begin VB.CommandButton cmdMantenimiento 
         Caption         =   "Man&tenimiento"
         Height          =   330
         Left            =   8310
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3900
         Width           =   1410
      End
      Begin VB.CommandButton cmdVotacion 
         Caption         =   "Votac&ión"
         Height          =   330
         Left            =   11730
         TabIndex        =   25
         Top             =   2160
         Width           =   3210
      End
      Begin VB.CommandButton cmdReconsiderar 
         Caption         =   "Reconsiderar"
         Height          =   330
         Left            =   11730
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Período legislativo y Tipo de Sesión"
         Top             =   3270
         Width           =   3210
      End
      Begin VB.CommandButton Salir 
         Caption         =   "&Salir"
         Height          =   330
         Left            =   11730
         TabIndex        =   23
         Top             =   3900
         Width           =   3210
      End
      Begin VB.CommandButton cmdAbstenciones 
         Caption         =   "Selector de &Abstenciones"
         Height          =   330
         Left            =   11730
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Período legislativo y Tipo de Sesión"
         Top             =   2910
         Width           =   3210
      End
      Begin VB.CommandButton cmdCancelarVotacion 
         Caption         =   "Cance&lar Votación"
         Height          =   330
         Left            =   11730
         TabIndex        =   21
         ToolTipText     =   "Cancelar Votación"
         Top             =   2520
         Width           =   3210
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
         Left            =   8235
         Locked          =   -1  'True
         TabIndex        =   20
         ToolTipText     =   "Resultado de la Votación"
         Top             =   840
         Width           =   3495
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
         Left            =   14220
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   " "
         ToolTipText     =   "Cantidad de Legisladores Pendientes de Emitir Voto"
         Top             =   495
         Width           =   435
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   " "
         ToolTipText     =   "Tiempo de Votación"
         Top             =   495
         Width           =   690
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
         Left            =   8250
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "Cantidad de Votos Afirmativos"
         Top             =   495
         Width           =   450
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
         Left            =   11280
         Locked          =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "Cantidad de Abstenciones"
         Top             =   495
         Width           =   450
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
         Left            =   3390
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   " "
         Top             =   840
         Width           =   1785
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
         Left            =   14220
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Legisladores ocupadando banca sin identificarse"
         Top             =   840
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
         Left            =   4785
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   " "
         Top             =   495
         Width           =   405
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
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   " "
         Top             =   495
         Width           =   435
      End
      Begin VB.TextBox txtHora 
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   840
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
         Left            =   9750
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Cantidad de votos Negativos"
         Top             =   495
         Width           =   450
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   495
         Width           =   975
      End
      Begin MSDataListLib.DataCombo dcTipoMayoria 
         Height          =   315
         Left            =   8580
         TabIndex        =   40
         Top             =   1800
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
         Left            =   8580
         TabIndex        =   41
         Top             =   2160
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
         Left            =   12660
         TabIndex        =   42
         Top             =   1770
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcAbstencion 
         Height          =   315
         Left            =   8580
         TabIndex        =   43
         Top             =   2880
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
         Left            =   8580
         TabIndex        =   44
         Top             =   2520
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
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
         Left            =   1500
         TabIndex        =   65
         Top             =   2190
         Width           =   4395
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   4
         Left            =   540
         Picture         =   "frmConsolaOperacionAEBFranes.frx":B6004
         Top             =   3450
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   3
         Left            =   540
         Picture         =   "frmConsolaOperacionAEBFranes.frx":B66EE
         Top             =   60
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   2
         Left            =   6630
         Picture         =   "frmConsolaOperacionAEBFranes.frx":B6DD8
         Top             =   3450
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   1
         Left            =   540
         Picture         =   "frmConsolaOperacionAEBFranes.frx":B74C2
         Top             =   1350
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   0
         Left            =   11760
         Picture         =   "frmConsolaOperacionAEBFranes.frx":B7BAC
         Top             =   1350
         Width           =   360
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
         Left            =   6780
         TabIndex        =   64
         Top             =   2595
         Width           =   1710
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
         Left            =   6780
         TabIndex        =   63
         Top             =   2940
         Width           =   1710
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
         Left            =   6750
         TabIndex        =   62
         Top             =   1890
         Width           =   1740
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
         Left            =   6750
         TabIndex        =   61
         Top             =   2235
         Width           =   1740
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
         Left            =   810
         TabIndex        =   60
         Top             =   2520
         Width           =   645
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
         Left            =   540
         TabIndex        =   59
         Top             =   2220
         Width           =   930
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Modalidad Información"
         Height          =   255
         Left            =   900
         TabIndex        =   58
         Top             =   3930
         Width           =   1875
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
         Left            =   11970
         TabIndex        =   57
         Top             =   540
         Width           =   2205
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
         Left            =   5565
         TabIndex        =   56
         Top             =   540
         Width           =   630
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
         Left            =   540
         TabIndex        =   55
         Top             =   540
         Width           =   540
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
         Left            =   645
         TabIndex        =   54
         Top             =   885
         Width           =   420
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
         Left            =   10290
         TabIndex        =   53
         Top             =   540
         Width           =   960
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
         Left            =   8835
         TabIndex        =   52
         Top             =   540
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
         Left            =   7275
         TabIndex        =   51
         Top             =   540
         Width           =   945
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
         Left            =   7320
         TabIndex        =   50
         Top             =   885
         Width           =   870
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
         Left            =   2685
         TabIndex        =   49
         Top             =   885
         Width           =   660
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
         Left            =   2505
         TabIndex        =   48
         Top             =   540
         Width           =   855
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
         Left            =   3945
         TabIndex        =   47
         Top             =   540
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
         Left            =   12765
         TabIndex        =   46
         Top             =   885
         Width           =   1410
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
         Left            =   11730
         TabIndex        =   45
         Top             =   1830
         Width           =   1665
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
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      Picture         =   "frmConsolaOperacionAEBFranes.frx":B8296
      ScaleHeight     =   780
      ScaleWidth      =   15105
      TabIndex        =   5
      Top             =   0
      Width           =   15105
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
         Left            =   420
         TabIndex        =   7
         Top             =   435
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   330
         TabIndex        =   6
         Top             =   210
         Width           =   14430
      End
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

Private datBanca(71) As tpBanca
Private datLista(71) As tpBanca
Private datBloque() As tpBloque
Private flgRepintarLista As Boolean
Private cntRepintarLista As Long

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
    cntRepintarLista = 8
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
        txtAusentes.Text = rstVector!Ausentes
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
    Dim bIdx As Long
    Dim strBloque   As String
    Dim C As Integer
    Dim col As Single
    Dim row As Single
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
    C = 0
    For i = 0 To UBound(mVectorColores)
        clave = i
        shpBanca(i).FillColor = mColores(mVectorColores(clave))
        ctrBanca(i).ForeColor = mColoresFuentes(mVectorColores(clave))
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
                    col = C \ Rows
                    row = C Mod Rows
                    Xp = Xo + (col * ColWidth)
                    Yp = Yo + (row * RowHeight)
                    strBloque = datBloque(bIdx).Nombre
                    lblBloque(bIdx).Caption = strBloque
                    lblBloque(bIdx).Width = ColWidth
                    lblBloque(bIdx).AutoSize = True
                    lblBloque(bIdx).Left = Xp
                    lblBloque(bIdx).Top = Yp
                    lblBloque(bIdx).Visible = True
                    C = C + 1
                    col = C \ Rows
                    row = C Mod Rows
                    Xp = Xo + (col * ColWidth)
                    Yp = Yo + (row * RowHeight)
                    lblBloqueInfo(bIdx).Width = ColWidth
                    lblBloqueInfo(bIdx).AutoSize = True
                    lblBloqueInfo(bIdx).Left = Xp
                    lblBloqueInfo(bIdx).Top = Yp
                    lblBloqueInfo(bIdx).Visible = True
                    lblBloqueInfo(bIdx).FontBold = False
                    lblBloqueInfo(bIdx).FontSize = 10
                    C = C + 1
                End If
                If (lblLegis(i).Caption <> "") And Not (flPresidenteLegislador And bcIdx = 0) Then
                    lblBanca(i).Tag = strBloque
                    lblBanca(i).Caption = datLista(i).BancaDefecto
                    If C Mod Rows = 0 Then
                        C = C + 1
                    End If
                    col = C \ Rows
                    row = C Mod Rows
                    Xp = Xo + (col * ColWidth)
                    Yp = Yo + (row * RowHeight)
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
                    C = C + 1
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
    flgRepintarLista = False
    If MR2.Visible Then
        For i = 0 To UBound(datBloque)
            lblBloque(i).Caption = datBloque(i).Nombre
            lblBloqueInfo(i).Caption = "Presentes " & datBloque(i).Presentes & "/" & datBloque(i).TotalLegisladores & ". Identificados: " & datBloque(i).Identificados
        Next
    End If
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
        If datBanca(Index).legislador <> "" Then
            info.MostrarDatos Index, datBanca(Index).legislador, PermitirPrueba, PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
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

Private Sub ctrBanca_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim NumeroBanca As Long
    Dim PermitirPrueba      As Boolean
    Dim PermitirIdentificar As Boolean
    Dim PermitirCambioVoto  As Boolean
    Dim PermitirAbstener    As Boolean
    Dim legislador As String
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
            TraerInfoLegislador datBanca(Index).legislador, legislador, Bloque, Departamento
            mnuID.Caption = legislador
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
    CargarVectorLegisladoresPorBloque
    flgRepintarLista = True
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


Private Sub imgA_Click(Index As Integer)
    ML(0).Visible = True
    ML(1).Visible = False
    MR1.Visible = True
    MR2.Visible = False
End Sub

Private Sub imgB_Click(Index As Integer)
    ML(0).Visible = False
    ML(1).Visible = True
    MR1.Visible = False
    MR2.Visible = True
    cntRepintarLista = 9
End Sub

Private Sub lblBanca_DblClick(Index As Integer)
    Dim PermitirPrueba      As Boolean
    Dim PermitirIdentificar As Boolean
    Dim PermitirCambioVoto  As Boolean
    Dim PermitirAbstener    As Boolean
    
    Index = Val(lblBanca(Index).Caption)
    If chkInformacion.Value = vbChecked Then
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

Private Sub lblBanca_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim NumeroBanca As Long
    Dim PermitirPrueba      As Boolean
    Dim PermitirIdentificar As Boolean
    Dim PermitirCambioVoto  As Boolean
    Dim PermitirAbstener    As Boolean
    Dim legislador As String
    Dim Bloque As String
    Dim Departamento As String
    
    If Button = 2 Then
        IndexBanca.Caption = lblBanca(Index).Caption
        If chkInformacion.Value = vbChecked Then
            NumeroBanca = Val(lblBanca(Index).Caption)
            EvaluarPermisosOperacionesBanca Val(lblBanca(Index).Caption), PermitirPrueba, PermitirIdentificar, PermitirCambioVoto, PermitirAbstener
            mnuAsignarID.Enabled = PermitirIdentificar
            mnuAbstener.Enabled = PermitirAbstener
            DeterminarLeyendaMenuAbstencion NumeroBanca
            Sep00.Visible = True
            TraerInfoLegislador datBanca(Index).legislador, legislador, Bloque, Departamento
            mnuID.Caption = legislador
            mnuBloque.Visible = (Bloque <> "")
            mnuBloque.Caption = Bloque
            mnuDepartamento.Visible = (Departamento <> "")
            mnuDepartamento.Caption = Departamento
            Me.PopupMenu mnuPopUP, , , , mnuID
        End If
    End If
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
            xObjeto = (Me.Width - ML(0).Width - MR1.Left) / 2
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


Private Sub TraerInfoLegislador(pIdLegislador As String, legislador As String, Optional Bloque As String, Optional Departamento As String, Optional Agrupacion As String)
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
                Bloque = Trim(rstAux!bloque_politico)
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
    Dim B As Integer
    Dim C As Integer
    Dim idxB As Integer
    Dim bolPresidente As Boolean
    sql = "SELECT Count(legisladores_activos.ID) AS TotalLegisladores, Legisladores.bloque_politico AS BLOQUE "
    sql = sql & "FROM legisladores_activos "
    sql = sql & "INNER JOIN Legisladores ON legisladores_activos.ID = Legisladores.id "
    sql = sql & "GROUP BY Legisladores.bloque_politico "
    SetearRs sql, rstB
    rstB.MoveLast
    BCont = rstB.RecordCount
    ReDim datBloque(BCont)
    BCont = 0
    rstB.MoveFirst
    C = 0
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
            datLista(C).BloqueIndex = BCont
            datLista(C).BloqueNombre = rstB("Bloque").Value & ""
            datLista(C).LegisladorDefecto = rstAux("ID").Value & ""
            datLista(C).LegisladorDefectoNombre = UCase(rstAux("APELLIDO").Value) & ", " & rstAux("NOMBRE").Value
            datLista(C).BancaDefecto = rstAux!Banca
            rstAux.MoveNext
            C = C + 1
        Loop
        datBloque(BCont).Nombre = rstB("Bloque").Value
        datBloque(BCont).TotalLegisladores = rstB("TotalLegisladores").Value + IIf(bolPresidente, -1, 0)
        bolPresidente = False
        BCont = BCont + 1
        rstB.MoveNext
    Loop
End Sub


