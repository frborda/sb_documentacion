VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReportesExportar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar Reporte"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportesExportar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   3330
      TabIndex        =   77
      Top             =   2610
      Width           =   1125
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   2130
      TabIndex        =   76
      Top             =   2610
      Width           =   1125
   End
   Begin VB.PictureBox picOptions 
      Height          =   855
      Index           =   2
      Left            =   210
      ScaleHeight     =   795
      ScaleWidth      =   4155
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4215
      Begin VB.CommandButton btnBrowse 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   74
         Top             =   375
         Width           =   375
      End
      Begin VB.CommandButton btnHTMLFolder 
         Caption         =   "..."
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   855
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkSinglePage 
         Caption         =   "Single Page Output"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   3360
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CheckBox chkMHTArchive 
         Caption         =   "Create MIME Archive"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3360
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CheckBox chkCreateFrameset 
         Caption         =   "Create Frames"
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   3120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CheckBox chkCreateCSS 
         Caption         =   "Create CSS"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox cboTOCFormat 
         Height          =   315
         ItemData        =   "frmReportesExportar.frx":058A
         Left            =   1440
         List            =   "frmReportesExportar.frx":059A
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2640
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox cboHTMLVersion 
         Height          =   315
         ItemData        =   "frmReportesExportar.frx":05D4
         Left            =   1440
         List            =   "frmReportesExportar.frx":05DE
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2280
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtHTMLTitle 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   1560
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtHTMLCharset 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   1920
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton btnAuxFolder 
         Caption         =   "..."
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   1230
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAuxFolder 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   1215
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtHTMLFolder 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtFilename 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblTag 
         Caption         =   "Nombre de Archivo:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   75
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblTag 
         Caption         =   "TOC Format:"
         Height          =   315
         Index           =   13
         Left            =   120
         TabIndex        =   57
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTag 
         Caption         =   "HTML Format:"
         Height          =   315
         Index           =   12
         Left            =   120
         TabIndex        =   56
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTag 
         Caption         =   "Title:"
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   55
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTag 
         Caption         =   "Charset:"
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   54
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTag 
         Caption         =   "Images Folder:"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   53
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTag 
         Caption         =   "HTML Folder:"
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   52
         Top             =   825
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblTag 
         Caption         =   "File Prefix:"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   51
         Top             =   585
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3660
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   77
      ImageHeight     =   93
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportesExportar.frx":0602
            Key             =   "pic0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportesExportar.frx":0E8B
            Key             =   "pic2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportesExportar.frx":1712
            Key             =   "pic3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportesExportar.frx":1F0C
            Key             =   "pic1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportesExportar.frx":2606
            Key             =   "pic4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportesExportar.frx":2CEF
            Key             =   "pic5"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   5160
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboExportFormat 
      Height          =   315
      ItemData        =   "frmReportesExportar.frx":354F
      Left            =   180
      List            =   "frmReportesExportar.frx":3565
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   4215
   End
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   4500
      TabIndex        =   43
      Top             =   0
      Width           =   4500
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Exportar Reporte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label lblSubtitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione el formato del archivo a exportar"
         Height          =   375
         Left            =   360
         TabIndex        =   45
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.PictureBox picOptions 
      Height          =   855
      Index           =   1
      Left            =   225
      ScaleHeight     =   795
      ScaleWidth      =   4155
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4215
      Begin VB.ListBox lstPDFFonts 
         Appearance      =   0  'Flat
         Height          =   1740
         IntegralHeight  =   0   'False
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.ComboBox cboPDFJPGQuality 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
         Height          =   315
         ItemData        =   "frmReportesExportar.frx":3618
         Left            =   1800
         List            =   "frmReportesExportar.frx":3628
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1170
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox cboAcrobatVersion 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmReportesExportar.frx":363D
         Left            =   1800
         List            =   "frmReportesExportar.frx":364A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton btnBrowse 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   6
         Top             =   375
         Width           =   375
      End
      Begin VB.TextBox txtFilename 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblTag 
         Caption         =   "No Embedding Fonts:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   50
         Top             =   1560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblTag 
         Caption         =   "JPG Quality:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   49
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblTag 
         Caption         =   "Acrobat Version:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   48
         Top             =   750
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblTag 
         Caption         =   "Nombre de Archivo:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox picOptions 
      Height          =   855
      Index           =   0
      Left            =   225
      ScaleHeight     =   795
      ScaleWidth      =   4155
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4215
      Begin VB.CommandButton btnBrowse 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   3
         Top             =   375
         Width           =   375
      End
      Begin VB.TextBox txtFilename 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblTag 
         Caption         =   "Nombre de Archivo:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.PictureBox picOptions 
      Height          =   855
      Index           =   5
      Left            =   195
      ScaleHeight     =   795
      ScaleWidth      =   4155
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4215
      Begin VB.CheckBox chkTXTUnicode 
         Caption         =   "Unicode"
         Height          =   255
         Left            =   1560
         TabIndex        =   71
         Top             =   1320
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox chkTXTSupressEmptyLines 
         Caption         =   "Supress Empty Lines"
         Height          =   255
         Left            =   1560
         TabIndex        =   70
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtTXTTextDelimiter 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   69
         Text            =   ","
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtFilename 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton btnBrowse 
         Caption         =   "..."
         Height          =   255
         Index           =   5
         Left            =   3720
         TabIndex        =   42
         Top             =   375
         Width           =   375
      End
      Begin VB.Label lblTag 
         Caption         =   "Text Delimiter:"
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   68
         Top             =   735
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblTag 
         Caption         =   "Nombre de Archivo:"
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   67
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.PictureBox picOptions 
      Height          =   855
      Index           =   4
      Left            =   195
      ScaleHeight     =   795
      ScaleWidth      =   4155
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4215
      Begin VB.CommandButton btnBrowse 
         Caption         =   "..."
         Height          =   255
         Index           =   4
         Left            =   3720
         TabIndex        =   39
         Top             =   375
         Width           =   375
      End
      Begin VB.TextBox txtFilename 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblTag 
         Caption         =   "Nombre de Archivo:"
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   66
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.PictureBox picOptions 
      Height          =   855
      Index           =   3
      Left            =   195
      ScaleHeight     =   795
      ScaleWidth      =   4155
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4215
      Begin VB.CheckBox chkXLSTrimEmptySpace 
         Caption         =   "Trim Empty Space"
         Height          =   255
         Left            =   1560
         TabIndex        =   36
         Top             =   3360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox chkXLSShowMarginSpace 
         Caption         =   "Show Margin Space"
         Height          =   255
         Left            =   1560
         TabIndex        =   35
         Top             =   3120
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox chkXLSMultisheet 
         Caption         =   "Generate Multiple Sheets"
         Height          =   255
         Left            =   1560
         TabIndex        =   34
         Top             =   2880
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox chkXLSGenPageBreaks 
         Caption         =   "Generate Page Breaks"
         Height          =   255
         Left            =   1560
         TabIndex        =   33
         Top             =   2640
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox chkXLSDoubleBoundaries 
         Caption         =   "Double Boundaries"
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   2400
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox chkXLSAutoRowHeight 
         Caption         =   "Auto Row Height"
         Height          =   255
         Left            =   1560
         TabIndex        =   31
         Top             =   2160
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtXLSMinRowHeight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Text            =   "128"
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtXLSMinColWidth 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Text            =   "1011"
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtXLSBorderSpace 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   28
         Text            =   "59"
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cboXLSVersion 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmReportesExportar.frx":3681
         Left            =   1560
         List            =   "frmReportesExportar.frx":3697
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton btnBrowse 
         Caption         =   "..."
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   26
         Top             =   375
         Width           =   375
      End
      Begin VB.TextBox txtFilename 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblTag 
         Caption         =   "Twips"
         Height          =   240
         Index           =   21
         Left            =   2640
         TabIndex        =   65
         Top             =   1815
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblTag 
         Caption         =   "Min. row height:"
         Height          =   240
         Index           =   20
         Left            =   120
         TabIndex        =   64
         Top             =   1815
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTag 
         Caption         =   "Twips"
         Height          =   240
         Index           =   19
         Left            =   2640
         TabIndex        =   63
         Top             =   1455
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblTag 
         Caption         =   "Min. col. width:"
         Height          =   240
         Index           =   18
         Left            =   120
         TabIndex        =   62
         Top             =   1455
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTag 
         Caption         =   "Twips"
         Height          =   240
         Index           =   17
         Left            =   2640
         TabIndex        =   61
         Top             =   1095
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblTag 
         Caption         =   "Border Space:"
         Height          =   240
         Index           =   16
         Left            =   120
         TabIndex        =   60
         Top             =   1095
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTag 
         Caption         =   "Version:"
         Height          =   240
         Index           =   15
         Left            =   120
         TabIndex        =   59
         Top             =   757
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblTag 
         Caption         =   "Nombre de Archivo:"
         Height          =   315
         Index           =   14
         Left            =   120
         TabIndex        =   58
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6960
      Y1              =   865
      Y2              =   865
   End
   Begin VB.Label lblDesc 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Adentro de los pictures hay mas opciones ocultas"
      Height          =   495
      Left            =   -1260
      TabIndex        =   72
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -1260
      X2              =   4380
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblTag 
      Caption         =   "Formatos:"
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   44
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmReportesExportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Pages As DDActiveReports2.Pages

Private Sub btnAuxFolder_Click()
Dim sFOlder As String
    On Error GoTo ehHTMLFolder

    If Not hlpBrowseForFolder(Me.hwnd, "Select the HTML images folder", sFOlder) Then Exit Sub
    txtAuxFolder = sFOlder
    Exit Sub
    
ehHTMLFolder:
    MsgBox Str(err.Number) & " - " & err.Description, vbOKOnly, "Images Folder Select Error"
End Sub

Private Sub btnBrowse_Click(Index As Integer)
    On Error GoTo ehBrowse
    
    Select Case Index
    Case 0  ' RTF
        dlg.Filter = "Formato RTF (*.RTF)|*.rtf"
    Case 1  ' PDF
        dlg.Filter = "Formato Acrobat (*.pdf)|*.pdf"
    Case 2  ' HTML
        dlg.Filter = "Documento de Internet (*.htm)|*.htm"
    Case 3  ' Excel
        dlg.Filter = "Documento Microsoft Excel (*.xls)|*.xls"
    Case 4  ' TIF
        dlg.Filter = "Formato TIF (*.tif)|*.tif"
    Case 5  ' Text
        dlg.Filter = "Archivos de Texto (*.txt)|*.txt"
    End Select

    dlg.ShowSave

    If dlg.FileName <> "" Then txtFilename(Index).Text = dlg.FileName
        
    Exit Sub
ehBrowse:
    MsgBox Str(err.Number) & " - " & err.Description, vbOKOnly, "Error browsing for filename"
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnHTMLFolder_Click()
Dim sFOlder As String
    On Error GoTo ehHTMLFolder
    If Not hlpBrowseForFolder(Me.hwnd, "Select the HTML pages folder", sFOlder) Then Exit Sub
    txtHTMLFolder.Text = sFOlder
    Exit Sub
    
ehHTMLFolder:
    MsgBox Str(err.Number) & " - " & err.Description, vbOKOnly, "HTML Folder Select Error"
End Sub

Private Sub btnOK_Click()
    On Error GoTo ehExport

    If Pages Is Nothing Then Exit Sub
    If cboExportFormat.ListIndex < 0 Then Exit Sub
    If txtFilename(cboExportFormat.ListIndex) = "" Then
        MsgBox "Especifique un nombre al archivo.", vbInformation
        Exit Sub
    End If
    
    Select Case cboExportFormat.ListIndex
    Case 0  ' RTF
        Dim rtf As ARExportRTF
        Set rtf = New ARExportRTF
        rtf.FileName = txtFilename(cboExportFormat.ListIndex).Text
        rtf.Export Pages

        Set rtf = Nothing
    Case 1  ' PDF
        Dim lFont As Long
        Dim sFonts As String
        sFonts = ""
        
        Dim pdf As ARExportPDF
        Set pdf = New ARExportPDF
        
        pdf.FileName = txtFilename(cboExportFormat.ListIndex).Text
        pdf.AcrobatVersion = cboAcrobatVersion.ListIndex
        
        Select Case cboPDFJPGQuality.ListIndex
        Case 0: pdf.JPGQuality = 25
        Case 1: pdf.JPGQuality = 50
        Case 2: pdf.JPGQuality = 75
        Case 3: pdf.JPGQuality = 100    ' No Compression
        End Select
        
        ' Create a semi-color delimited string of the fonts
        ' that shouldn't be embedded in the PDF file
        For lFont = 0 To lstPDFFonts.ListCount - 1
            If lstPDFFonts.Selected(lFont) Then
                sFonts = sFonts & lstPDFFonts.List(lFont) & ";"
            End If
        Next
        pdf.SemiDelimitedNeverEmbedFonts = sFonts
        pdf.Export Pages
        Set pdf = Nothing
    Case 2  ' HTML
        Dim html As HTMLexport
        Dim Posicion As Integer
        
        Set html = New HTMLexport
        'Obtengo el prefijo
        Posicion = InStrRev(txtFilename(cboExportFormat.ListIndex).Text, "\", , vbTextCompare)
        
        
        html.FileNamePrefix = txtFilename(cboExportFormat.ListIndex).Text
        html.FileNamePrefix = Right(txtFilename(cboExportFormat.ListIndex).Text, Len(txtFilename(cboExportFormat.ListIndex).Text) - Posicion)
        txtHTMLFolder.Text = Left(txtFilename(cboExportFormat.ListIndex).Text, Posicion)
        txtAuxFolder.Text = Left(txtFilename(cboExportFormat.ListIndex).Text, Posicion)
        'Seteo directorio
        html.HTMLOutputPath = txtHTMLFolder.Text
        html.AuxOutputPath = txtAuxFolder.Text
        
        If txtHTMLCharset.Text = "" Then
            ' Set default value
        Else
            html.CharacterSet = txtHTMLCharset.Text
        End If
        
        html.Title = txtHTMLTitle.Text
        html.HTMLVersion = cboHTMLVersion.ListIndex
        html.TableOfContents = cboTOCFormat.ListIndex
        html.CreateCSSFile = (chkCreateCSS.Value = 1)
        html.CreateFramesetPage = (chkCreateFrameset.Value = 1)
        html.MHTOutput = (chkMHTArchive.Value = 1)
        html.MultiPageOutput = (chkSinglePage.Value = 0)
        html.Export Pages
        Set html = Nothing
    Case 3  ' XLS
        Dim xls As ARExportExcel
        Set xls = New ARExportExcel
        xls.FileName = txtFilename(3).Text
        Select Case cboXLSVersion.ListIndex
        Case 0: xls.Version = 2
        Case 1: xls.Version = 3
        Case 2: xls.Version = 4
        Case 3: xls.Version = 5
        Case 4: xls.Version = 7
        Case 5: xls.Version = 8
        End Select
        xls.AutoRowHeight = (chkXLSAutoRowHeight.Value = 1)
        xls.BorderSpace = Val(txtXLSBorderSpace.Text)
        xls.DoubleBoundaries = (chkXLSDoubleBoundaries.Value = 1)
        xls.GenPagebreaks = (chkXLSGenPageBreaks.Value = 1)
        xls.MinColumnWidth = Val(txtXLSMinColWidth.Text)
        xls.MinRowHeight = Val(txtXLSMinRowHeight.Text)
        xls.MultiSheet = (chkXLSMultisheet.Value = 1)
        xls.ShowMarginSpace = (chkXLSShowMarginSpace.Value = 1)
        xls.TrimEmptySpace = (chkXLSTrimEmptySpace.Value = 1)
        xls.Export Pages
        Set xls = Nothing
    Case 4  ' TIF
        Dim tif As TIFFExport
        Set tif = New TIFFExport
        
        tif.FileName = txtFilename(4).Text
        tif.Export Pages
        Set tif = Nothing
    Case 5  ' TXT
        Dim Txt As ARExportText
        Set Txt = New ARExportText
        Txt.FileName = txtFilename(5).Text
        Txt.TextDelimiter = txtTXTTextDelimiter.Text
        Txt.SuppressEmptyLines = (chkTXTSupressEmptyLines.Value = 1)
        Txt.Unicode = (chkTXTUnicode.Value = 1)
        Txt.Export Pages
        Set Txt = Nothing
    End Select
    Unload Me
    Exit Sub
ehExport:
    MsgBox Str(err.Number) & " - " & err.Description, vbOKOnly, "Error Exporting Document"
End Sub

Private Sub cboExportFormat_Click()
    On Error GoTo ehExportFormatClick
    If cboExportFormat.ListIndex < 0 Then Exit Sub
    
    Dim i As Integer
    For i = 0 To 5
        picOptions(i).Left = -10000
        picOptions(i).Visible = False
    Next
    picOptions(cboExportFormat.ListIndex).Visible = True
    picOptions(cboExportFormat.ListIndex).Left = 195
    'Inserto el picture del formato
'    picFormato.Picture = ImageList1.ListImages("pic" & cboExportFormat.ListIndex).Picture
    
    Exit Sub
ehExportFormatClick:
    MsgBox Str(err.Number) & " - " & err.Description, vbOKOnly, "Error ExportFormat_Click"
End Sub


Private Sub Form_Load()
Dim lFont As Long
    
    ' Load PDF Fonts
    ' This routine can be optimized by using API functions to
    ' enumerate fonts
    For lFont = 1 To Screen.FontCount
        lstPDFFonts.AddItem Screen.Fonts(lFont)
    Next
    
    cboAcrobatVersion.ListIndex = 1 ' 3.x
    cboHTMLVersion.ListIndex = 1    ' DHTML
    cboPDFJPGQuality.ListIndex = 2  ' 75%
    cboTOCFormat.ListIndex = 0  ' None
    cboXLSVersion.ListIndex = 5 ' 8.x
    

    
End Sub

Public Function ExportReport(ByVal Report As Object) As Boolean
    'Corro el reporte sin mostrarlo
    Report.Run False
    'Seteo las paginas del reporte
    Set Pages = Report.Pages
    'Pongo el nombre al reporte HTML
    txtHTMLTitle.Text = Report.Caption
    'Selecciono el primer formato
    cboExportFormat.ListIndex = 0
    Me.Show vbModal
End Function


