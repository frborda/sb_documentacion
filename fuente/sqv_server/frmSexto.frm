VERSION 5.00
Begin VB.Form frmSexto 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmUpdate 
      Interval        =   500
      Left            =   1680
      Top             =   4260
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
      Left            =   360
      TabIndex        =   10
      Top             =   2730
      Width           =   1065
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
      Left            =   1470
      TabIndex        =   9
      Top             =   2700
      Visible         =   0   'False
      Width           =   345
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
      Left            =   1950
      TabIndex        =   8
      Top             =   2730
      Width           =   2025
   End
   Begin VB.Shape shpHora 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1005
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7185
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
      Left            =   5400
      TabIndex        =   7
      Top             =   180
      Width           =   1905
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
      Left            =   300
      TabIndex        =   6
      Top             =   1350
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
      Left            =   300
      TabIndex        =   5
      Top             =   2100
      Width           =   1065
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
      Left            =   1410
      TabIndex        =   4
      Top             =   1320
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
      Left            =   1410
      TabIndex        =   3
      Top             =   2070
      Visible         =   0   'False
      Width           =   345
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
      Left            =   1890
      TabIndex        =   2
      Top             =   1350
      Width           =   6075
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
      Left            =   1890
      TabIndex        =   1
      Top             =   2100
      Width           =   5985
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
      Left            =   360
      TabIndex        =   0
      Top             =   180
      Width           =   2865
   End
End
Attribute VB_Name = "frmSexto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Dim r As Integer
        r = MsgBox("¿Desea cerrar el servidor?", vbYesNo)
        If (r = vbYes) Then
            End
        End If
    End If
End Sub

Private Sub Form_Load()
Me.lblFecha.Caption = frmCartel2011.lblFecha.Caption
Me.lblHora.Caption = frmCartel2011.lblHora.Caption
Me.lblLeyendaReunion.Caption = frmCartel2011.lblLeyendaReunion.Caption
Me.lblNumeroPeriodo.Caption = frmCartel2011.lblNumeroPeriodo.Caption
Me.lblNumeroReunion.Caption = frmCartel2011.lblNumeroReunion.Caption
Me.lblNumeroSesion.Caption = frmCartel2011.lblNumeroSesion.Caption
Me.lblTipoPeriodo.Caption = frmCartel2011.lblTipoPeriodo.Caption
Me.lblTipoSesion.Caption = frmCartel2011.lblTipoSesion.Caption
End Sub

Private Sub tmUpdate_Timer()
Me.lblFecha.Caption = frmCartel2011.lblFecha.Caption
Me.lblHora.Caption = frmCartel2011.lblHora.Caption
Me.lblLeyendaReunion.Caption = frmCartel2011.lblLeyendaReunion.Caption
Me.lblNumeroPeriodo.Caption = frmCartel2011.lblNumeroPeriodo.Caption
Me.lblNumeroReunion.Caption = frmCartel2011.lblNumeroReunion.Caption
Me.lblNumeroSesion.Caption = frmCartel2011.lblNumeroSesion.Caption
Me.lblTipoPeriodo.Caption = frmCartel2011.lblTipoPeriodo.Caption
Me.lblTipoSesion.Caption = frmCartel2011.lblTipoSesion.Caption
Me.lblNumeroSesion.Visible = frmCartel2011.lblNumeroSesion.Visible
End Sub
