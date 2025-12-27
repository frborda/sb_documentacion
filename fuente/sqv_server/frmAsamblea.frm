VERSION 5.00
Begin VB.Form frmAsamblea 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Asamblea"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2460
      Top             =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   7860
      X2              =   7860
      Y1              =   420
      Y2              =   1140
   End
   Begin VB.Label lblQuorum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ASAMBLEA   LEGISLATIVA"
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
      Left            =   1860
      TabIndex        =   2
      Top             =   360
      Width           =   12195
   End
   Begin VB.Shape shpQuorum 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1005
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   270
      Width           =   14085
   End
   Begin VB.Shape shpHora 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1005
      Left            =   3930
      Shape           =   4  'Rounded Rectangle
      Top             =   1380
      Width           =   7185
   End
   Begin VB.Label lblHora 
      BackStyle       =   0  'Transparent
      Caption         =   "11:20"
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
      Left            =   9060
      TabIndex        =   1
      Top             =   1470
      Width           =   1905
   End
   Begin VB.Label lblFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "21/03/2015"
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
      Left            =   4050
      TabIndex        =   0
      Top             =   1470
      Width           =   4365
   End
End
Attribute VB_Name = "frmAsamblea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 76 Or KeyAscii = 108) Then
        Me.Hide
        frmCartel2011.Show
    End If
End Sub

Private Sub Form_Load()
Me.lblFecha = Format(Now, "dd/mm/yyyy")
Me.lblHora = frmCartel2011.lblHora
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Me.lblFecha = Format(Now, "dd/mm/yyyy")
Me.lblHora = frmCartel2011.lblHora
End Sub
