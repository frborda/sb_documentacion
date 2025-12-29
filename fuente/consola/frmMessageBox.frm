VERSION 5.00
Begin VB.Form frmMessageBox 
   BorderStyle     =   0  'None
   Caption         =   "HCDN"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   2715
      Left            =   0
      Top             =   0
      Width           =   3795
   End
   Begin VB.Label lblDescripcion 
      Caption         =   "Descripción del error"
      Height          =   2055
      Left            =   60
      TabIndex        =   1
      Top             =   240
      Width           =   3675
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3795
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Desc As String
Public Titulo As String
Private Sub cmdAceptar_Click()
Unload Me
End Sub
Private Sub Form_Load()
lblTitulo.Caption = Titulo
lblDescripcion.Caption = Desc
End Sub
