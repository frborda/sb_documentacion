VERSION 5.00
Begin VB.Form frmControlSistema 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control de Sistema"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver a Consola"
      Height          =   465
      Left            =   150
      TabIndex        =   2
      Top             =   780
      Width           =   3105
   End
   Begin VB.CommandButton cmdIdentificacionPorTeclado 
      Caption         =   "Identificación General por Teclado"
      Height          =   465
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   3105
   End
   Begin VB.Label Label2 
      Caption         =   "La modalidad de indentificación de legisladores en el recinto pasa a ser a traves del escaner dactilar de la unidad de banca."
      Height          =   465
      Left            =   3420
      TabIndex        =   1
      Top             =   180
      Width           =   4485
   End
End
Attribute VB_Name = "frmControlSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdIdentificacionPorTeclado_Click()
    MensajesSQV.IdentificacionPorTeclado "brc"
End Sub

Private Sub cmdVolver_Click()
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

