VERSION 5.00
Begin VB.Form frmMensajesCartel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cartel"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdApagar 
      Caption         =   "Apagar "
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdControlSqv 
      Caption         =   "Control Sqv"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdGeneral 
      Caption         =   "General"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMensajesCartel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdApagar_Click()
    strCartelSQV = "0"
    Unload Me
End Sub

Private Sub cmdControlSqv_Click()
    strCartelSQV = "1"
    Unload Me
End Sub

Private Sub cmdGeneral_Click()
    strCartelSQV = "2"
    Unload Me
End Sub

Private Sub Command1_Click()
    strCartelSQV = ""
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

