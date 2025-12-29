VERSION 5.00
Begin VB.Form frmFlotante 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Información Rápida"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6255
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   2115
   End
End
Attribute VB_Name = "frmFlotante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Visible = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Visible = False
End Sub
