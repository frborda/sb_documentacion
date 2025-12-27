VERSION 5.00
Begin VB.Form frmNegro 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmNegro"
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
Me.Width = frmMain.Width
Me.Height = frmMain.Height
Me.Left = frmMain.Left
Me.top = frmMain.top
End Sub
