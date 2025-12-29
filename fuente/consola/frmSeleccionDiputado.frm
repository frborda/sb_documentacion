VERSION 5.00
Begin VB.Form frmSeleccionDiputado 
   BackColor       =   &H00404040&
   Caption         =   "Seleccione un diputado"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFiltro 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   4020
      Width           =   5115
   End
   Begin VB.ListBox lstDiputados 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   5115
   End
End
Attribute VB_Name = "frmSeleccionDiputado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ids() As String
Public mPeriodo As String
Public mSesion As String
Public mActa As String
Public mVersion As String
Public Result As Integer
Private Sub Form_Load()
Dim s As String
Me.Result = 0
s = "SELECT id, apellido + ', ' + nombre AS diputado " & _
"From Legisladores WHERE (id IN (SELECT     Legislador_asignado From detalleactas " & _
"WHERE      (Período_Legislativo = '" & mPeriodo & "') AND (Sesión = " & mSesion & ") AND (Nro_de_Acta = " & mActa & ") " & _
"AND (Versión_Acta = " & mVersion & "))) ORDER BY diputado"
Dim rs As New ADODB.Recordset
SetearRs s, rs
If (rs.EOF) Then
    MsgBox "No se pudieron obtener los diputados de la sesión"
    Unload Me
End If
While Not rs.EOF
    lstDiputados.AddItem rs.Fields("diputado")
    ReDim Preserve ids(0 To (lstDiputados.ListCount - 1))
    ids(lstDiputados.ListCount - 1) = rs.Fields("id")
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
End Sub

Private Sub lstDiputados_DblClick()
SubmitDiputado
End Sub

Private Sub txtFiltro_Change()
Dim mItem As Object
Dim currentItem As String
Dim currentFilter As String
Dim currentLength As Integer
Dim piece1 As String
Dim piece2 As String
currentFilter = UCase(txtFiltro.Text)
currentLength = Len(txtFiltro.Text)
piece2 = Left(currentFilter, currentLength)
Dim r As Long
For i = 0 To (lstDiputados.ListCount - 1)
    currentItem = UCase(lstDiputados.List(i))
    piece1 = Left(currentItem, currentLength)
    If (piece1 = piece2) Then
        lstDiputados.ListIndex = i
        Exit For
    End If
Next i
End Sub

Private Sub SubmitDiputado()
If (lstDiputados.ListIndex > -1) Then
    Result = ids(lstDiputados.ListIndex)
    Unload Me
Else
    MsgBox "No ha seleccionado ningún diputado"
End If
End Sub

Private Sub txtFiltro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SubmitDiputado
End If
End Sub
