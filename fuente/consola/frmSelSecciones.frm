VERSION 5.00
Begin VB.Form frmSelSecciones 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmSelSecciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql   As String
Dim strOrder As String

Private Sub CargarGrilla()
    Dim RsTemp As ADODB.Recordset
    Dim xFila  As Long
    Set RsTemp = New ADODB.Recordset
    SetearRs strSql + strOrder, RsTemp
    xFila = 1
    vsGrilla.Rows = 1
    With RsTemp
        If .RecordCount > 0 Then
            .MoveFirst
            vsGrilla.Rows = .RecordCount + 1
            While Not .EOF
                vsGrilla.TextMatrix(xFila, 0) = .Fields(0).Value
                xFila = xFila + 1
                .MoveNext
            Wend
        End If
        .Close
    End With
    Set RsTemp = Nothing
End Sub
Private Sub Ok()
    frmABMSecciones.lblid.Caption = vsGrilla.TextMatrix(vsGrilla.row, 0)
     
    Unload Me
End Sub
Private Sub Cancelar()
    frmABMSecciones.lblid.Caption = "nothing"
    Unload Me
End Sub
Private Sub Command1_Click()
    Call Ok
End Sub
Private Sub Command2_Click()
    Call Cancelar
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    strSql = "SELECT * FROM secciones "
    strOrder = "ORDER BY seccion"
    Call CargarGrilla
End Sub
Private Sub vsGrilla_DblClick()
    Call Ok
End Sub
Private Sub vsGrilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Ok
    End If
End Sub

