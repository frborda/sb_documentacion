Attribute VB_Name = "Pendientes"
Type DiputadoPendiente
    diputado(0 To 27) As String
End Type
Public paginaActualPendientes As Integer

Public Function getPaginasPendientes() As DiputadoPendiente()
'10 pendientes nada
Dim z As Integer
Dim Init As Integer
Dim strIn As String
Dim diputados() As DiputadoPendiente
strIn = ""
Init = 1
If EstadoActual.PresidenteHabilitadoParaVotar = True Then
    Init = 0
End If
For z = Init To 256
    If EstadoActual.VectorIdentificacion(z) <> NO_IDENTIFICADO And EstadoActual.VectorResultados(z) <> AFIRMATIVO And EstadoActual.VectorResultados(z) <> NEGATIVO Then
        If (strIn = "") Then
            strIn = EstadoActual.VectorIdentificacion(z)
        Else
            strIn = strIn & "," & EstadoActual.VectorIdentificacion(z)
        End If
    End If
Next z
If strIn = "" Then
    strIn = "-1"
End If
Dim rs As New ADODB.Recordset
Dim current As Integer
Dim soFar As Integer
Dim currentIndex As Integer
currentIndex = -1
soFar = 28
current = -1
frmMain.SetearRsAux "SELECT apellido + ', ' + nombre FROM legisladores_activos WHERE id IN (" & strIn & ") ORDER BY apellido, nombre", rs
While Not rs.EOF
    If soFar = 28 Then
        soFar = 0
        current = -1
        currentIndex = currentIndex + 1
        ReDim Preserve diputados(0 To currentIndex)
    End If
    current = current + 1
    soFar = soFar + 1
    diputados(currentIndex).diputado(current) = rs.Fields(0)
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
If (currentIndex <= -1) Then
    ReDim diputados(0 To 0)
    currentIndex = 0
End If
'Lleno los que faltan con vacio
While (current < 27)
    current = current + 1
    diputados(currentIndex).diputado(current) = ""
Wend
'Devuelvo
getPaginasPendientes = diputados
End Function
