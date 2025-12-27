Attribute VB_Name = "VL"
Public BackupVectorPresencia(0 To 256) As String
Public BackupVectorIdentificacion(0 To 256) As String
Public BackupVectorResultados(0 To 256) As String
Public BancasEnCuenta(0 To 256) As Boolean
Public PresenciaReal(0 To 256) As String
Public PerdioIdentificacion(0 To 256) As Boolean
Public modoExtendido As Boolean

Public Sub guardarEstado()
Dim i As Integer
Call limpiaEstado
For i = 0 To 256
    If i = 0 Or EstadoActual.VectorPresencia(i) = PRESENTE Then
        BancasEnCuenta(i) = True
    End If
    BackupVectorPresencia(i) = EstadoActual.VectorPresencia(i)
    BackupVectorIdentificacion(i) = EstadoActual.VectorIdentificacion(i)
    BackupVectorResultados(i) = EstadoActual.VectorResultados(i)
    PresenciaReal(i) = EstadoActual.VectorPresencia(i)
Next i
End Sub

Private Sub limpiaEstado()
Dim i As Integer
For i = 0 To 256
    BancasEnCuenta(i) = False
    BackupVectorPresencia(i) = AUSENTE
    BackupVectorIdentificacion(i) = NO_IDENTIFICADO
    BackupVectorResultados(i) = ABSTENCION
    PresenciaReal(i) = AUSENTE
    PerdioIdentificacion(i) = False
Next i
End Sub

Public Function bancaValida(banca As Integer) As Boolean
Dim ret As Boolean
ret = False
If Not (modoExtendido = True And BancasEnCuenta(banca) = True Or (modoExtendido = True And BancasEnCuenta(banca) = False And EstadoActual.VectorPresencia(banca) = AUSENTE)) Then
    ret = True
End If
bancaValida = ret
End Function

Public Function votosPendientes() As Boolean
Dim res As Boolean
Dim i As Integer
res = False
For i = 1 To 256
    If (EstadoActual.VectorIdentificacion(i) <> NO_IDENTIFICADO And EstadoActual.VectorResultados(i) <> AFIRMATIVO And EstadoActual.VectorResultados(i) <> NEGATIVO) Then
        'Si esta identificado y el voto es distinto de afirmativo y negativo (abstenido)
        res = True
        i = 256
    End If
Next i
votosPendientes = res
End Function

Public Sub log(text As String)
On Error Resume Next
Dim line As String
line = Now() & " - " & text
Open App.Path & "\log_cierre.txt" For Append As #1
Print #1, line
Close #1
End Sub
