Attribute VB_Name = "VotoRemoto"
Public ExtraQuorum() As String
Private VRDiputados(256) As VRDiputado
Private VRPresentes(256) As VRDiputado
Public Loaded As Boolean
Public DatabaseOpened As Boolean

Private Type VRDiputado
    id As Integer
    apellido As String
    nombre As String
    cuil As String
End Type

Public Function getPresentes() As Integer
Call CheckInit
Dim i As Integer
Dim totalPresentes As Integer
totalPresentes = 0
If (Not tienePresentes) Then
    getPresentes = 0
    Exit Function
End If
For i = 0 To UBound(VRPresentes)
    If (VRPresentes(i).id <> 0) Then
        totalPresentes = totalPresentes + 1
    End If
Next i
getPresentes = totalPresentes
End Function

Public Function identificar(cuil As String) As Boolean
Call CheckInit
Dim i As Integer
Dim vFound As VRDiputado
Dim firstEmptyIndex As Integer
firstEmptyIndex = -1
'Busco diputado
For i = 0 To UBound(VRDiputados)
    If (VRDiputados(i).cuil = cuil) Then
        vFound = VRDiputados(i)
        Exit For
    End If
Next i
If (vFound.id = 0) Then
    identificar = False
    Exit Function
End If
'Diputado encontrado, recorro presentes extra a ver si ya está identificado
For i = 0 To UBound(VRPresentes)
    If (firstEmptyIndex = -1) Then
        If (VRPresentes(i).id = 0) Then
            firstEmptyIndex = i
        End If
    End If
    If (VRPresentes(i).id = vFound.id) Then
        identificar = False
        Exit Function
    End If
Next i
VRPresentes(firstEmptyIndex) = vFound
identificar = True
End Function

Public Function limpiar(cuil As String) As Boolean
Call CheckInit
Dim i As Integer
Dim deleted As Boolean
Dim firstEmptyIndex As Integer
firstEmptyIndex = -1
deleted = False
'Busco diputado
For i = 0 To UBound(VRPresentes)
    If (VRPresentes(i).cuil = cuil) Then
        VRPresentes(i).id = 0
        VRPresentes(i).apellido = ""
        VRPresentes(i).nombre = ""
        VRPresentes(i).cuil = ""
        deleted = True
        Exit For
    End If
Next i
limpiar = deleted
End Function

Private Function tienePresentes() As Boolean
If VRPresentes(0).id <> 0 Then
    tienePresentes = True
    Exit Function
End If
tienePresentes = False
End Function

Private Sub CheckInit()
On Error GoTo checkErr
If (VRDiputados(0).id <> 0 And VRDiputados(1).id <> 0) Then
    Exit Sub
End If
If (VotoRemoto.DatabaseOpened = False) Then
    Exit Sub
End If
Dim rs As New ADODB.Recordset
frmMain.SetearRsAux "SELECT legisladores_activos.id, legisladores_activos.apellido, legisladores_activos.nombre, DiputadosCuil.cuil AS cuil, ISNULL(BancasProbables.banca, 300) AS banca FROM legisladores_activos LEFT JOIN DiputadosCuil ON DiputadosCuil.id = legisladores_activos.id LEFT JOIN BancasProbables ON BancasProbables.id_legislador = legisladores_activos.id", rs
Dim i As Integer
i = -1
While Not rs.EOF
    Dim svId As Integer
    Dim svCuil As String
    Dim svBanca As String
    svId = rs.Fields(0)
    svCuil = rs.Fields(3)
    svBanca = rs.Fields(4)
    i = i + 1
    VRDiputados(i).id = rs.Fields(0)
    VRDiputados(i).apellido = rs.Fields(1)
    VRDiputados(i).nombre = rs.Fields(2)
    VRDiputados(i).cuil = rs.Fields(3)
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
'Hardcore Init
'VRPresentes(0) = VRDiputados(0)
'VRPresentes(1) = VRDiputados(1)
'VRPresentes(2) = VRDiputados(2)
Dim s As Boolean
s = VotoRemoto.identificar("27108513727")
VotoRemoto.limpiar ("27108513727")
checkErr:
    Exit Sub
End Sub
