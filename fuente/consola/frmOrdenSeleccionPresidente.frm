VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOrdenSeleccionPresidente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Orden Asignado por el Presidente"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   9495
   End
   Begin VB.CommandButton cmdMaximo 
      Caption         =   "A&signar máximo a todos"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   5700
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   5700
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla 
      Height          =   4515
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7964
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Buscar apellido"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "Doble click para modificar orden"
      Height          =   195
      Left            =   7080
      TabIndex        =   2
      Top             =   240
      Width           =   2280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Órden de Presidente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2490
   End
End
Attribute VB_Name = "frmOrdenSeleccionPresidente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strconexion
Private Const MaximoIniciaEn  As Integer = 300

Private Sub armarGrilla()
    With Grilla
        .Cols = 7
        .TextMatrix(0, 0) = "Orden"
        .TextMatrix(0, 1) = "Apellido"
        .TextMatrix(0, 2) = "Nombre"
        .TextMatrix(0, 3) = "Agrupación Política"
        .TextMatrix(0, 4) = "Distrito y Sección"
        .TextMatrix(0, 5) = ""
        .TextMatrix(0, 6) = "" 'orden
        .ColWidth(0) = 800
        .ColWidth(1) = 2000
        .ColWidth(2) = 1500
        .ColWidth(3) = IIf(AGRUPACION_POLITICA_HABILITADA, 2000, 0)
        .ColWidth(4) = IIf(DISTRITO_HABILITADO, 2700, 0)
        .ColWidth(5) = 0
        .ColWidth(6) = 0
    End With
End Sub
Private Sub CargarGrilla()
    Dim strSql As String
    Dim xFila  As Long
    Dim Rs As New ADODB.Recordset
'    strSql = "SELECT Personas.id, Personas.apellido, " _
           & "Personas.nombre, Personas.bloquepolitico, " _
           & "Personas.provincia, ordpres.Orden_asignado_para_presidente " _
           & "FROM Personas INNER JOIN Ordpres ON Personas.id = dbo.ordpres.Identificador_de_Persona " _
           & "ORDER BY ordpres.Orden_asignado_para_presidente"
    'strSql = "SELECT Legisladores.*, legisladores_activos.OrdenPresidente, REPLACE(Legisladores.apellido, 'Ñ', 'NZ') + ', ' + REPLACE(Legisladores.nombre, 'Ñ', 'NZ') as orden " _
    '         & " From Legisladores INNER JOIN legisladores_activos ON Legisladores.id = legisladores_activos.ID " _
    '         & " ORDER BY legisladores_activos.OrdenPresidente"
    'strSql = "SELECT     Legisladores.apellido, Legisladores.nombre, Legisladores.bloque_politico, Legisladores.departamento, Legisladores.id, " & _
                      " legisladores_activos.OrdenPresidente, REPLACE(Legisladores.apellido, 'Ñ', 'NZ') + ', ' + REPLACE(Legisladores.nombre, 'Ñ', 'NZ') AS orden, " & _
                      " distritos.distrito + ', ' + secciones.seccion + ' Seccion ' AS departamento " & _
                      " FROM         secciones RIGHT OUTER JOIN " & _
                      " distritos ON secciones.id_seccion = distritos.seccion RIGHT OUTER JOIN " & _
                      " Legisladores INNER JOIN " & _
                      " legisladores_activos ON Legisladores.id = legisladores_activos.ID ON distritos.id_distrito = Legisladores.distrito " & _
                      " ORDER BY legisladores_activos.Apellido, legisladores_activos.Nombre"
    strSql = "SELECT     legisladores_activos.apellido, legisladores_activos.nombre, legisladores_activos.bloque_politico, 'x' AS depto, legisladores_activos.id, " & _
                      "legisladores_activos.OrdenPresidente, legisladores_activos.apellido + ', ' + legisladores_activos.nombre AS orden, distritos.distrito AS departamento " & _
"FROM         legisladores_activos INNER JOIN " & _
                      "Legisladores ON Legisladores.id = legisladores_activos.id LEFT OUTER JOIN " & _
                      "distritos ON distritos.id_distrito = Legisladores.distrito " & _
"ORDER BY legisladores_activos.apellido, legisladores_activos.nombre"
    
    SetearRs strSql, Rs
    xFila = 1
    With Rs
        Grilla.Rows = .RecordCount + 1
        While Not .EOF
            Grilla.TextMatrix(xFila, 0) = .Fields("OrdenPresidente").Value
            Grilla.TextMatrix(xFila, 1) = NullCadena(.Fields("apellido").Value)
            Grilla.TextMatrix(xFila, 2) = NullCadena(.Fields("nombre").Value)
            Grilla.TextMatrix(xFila, 3) = NullCadena(.Fields("bloque_politico").Value)
            Grilla.TextMatrix(xFila, 4) = NullCadena(.Fields("departamento").Value)
            Grilla.TextMatrix(xFila, 5) = .Fields("Id").Value
            Grilla.TextMatrix(xFila, 6) = .Fields("orden")
            .MoveNext
            xFila = xFila + 1
        Wend
    End With
    If Rs.State = adStateOpen Then
        Rs.Close
    End If
    Set Rs = Nothing
End Sub
Private Sub cmdMaximo_Click()
    If MsgBox("Está Ud. seguro de modificar el orden de TODOS los legisladores?", vbQuestion + vbYesNo, "Confirma la operación?") = vbYes Then
        asignarMaximo
    End If
End Sub
Private Sub asignarMaximo()
    Dim orden As Integer
    Dim legislador As String
    Dim i As Integer
    orden = MaximoIniciaEn
    Grilla.ColSel = 6
    Grilla.Sort = 4
    Datos.IniciarTransaccion
    On Error GoTo ErrorDatos
    For i = 1 To Grilla.Rows - 1
        orden = orden + 1
        legislador = Grilla.TextMatrix(i, 5)
        actualizarOrden legislador, 300
    Next i
    Datos.FinalizarTransaccion True
    CargarGrilla
Exit Sub
ErrorDatos:
    MsgBox "Ha ocurrido un error al registrar los datos.", vbInformation + vbOKOnly
    Datos.FinalizarTransaccion False
    CargarGrilla
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Call armarGrilla
    Call CargarGrilla
End Sub
Private Sub actualizarOrden(pLegislador As String, pOrden As Integer)
    Dim strSql     As String
   strSql = "UPDATE legisladores_activos SET OrdenPresidente = " & Str(pOrden) _
               & "WHERE  Id = '" & Trim(pLegislador) & "'"
   Datos.SenteciaSQl strSql
End Sub
Private Sub Grilla_DblClick()
    On Error GoTo Trap_Error
    
    Dim xOrden     As Integer
    Dim strPersona As String
    
    If Grilla.Row > 0 Then
        strPersona = Grilla.TextMatrix(Grilla.Row, 5)
        xOrden = InputBox("Ingrese el orden para selección de presidente que se aplica al legislador seleccionado", "Orden para selección de presidente", 1)
        If Not IsNumeric(xOrden) Then
            MsgBox "El orden para la selección depresidente debe ser un número entero", vbInformation + vbOKOnly
            Exit Sub
        End If
        xOrden = Int(xOrden)
        actualizarOrden strPersona, xOrden
        Call CargarGrilla
    End If
Exit Sub
Trap_Error:
    Select Case err.Number
        Case 13
            MsgBox "El orden para la selección depresidente debe ser un número entero", vbInformation + vbOKOnly
            Exit Sub
        Case Else
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            Resume
    End Select
End Sub
Private Sub txtBuscar_GotFocus()
    seleccionadoTxt txtBuscar
End Sub
Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim Col As Integer
        Dim Row As Integer
        Dim X   As Long
        Funciones.BuscarEnGrilla Grilla, 2, txtBuscar.Text, Col, Row
        If (Col <> -1) And (Row <> -1) Then
            Grilla.SetFocus
            Grilla.Row = Row
            Grilla.RowSel = Row
            Grilla.ColSel = 5
            Grilla.TopRow = Row
        Else
            MsgBox "No se ha encontrado el texto deseado." & Chr(13) & "Intente con otra búsqueda.", vbInformation + vbOKOnly
        End If
    End If
End Sub
