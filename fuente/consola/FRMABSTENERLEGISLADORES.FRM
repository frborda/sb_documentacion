VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAbstenerLegisladores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Abstener Legisladores"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTodos 
      Caption         =   "Todos"
      Height          =   195
      Left            =   8280
      TabIndex        =   6
      Top             =   150
      Width           =   825
   End
   Begin VB.CommandButton cmdAbstenerBloque 
      Caption         =   "&Seleccionar"
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      ToolTipText     =   "Seleccionar todos los legisladores del bloque"
      Top             =   8370
      Width           =   1600
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   915
      TabIndex        =   4
      Top             =   8340
      Width           =   1800
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla 
      Height          =   7875
      Left            =   60
      TabIndex        =   2
      Top             =   420
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   13891
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   315
      Left            =   7530
      TabIndex        =   1
      Top             =   8370
      Width           =   1600
   End
   Begin VB.CommandButton cmdRegistrarAbstencion 
      Caption         =   "&Emitir Petición"
      Height          =   315
      Left            =   4380
      TabIndex        =   0
      ToolTipText     =   "Enviar mensaje para permitir abstención o quitar estado de abstención de legisladores"
      Top             =   8370
      Width           =   1600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bloque : "
      Height          =   315
      Left            =   210
      TabIndex        =   3
      Top             =   8400
      Width           =   630
   End
End
Attribute VB_Name = "frmAbstenerLegisladores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsBloques                 As ADODB.Recordset
Private strVectorIdentificacion() As String
Private strVectorAbstencion()     As String


Private Sub LeerVectores()

    Dim strSql    As String
    Dim rsVect    As ADODB.Recordset
    Dim strCadena As String
   
    Set rsVect = New ADODB.Recordset
    ' ----------------------------------------------------------------------
    ' Leer vector identificacion
    ' ----------------------------------------------------------------------
    strSql = "SELECT TOP 1 Vector_Identificacion FROM Vector"
    Datos.SetearRs strSql, rsVect
    strCadena = Trim(rsVect.Fields("Vector_Identificacion").Value)
    rsVect.Close
    strVectorIdentificacion = Split(strCadena, ";")
    ' ----------------------------------------------------------------------
    ' Leer vector Abstencion
    ' ----------------------------------------------------------------------
    strSql = "SELECT TOP 1 Vector_Abstencion FROM Vector"
    Datos.SetearRs strSql, rsVect
    strCadena = Trim(rsVect.Fields("Vector_Abstencion").Value)
    rsVect.Close
    strVectorAbstencion = Split(strCadena, ";")
End Sub
Private Sub armarGrilla()
    With Grilla
        .Cols = 6
        .TextMatrix(0, 0) = "Banca"
        .TextMatrix(0, 1) = "Apellido"
        .TextMatrix(0, 2) = "Nombre"
        .TextMatrix(0, 3) = "Bloque"
        .TextMatrix(0, 4) = "Se Abstiene"
        .TextMatrix(0, 5) = "id_leg"
        .ColWidth(0) = 600
        .ColWidth(1) = 2300
        .ColWidth(2) = 2300
        .ColWidth(3) = 2300
        .ColWidth(4) = 1300
        .ColWidth(5) = 0
    End With
End Sub

Private Sub cargarGrilla()
    Dim strSql         As String
    Dim rsLegisladores As ADODB.Recordset
    Dim xFila          As Long
    
    Set rsLegisladores = New ADODB.Recordset
    
    strSql = "SELECT Legisladores.id, legisladores_activos.DESKID AS Banca, Legisladores.apellido, Legisladores.nombre, " _
           & "Legisladores.bloque_politico FROM Legisladores INNER JOIN legisladores_activos ON Legisladores.id = legisladores_activos.ID " _
           & "ORDER BY banca"
    xFila = 1
    Datos.SetearRs strSql, rsLegisladores
    
    With rsLegisladores
        If .RecordCount > 0 Then
            Grilla.Rows = .RecordCount + 2
            .MoveFirst
            While Not .EOF
                Grilla.TextMatrix(xFila, 0) = .Fields("Banca").Value
                Grilla.TextMatrix(xFila, 1) = .Fields("Apellido").Value
                Grilla.TextMatrix(xFila, 2) = .Fields("Nombre").Value
                Grilla.TextMatrix(xFila, 3) = .Fields("bloque_politico").Value
                Grilla.TextMatrix(xFila, 4) = "No"
                Grilla.TextMatrix(xFila, 5) = .Fields("id").Value
                xFila = xFila + 1
                .MoveNext
            Wend
        End If
    End With
End Sub
Private Sub SetearValoresGrilla()
    Dim xFilas As Long
    Dim x      As Long
    Dim y      As Long
    Dim strLeg As String ' id de legislador
    ' ------------------------------------------------------------------------------
    ' Recorrer el vector Abstencion
    ' ------------------------------------------------------------------------------
    For y = 0 To UBound(strVectorAbstencion)
        ' Tomo un legislador del vector de abstenciones
        strLeg = Trim(strVectorAbstencion(y))
        If Not IsNull(strLeg) Then ' Buscarlo en la grilla y ponerlo en Sí
            For x = 1 To Grilla.Rows - 1
                If Trim(LCase(Grilla.TextMatrix(x, 5))) = Trim(LCase(strLeg)) Then
                    Grilla.TextMatrix(x, 4) = "Sí"
                    Grilla.Refresh
                    Exit For
                End If
            Next x
        End If
    Next y
End Sub
Private Sub CargarCombo()
    Dim strSql    As String
    Dim rsBloques As ADODB.Recordset
    ' ------------------------------------------------------------------------------
    ' Buscar todos los bloques definidos en SQV
    ' ------------------------------------------------------------------------------
    strSql = "SELECT * From Bloques ORDER BY Bloque_Político"
    Set rsBloques = New ADODB.Recordset
    Datos.SetearRs strSql, rsBloques
    ' ------------------------------------------------------------------------------
    ' Cargar bloques en el combo IU
    ' ------------------------------------------------------------------------------
    Combo1.Clear
    With rsBloques
        .MoveFirst
        While Not .EOF
            Combo1.AddItem .Fields("Bloque_Político").Value
            .MoveNext
        Wend
        .Close
    End With
    Set rsBloques = Nothing
End Sub
Private Sub chkTodos_Click()
    Dim x As Long
    ' ------------------------------------------------------------------------------
    ' Seleccionar todos los elementos de la grilla, o deseleccionar a todos
    ' ------------------------------------------------------------------------------
    For x = 1 To Grilla.Rows - 1
        If chkTodos = 1 Then
            Grilla.TextMatrix(x, 4) = "Sí"
        Else
            Grilla.TextMatrix(x, 4) = "No"
        End If
    Next x
End Sub
Private Sub cmdAbstenerBloque_Click()
    ' ------------------------------------------------------------------------------
    ' Seleccionar todos los legisladores de un bloque
    ' ------------------------------------------------------------------------------
    Dim x As Long
    For x = 1 To Grilla.Rows - 1
        If Trim(LCase(Grilla.TextMatrix(x, 3))) = Trim(LCase(Combo1.Text)) Then
            Grilla.TextMatrix(x, 4) = "Sí"
        Else
            Grilla.TextMatrix(x, 4) = "No"
        End If
    Next x
End Sub

Private Sub cmdRegistrarAbstencion_Click()
    Dim strLista As String
    Dim x        As Long
    
    ' ------------------------------------------------------------------------------
    ' Armar vector de abstencion y enviar mensaje a SQV
    ' ------------------------------------------------------------------------------
    strLista = ""
    For x = 1 To Grilla.Rows - 1
        If Grilla.TextMatrix(x, 4) = "Sí" Then
            strLista = strLista & Grilla.TextMatrix(x, 5) & ";"
        End If
    Next x
    ' strLista = "00000118;00000119;00000101;00000100;00000103;"
    MensajesSQV.AbstenerLeg (strLista)
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Call LeerVectores
    Call armarGrilla
    Call CargarCombo
    Call cargarGrilla
    Call SetearValoresGrilla
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Grilla_DblClick()
    With Grilla
        If .TextMatrix(.row, 4) = "No" Then
            .TextMatrix(.row, 4) = "Sí"
        Else
            .TextMatrix(.row, 4) = "No"
        End If
    End With
End Sub
