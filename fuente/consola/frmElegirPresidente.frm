VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmElegirPresidente2 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccionar Presidente"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmElegir 
      Interval        =   500
      Left            =   4020
      Top             =   4500
   End
   Begin VB.Frame frmBusqueda 
      BackColor       =   &H00404040&
      Caption         =   "Búsqueda por aproximación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   8055
      Begin Proyecto1.ButtonOffice cmdAplicar 
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   300
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BackColor       =   12230304
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin VB.TextBox txtApellido 
         Height          =   345
         Left            =   840
         TabIndex        =   5
         Top             =   300
         Width           =   4935
      End
      Begin VB.TextBox txtNombre 
         Height          =   345
         Left            =   3300
         TabIndex        =   3
         Top             =   300
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   2700
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid dgPresidente 
      Height          =   3195
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5636
      _Version        =   393216
      Cols            =   6
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin Proyecto1.ButtonOffice cmdCancelar 
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   4500
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   12230304
      Caption         =   "&Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdMantenerPresidente 
      Height          =   300
      Left            =   180
      TabIndex        =   9
      Top             =   4560
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   529
      BackColor       =   12230304
      Caption         =   "Mantener el mismo Presidente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin VB.Shape shpMantienePresidente 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   16
      Height          =   195
      Left            =   210
      Top             =   4620
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image picDiputado 
      Height          =   3195
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label lbldClick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doble click sobre el nombre para elegir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   3330
   End
End
Attribute VB_Name = "frmElegirPresidente2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstLista As New ADODB.Recordset
Private Cargo As Boolean
Private PrimeraCarga As Boolean
Private xCargo As Boolean
Private Sub cmdAplicar_Click()
If txtApellido.Text <> "" Then
    Call TitulosGRilla
    Call CargarGrillaFiltrada(Trim(txtNombre.Text), Trim(txtApellido.Text))
    If dgPresidente.Rows = 2 Then
        If dgPresidente.TextMatrix(dgPresidente.Row, 3) <> "" Then
            dgPresidente_EnterCell
        End If
    Else
        Set picDiputado.Picture = Nothing
    End If
Else
    Call MsgBox("Para utilizar la búsqueda escriba una palabra", vbInformation)
End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdMantenerPresidente_Click()
Unload Me
End Sub
Private Sub dgPresidente_Click()
If dgPresidente.Rows = 2 Then
    If dgPresidente.TextMatrix(dgPresidente.Row, 3) <> "" Then
        dgPresidente_EnterCell
    End If
End If
End Sub

Private Sub dgPresidente_DblClick()
    If dgPresidente.Row > 0 Then
        If dgPresidente.CellBackColor <> &H8000000F Then
            If (gTipoUsuario <> 1) And (gTipoUsuario <> 4) Then
                MensajesSQV.cambiarPresidente dgPresidente.TextMatrix(dgPresidente.Row, 3)
                If dgPresidente.TextMatrix(dgPresidente.Row, 4) = 1 Then
                    flPresidenteLegislador = True
                Else
                    flPresidenteLegislador = False
                End If
                ' frmConsolaOperacion.lblPresidente.Caption = dgPresidente.TextMatrix(dgPresidente.row, 2)
                cmdCancelar_Click
            Else
                MsgBox "Ud. no dispone de permisos para realizar esta acción.", vbInformation + vbOKOnly
            End If
        End If
    End If
End Sub
Private Sub dgPresidente_EnterCell()
On Error Resume Next
Dim pic As New ADODB.Stream
Dim Rinfo As New ADODB.Recordset
If (Cargo = True Or (PrimeraCarga = True And Cargo = False)) Or Presidente_Label = "Seleccione el presidente antes de continuar" Then
    If mVectorIdentificacion(0) = dgPresidente.TextMatrix(dgPresidente.Row, 3) Or xCargo = True Then
        SetearRs "SELECT PICTURE FROM legisladores WHERE id = " & dgPresidente.TextMatrix(dgPresidente.Row, 3), Rinfo
        If Not Rinfo.EOF Then
            If Not IsNull(Rinfo.Fields(0)) Then
                Set pic = New ADODB.Stream
                pic.Type = adTypeBinary
                pic.Open
                pic.Write Rinfo.Fields(0)
                pic.SaveToFile App.Path & "\temp.jpg", adSaveCreateOverWrite
                picDiputado.Picture = LoadPicture(App.Path & "\temp.jpg")
            Else
                picDiputado.Picture = Nothing
            End If
        Else
            picDiputado.Picture = Nothing
        End If
        PrimeraCarga = False
    End If
End If
'MsgBox dgPresidente.RowSel
End Sub
Private Sub dgPresidente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dgPresidente_DblClick
    End If
End Sub
Private Sub dgPresidente_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If xCargo = False Then
    xCargo = True
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 27 Then
    '    Unload Me
    'End If
End Sub
Private Sub Form_Load()
    xCargo = False
    Cargo = False
    PrimeraCarga = True
    TitulosGRilla
    Call CargarGrilla
    Cargo = True
    'If frmConsolaOperacion.lblPresidente.Caption = "Seleccione el presidente antes de continuar" Then
        dgPresidente_EnterCell
    'End If
End Sub
Private Sub TitulosGRilla()
    With dgPresidente
        .Clear
        .Cols = 5
        .ColWidth(0) = 100
        .ColWidth(1) = 1000
        .ColWidth(2) = 6000
        .ColWidth(3) = 0 'codigo
        .ColWidth(4) = 0 ' Es_Legislador
        .TextMatrix(0, 1) = "Nº Orden"
        .TextMatrix(0, 2) = "Legislador"
    End With
End Sub
Private Sub CargarGrilla()
    Dim i           As Integer
    Dim j           As Integer
    Dim Posicion    As Integer
    Dim strSql      As String
    Dim strVector() As String
    Dim strCadena   As String
    Dim blFind      As Boolean
    Dim xFila       As Long
    Dim strIdLista  As String
    
    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Levantar Vector Identificación
    ' ----------------------------------------------------------------------------------------------------------------------------------
    strSql = "SELECT vector_identificacion FROM vector"
    Datos.SetearRs strSql, rstLista
    strCadena = Trim(rstLista.Fields(0).Value)
    strCadena = ";" & strCadena
    rstLista.Close
    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Listar legisladores activos
    ' ----------------------------------------------------------------------------------------------------------------------------------
    strSql = "SELECT Legisladores.id, Legisladores.apellido + ' ' + dbo.Legisladores.nombre AS Legislador, " _
           & "legisladores_activos.ordenpresidente AS OPresidencia, Legisladores.Es_Legislador FROM Legisladores INNER JOIN legisladores_activos ON " _
           & "Legisladores.ID = legisladores_activos.ID WHERE Legisladores.tipo = 1 ORDER BY OPresidencia,Legisladores.apellido"
    
    Datos.SetearRs strSql, rstLista
    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Cargar en la grilla Todos los legisladores
    ' ----------------------------------------------------------------------------------------------------------------------------------
    xFila = 0
    dgPresidente.Rows = 1 ' Asegurar tener la misma cantidad de filas que de legisladores (en la grilla)
    If rstLista.EOF = False Then
        rstLista.MoveFirst
        Do While Not (rstLista.EOF)
            ' Cargar legislador en la grilla (los datos se sacan del recordset)
            dgPresidente.AddItem vbTab & rstLista!OPresidencia & vbTab & rstLista!legislador & vbTab & rstLista!id & vbTab & rstLista!Es_Legislador
            ' Buscar al legislador en strVector: Si se lo encuentra, se pinta de gris esa fila
            strIdLista = Trim(";" & Trim(rstLista!id) & ";")
            If InStr(1, strCadena, strIdLista) > 0 Then
                dgPresidente.Row = xFila + 1
                For j = 1 To dgPresidente.Cols - 1
                    dgPresidente.Col = j
                    dgPresidente.CellBackColor = &H8000000F
                Next j
            End If
             xFila = xFila + 1
            rstLista.MoveNext
        Loop
    Else
        MsgBox "No hay legisladores disponibles para presidir el recinto.", vbInformation, "Consola SQV"
        Unload Me
    End If
End Sub
Private Sub CargarGrillaFiltrada(v_Nombre As String, v_Apellido As String)
    Dim i           As Integer
    Dim j           As Integer
    Dim Posicion    As Integer
    Dim strSql      As String
    Dim strVector() As String
    Dim strCadena   As String
    Dim blFind      As Boolean
    Dim xFila       As Long
    Dim strIdLista  As String
    
    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Levantar Vector Identificación
    ' ----------------------------------------------------------------------------------------------------------------------------------
    strSql = "SELECT vector_identificacion FROM vector"
    Datos.SetearRs strSql, rstLista
    strCadena = Trim(rstLista.Fields(0).Value)
    strCadena = ";" & strCadena
    rstLista.Close
    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Listar legisladores activos
    ' ----------------------------------------------------------------------------------------------------------------------------------
    strSql = "SELECT Legisladores.id, Legisladores.apellido + ' ' + dbo.Legisladores.nombre AS Legislador, " _
           & "legisladores_activos.ordenpresidente AS OPresidencia, Legisladores.Es_Legislador FROM Legisladores INNER JOIN legisladores_activos ON " _
           & "Legisladores.ID = legisladores_activos.ID WHERE Legisladores.tipo = 1 AND Legisladores.nombre LIKE '" & v_Nombre & "%' AND Legisladores.apellido LIKE '" & v_Apellido & "%' ORDER BY OPresidencia,Legisladores.Apellido"
    
    Datos.SetearRs strSql, rstLista
    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Cargar en la grilla Todos los legisladores
    ' ----------------------------------------------------------------------------------------------------------------------------------
    xFila = 0
    dgPresidente.Rows = 1 ' Asegurar tener la misma cantidad de filas que de legisladores (en la grilla)
    If rstLista.EOF = False Then
        rstLista.MoveFirst
        Do While Not (rstLista.EOF)
            ' Cargar legislador en la grilla (los datos se sacan del recordset)
            dgPresidente.AddItem vbTab & rstLista!OPresidencia & vbTab & rstLista!legislador & vbTab & rstLista!id & vbTab & rstLista!Es_Legislador
            ' Buscar al legislador en strVector: Si se lo encuentra, se pinta de gris esa fila
            strIdLista = Trim(";" & Trim(rstLista!id) & ";")
            If InStr(1, strCadena, strIdLista) > 0 Then
                dgPresidente.Row = xFila + 1
                For j = 1 To dgPresidente.Cols - 1
                    dgPresidente.Col = j
                    dgPresidente.CellBackColor = &H8000000F
                Next j
            End If
             xFila = xFila + 1
            rstLista.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If rstLista.State = adStateOpen Then
        rstLista.Close
    End If
    Set rstLista = Nothing
End Sub

Private Sub tmElegir_Timer()
If lbldClick.ForeColor = vbWhite Then
    lbldClick.ForeColor = vbRed
Else
    lbldClick.ForeColor = vbWhite
End If
End Sub

Private Sub txtApellido_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAplicar_Click
End If
End Sub
Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAplicar_Click
End If
End Sub
