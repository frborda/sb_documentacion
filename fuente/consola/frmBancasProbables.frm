VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmBancasProbables 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administración de Bancas Probables"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      Caption         =   "Filtro por Apellido"
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
      TabIndex        =   12
      Top             =   1320
      Width           =   3885
      Begin VB.TextBox txtApellido 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   0
         Top             =   330
         Width           =   3045
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   "Filtro por Banca"
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
      Left            =   4080
      TabIndex        =   9
      Top             =   1320
      Width           =   4815
      Begin VB.ComboBox cmbBanca 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         TabIndex        =   10
         Text            =   "-Seleccione una banca-"
         Top             =   330
         Width           =   2445
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00404040&
      Caption         =   "Opciones Masivas"
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
      Left            =   4080
      TabIndex        =   6
      Top             =   30
      Width           =   4815
      Begin Proyecto1.ButtonOffice cmdResetMasivo 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   661
         BackColor       =   12230304
         Caption         =   "Asignar 300 a todos los diputados"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Orden"
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
      Height          =   1155
      Left            =   120
      TabIndex        =   3
      Top             =   30
      Width           =   3885
      Begin VB.OptionButton optBloque 
         BackColor       =   &H00404040&
         Caption         =   "Por Bloque Político"
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
         Left            =   90
         TabIndex        =   13
         Top             =   690
         Width           =   2145
      End
      Begin VB.OptionButton optBancaProbable 
         BackColor       =   &H00404040&
         Caption         =   "Por Banca Probable"
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
         Left            =   1590
         TabIndex        =   5
         Top             =   330
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optApellido 
         BackColor       =   &H00404040&
         Caption         =   "Por Apellido"
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
         Left            =   90
         TabIndex        =   4
         Top             =   330
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Listado de Bancas Probables"
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
      Height          =   6705
      Left            =   120
      TabIndex        =   1
      Top             =   2310
      Width           =   8715
      Begin MSFlexGridLib.MSFlexGrid flexDatos 
         Height          =   6255
         Left            =   120
         TabIndex        =   2
         Top             =   330
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   11033
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Proyecto1.ButtonOffice cmdVolver 
      Height          =   375
      Left            =   5670
      TabIndex        =   8
      Top             =   9120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      BackColor       =   12230304
      Caption         =   "&Volver a Configuraciones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdImprimirListado 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   9120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      BackColor       =   12230304
      Caption         =   "Imprimir listado actual"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
End
Attribute VB_Name = "frmBancasProbables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ListaNombres() As String
Dim ListaApellidos() As String
Dim rsLegisladores As ADODB.Recordset
Public BancaAAsignar As Integer
Private Sub cmbBancasProbables_Click()
If cmbBancasProbables.ListIndex <> -1 Then
    lblBancaProbable.Caption = cmbBancasProbables.List(cmbBancasProbables.ListIndex)
End If
End Sub
Private Sub AplicarFiltro()
Dim consulta As String
Dim filtro As String
txtApellido.Enabled = False
LimpiarFlex
filtro = " WHERE legisladores_activos.apellido LIKE '" & txtApellido.Text & "%'"
consulta = "SELECT BancasProbables.banca,id_legislador,legisladores_activos.nombre,legisladores_activos.apellido, legisladores_activos.bloque_politico FROM BancasProbables INNER JOIN legisladores_activos ON legisladores_activos.id = BancasProbables.id_legislador "
If cmbBanca.ListIndex > 0 Then
    filtro = filtro & " AND BancasProbables.banca = " & cmbBanca.List(cmbBanca.ListIndex)
End If
consulta = consulta & filtro
If optBloque.Value = True Then
    consulta = consulta & " ORDER BY bloque_politico,apellido, nombre"
ElseIf optApellido.Value = True Then
    consulta = consulta & " ORDER BY apellido, nombre"
ElseIf optBancaProbable.Value = True Then
    consulta = consulta & " ORDER BY banca"
End If
Set rsLegisladores = New ADODB.Recordset
SetearRs consulta, rsLegisladores
While Not rsLegisladores.EOF
    flexDatos.AddItem ""
    campo = campo + 1
    flexDatos.TextMatrix(campo, 0) = rsLegisladores.Fields("apellido")
    flexDatos.TextMatrix(campo, 1) = rsLegisladores.Fields("nombre")
    flexDatos.TextMatrix(campo, 2) = Trim(Str(rsLegisladores.Fields("banca")))
    flexDatos.TextMatrix(campo, 3) = IIf(IsNull(rsLegisladores.Fields("bloque_politico")), "", rsLegisladores.Fields("bloque_politico"))
    flexDatos.TextMatrix(campo, 4) = rsLegisladores.Fields("id_legislador")
    rsLegisladores.MoveNext
Wend
rsLegisladores.Close
Set rsLegisladores = Nothing
txtApellido.Enabled = True
txtApellido.SetFocus
End Sub
Private Function GetFiltro() As String
Dim consulta As String
Dim filtro As String
filtro = " WHERE legisladores_activos.apellido LIKE '" & txtApellido.Text & "%'"
consulta = "SELECT GETDATE() AS fecha_actual, BancasProbables.banca ,id_legislador,legisladores_activos.apellido,legisladores_activos.nombre, legisladores_activos.apellido + ', ' + legisladores_activos.nombre AS ApellidoNombre, legisladores_activos.bloque_politico FROM BancasProbables INNER JOIN legisladores_activos ON legisladores_activos.id = BancasProbables.id_legislador "
If cmbBanca.ListIndex > 0 Then
    filtro = filtro & " AND BancasProbables.banca = " & cmbBanca.List(cmbBanca.ListIndex)
End If
consulta = consulta & filtro
If optBloque.Value = True Then
    consulta = consulta & " ORDER BY bloque_politico,apellido, nombre"
ElseIf optApellido.Value = True Then
    consulta = consulta & " ORDER BY apellido, nombre"
ElseIf optBancaProbable.Value = True Then
    consulta = consulta & " ORDER BY banca"
End If
GetFiltro = consulta
End Function

Private Sub cmbBanca_Click()
AplicarFiltro
End Sub

Private Sub cmdImprimirListado_Click()
Dim rptB As New rptBancasProbables
Dim Rs As New Recordset
SetearRs GetFiltro, Rs
rptB.DataControl1.Recordset = Rs
For i = 0 To rptB.Pages.Count - 1
    rptB.Pages(i).Width = 300
Next i
rptB.PrintReport True
End Sub

Private Sub cmdResetMasivo_Click()
Dim r As Integer
r = MsgBox("¿Está seguro de que desea asignar el número 300 a todos los diputados?", vbYesNo, "Alerta")
If r = vbYes Then
    EjecutarSQL ("UPDATE BancasProbables SET banca = 300")
    AplicarFiltro
End If
End Sub

Private Sub cmdVolver_Click()
Unload Me
End Sub
Private Sub flexDatos_DblClick()
Dim i As Integer
Dim Rs As ADODB.Recordset
Dim fS As frmSeleccionarBancaProbable
Set fS = New frmSeleccionarBancaProbable
fS.lblDiputado = flexDatos.TextMatrix(flexDatos.RowSel, 0) & " " & flexDatos.TextMatrix(flexDatos.RowSel, 1)
fS.cmbBancasProbables.AddItem "300"
For i = 0 To 256
    Set Rs = New ADODB.Recordset
    SetearRs "SELECT * FROM BancasProbables WHERE banca = " & i, Rs
    If Rs.EOF Then
        fS.cmbBancasProbables.AddItem Trim(Str(i))
    End If
    Rs.Close
    Set Rs = Nothing
Next i
fS.Show vbModal
If BancaAAsignar <> -1 Then
    EjecutarSQL ("UPDATE BancasProbables SET banca = " & BancaAAsignar & " WHERE id_legislador = " & flexDatos.TextMatrix(flexDatos.RowSel, 4))
    flexDatos.TextMatrix(flexDatos.RowSel, 2) = Trim(Str(BancaAAsignar))
    'cmdAplicarFiltro_Click
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim RsTemp As ADODB.Recordset
Set rsLegisladores = New ADODB.Recordset
SetearRs "SELECT id FROM legisladores_activos ORDER BY id", rsLegisladores
While Not rsLegisladores.EOF
    Set RsTemp = New ADODB.Recordset
    SetearRs "SELECT * FROM BancasProbables WHERE id_legislador = " & rsLegisladores.Fields(0), RsTemp
    If RsTemp.EOF Then
        consulta = "INSERT INTO BancasProbables(id_legislador,banca) VALUES(" & rsLegisladores.Fields(0) & ",300)"
        EjecutarSQL (consulta)
    End If
    RsTemp.Close
    Set RsTemp = Nothing
    rsLegisladores.MoveNext
Wend
rsLegisladores.Close
Set rsLegisladores = Nothing
LimpiarFlex
LlenarFlex ("ORDER BY BancasProbables.banca")
'Filtro de bancas
cmbBanca.AddItem "-Seleccione una banca-"
For i = 0 To 256
    cmbBanca.AddItem Str(i)
Next i
End Sub
Private Sub LimpiarFlex()
flexDatos.Clear
flexDatos.Rows = 1
flexDatos.Cols = 5
flexDatos.ColWidth(0) = 2500
flexDatos.TextMatrix(0, 0) = "Apellido"
flexDatos.ColWidth(1) = 3000
flexDatos.TextMatrix(0, 1) = "Nombre"
flexDatos.ColWidth(2) = 1800
flexDatos.TextMatrix(0, 2) = "Banca Probable"
flexDatos.ColWidth(3) = 1800
flexDatos.TextMatrix(0, 3) = "Bloque"
flexDatos.ColWidth(4) = 0
End Sub
Private Sub LlenarFlex(OrderBy As String)
Dim RsTemp As ADODB.Recordset
Dim Cantidad As Integer
Dim i As Integer
Dim consulta As String
Dim campo As Integer
LimpiarFlex
campo = 0
Set rsLegisladores = New ADODB.Recordset
SetearRs "SELECT BancasProbables.banca,id_legislador,legisladores_activos.nombre,legisladores_activos.apellido, legisladores_activos.bloque_politico FROM BancasProbables INNER JOIN legisladores_activos ON legisladores_activos.id = BancasProbables.id_legislador " & OrderBy, rsLegisladores
rsLegisladores.MoveFirst
While Not rsLegisladores.EOF
    flexDatos.AddItem ""
    campo = campo + 1
    flexDatos.TextMatrix(campo, 0) = rsLegisladores.Fields("apellido")
    flexDatos.TextMatrix(campo, 1) = rsLegisladores.Fields("nombre")
    flexDatos.TextMatrix(campo, 2) = Trim(Str(rsLegisladores.Fields("banca")))
    flexDatos.TextMatrix(campo, 3) = IIf(IsNull(rsLegisladores.Fields("bloque_politico")), "", rsLegisladores.Fields("bloque_politico"))
    flexDatos.TextMatrix(campo, 4) = rsLegisladores.Fields("id_legislador")
    rsLegisladores.MoveNext
Wend
rsLegisladores.Close
Set rsLegisladores = Nothing
End Sub
Private Sub lstApellidos_Click()
If lstApellidos.ListIndex <> -1 Then
    txtApellidos.Text = lstApellidos.List(lstApellidos.ListIndex)
    txtApellidos_KeyUp 0, 0
End If
End Sub
Private Sub lstNombres_Click()
If lstNombres.ListIndex <> -1 Then
    txtNombre.Text = lstNombres.List(lstNombres.ListIndex)
    txtNombre_KeyUp 0, 0
End If
End Sub
Private Sub txtApellidos_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer
lstApellidos.Clear
If txtApellidos.Text = "" Then
    For i = LBound(ListaApellidos) To UBound(ListaApellidos)
        lstApellidos.AddItem ListaApellidos(i)
    Next i
Else
    For i = LBound(ListaApellidos) To UBound(ListaApellidos)
        If (InStr(LCase(ListaApellidos(i)), LCase(txtApellidos.Text)) = 1) Then
            lstApellidos.AddItem ListaApellidos(i)
        End If
    Next i
End If
End Sub
Private Sub txtNombre_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer
lstNombres.Clear
If txtNombre.Text = "" Then
    For i = LBound(ListaNombres) To UBound(ListaNombres)
        lstNombres.AddItem ListaNombres(i)
    Next i
Else
    For i = LBound(ListaNombres) To UBound(ListaNombres)
        If (InStr(LCase(ListaNombres(i)), LCase(txtNombre.Text)) = 1) Then
            lstNombres.AddItem ListaNombres(i)
        End If
    Next i
End If
End Sub
Private Sub optApellido_Click()
AplicarFiltro
End Sub
Private Sub optBancaProbable_Click()
AplicarFiltro
End Sub

Private Sub optBloque_Click()
AplicarFiltro
End Sub

Private Sub txtApellido_Change()
AplicarFiltro
End Sub
