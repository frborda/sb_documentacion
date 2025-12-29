VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmHistorico 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estado de Legisladores"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmGrupos 
      BackColor       =   &H00404040&
      Caption         =   "Filtro por Provincia"
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
      Height          =   2295
      Left            =   7230
      TabIndex        =   8
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtGrupos 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   60
         TabIndex        =   10
         Text            =   "Escriba una provincia o parte de ella"
         Top             =   240
         Width           =   3315
      End
      Begin VB.ListBox lstGrupos 
         Height          =   1620
         ItemData        =   "frmHistorico.frx":0000
         Left            =   60
         List            =   "frmHistorico.frx":0002
         TabIndex        =   9
         Top             =   540
         Width           =   3315
      End
   End
   Begin VB.Frame frmBloques 
      BackColor       =   &H00404040&
      Caption         =   "Filtro por Bloque Político"
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
      Height          =   2295
      Left            =   3690
      TabIndex        =   5
      Top             =   120
      Width           =   3495
      Begin VB.ListBox lstBloques 
         Height          =   1620
         ItemData        =   "frmHistorico.frx":0004
         Left            =   60
         List            =   "frmHistorico.frx":0006
         TabIndex        =   7
         Top             =   540
         Width           =   3315
      End
      Begin VB.TextBox txtBloques 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   60
         TabIndex        =   6
         Text            =   "Escriba un bloque o parte de él para buscar"
         Top             =   240
         Width           =   3315
      End
   End
   Begin VB.Frame frmApellidos 
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
      Height          =   1575
      Left            =   150
      TabIndex        =   2
      Top             =   840
      Width           =   3495
      Begin VB.TextBox txtApellido 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Escriba un apellido o parte de él para buscar"
         Top             =   240
         Width           =   3255
      End
      Begin VB.ListBox lstApellidos 
         Height          =   840
         ItemData        =   "frmHistorico.frx":0008
         Left            =   120
         List            =   "frmHistorico.frx":000A
         TabIndex        =   3
         Top             =   570
         Width           =   3255
      End
   End
   Begin VB.Frame frmFiltros 
      BackColor       =   &H00404040&
      Caption         =   "Filtro por Estado"
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
      Height          =   675
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.ComboBox cmbEstado 
         Height          =   315
         ItemData        =   "frmHistorico.frx":000C
         Left            =   120
         List            =   "frmHistorico.frx":000E
         TabIndex        =   1
         Text            =   "-Seleccione un Estado-"
         Top             =   240
         Width           =   3255
      End
   End
   Begin Proyecto1.ButtonOffice cmdAplicarFiltro 
      Height          =   375
      Left            =   180
      TabIndex        =   11
      Top             =   2550
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   661
      BackColor       =   12230304
      Caption         =   "&Aplicar Filtro"
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
   Begin TabDlg.SSTab tbHistorico 
      Height          =   5460
      Left            =   540
      TabIndex        =   12
      Top             =   3030
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9631
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   4210752
      TabCaption(0)   =   "Actualización de estado"
      TabPicture(0)   =   "frmHistorico.frx":0010
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCantidad"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "asfasdfasd"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "flexEstados"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "frmHistorico.frx":002C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flexHistorico"
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid flexEstados 
         Height          =   4395
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7752
         _Version        =   393216
         Rows            =   1
         Cols            =   11
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid flexHistorico 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   8281
         _Version        =   393216
         Rows            =   1
         Cols            =   11
         FixedCols       =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Doble click sobre un legislador para actualizar su estado"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   4950
         Width           =   4035
      End
      Begin VB.Label asfasdfasd 
         Caption         =   "Cantidad de Legisladores:"
         Height          =   255
         Left            =   7320
         TabIndex        =   16
         Top             =   4950
         Width           =   1875
      End
      Begin VB.Label lblCantidad 
         Caption         =   "0"
         Height          =   195
         Left            =   9240
         TabIndex        =   15
         Top             =   4950
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsDatos As ADODB.Recordset
Dim ItemsApellidos() As String
Dim ItemsBloques() As String
Dim ItemsGrupos() As String
Public EstadoACambiar As String
Public ObservacionesACambiar As String
Public OrdenPresidente As String

Private Sub cmdActivoMasivo_Click()

End Sub

Private Sub cmdAplicarFiltro_Click()
Dim Consulta As String
Dim Fila As Integer
Dim nOrder As String
flexEstados.Clear
flexEstados.Rows = 1
flexEstados.Cols = 11
flexEstados.TextMatrix(0, 0) = "Orden"
flexEstados.TextMatrix(0, 1) = "Nombre"
flexEstados.TextMatrix(0, 2) = "Apellido"
flexEstados.TextMatrix(0, 3) = "Bloque Político"
flexEstados.TextMatrix(0, 4) = "Provincia"
flexEstados.TextMatrix(0, 5) = "Estado"
flexEstados.TextMatrix(0, 6) = "Fecha"
flexEstados.TextMatrix(0, 7) = "Hora"
flexEstados.TextMatrix(0, 8) = "IDI"
flexEstados.TextMatrix(0, 9) = "IDL"
flexEstados.TextMatrix(0, 10) = "ID"
flexEstados.ColWidth(0) = 600
flexEstados.ColWidth(3) = 2000
flexEstados.ColWidth(4) = 2000
flexEstados.ColWidth(5) = 1600
flexEstados.ColWidth(8) = 0 '400
flexEstados.ColWidth(9) = 0 '400
flexHistorico.Clear
flexHistorico.Rows = 1
flexHistorico.Cols = 11
flexHistorico.TextMatrix(0, 0) = "Orden"
flexHistorico.TextMatrix(0, 1) = "Nombre"
flexHistorico.TextMatrix(0, 2) = "Apellido"
flexHistorico.TextMatrix(0, 3) = "Bloque Político"
flexHistorico.TextMatrix(0, 4) = "Provincia"
flexHistorico.TextMatrix(0, 5) = "Estado"
flexHistorico.TextMatrix(0, 6) = "Fecha"
flexHistorico.TextMatrix(0, 7) = "Hora"
flexHistorico.TextMatrix(0, 8) = "IDI"
flexHistorico.TextMatrix(0, 9) = "IDL"
flexHistorico.TextMatrix(0, 10) = "ID"
flexHistorico.ColWidth(0) = 600
flexHistorico.ColWidth(3) = 2000
flexHistorico.ColWidth(4) = 2000
flexHistorico.ColWidth(5) = 1600
flexHistorico.ColWidth(8) = 0 '400
flexHistorico.ColWidth(9) = 0 '400
flexHistorico.ColWidth(6) = 0 '400
flexHistorico.ColWidth(7) = 0 '400
'Consulta = "SELECT Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, Legisladores.grupo_politico, estados.descripcion AS estado_actual," & _
            "legisladores_estado.fecha_vigencia, legisladores_estado.hora" & _
            " FROM Legisladores INNER JOIN " & _
            "legisladores_estado ON legisladores_estado.id_legislador = Legisladores.id INNER JOIN " & _
            "estados ON estados.id_estado = legisladores_estado.estado "
Consulta = "SELECT legisladores_para_actualizar.*, Legisladores.codigo_persona, distritos.distrito FROM legisladores_para_actualizar INNER JOIN Legisladores ON Legisladores.id = legisladores_para_actualizar.id LEFT JOIN distritos ON distritos.id_distrito = Legisladores.distrito"
Consulta = Consulta & " " & AplicarFiltros1() & " ORDER BY legisladores_para_actualizar.apellido, legisladores_para_actualizar.nombre"
Fila = 0
Set rsDatos = New ADODB.Recordset
SetearRs Consulta, rsDatos
While (rsDatos.EOF = False)
    With rsDatos
        Fila = Fila + 1
        flexEstados.AddItem ""
        flexEstados.TextMatrix(Fila, 0) = .Fields("deskid")
        flexEstados.TextMatrix(Fila, 1) = .Fields("nombre")
        flexEstados.TextMatrix(Fila, 2) = .Fields("apellido")
        flexEstados.TextMatrix(Fila, 3) = IIf(IsNull(.Fields("bloque_politico")), "", .Fields("bloque_politico"))
        flexEstados.TextMatrix(Fila, 4) = IIf(IsNull(.Fields("distrito")), "", .Fields("distrito"))
        flexEstados.TextMatrix(Fila, 5) = .Fields("descripcion")
        flexEstados.TextMatrix(Fila, 6) = .Fields("fecha_vigencia")
        flexEstados.TextMatrix(Fila, 7) = .Fields("hora")
        flexEstados.TextMatrix(Fila, 8) = .Fields("actualizacion")
        flexEstados.TextMatrix(Fila, 9) = .Fields("id")
        If IsNull(.Fields("codigo_persona")) Then
            flexEstados.TextMatrix(Fila, 10) = "Sin ID"
        Else
            flexEstados.TextMatrix(Fila, 10) = .Fields("codigo_persona")
        End If
    End With
    rsDatos.MoveNext
Wend
lblCantidad.Caption = rsDatos.RecordCount
Set rsDatos = Nothing
Consulta = "SELECT legisladores_historico.*, Legisladores.codigo_persona FROM legisladores_historico INNER JOIN Legisladores ON Legisladores.id = legisladores_historico.id"
Consulta = Consulta & " " & AplicarFiltros2() & " ORDER BY legisladores_historico.id,legisladores_historico.apellido,legisladores_historico.actualizacion DESC,legisladores_historico.nombre"
Fila = 0
Set rsDatos = New ADODB.Recordset
SetearRs Consulta, rsDatos
While (rsDatos.EOF = False)
    With rsDatos
        Fila = Fila + 1
        flexHistorico.AddItem ""
        flexHistorico.TextMatrix(Fila, 0) = .Fields("deskid")
        flexHistorico.TextMatrix(Fila, 1) = .Fields("nombre")
        flexHistorico.TextMatrix(Fila, 2) = .Fields("apellido")
        flexHistorico.TextMatrix(Fila, 3) = IIf(IsNull(.Fields("bloque_politico")), "", .Fields("bloque_politico"))
        flexHistorico.TextMatrix(Fila, 4) = IIf(IsNull(.Fields("grupo_politico")), "", .Fields("grupo_politico"))
        flexHistorico.TextMatrix(Fila, 5) = .Fields("descripcion")
        flexHistorico.TextMatrix(Fila, 6) = .Fields("fecha_vigencia")
        flexHistorico.TextMatrix(Fila, 7) = .Fields("hora")
        flexHistorico.TextMatrix(Fila, 8) = .Fields("actualizacion")
        flexHistorico.TextMatrix(Fila, 9) = .Fields("id")
        flexHistorico.TextMatrix(Fila, 10) = IIf(IsNull(.Fields("codigo_persona")), "NULO", .Fields("codigo_persona"))
    End With
    rsDatos.MoveNext
Wend
Set rsDatos = Nothing
End Sub
Private Function AplicarFiltros1() As String
Dim cApellido As String
Dim cBloquePolitico As String
Dim cGrupoPolitico As String
Dim cEstado As String
Dim TotalString As String
If txtApellido.Text = "Escriba un apellido o parte de él para buscar" Then
    cApellido = ""
Else
    cApellido = txtApellido.Text
End If
TotalString = TotalString & " WHERE legisladores_para_actualizar.apellido LIKE '" & cApellido & "%'"
If txtBloques.Text <> "Escriba un bloque o parte de él para buscar" And txtBloques.Text <> "" Then
    cBloquePolitico = txtBloques.Text
    TotalString = TotalString & " AND legisladores_para_actualizar.bloque_politico = '" & cBloquePolitico & "'"
End If
If txtGrupos.Text <> "Escriba una provincia o parte de ella" And txtGrupos.Text <> "" Then
    cGrupoPolitico = txtGrupos.Text
    TotalString = TotalString & " AND distritos.distrito = '" & cGrupoPolitico & "'"
End If
If cmbEstado.ListIndex <> -1 And cmbEstado.Text <> "Cualquier estado" Then
    TotalString = TotalString & " AND legisladores_para_actualizar.descripcion = '" & cmbEstado.Text & "'"
End If
AplicarFiltros1 = TotalString
End Function
Private Function AplicarFiltros2() As String
Dim cApellido As String
Dim cBloquePolitico As String
Dim cGrupoPolitico As String
Dim cEstado As String
Dim TotalString As String
If txtApellido.Text = "Escriba un apellido o parte de él para buscar" Then
    cApellido = ""
Else
    cApellido = txtApellido.Text
End If
TotalString = TotalString & " WHERE legisladores_historico.apellido LIKE '" & cApellido & "%'"
If txtBloques.Text <> "Escriba un bloque o parte de él para buscar" And txtBloques.Text <> "" Then
    cBloquePolitico = txtBloques.Text
    TotalString = TotalString & " AND legisladores_historico.bloque_politico = '" & cBloquePolitico & "'"
End If
If txtGrupos.Text <> "Escriba un grupo o parte de él para buscar" And txtGrupos.Text <> "" Then
    cGrupoPolitico = txtGrupos.Text
    TotalString = TotalString & " AND legisladores_historico.grupo_politico = '" & cGrupoPolitico & "'"
End If
If cmbEstado.ListIndex <> -1 And cmbEstado.Text <> "Cualquier estado" Then
    TotalString = TotalString & " AND legisladores_historico.descripcion = '" & cmbEstado.Text & "'"
End If
AplicarFiltros2 = TotalString
End Function
Private Sub flexEstados_DblClick()
Dim rSelected As Integer
Dim Consulta As String
Dim EraActivo As Boolean
EstadoACambiar = ""
ObservacionesACambiar = ""
OrdenPresidente = ""
rSelected = flexEstados.RowSel
If flexEstados.Rows > 1 Then
    Dim xF As Form
    Dim i As Integer
    Dim Activo As Boolean
    Set xF = New frmCambiarEstado
    xF.Show vbModal
    If EstadoACambiar <> "" Then
        Dim NOrden As Integer
        Dim nIDEstado As Integer
        Dim cCons As String
        Set rsDatos = New ADODB.Recordset
        Consulta = "SELECT activo FROM estados WHERE descripcion = '" & flexEstados.TextMatrix(rSelected, 5) & "'"
        SetearRs Consulta, rsDatos
        If (rsDatos.EOF = False) Then
            If rsDatos.Fields("activo") = 1 Then
                EraActivo = True
            Else
                EraActivo = False
            End If
        Else
            EraActivo = False
        End If
        rsDatos.Close
        Set rsDatos = Nothing
        Set rsDatos = New ADODB.Recordset
        Consulta = "SELECT id_estado,activo FROM estados WHERE descripcion = '" & EstadoACambiar & "'"
        SetearRs Consulta, rsDatos
        If (rsDatos.EOF = False) Then
            If rsDatos.Fields("activo") = 1 Then
                Activo = True
            Else
                Activo = False
            End If
            nIDEstado = rsDatos.Fields("id_estado")
        Else
            MsgBox ("Error. Se modificó manualmente la base de datos y provocó un error crítico")
        End If
        If Activo = True Then
            If EraActivo = True Then
                NOrden = Int(flexEstados.TextMatrix(rSelected, 0))
            Else
                NOrden = ObtenerSlotLibre()
            End If
        Else
            cCons = "INSERT INTO legisladores_estado(id_legislador,fecha_vigencia,hora,estado,numero_orden_activacion,observaciones,OrdenPresidente) " & _
            " VALUES('" & flexEstados.TextMatrix(rSelected, 9) & "',CONVERT(nvarchar(11), DAY({ fn NOW() })) + '/' + CONVERT(nvarchar(11), MONTH({ fn NOW() })) + '/' + CONVERT(nvarchar(11), YEAR({ fn NOW() }))," & _
            " CONVERT(nvarchar(2), DATEPART(hh, GETDATE())) + ':' + CONVERT(nvarchar(2), DATEPART(mi, GETDATE())) + ':' + CONVERT(nvarchar(2),DATEPART(ss, GETDATE()))," & _
            nIDEstado & "," & "-1" & ",'" & ObservacionesACambiar & "'," & Int(OrdenPresidente) & ")"
            EjecutarSQL (cCons)
            'EjecutarSQL ("EXEC NECAR_HCDN.HCDN_UpdatePersonStatus " & flexEstados.TextMatrix(rSelected, 9) & "," & Str(nIDEstado))
        End If
        If Activo = True Then
            If NOrden <> -1 Then
                cCons = "INSERT INTO legisladores_estado(id_legislador,fecha_vigencia,hora,estado,numero_orden_activacion,observaciones,OrdenPresidente) " & _
                " VALUES('" & flexEstados.TextMatrix(rSelected, 9) & "',CONVERT(nvarchar(11), DAY({ fn NOW() })) + '/' + CONVERT(nvarchar(11), MONTH({ fn NOW() })) + '/' + CONVERT(nvarchar(11), YEAR({ fn NOW() }))," & _
                " CONVERT(nvarchar(2), DATEPART(hh, GETDATE())) + ':' + CONVERT(nvarchar(2), DATEPART(mi, GETDATE())) + ':' + CONVERT(nvarchar(2),DATEPART(ss, GETDATE()))," & _
                nIDEstado & "," & NOrden & ",'" & ObservacionesACambiar & "'," & Int(OrdenPresidente) & ")"
                EjecutarSQL (cCons)
            Else
                MsgBox "Ya tiene 257 diputados activos!", vbCritical
            End If
            'EjecutarSQL ("EXEC NECAR_HCDN.dbo.HCDN_UpdatePersonStatus " & flexEstados.TextMatrix(rSelected, 9) & "," & Str(nIDEstado))
        End If
    End If
End If
End Sub
Private Function ObtenerSlotLibre() As Integer
Dim num As Integer
Dim Numero As Integer
Dim rsSlots As New ADODB.Recordset
Numero = -1
For num = 1 To 257
    SetearRs "SELECT * FROM legisladores_activos WHERE deskid = " & num, rsSlots
    If rsSlots.RecordCount <= 0 Then
        Numero = num
        num = 257
    End If
    rsSlots.Close
    Set rsSlots = Nothing
Next num
ObtenerSlotLibre = Numero
End Function
Private Sub Form_Load()
Dim rsTemp As ADODB.Recordset
Dim cCons As String
Dim r As Integer
Dim nDefault As Integer
Dim RsDip As ADODB.Recordset
Dim Buff As String
Buff = vbCrLf
flexEstados.ColWidth(6) = 0
flexEstados.ColWidth(7) = 0
flexHistorico.ColWidth(6) = 0
flexHistorico.ColWidth(7) = 0
Set RsDip = New ADODB.Recordset
'SetearRs "SELECT id AS id_legislador ,apellido,nombre FROM Legisladores WHERE (id NOT IN (SELECT id_legislador FROM legisladores_estado))", RsDip
SetearRs "SELECT id AS id_legislador ,apellido,nombre FROM Legisladores WHERE (id NOT IN (SELECT id_legislador FROM legisladores_estado)) UNION SELECT id AS id_legislador,apellido,nombre FROM legisladores_para_actualizar WHERE descripcion = 'Nuevo'", RsDip
If Not RsDip.EOF Then
    Dim xFrm As frmAvisos
    Set xFrm = New frmAvisos
    While Not RsDip.EOF
        xFrm.lstNuevosDiputados.AddItem RsDip.Fields("apellido") & ", " & RsDip.Fields("nombre")
        RsDip.MoveNext
    Wend
    xFrm.Show vbModal, Me
    If RTA_Activar = True Then
        r = vbYes
    Else
        r = vbNo
    End If
    'r = MsgBox("Se detectaron nuevos diputados(" & RsDip.RecordCount & "): " & Buff & vbCrLf & vbCrLf & "¿Desea iniciarlos como ACTIVOS?" & vbCrLf & "De lo contrario figurarán como NUEVOS." & vbCrLf & "Nota: Si uno o más de los diputados excede el límite de 257 activos automáticamente será asignado como NUEVO.", vbYesNo)
    If r = vbYes Then
        nDefault = 1
    Else
        nDefault = 8
    End If
    RsDip.Close
    Dim estadoNuevo As Boolean
    Dim cConsulta As String
    If nDefault = 8 Then
        cConsulta = "SELECT id AS id_legislador ,apellido,nombre FROM Legisladores WHERE (id NOT IN (SELECT id_legislador FROM legisladores_estado))"
    Else
        cConsulta = "SELECT id AS id_legislador ,apellido,nombre FROM Legisladores WHERE (id NOT IN (SELECT id_legislador FROM legisladores_estado)) UNION SELECT id AS id_legislador,apellido,nombre FROM legisladores_para_actualizar WHERE descripcion = 'Nuevo'"
    End If
    SetearRs cConsulta, RsDip
    While Not RsDip.EOF
        If nDefault = 8 Or ObtenerSlotLibre = -1 Then
            cCons = "INSERT INTO legisladores_estado(id_legislador,fecha_vigencia,hora,estado,numero_orden_activacion,observaciones,OrdenPresidente) " & _
            " VALUES('" & RsDip.Fields("id_legislador") & "',CONVERT(nvarchar(11), DAY({ fn NOW() })) + '/' + CONVERT(nvarchar(11), MONTH({ fn NOW() })) + '/' + CONVERT(nvarchar(11), YEAR({ fn NOW() }))," & _
            " CONVERT(nvarchar(2), DATEPART(hh, GETDATE())) + ':' + CONVERT(nvarchar(2), DATEPART(mi, GETDATE())) + ':' + CONVERT(nvarchar(2),DATEPART(ss, GETDATE()))," & _
            "8" & "," & "-1" & ",'" & "Observacion" & "',99)"
        Else
            cCons = "INSERT INTO legisladores_estado(id_legislador,fecha_vigencia,hora,estado,numero_orden_activacion,observaciones,OrdenPresidente) " & _
            " VALUES('" & RsDip.Fields("id_legislador") & "',CONVERT(nvarchar(11), DAY({ fn NOW() })) + '/' + CONVERT(nvarchar(11), MONTH({ fn NOW() })) + '/' + CONVERT(nvarchar(11), YEAR({ fn NOW() }))," & _
            " CONVERT(nvarchar(2), DATEPART(hh, GETDATE())) + ':' + CONVERT(nvarchar(2), DATEPART(mi, GETDATE())) + ':' + CONVERT(nvarchar(2),DATEPART(ss, GETDATE()))," & _
            nDefault & "," & ObtenerSlotLibre() & ",'" & "Observacion" & "',99)"
        End If
        EjecutarSQL (cCons)
        RsDip.MoveNext
    Wend
End If
RsDip.Close
Set RsDip = Nothing
Dim i As Integer
LlenaCombo cmbEstado, "descripcion", "estados", "orden"
cmbEstado.AddItem "Cualquier estado"
LlenaListBox lstApellidos, "apellido", "Legisladores", "apellido"
ReDim ItemsApellidos(lstApellidos.ListCount - 1)
For i = LBound(ItemsApellidos) To UBound(ItemsApellidos)
    ItemsApellidos(i) = lstApellidos.List(i)
Next i
LlenaListBox lstBloques, "Bloque_Político", "Bloques", "Bloque_Político"
ReDim ItemsBloques(lstBloques.ListCount - 1)
For i = LBound(ItemsBloques) To UBound(ItemsBloques)
    ItemsBloques(i) = lstBloques.List(i)
Next i
LlenaListBox lstGrupos, "distrito", "distritos", "distrito"
ReDim ItemsGrupos(lstGrupos.ListCount - 1)
For i = LBound(ItemsGrupos) To UBound(ItemsGrupos)
    ItemsGrupos(i) = lstGrupos.List(i)
Next i
flexEstados.TextMatrix(0, 0) = "Orden"
flexEstados.TextMatrix(0, 1) = "Nombre"
flexEstados.TextMatrix(0, 2) = "Apellido"
flexEstados.TextMatrix(0, 3) = "Bloque Político"
flexEstados.TextMatrix(0, 4) = "Provincia"
flexEstados.TextMatrix(0, 5) = "Estado"
flexEstados.TextMatrix(0, 6) = "Fecha"
flexEstados.TextMatrix(0, 7) = "Hora"
flexEstados.TextMatrix(0, 8) = "IDI"
flexEstados.TextMatrix(0, 9) = "IDL"
flexEstados.TextMatrix(0, 10) = "ID"
flexEstados.ColWidth(0) = 600
flexEstados.ColWidth(3) = 2000
flexEstados.ColWidth(4) = 2000
flexEstados.ColWidth(5) = 1400
flexEstados.ColWidth(8) = 0 '400
flexEstados.ColWidth(9) = 0 '400
flexHistorico.Clear
flexHistorico.Rows = 1
flexHistorico.Cols = 11
flexHistorico.TextMatrix(0, 0) = "Orden"
flexHistorico.TextMatrix(0, 1) = "Nombre"
flexHistorico.TextMatrix(0, 2) = "Apellido"
flexHistorico.TextMatrix(0, 3) = "Bloque Político"
flexHistorico.TextMatrix(0, 4) = "Provincia"
flexHistorico.TextMatrix(0, 5) = "Estado"
flexHistorico.TextMatrix(0, 6) = "Fecha"
flexHistorico.TextMatrix(0, 7) = "Hora"
flexHistorico.TextMatrix(0, 8) = "IDI"
flexHistorico.TextMatrix(0, 9) = "IDL"
flexHistorico.TextMatrix(0, 10) = "ID"
flexHistorico.ColWidth(0) = 600
flexHistorico.ColWidth(3) = 2000
flexHistorico.ColWidth(4) = 2000
flexHistorico.ColWidth(5) = 1600
flexHistorico.ColWidth(8) = 0 '400
flexHistorico.ColWidth(9) = 0 '400
flexHistorico.ColWidth(6) = 0 '400
flexHistorico.ColWidth(7) = 0 '400
End Sub
Public Sub LlenaCombo(CMB As ComboBox, campo As String, tabla As String, order As String)
Set rsDatos = New ADODB.Recordset
SetearRs "SELECT " & campo & " FROM " & tabla & " ORDER BY " & order, rsDatos
While Not rsDatos.EOF
    CMB.AddItem rsDatos.Fields(0)
    rsDatos.MoveNext
Wend
rsDatos.Close
Set rsDatos = Nothing
End Sub
Public Sub LlenaListBox(CMB As ListBox, campo As String, tabla As String, order As String)
Set rsDatos = New ADODB.Recordset
SetearRs "SELECT " & campo & " FROM " & tabla & " ORDER BY " & order, rsDatos
While Not rsDatos.EOF
    CMB.AddItem rsDatos.Fields(0)
    rsDatos.MoveNext
Wend
rsDatos.Close
Set rsDatos = Nothing
End Sub

Private Sub Image1_Click()

End Sub
Private Sub lstApellidos_Click()
If lstApellidos.ListIndex <> -1 Then
    txtApellido.Text = lstApellidos.List(lstApellidos.ListIndex)
    txtApellido.ForeColor = vbBlack
End If
End Sub
Private Sub lstBloques_Click()
If lstBloques.ListIndex <> -1 Then
    txtBloques.Text = lstBloques.List(lstBloques.ListIndex)
    txtBloques.ForeColor = vbBlack
End If
End Sub
Private Sub lstGrupos_Click()
If lstGrupos.ListIndex <> -1 Then
    txtGrupos.Text = lstGrupos.List(lstGrupos.ListIndex)
    txtGrupos.ForeColor = vbBlack
End If
End Sub

Private Sub txtApellido_Click()
If txtApellido.Text = "Escriba un apellido o parte de él para buscar" Then
    txtApellido.Text = ""
    txtApellido.ForeColor = vbBlack
End If
End Sub
Private Sub txtApellido_KeyPress(KeyAscii As Integer)
If txtApellido.Text = "" Or txtApellido.Text = "Escriba un apellido o parte de él para buscar" Then
    txtApellido.Text = ""
    txtApellido.ForeColor = vbBlack
End If
End Sub
Private Sub txtApellido_KeyUp(KeyCode As Integer, Shift As Integer)
Dim cPos As Integer
If txtApellido.Text = "" Then
    lstApellidos.Clear
    For i = LBound(ItemsApellidos) To UBound(ItemsApellidos)
        lstApellidos.AddItem ItemsApellidos(i)
    Next i
    txtApellido.ForeColor = &H808080
    txtApellido.Text = "Escriba un apellido o parte de él para buscar"
End If
If txtApellido.Text <> "Escriba un apellido o parte de él para buscar" Then
    lstApellidos.Clear
    For i = LBound(ItemsApellidos) To UBound(ItemsApellidos)
        cPos = InStr(LCase(ItemsApellidos(i)), LCase(txtApellido.Text))
        If cPos = 1 Then
            lstApellidos.AddItem ItemsApellidos(i)
        End If
    Next i
End If
End Sub
Private Sub txtBloques_Click()
If txtBloques.Text = "Escriba un bloque o parte de él para buscar" Then
    txtBloques.Text = ""
    txtBloques.ForeColor = vbBlack
End If
End Sub
Private Sub txtBloques_KeyPress(KeyAscii As Integer)
If txtBloques.Text = "" Or txtBloques.Text = "Escriba un bloque o parte de él para buscar" Then
    txtBloques.Text = ""
    txtBloques.ForeColor = vbBlack
End If
End Sub
Private Sub txtBloques_KeyUp(KeyCode As Integer, Shift As Integer)
Dim cPos As Integer
If txtBloques.Text = "" Then
    lstBloques.Clear
    For i = LBound(ItemsBloques) To UBound(ItemsBloques)
        lstBloques.AddItem ItemsBloques(i)
    Next i
    txtBloques.ForeColor = &H808080
    txtBloques.Text = "Escriba un bloque o parte de él para buscar"
End If
If txtBloques.Text <> "Escriba un bloque o parte de él para buscar" Then
    lstBloques.Clear
    For i = LBound(ItemsBloques) To UBound(ItemsBloques)
        cPos = InStr(LCase(ItemsBloques(i)), LCase(txtBloques.Text))
        If cPos = 1 Then
            lstBloques.AddItem ItemsBloques(i)
        End If
    Next i
End If
End Sub
Private Sub txtGrupos_Click()
If txtGrupos.Text = "Escriba una provincia o parte de ella" Then
    txtGrupos.Text = ""
    txtGrupos.ForeColor = vbBlack
End If
End Sub
Private Sub txtGrupos_KeyPress(KeyAscii As Integer)
If txtGrupos.Text = "" Or txtGrupos.Text = "Escriba una provincia o parte de ella" Then
    txtGrupos.Text = ""
    txtGrupos.ForeColor = vbBlack
End If
End Sub
Private Sub txtGrupos_KeyUp(KeyCode As Integer, Shift As Integer)
Dim cPos As Integer
If txtGrupos.Text = "" Then
    lstGrupos.Clear
    For i = LBound(ItemsGrupos) To UBound(ItemsGrupos)
        lstGrupos.AddItem ItemsGrupos(i)
    Next i
    txtGrupos.ForeColor = &H808080
    txtGrupos.Text = "Escriba una provincia o parte de ella"
End If
If txtGrupos.Text <> "Escriba una provincia o parte de ella" Then
    lstGrupos.Clear
    For i = LBound(ItemsGrupos) To UBound(ItemsGrupos)
        cPos = InStr(LCase(ItemsGrupos(i)), LCase(txtGrupos.Text))
        If cPos = 1 Then
            lstGrupos.AddItem ItemsGrupos(i)
        End If
    Next i
End If
End Sub
