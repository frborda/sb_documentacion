VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmEstadisticasGenerales 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estadísticas Generales"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00404040&
      Caption         =   "Exclusión / Inclusión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1185
      Left            =   210
      TabIndex        =   17
      Top             =   6690
      Visible         =   0   'False
      Width           =   5610
      Begin VB.OptionButton optExcluir 
         BackColor       =   &H00404040&
         Caption         =   "Excluír a los diputados listados"
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   210
         TabIndex        =   19
         Top             =   330
         Value           =   -1  'True
         Width           =   3465
      End
      Begin VB.OptionButton optIncluir 
         BackColor       =   &H00404040&
         Caption         =   "Únicamente listar a los diputados listados"
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   210
         TabIndex        =   18
         Top             =   630
         Width           =   3915
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1185
      Left            =   90
      TabIndex        =   11
      Top             =   5250
      Width           =   5985
      Begin Proyecto1.ButtonOffice cmdCopyFecha 
         Height          =   315
         Left            =   3180
         TabIndex        =   12
         Top             =   720
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BackColor       =   12230304
         Caption         =   "&C"
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
      Begin MSComCtl2.DTPicker dtDesde 
         Height          =   315
         Left            =   780
         TabIndex        =   13
         Top             =   300
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         Format          =   64028673
         CurrentDate     =   40666
      End
      Begin MSComCtl2.DTPicker dtHasta 
         Height          =   315
         Left            =   780
         TabIndex        =   14
         Top             =   720
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         Format          =   64028673
         CurrentDate     =   40666
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Left            =   180
         TabIndex        =   16
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Left            =   240
         TabIndex        =   15
         Top             =   780
         Width           =   555
      End
   End
   Begin ComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   7200
      TabIndex        =   10
      Top             =   6210
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin Proyecto1.ButtonOffice cmdImprimir 
      Height          =   495
      Left            =   7200
      TabIndex        =   9
      Top             =   5700
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "Imprimir"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   "Ordenamiento"
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
      Height          =   915
      Left            =   90
      TabIndex        =   5
      Top             =   4260
      Width           =   5985
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
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   390
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.OptionButton optAfirmativos 
         BackColor       =   &H00404040&
         Caption         =   "Votos Afirmativos"
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
         Height          =   285
         Left            =   1500
         TabIndex        =   7
         Top             =   390
         Width           =   1965
      End
      Begin VB.OptionButton optAusentes 
         BackColor       =   &H00404040&
         Caption         =   "Ausencias"
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
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Top             =   390
         Width           =   1575
      End
   End
   Begin Proyecto1.ButtonOffice cmdMoverDerecha 
      Height          =   435
      Left            =   6180
      TabIndex        =   4
      Top             =   120
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   767
      BackColor       =   12230304
      Caption         =   ">"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicOpacity      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Lista de Exclusión"
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
      Height          =   5595
      Left            =   7200
      TabIndex        =   2
      Top             =   30
      Width           =   5535
      Begin VB.ListBox lstSeleccionados 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4890
         ItemData        =   "frmEstadisticasGenerales.frx":0000
         Left            =   60
         List            =   "frmEstadisticasGenerales.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   5415
      End
      Begin Proyecto1.ButtonOffice cmdLimpiar 
         Height          =   315
         Left            =   90
         TabIndex        =   20
         Top             =   5190
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         BackColor       =   8421631
         Caption         =   "Limpiar lista"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         PicOpacity      =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Diputados"
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
      Height          =   4125
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   5985
      Begin VB.CheckBox chkActivos 
         BackColor       =   &H00404040&
         Caption         =   "Listar únicamente diputados activos"
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
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   270
         Width           =   4395
      End
      Begin VB.ListBox lstDiputados 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         ItemData        =   "frmEstadisticasGenerales.frx":0004
         Left            =   120
         List            =   "frmEstadisticasGenerales.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   660
         Width           =   5775
      End
   End
   Begin Proyecto1.ButtonOffice cmdMoverIzquierda 
      Height          =   435
      Left            =   6180
      TabIndex        =   21
      Top             =   660
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   767
      BackColor       =   12230304
      Caption         =   "<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicOpacity      =   0
   End
End
Attribute VB_Name = "frmEstadisticasGenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonOffice1_Click()
CargaDiputados
End Sub
Private Sub cmdBorrarItem_Click()
If lstSeleccionados.ListIndex <> -1 Then
    lstSeleccionados.RemoveItem lstSeleccionados.ListIndex
Else
    MsgBox "Ningún diputado seleccionado", vbCritical
End If
End Sub

Private Sub chkActivos_Click()
If chkActivos.Value = vbChecked Then
    CargaDiputadosActivos
Else
    CargaDiputados
End If
RevisaListas
End Sub
Private Sub cmdCopyFecha_Click()
dtHasta.Day = dtDesde.Day
dtHasta.Month = dtDesde.Month
dtHasta.Year = dtDesde.Year
End Sub
Private Sub cmdImprimir_Click()
Dim X As String
X = GetFiltro
cmdImprimir.Enabled = False
DoEvents
PreSincronizacion
Dim RsTemp As ADODB.Recordset
Dim rpt As New rptEstadisticasGenerales
Dim i As Integer
rpt.lblVotacion.Caption = "Datos Estadísticos desde el " & dtDesde.Day & "/" & dtDesde.Month & "/" & dtDesde.Year
rpt.lblVotacion.Caption = rpt.lblVotacion.Caption & " hasta el " & dtHasta.Day & "/" & dtHasta.Month & "/" & dtHasta.Year
If optApellido.Value = True Then
    rpt.lblVotacion.Caption = rpt.lblVotacion.Caption + " ordenado alfabéticamente"
ElseIf optAfirmativos.Value = True Then
    rpt.lblVotacion.Caption = rpt.lblVotacion.Caption + " ordenado por votos afirmativos"
ElseIf optAusentes.Value = True Then
    rpt.lblVotacion.Caption = rpt.lblVotacion.Caption + " ordenado por ausencias"
End If
Set RsTemp = New ADODB.Recordset
SetearRs "SELECT tbEstadisticas.*,SUBSTRING(tbEstadisticas.Bloque_Politico,1,34) AS Bloque_Politico2, Legisladores.apellido + ', ' + Legisladores.nombre AS CDiputado, tbEstadisticas.Afirmativos + tbEstadisticas.Negativos + tbEstadisticas.Abstenciones + tbEstadisticas.Ausentes + tbEstadisticas.Presidente_Afirmativos + tbEstadisticas.Presidente_Negativos + tbEstadisticas.Presidente_SinVoto + tbEstadisticas.PDL_Presentes + tbEstadisticas.PDL_Ausentes AS Total FROM tbEstadisticas INNER JOIN Legisladores ON Legisladores.id = tbEstadisticas.id INNER JOIN distritos on distritos.id_distrito = Legisladores.distrito " & GetFiltro & GetOrden, RsTemp
If Not RsTemp.EOF Then
    Set rpt.DataControl1.Recordset = RsTemp
    rpt.Run False
    For i = 0 To (rpt.Pages.Count - 1)
        rpt.Pages(i).Width = 300
    Next i
End If
RsTemp.Close
Set RsTemp = Nothing
rpt.PrintReport True
cmdImprimir.Enabled = True
DoEvents
End Sub
Private Function GetFiltro() As String
Dim Buff As String
Dim id() As String
Dim Operador As String
Dim Operador2 As String
If lstSeleccionados.ListCount = 0 Then
    GetFiltro = ""
    Exit Function
End If
Buff = " WHERE ("
If optExcluir.Value = True Then
    Operador = "<>"
    Operador2 = "AND"
Else
    Operador = "="
    Operador2 = "OR"
End If
Dim i As Integer
For i = 0 To lstSeleccionados.ListCount - 1
    id = Split(lstSeleccionados.List(i), ";")
    Buff = Buff & "tbEstadisticas.id " & Operador & "  '" & id(1) & "' "
    If i <> lstSeleccionados.ListCount - 1 Then
        Buff = Buff & Operador2 & " "
    End If
Next i
Buff = Buff & ") "
GetFiltro = Buff
End Function
Private Sub cmdLimpiar_Click()
If chkActivos.Value = vbChecked Then
    CargaDiputadosActivos
Else
    CargaDiputados
End If
lstSeleccionados.Clear
lstDiputados.ListIndex = 0
lstDiputados.SetFocus
CheckLista
End Sub
Private Sub cmdMoverDerecha_Click()
Dim i As Integer
If lstDiputados.ListIndex > -1 Then
    i = lstDiputados.ListIndex
    lstSeleccionados.AddItem (lstDiputados.List(lstDiputados.ListIndex))
    lstDiputados.RemoveItem lstDiputados.ListIndex
    If i > lstDiputados.ListCount - 1 Then
        i = lstDiputados.ListCount - 1
    End If
    lstDiputados.ListIndex = i
    lstDiputados.SetFocus
End If
CheckLista
RevisaListas
End Sub
Private Sub CheckLista()
If lstSeleccionados.ListCount = 0 Then
    optExcluir.Enabled = False
    optIncluir.Enabled = False
Else
    optExcluir.Enabled = True
    optIncluir.Enabled = True
End If
End Sub
Private Sub cmdMoverIzquierda_Click()
Dim i As Integer
If lstSeleccionados.ListIndex > -1 Then
    lstDiputados.AddItem lstSeleccionados.List(lstSeleccionados.ListIndex)
    lstSeleccionados.RemoveItem lstSeleccionados.ListIndex
End If
CheckLista
RevisaListas
End Sub
Private Sub Form_Activate()
lstDiputados.SetFocus
lstDiputados.ListIndex = 0
End Sub
Private Sub Form_Load()
CargaDiputados
dtDesde.Day = Format(Now(), "dd")
dtDesde.Month = Format(Now(), "mm")
dtDesde.Year = Format(Now(), "yyyy")
dtHasta.Day = dtDesde.Day
dtHasta.Month = dtDesde.Month
dtHasta.Year = dtDesde.Year
End Sub
Private Sub CargaDiputados()
Dim RsTemp As ADODB.Recordset
Set RsTemp = New ADODB.Recordset
lstDiputados.Clear
SetearRs "SELECT Legisladores.apellido + ', ' + Legisladores.nombre AS DFull, Legisladores.id FROM Legisladores WHERE Legisladores.tipo = 1 ORDER BY Apellido", RsTemp
While Not RsTemp.EOF
    lstDiputados.AddItem RsTemp.Fields("DFull") & Space(200) & ";" & RsTemp.Fields("id")
    RsTemp.MoveNext
Wend
RsTemp.Close
Set RsTemp = Nothing
End Sub
Private Sub CargaDiputadosActivos()
Dim RsTemp As ADODB.Recordset
Set RsTemp = New ADODB.Recordset
lstDiputados.Clear
SetearRs "SELECT Legisladores_activos.apellido + ', ' + Legisladores_activos.nombre AS DFull, Legisladores_activos.id FROM Legisladores_activos ORDER BY Apellido", RsTemp
While Not RsTemp.EOF
    lstDiputados.AddItem RsTemp.Fields("DFull") & Space(200) & ";" & RsTemp.Fields("id")
    RsTemp.MoveNext
Wend
RsTemp.Close
Set RsTemp = Nothing
End Sub
Private Sub CargaReuniones()
Dim RsTemp As ADODB.Recordset
Dim consulta As String
consulta = "SELECT     Legisladores.id,Legisladores.Bloque_Politico, distritos.distrito, Legisladores.apellido + ', ' + Legisladores.nombre AS Diputado, " & _
" (SELECT     COUNT(DISTINCT Reunion) AS Expr1 From actas " & _
" WHERE      (Período_Legislativo = Período_Legislativo) AND (Número_de_Acta = Número_de_Acta) AND (Sesión = Sesión)) AS Votos " & _
" FROM         Legisladores " & _
" INNER JOIN distritos ON Legisladores.distrito = distritos.id_distrito INNER JOIN " & _
                      "detalleactas ON Legisladores.id = detalleactas.Legislador_asignado " & _
"INNER JOIN actas ON actas.Votacion NOT LIKE '%ANULADA%' AND detalleactas.Versión_Acta = actas.Versión_Acta AND detalleactas.Sesión = actas.Sesión AND detalleactas.Nro_de_Acta = actas.Número_de_Acta AND detalleactas.Período_Legislativo = actas.Período_Legislativo " & _
" WHERE     (detalleactas.Versión_Acta = 0) AND Fecha BETWEEN " & GetFecha & _
" AND detalleactas.Operación = 'votnom' GROUP BY Legisladores.id, Legisladores.apellido, Legisladores.nombre, Legisladores.Bloque_Politico, distritos.distrito " & _
" ORDER BY Legisladores.id DESC "
Set RsTemp = New ADODB.Recordset
SetearRs consulta, RsTemp
If Not RsTemp.EOF Then
    prgBar.max = RsTemp.RecordCount
End If
prgBar.Value = 0
While Not RsTemp.EOF
    If RsTemp.Fields("id") = "540" Then
        consulta = consulta
    End If
'    Me.Caption = "Cargando " & UCase(tipo) & " " & rsTemp.Fields("id")
    DoEvents
    consulta = "UPDATE tbEstadisticas SET Reuniones = " & RsTemp.Fields("Votos") & " WHERE id = '" & RsTemp.Fields("id") & "'"
    EjecutarSQL (consulta)
    prgBar.Value = prgBar.Value + 1
    RsTemp.MoveNext
Wend
RsTemp.Close
Set RsTemp = Nothing
'Me.Caption = UCase(tipo) & " OK"
DoEvents
End Sub
Private Sub CargaVotos(campo As String, tipo As String)
Dim RsTemp As ADODB.Recordset
Dim consulta As String
If campo = "Ausentes" Then
    Dim X As String
    X = ""
End If
If (campo = "Ausentes") Then
    consulta = "SELECT     Legisladores.id,Legisladores.Bloque_Politico, distritos.distrito, Legisladores.apellido + ', ' + Legisladores.nombre AS Diputado, COUNT(detalleactas.Resultado) AS Votos" & _
    " FROM         Legisladores " & _
    " INNER JOIN distritos ON Legisladores.distrito = distritos.id_distrito INNER JOIN " & _
                          "detalleactas ON Legisladores.id = detalleactas.Legislador_asignado " & _
    "INNER JOIN actas ON actas.Votacion NOT LIKE '%ANULADA%' AND detalleactas.Versión_Acta = actas.Versión_Acta AND detalleactas.Sesión = actas.Sesión AND detalleactas.Nro_de_Acta = actas.Número_de_Acta AND detalleactas.Período_Legislativo = actas.Período_Legislativo " & _
    " WHERE (detalleactas.Versión_Acta = 0) AND Fecha BETWEEN " & GetFecha & _
    " AND detalleactas.Operación = 'votnom' " & _
    " AND (" & _
    " (detalleactas.Resultado = 'AUSENTE' AND actas.presidente = Legisladores.id AND actas.presidente_habilitado_votar = 0) OR " & _
    " (detalleactas.Resultado = 'AUSENTE' AND actas.presidente <> Legisladores.id) " & _
    " )" & _
    " GROUP BY Legisladores.id, Legisladores.apellido, Legisladores.nombre, Legisladores.Bloque_Politico, distritos.distrito " & _
    " ORDER BY Legisladores.id DESC "
Else
    consulta = "SELECT     Legisladores.id,Legisladores.Bloque_Politico, distritos.distrito, Legisladores.apellido + ', ' + Legisladores.nombre AS Diputado, COUNT(detalleactas.Resultado) AS Votos" & _
    " FROM         Legisladores " & _
    " INNER JOIN distritos ON Legisladores.distrito = distritos.id_distrito INNER JOIN " & _
                          "detalleactas ON Legisladores.id = detalleactas.Legislador_asignado " & _
    "INNER JOIN actas ON actas.Votacion NOT LIKE '%ANULADA%' AND detalleactas.Versión_Acta = actas.Versión_Acta AND detalleactas.Sesión = actas.Sesión AND detalleactas.Nro_de_Acta = actas.Número_de_Acta AND detalleactas.Período_Legislativo = actas.Período_Legislativo " & _
    " WHERE (detalleactas.Versión_Acta = 0) AND detalleactas.Resultado = '" & tipo & "' AND Fecha BETWEEN " & GetFecha & _
    " AND detalleactas.Operación = 'votnom' GROUP BY Legisladores.id, Legisladores.apellido, Legisladores.nombre, Legisladores.Bloque_Politico, distritos.distrito " & _
    " ORDER BY Legisladores.id DESC "
End If
Set RsTemp = New ADODB.Recordset
SetearRs consulta, RsTemp
If Not RsTemp.EOF Then
    prgBar.max = RsTemp.RecordCount
End If
prgBar.Value = 0
While Not RsTemp.EOF
    If RsTemp.Fields("id") = "540" Then
        consulta = consulta
    End If
'    Me.Caption = "Cargando " & UCase(tipo) & " " & rsTemp.Fields("id")
    DoEvents
    consulta = "UPDATE tbEstadisticas SET " & campo & " = " & RsTemp.Fields("Votos") & " WHERE id = '" & RsTemp.Fields("id") & "'"
    EjecutarSQL (consulta)
    prgBar.Value = prgBar.Value + 1
    RsTemp.MoveNext
Wend
RsTemp.Close
Set RsTemp = Nothing
'Me.Caption = UCase(tipo) & " OK"
DoEvents
End Sub
Private Sub CargaVotosPasLis(campo As String, tipo As String)
Dim RsTemp As ADODB.Recordset
Dim consulta As String
consulta = "SELECT     Legisladores.id,Legisladores.Bloque_Politico, distritos.distrito, Legisladores.apellido + ', ' + Legisladores.nombre AS Diputado, COUNT(detalleactas.Resultado) AS Votos" & _
" FROM         Legisladores " & _
" INNER JOIN distritos ON Legisladores.distrito = distritos.id_distrito INNER JOIN " & _
                      "detalleactas ON Legisladores.id = detalleactas.Legislador_asignado " & _
"INNER JOIN actas ON actas.Votacion NOT LIKE '%ANULADA%' AND detalleactas.Versión_Acta = actas.Versión_Acta AND detalleactas.Sesión = actas.Sesión AND detalleactas.Nro_de_Acta = actas.Número_de_Acta AND detalleactas.Período_Legislativo = actas.Período_Legislativo " & _
" WHERE     (detalleactas.Versión_Acta = 0) AND detalleactas.Resultado = '" & tipo & "' AND Fecha BETWEEN " & GetFecha & _
" AND detalleactas.Operación = 'paslis' GROUP BY Legisladores.id, Legisladores.apellido, Legisladores.nombre, Legisladores.Bloque_Politico, distritos.distrito " & _
" ORDER BY Legisladores.id DESC "
Set RsTemp = New ADODB.Recordset
SetearRs consulta, RsTemp
If Not RsTemp.EOF Then
    prgBar.max = RsTemp.RecordCount
End If
prgBar.Value = 0
While Not RsTemp.EOF
    'Me.Caption = "Cargando Pases de Lista " & UCase(tipo) & " " & rsTemp.Fields("id")
    DoEvents
    consulta = "UPDATE tbEstadisticas SET " & campo & " = " & RsTemp.Fields("Votos") & " WHERE id = '" & RsTemp.Fields("id") & "'"
    EjecutarSQL (consulta)
    prgBar.Value = prgBar.Value + 1
    RsTemp.MoveNext
Wend
RsTemp.Close
Set RsTemp = Nothing
'Me.Caption = UCase(tipo) & " OK"
DoEvents
End Sub
Private Sub CargaVotosPresi(campo As String, tipo As String)
Dim RsTemp As ADODB.Recordset
Dim rsGeneral As ADODB.Recordset
Dim consulta As String
Dim BConsulta As String
Set rsGeneral = New ADODB.Recordset
BConsulta = "SELECT id FROM tbEstadisticas ORDER BY id"
SetearRs BConsulta, rsGeneral
If Not rsGeneral.EOF Then
    prgBar.max = rsGeneral.RecordCount
End If
prgBar.Value = 0
While Not rsGeneral.EOF
    Set RsTemp = New ADODB.Recordset
    consulta = "SELECT     COUNT(resultado_voto_presidente) AS Votos " & _
    " From actas " & _
    " WHERE Fecha BETWEEN " & GetFecha & " AND Versión_Acta = 0 AND Tipo_de_Operación = 'votnom' AND (resultado_voto_presidente = '" & tipo & "') AND (Presidente = '" & rsGeneral.Fields(0) & "') "
    SetearRs consulta, RsTemp
    If Not RsTemp.EOF Then
        EjecutarSQL ("UPDATE tbEstadisticas SET " & campo & " = " & RsTemp.Fields(0) & " WHERE id = " & rsGeneral.Fields(0))
    End If
    RsTemp.Close
    Set RsTemp = Nothing
    'Me.Caption = "Actualizando Presidente Votos " & tipo & " " & rsGeneral.Fields(0)
    DoEvents
    rsGeneral.MoveNext
    prgBar.Value = prgBar.Value + 1
Wend
rsGeneral.Close
Set rsGeneral = Nothing
'Me.Caption = "PRESIDENTE " & UCase(tipo) & " OK"
DoEvents
End Sub
Private Sub CargaVotosDesempate(campo As String, filtro As String, tipo As String)
Dim RsTemp As ADODB.Recordset
Dim rsGeneral As ADODB.Recordset
Dim consulta As String
Dim BConsulta As String
Set rsGeneral = New ADODB.Recordset
BConsulta = "SELECT id," & campo & " FROM tbEstadisticas ORDER BY id"
SetearRs BConsulta, rsGeneral
If Not rsGeneral.EOF Then
    prgBar.max = rsGeneral.RecordCount
End If
prgBar.Value = 0
While Not rsGeneral.EOF
    Set RsTemp = New ADODB.Recordset
    consulta = "SELECT     COUNT(actas.Desempate) AS Votos " & _
    " From actas " & _
    " WHERE Fecha BETWEEN " & GetFecha & " AND Versión_Acta = 0 AND Tipo_de_Operación = 'votnom' AND (Desempate = '" & filtro & "') AND (Presidente = '" & rsGeneral.Fields(0) & "' AND Votacion = '" & tipo & "') "
    SetearRs consulta, RsTemp
    If Not RsTemp.EOF Then
        BConsulta = "UPDATE tbEstadisticas SET " & campo & " = (" & RsTemp.Fields(0) & " + " & rsGeneral.Fields(1) & ") WHERE id = " & rsGeneral.Fields(0)
        EjecutarSQL (BConsulta)
    End If
    RsTemp.Close
    Set RsTemp = Nothing
    prgBar.Value = prgBar.Value + 1
    DoEvents
    rsGeneral.MoveNext
Wend
rsGeneral.Close
Set rsGeneral = Nothing
DoEvents
End Sub
Private Sub CargaPresidenteSinVoto()
Dim RsTemp As ADODB.Recordset
Dim rsGeneral As ADODB.Recordset
Dim consulta As String
Dim BConsulta As String
Set rsGeneral = New ADODB.Recordset
BConsulta = "SELECT id FROM tbEstadisticas ORDER BY id"
SetearRs BConsulta, rsGeneral
If Not rsGeneral.EOF Then
    prgBar.max = rsGeneral.RecordCount
End If
prgBar.Value = 0
While Not rsGeneral.EOF
    Set RsTemp = New ADODB.Recordset
    If rsGeneral.Fields(0) = "540" Then
        consulta = Consultas
    End If
    consulta = "SELECT     COUNT(Presidente) AS Votos " & _
    " From actas " & _
    " WHERE Fecha BETWEEN " & GetFecha & " AND Tipo_de_Operación = 'votnom' AND Versión_Acta = 0 AND  (Desempate = 'No') AND (Presidente = '" & rsGeneral.Fields(0) & "' AND presidente_habilitado_votar = 0)"
    SetearRs consulta, RsTemp
    If Not RsTemp.EOF Then
        BConsulta = "UPDATE tbEstadisticas SET Presidente_SinVoto = (" & RsTemp.Fields(0) & ") WHERE id = " & rsGeneral.Fields(0)
        EjecutarSQL (BConsulta)
    End If
    RsTemp.Close
    Set RsTemp = Nothing
    prgBar.Value = prgBar.Value + 1
    DoEvents
    rsGeneral.MoveNext
Wend
rsGeneral.Close
Set rsGeneral = Nothing
DoEvents
End Sub
Private Sub FixReuniones()
Dim RsTemp As ADODB.Recordset
Dim rsGeneral As ADODB.Recordset
Dim consulta As String
Dim BConsulta As String
Set rsGeneral = New ADODB.Recordset
BConsulta = "SELECT id FROM tbEstadisticas ORDER BY id"
SetearRs BConsulta, rsGeneral
If Not rsGeneral.EOF Then
    prgBar.max = rsGeneral.RecordCount
End If
prgBar.Value = 0
While Not rsGeneral.EOF
    Set RsTemp = New ADODB.Recordset
    consulta = "SELECT     COUNT(DISTINCT Reunion) AS Votos " & _
    " From actas " & _
    " WHERE Fecha BETWEEN " & GetFecha & " AND Tipo_de_Operación = 'votnom' AND presidente_habilitado_votar = 0 AND Versión_Acta = 0 AND Presidente = '" & rsGeneral.Fields(0) & "'"
    'Nueva
    consulta = "SELECT     COUNT(DISTINCT CAST(actas.Reunion AS varchar(50)) + actas.Período_Legislativo) AS Votos " & _
"FROM         actas INNER JOIN " & _
                      "detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND " & _
                      "actas.Sesión = detalleactas.Sesión " & _
" WHERE Actas.Fecha BETWEEN " & GetFecha & " AND (detalleactas.Legislador_asignado = '" & rsGeneral.Fields("id") & "')"
    SetearRs consulta, RsTemp
    If Not RsTemp.EOF Then
        If rsGeneral.Fields(0) = "540" Then
            consulta = consulta
        End If
        'BConsulta = "UPDATE tbEstadisticas SET PDL_Ausentes = (" & rsGeneral.Fields("PDL_Ausentes") & " - " & rsTemp.Fields(0) & ") WHERE id = " & rsGeneral.Fields(0)
        BConsulta = "UPDATE tbEstadisticas SET Reuniones = " & RsTemp.Fields("Votos") & " WHERE id = '" & rsGeneral.Fields("id") & "'"
        EjecutarSQL (BConsulta)
    End If
    RsTemp.Close
    Set RsTemp = Nothing
    prgBar.Value = prgBar.Value + 1
    DoEvents
    rsGeneral.MoveNext
Wend
rsGeneral.Close
Set rsGeneral = Nothing
DoEvents
End Sub
Private Sub FixPresidencias()
Dim RsTemp As ADODB.Recordset
Dim rsGeneral As ADODB.Recordset
Dim consulta As String
Dim BConsulta As String
Set rsGeneral = New ADODB.Recordset
BConsulta = "SELECT id,Ausentes,PDL_Ausentes FROM tbEstadisticas ORDER BY id"
SetearRs BConsulta, rsGeneral
If Not rsGeneral.EOF Then
    prgBar.max = rsGeneral.RecordCount
End If
prgBar.Value = 0
While Not rsGeneral.EOF
    Set RsTemp = New ADODB.Recordset
    consulta = "SELECT     COUNT(Presidente) AS Votos " & _
    " From actas " & _
    " WHERE Fecha BETWEEN " & GetFecha & " AND Tipo_de_Operación = 'votnom' AND presidente_habilitado_votar = 0 AND Versión_Acta = 0 AND Presidente = '" & rsGeneral.Fields(0) & "'"
    SetearRs consulta, RsTemp
    If Not RsTemp.EOF Then
        If rsGeneral.Fields(0) = "540" Then
            consulta = consulta
        End If
        BConsulta = "UPDATE tbEstadisticas SET Ausentes = (" & rsGeneral.Fields("Ausentes") & " - " & RsTemp.Fields(0) & ") WHERE id = " & rsGeneral.Fields(0)
        'BConsulta = "UPDATE tbEstadisticas SET PDL_Ausentes = (" & rsGeneral.Fields("PDL_Ausentes") & " - " & rsTemp.Fields(0) & ") WHERE id = " & rsGeneral.Fields(0)
        EjecutarSQL (BConsulta)
    End If
    RsTemp.Close
    Set RsTemp = Nothing
    prgBar.Value = prgBar.Value + 1
    DoEvents
    rsGeneral.MoveNext
Wend
rsGeneral.Close
Set rsGeneral = Nothing
DoEvents
End Sub
Private Sub PreSincronizacion()
Dim RsTemp As ADODB.Recordset
Dim consulta As String
Me.Caption = "Cargando..."
DoEvents
EjecutarSQL ("DELETE FROM tbEstadisticas")
Set RsTemp = New ADODB.Recordset
consulta = "SELECT     Legisladores.id,(SELECT TOP 1 SUBSTRING(bloque_politico, 0, 36) FROM bloques_voto WHERE (Fecha BETWEEN " & GetFecha & ")AND legislador_asignado = Legisladores.id ORDER BY fecha DESC) AS Bloque_Politico, distritos.distrito, Legisladores.apellido + ', ' + Legisladores.nombre AS Diputado, COUNT(detalleactas.Resultado) AS Votos_Negativos" & _
" FROM         Legisladores " & _
" INNER JOIN distritos ON Legisladores.distrito = distritos.id_distrito INNER JOIN " & _
                      "detalleactas ON Legisladores.id = detalleactas.Legislador_asignado " & _
"INNER JOIN actas ON actas.Votacion NOT LIKE '%ANULADA%' AND detalleactas.Versión_Acta = actas.Versión_Acta AND detalleactas.Sesión = actas.Sesión AND detalleactas.Nro_de_Acta = actas.Número_de_Acta AND detalleactas.Período_Legislativo = actas.Período_Legislativo " & _
" WHERE (detalleactas.Versión_Acta = 0) AND (Fecha BETWEEN " & GetFecha & ")" & _
" AND detalleactas.Operación <> 'paslis' GROUP BY Legisladores.id, Legisladores.apellido, Legisladores.nombre, Legisladores.Bloque_Politico, distritos.distrito " & _
" ORDER BY Legisladores.id DESC "
SetearRs consulta, RsTemp
If RsTemp.RecordCount > 0 Then
    prgBar.max = RsTemp.RecordCount
    While Not RsTemp.EOF
        consulta = "INSERT INTO tbEstadisticas(id,Bloque_Politico,Provincia,Afirmativos,Negativos,Abstenciones,Ausentes,Presidente_Afirmativos,Presidente_Negativos,Presidente_SinVoto,PDL_Presentes,PDL_Ausentes) VALUES(" & _
        "'" & RsTemp.Fields("id") & "','" & RsTemp.Fields("Bloque_Politico") & "','" & RsTemp.Fields("distrito") & "',0,0,0,0,0,0,0,0,0)"
        'frmPreEstadisticas.Caption = "Cargando diputados..." & rsTemp.Fields("id")
        DoEvents
        EjecutarSQL (consulta)
        RsTemp.MoveNext
        prgBar.Value = prgBar.Value + 1
    Wend
    RsTemp.Close
    Set RsTemp = Nothing
    frmPreEstadisticas.Caption = "Diputados OK"
    DoEvents
    Me.Caption = "Cargando votos afirmativos..."
    CargaVotos "Afirmativos", "AFIRMATIVO"
    Me.Caption = "Cargando votos negativos..."
    CargaVotos "Negativos", "NEGATIVO"
    Me.Caption = "Cargando abstenciones..."
    CargaVotos "Abstenciones", "ABSTENCION"
    Me.Caption = "Cargando ausencias..."
    CargaVotos "Ausentes", "AUSENTE"
    Me.Caption = "Cargando ausencias en pase de lista..."
    CargaVotosPasLis "PDL_Ausentes", "AUSENTE"
    Me.Caption = "Cargando presencias en pase de lista..."
    CargaVotosPasLis "PDL_Presentes", "PRESENTE"
    Me.Caption = "Cargando votos afirmativos del presidente..."
    CargaVotosPresi "Presidente_Afirmativos", "s"
    Me.Caption = "Cargando votos negativos del presidente..."
    CargaVotosPresi "Presidente_Negativos", "n"
    Me.Caption = "Cargando desempates afirmativos..."
    CargaVotosDesempate "Presidente_Afirmativos", "Si", "AFIRMATIVO"
    Me.Caption = "Cargando desempates negativos..."
    CargaVotosDesempate "Presidente_Negativos", "Si", "NEGATIVO"
    Me.Caption = "Cargando votaciones donde el presidente no votó..."
    CargaPresidenteSinVoto
    Me.Caption = "Cargando presidencias..."
    FixPresidencias
    Me.Caption = "Cargando reuniones..."
    FixReuniones
    prgBar.Value = 0
    Me.Caption = "Estadísticas"
Else
    MsgBox "No hay datos para las fechas seleccionadas"
End If
End Sub
Private Function GetFecha() As String
Dim Buff As String
Buff = "'" & dtDesde.Day & "/" & dtDesde.Month & "/" & dtDesde.Year & " 00:00:00'"
Buff = Buff & " AND "
Buff = Buff & "'" & dtHasta.Day & "/" & dtHasta.Month & "/" & dtHasta.Year & " 23:59:59'"
GetFecha = Buff
End Function
Private Function GetOrden() As String
Dim Buff As String
If optApellido.Value = True Then
    Buff = " ORDER BY Legisladores.apellido,Legisladores.nombre"
ElseIf optAfirmativos.Value = True Then
    Buff = " ORDER BY tbEstadisticas.Afirmativos DESC,Legisladores.apellido,Legisladores.nombre"
ElseIf optAusentes.Value = True Then
    Buff = " ORDER BY tbEstadisticas.Ausentes DESC,Legisladores.apellido,Legisladores.nombre"
End If
GetOrden = Buff
End Function
Private Sub RevisaListas()
Dim i As Integer
Dim b As Integer
If lstSeleccionados.ListCount > 0 Then
    For i = 0 To lstSeleccionados.ListCount - 1
        For b = 0 To lstDiputados.ListCount - 1
            If lstDiputados.List(b) = lstSeleccionados.List(i) Then
                lstDiputados.RemoveItem b
            End If
        Next b
    Next i
End If
End Sub
