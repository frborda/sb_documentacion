VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmLogIdentificaciones 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Log de Identificaciones"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   13350
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtInicioFecha 
      Height          =   315
      Left            =   5460
      TabIndex        =   13
      Top             =   360
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58130433
      CurrentDate     =   40666
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   13095
      Begin VB.OptionButton optDuplicidades 
         BackColor       =   &H00404040&
         Caption         =   "Mostrar Duplicidades"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   18
         Top             =   1050
         Width           =   3435
      End
      Begin VB.OptionButton optManuales 
         BackColor       =   &H00404040&
         Caption         =   "Mostrar Identificaciones Manuales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   17
         Top             =   750
         Value           =   -1  'True
         Width           =   3435
      End
      Begin VB.TextBox txtFinHora 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   11
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtInicioHora 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   10
         Top             =   300
         Width           =   2535
      End
      Begin VB.CheckBox chkCualquierHora 
         BackColor       =   &H00404040&
         Caption         =   "Cualquier Hora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   9360
         TabIndex        =   8
         Top             =   1140
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin Proyecto1.ButtonOffice cmdBuscar 
         Height          =   975
         Left            =   11940
         TabIndex        =   6
         Top             =   300
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1720
         BackColor       =   12230304
         Caption         =   "&Buscar"
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
      Begin VB.TextBox txtApellido 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1140
         TabIndex        =   0
         Top             =   300
         Width           =   2535
      End
      Begin Proyecto1.ButtonOffice cmdFastCopy 
         Height          =   345
         Left            =   7770
         TabIndex        =   12
         Top             =   720
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   609
         BackColor       =   12230304
         Caption         =   "&C"
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
      Begin MSComCtl2.DTPicker dtFinFecha 
         Height          =   315
         Left            =   5340
         TabIndex        =   14
         Top             =   720
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58130433
         CurrentDate     =   40666
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   4680
         TabIndex        =   9
         Top             =   750
         Width           =   555
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta las"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   8400
         TabIndex        =   5
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde las"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   8340
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde la fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3780
         TabIndex        =   3
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Diputado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   915
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexDiputados 
      Height          =   3915
      Left            =   120
      TabIndex        =   7
      Top             =   2100
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   6906
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   8421504
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Proyecto1.ButtonOffice cmdIprimirListado 
      Height          =   405
      Left            =   7320
      TabIndex        =   15
      Top             =   1620
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   714
      BackColor       =   12230304
      Caption         =   "Imprimir Ordenado Por Fecha"
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
   Begin Proyecto1.ButtonOffice cmdImprAlfab 
      Height          =   405
      Left            =   10080
      TabIndex        =   16
      Top             =   1620
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   714
      BackColor       =   12230304
      Caption         =   "Imprimir Ordenado Alfabéticamente"
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
   Begin Proyecto1.ButtonOffice cmdListadoReducido 
      Height          =   405
      Left            =   120
      TabIndex        =   19
      Top             =   1620
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   714
      BackColor       =   12230304
      Caption         =   "Imprimir listado reducido"
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
End
Attribute VB_Name = "frmLogIdentificaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonOffice1_Click()

End Sub

Private Sub chkCualquierHora_Click()
If chkCualquierHora.Value = vbChecked Then
    txtInicioHora.Enabled = False
    txtFinHora.Enabled = False
Else
    txtInicioHora.Enabled = True
    txtFinHora.Enabled = True
End If
End Sub
Private Sub cmdBuscar_Click()
Dim RsTemp As ADODB.Recordset
Dim consulta As String
Dim Fila As Integer
Dim fecha1 As String
Dim fecha2 As String
'01/12/2011 a 20111201
fecha1 = dtInicioFecha.Year & IIf(Len(dtInicioFecha.Month) = 1, "0" & dtInicioFecha.Month, dtInicioFecha.Month) & IIf(Len(dtInicioFecha.Day) = 1, "0" & dtInicioFecha.Day, dtInicioFecha.Day)
fecha2 = dtFinFecha.Year & IIf(Len(dtFinFecha.Month) = 1, "0" & dtFinFecha.Month, dtFinFecha.Month) & IIf(Len(dtFinFecha.Day) = 1, "0" & dtFinFecha.Day, dtFinFecha.Day)
SetGrilla
Set RsTemp = New ADODB.Recordset
If optDuplicidades.Value = True Then
    consulta = "SELECT CAST(LogIdentificaciones.banca AS varchar(3)) + '  -  ' + CAST(LogIdentificaciones.duplicidad AS varchar(3)) AS banca,LogIdentificaciones.hora, LogIdentificaciones.fecha, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, distritos.distrito AS Provincia FROM LogIdentificaciones INNER JOIN Legisladores ON Legisladores.id = LogIdentificaciones.id_diputado INNER JOIN distritos on Legisladores.distrito = distritos.id_distrito"
    consulta = consulta & " WHERE apellido LIKE '" & txtApellido.Text & "%'"
    consulta = consulta & " AND Fecha >= '" & fecha1 & "' AND Fecha <= '" & fecha2 & "'"
    If (txtInicioHora.Text <> "" And txtFinHora.Text <> "") And chkCualquierHora.Value = vbUnchecked Then
        consulta = consulta & " AND hora >= '" & txtInicioHora.Text & "' AND hora <= '" & txtFinHora.Text & "'"
    End If
    consulta = consulta & " AND duplicidad IS NOT NULL ORDER BY fecha DESC,hora DESC,apellido"
Else
    consulta = "SELECT LogIdentificaciones.banca,LogIdentificaciones.hora, LogIdentificaciones.fecha, Legisladores.nombre, Legisladores.apellido, Legisladores.bloque_politico, distritos.distrito AS Provincia FROM LogIdentificaciones INNER JOIN Legisladores ON Legisladores.id = LogIdentificaciones.id_diputado INNER JOIN distritos on Legisladores.distrito = distritos.id_distrito"
    consulta = consulta & " WHERE apellido LIKE '" & txtApellido.Text & "%'"
    consulta = consulta & " AND Fecha >= '" & fecha1 & "' AND Fecha <= '" & fecha2 & "'"
    If (txtInicioHora.Text <> "" And txtFinHora.Text <> "") And chkCualquierHora.Value = vbUnchecked Then
        consulta = consulta & " AND hora >= '" & txtInicioHora.Text & "' AND hora <= '" & txtFinHora.Text & "'"
    End If
    consulta = consulta & " AND duplicidad IS NULL ORDER BY fecha DESC,hora DESC,apellido"
End If
SetearRs consulta, RsTemp
If RsTemp.EOF Then
    Call MsgBox("No hay resultados", vbCritical, "Consola")
End If
While Not RsTemp.EOF
    flexDiputados.Rows = flexDiputados.Rows + 1
    Fila = flexDiputados.Rows - 2
    flexDiputados.TextMatrix(Fila, 0) = RsTemp.Fields("apellido")
    flexDiputados.TextMatrix(Fila, 1) = RsTemp.Fields("nombre")
    flexDiputados.TextMatrix(Fila, 2) = RsTemp.Fields("bloque_politico")
    flexDiputados.TextMatrix(Fila, 3) = RsTemp.Fields("provincia")
    flexDiputados.TextMatrix(Fila, 4) = RsTemp.Fields("banca")
    flexDiputados.TextMatrix(Fila, 5) = RsTemp.Fields("hora")
    flexDiputados.TextMatrix(Fila, 6) = RsTemp.Fields("fecha")
    RsTemp.MoveNext
Wend
flexDiputados.Rows = flexDiputados.Rows - 1
End Sub
Private Sub cmdFastCopy_Click()
dtFinFecha.Day = dtInicioFecha.Day
dtFinFecha.Month = dtInicioFecha.Month
dtFinFecha.Year = dtInicioFecha.Year
End Sub

Private Sub cmdImprAlfab_Click()
Dim X As New rptLogIdentificados
Dim rs As ADODB.Recordset
Dim consulta As String
Dim Fila As Integer
Dim fecha1 As String
Dim fecha2 As String
Dim Fecha1F As String
Dim Fecha2F As String
Dim i As Integer
Fecha1F = dtInicioFecha.Day & "/" & dtInicioFecha.Month & "/" & dtInicioFecha.Year
Fecha2F = dtFinFecha.Day & "/" & dtFinFecha.Month & "/" & dtFinFecha.Year
fecha1 = dtInicioFecha.Year & IIf(Len(dtInicioFecha.Month) = 1, "0" & dtInicioFecha.Month, dtInicioFecha.Month) & IIf(Len(dtInicioFecha.Day) = 1, "0" & dtInicioFecha.Day, dtInicioFecha.Day)
fecha2 = dtFinFecha.Year & IIf(Len(dtFinFecha.Month) = 1, "0" & dtFinFecha.Month, dtFinFecha.Month) & IIf(Len(dtFinFecha.Day) = 1, "0" & dtFinFecha.Day, dtFinFecha.Day)
Set RsTemp = New ADODB.Recordset
consulta = "SELECT '" & Now() & "' AS fecha_actual, '" & Fecha1F & " y " & Fecha2F & "' AS fechas , LogIdentificaciones.banca AS Banca_Diputado,LogIdentificaciones.hora, LogIdentificaciones.fecha, Legisladores.nombre, Legisladores.apellido,Legisladores.apellido + ' ' + Legisladores.nombre AS DetalleDiputado, Legisladores.bloque_politico, distritos.distrito AS Provincia, IsNull(VistaDedos.dedo,'-') AS dedo, IsNull(CAST(calidadMinima AS varchar(3)) + '/' + CAST(calidadMaxima AS varchar(3)),'-') AS calidad"
consulta = consulta & " FROM LogIdentificaciones INNER JOIN Legisladores ON Legisladores.id = LogIdentificaciones.id_diputado INNER JOIN distritos on Legisladores.distrito = distritos.id_distrito INNER JOIN VistaDedos ON VistaDedos.id = Legisladores.id"
consulta = consulta & " WHERE Legisladores.apellido LIKE '" & txtApellido.Text & "%'"
consulta = consulta & " AND Fecha >= '" & fecha1 & "' AND Fecha <= '" & fecha2 & "'"
If (txtInicioHora.Text <> "" And txtFinHora.Text <> "") And chkCualquierHora.Value = vbUnchecked Then
    consulta = consulta & " AND hora >= '" & txtInicioHora.Text & "' AND hora <= '" & txtFinHora.Text & "'"
End If
consulta = consulta
consulta = consulta & " ORDER BY Legisladores.apellido"
Set rs = New ADODB.Recordset
SetearRs consulta, rs
If rs.EOF Then
    MsgBox "No se encuentran resultados", vbCritical
Else
    X.DataControl1.Recordset = rs
    X.Run False
    For i = 0 To X.Pages.Count - 1
        X.Pages(i).Width = 300
    Next i
    X.PrintReport True
End If
rs.Close
Set rs = Nothing
Set X = Nothing
End Sub

Private Sub cmdIprimirListado_Click()
Dim X As New rptLogIdentificados
Dim rs As ADODB.Recordset
Dim consulta As String
Dim Fila As Integer
Dim fecha1 As String
Dim fecha2 As String
Dim Fecha1F As String
Dim Fecha2F As String
Dim i As Integer
Fecha1F = dtInicioFecha.Day & "/" & dtInicioFecha.Month & "/" & dtInicioFecha.Year
Fecha2F = dtFinFecha.Day & "/" & dtFinFecha.Month & "/" & dtFinFecha.Year
fecha1 = dtInicioFecha.Year & IIf(Len(dtInicioFecha.Month) = 1, "0" & dtInicioFecha.Month, dtInicioFecha.Month) & IIf(Len(dtInicioFecha.Day) = 1, "0" & dtInicioFecha.Day, dtInicioFecha.Day)
fecha2 = dtFinFecha.Year & IIf(Len(dtFinFecha.Month) = 1, "0" & dtFinFecha.Month, dtFinFecha.Month) & IIf(Len(dtFinFecha.Day) = 1, "0" & dtFinFecha.Day, dtFinFecha.Day)
Set RsTemp = New ADODB.Recordset
consulta = "SELECT '" & Now() & "' AS fecha_actual, '" & Fecha1F & " y " & Fecha2F & "' AS fechas , LogIdentificaciones.banca AS Banca_Diputado,LogIdentificaciones.hora, LogIdentificaciones.fecha, Legisladores.nombre, Legisladores.apellido,Legisladores.apellido + ' ' + Legisladores.nombre AS DetalleDiputado, Legisladores.bloque_politico, distritos.distrito AS Provincia FROM LogIdentificaciones INNER JOIN Legisladores ON Legisladores.id = LogIdentificaciones.id_diputado INNER JOIN distritos on Legisladores.distrito = distritos.id_distrito"
consulta = consulta & " WHERE apellido LIKE '" & txtApellido.Text & "%'"
consulta = consulta & " AND Fecha >= '" & fecha1 & "' AND Fecha <= '" & fecha2 & "'"
If (txtInicioHora.Text <> "" And txtFinHora.Text <> "") And chkCualquierHora.Value = vbUnchecked Then
    consulta = consulta & " AND hora >= '" & txtInicioHora.Text & "' AND hora <= '" & txtFinHora.Text & "'"
End If
consulta = consulta & " ORDER BY fecha DESC,hora DESC,apellido"
Set rs = New ADODB.Recordset
SetearRs consulta, rs
If rs.EOF Then
    MsgBox "No se encuentran resultados", vbCritical
Else
    X.DataControl1.Recordset = rs
    X.Run False
    For i = 0 To X.Pages.Count - 1
        X.Pages(i).Width = 300
    Next i
    X.PrintReport True
End If
rs.Close
Set rs = Nothing
Set X = Nothing
End Sub

Private Sub cmdListadoReducido_Click()
Dim X As New rptLogIdentificadosReducido
Dim rs As ADODB.Recordset
Dim consulta As String
Dim Fila As Integer
Dim fecha1 As String
Dim fecha2 As String
Dim Fecha1F As String
Dim Fecha2F As String
Dim i As Integer
Fecha1F = dtInicioFecha.Day & "/" & dtInicioFecha.Month & "/" & dtInicioFecha.Year
Fecha2F = dtFinFecha.Day & "/" & dtFinFecha.Month & "/" & dtFinFecha.Year
fecha1 = dtInicioFecha.Year & IIf(Len(dtInicioFecha.Month) = 1, "0" & dtInicioFecha.Month, dtInicioFecha.Month) & IIf(Len(dtInicioFecha.Day) = 1, "0" & dtInicioFecha.Day, dtInicioFecha.Day)
fecha2 = dtFinFecha.Year & IIf(Len(dtFinFecha.Month) = 1, "0" & dtFinFecha.Month, dtFinFecha.Month) & IIf(Len(dtFinFecha.Day) = 1, "0" & dtFinFecha.Day, dtFinFecha.Day)
Set RsTemp = New ADODB.Recordset
consulta = "SELECT DISTINCT '" & Now() & "' AS fecha_actual, '" & Fecha1F & " y " & Fecha2F & "' AS fechas, Legisladores.id, Legisladores.nombre, Legisladores.apellido,Legisladores.apellido + ' ' + Legisladores.nombre AS DetalleDiputado, Legisladores.bloque_politico, distritos.distrito AS Provincia, IsNull(VistaDedos.dedo,'-') AS dedo, IsNull(CAST(calidadMinima AS varchar(3)), '-') AS calidadMinima, IsNull(CAST(calidadMaxima AS varchar(3)),'-') AS calidadMaxima"
consulta = consulta & " FROM LogIdentificaciones INNER JOIN Legisladores ON Legisladores.id = LogIdentificaciones.id_diputado INNER JOIN distritos on Legisladores.distrito = distritos.id_distrito INNER JOIN VistaDedos ON VistaDedos.id = Legisladores.id"
consulta = consulta & " WHERE Legisladores.apellido LIKE '" & txtApellido.Text & "%'"
consulta = consulta & " AND Fecha >= '" & fecha1 & "' AND Fecha <= '" & fecha2 & "'"
If (txtInicioHora.Text <> "" And txtFinHora.Text <> "") And chkCualquierHora.Value = vbUnchecked Then
    consulta = consulta & " AND hora >= '" & txtInicioHora.Text & "' AND hora <= '" & txtFinHora.Text & "'"
End If
consulta = consulta
consulta = consulta & " ORDER BY Legisladores.apellido"
Set rs = New ADODB.Recordset
SetearRs consulta, rs
If rs.EOF Then
    MsgBox "No se encuentran resultados", vbCritical
Else
    X.DataControl1.Recordset = rs
    X.Run False
    For i = 0 To X.Pages.Count - 1
        X.Pages(i).Width = 300
    Next i
    X.PrintReport True
End If
rs.Close
Set rs = Nothing
Set X = Nothing
End Sub

Private Sub Form_Load()
Dim i As Integer
txtInicioHora.Enabled = False
txtFinHora.Enabled = False
dtInicioFecha.Day = Format(Now(), "dd")
dtInicioFecha.Month = Format(Now(), "mm")
dtInicioFecha.Year = Format(Now(), "YYYY")
dtFinFecha.Day = dtInicioFecha.Day
dtFinFecha.Month = dtInicioFecha.Month
dtFinFecha.Year = dtInicioFecha.Year
SetGrilla
'cmdBuscar_Click
End Sub
Private Sub SetGrilla()
Dim i As Integer
flexDiputados.Clear
flexDiputados.ScrollBars = flexScrollBarVertical
flexDiputados.Rows = 2
flexDiputados.Cols = 7
flexDiputados.FixedRows = 1
For i = 0 To 3
    flexDiputados.ColWidth(i) = 2350
Next i
flexDiputados.TextMatrix(0, 0) = "Apellido"
flexDiputados.TextMatrix(0, 1) = "Nombre"
flexDiputados.TextMatrix(0, 2) = "Bloque"
flexDiputados.TextMatrix(0, 3) = "Provincia"
flexDiputados.TextMatrix(0, 4) = "Banca"
flexDiputados.TextMatrix(0, 5) = "Hora"
flexDiputados.TextMatrix(0, 6) = "Fecha"
End Sub

Private Sub txtApellido_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar_Click
End If
End Sub
