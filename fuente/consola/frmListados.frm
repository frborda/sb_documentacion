VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmListados 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listados"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
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
      Height          =   1245
      Left            =   7200
      TabIndex        =   21
      Top             =   5190
      Width           =   2565
      Begin VB.OptionButton optIDInterno 
         BackColor       =   &H00404040&
         Caption         =   "Por ID Interno"
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
         Left            =   210
         TabIndex        =   24
         Top             =   840
         Width           =   2025
      End
      Begin VB.OptionButton optPorID 
         BackColor       =   &H00404040&
         Caption         =   "Por ID"
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
         Left            =   210
         TabIndex        =   23
         Top             =   540
         Width           =   1185
      End
      Begin VB.OptionButton optOrdenAlfabetico 
         BackColor       =   &H00404040&
         Caption         =   "Alfabético"
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
         Left            =   210
         TabIndex        =   22
         Top             =   270
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Otros"
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
      Height          =   1245
      Left            =   4560
      TabIndex        =   19
      Top             =   5190
      Width           =   2565
      Begin VB.CheckBox chkBloques 
         BackColor       =   &H00404040&
         Caption         =   "Todos los bloques"
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
         Left            =   210
         TabIndex        =   26
         Top             =   780
         Width           =   2175
      End
      Begin VB.CheckBox chkProvincias 
         BackColor       =   &H00404040&
         Caption         =   "Todas las provincias"
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
         Left            =   210
         TabIndex        =   25
         Top             =   510
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkIDInterno 
         BackColor       =   &H00404040&
         Caption         =   "Mostrar ID Interno"
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
         Left            =   210
         TabIndex        =   20
         Top             =   240
         Value           =   1  'Checked
         Width           =   1875
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00404040&
      Caption         =   "Selección de Bloques Políticos"
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
      Height          =   3795
      Left            =   60
      TabIndex        =   13
      Top             =   60
      Width           =   9735
      Begin VB.ListBox lstBloquesSeleccionados 
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
         Height          =   3390
         Left            =   5400
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   300
         Width           =   4215
      End
      Begin Proyecto1.ButtonOffice cmdMover 
         Height          =   375
         Left            =   4440
         TabIndex        =   15
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         BackColor       =   12230304
         Caption         =   ">"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin VB.ListBox lstBloques 
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
         Height          =   3390
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   300
         Width           =   4215
      End
      Begin Proyecto1.ButtonOffice cmdRemover 
         Height          =   375
         Left            =   4440
         TabIndex        =   16
         Top             =   720
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         BackColor       =   12230304
         Caption         =   "<"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Proyecto1.ButtonOffice cmdLimpiar 
         Height          =   375
         Left            =   4440
         TabIndex        =   17
         Top             =   3300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         BackColor       =   12230304
         Caption         =   "Limpiar"
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
   Begin VB.Frame frmOpt 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   675
      Left            =   90
      TabIndex        =   10
      Top             =   6510
      Width           =   9705
      Begin Proyecto1.ButtonOffice cmdImprimir 
         Height          =   555
         Left            =   30
         TabIndex        =   11
         Top             =   0
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   979
         BackColor       =   16744576
         Caption         =   "&Imprimir Listado"
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
      Begin Proyecto1.ButtonOffice cmdSalir 
         Height          =   555
         Left            =   7800
         TabIndex        =   12
         Top             =   0
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   979
         BackColor       =   16744576
         Caption         =   "&Salir"
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      Caption         =   "Listados"
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
      Height          =   1275
      Left            =   7200
      TabIndex        =   9
      Top             =   3900
      Width           =   2565
      Begin VB.OptionButton optRegular 
         BackColor       =   &H00404040&
         Caption         =   "Normal"
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
         Height          =   345
         Left            =   210
         TabIndex        =   5
         Top             =   660
         Value           =   -1  'True
         Width           =   2325
      End
      Begin VB.OptionButton optConFotos 
         BackColor       =   &H00404040&
         Caption         =   "Con Fotografías"
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
         Height          =   345
         Left            =   210
         TabIndex        =   4
         Top             =   300
         Width           =   1665
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   "Sexo"
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
      Height          =   1275
      Left            =   4560
      TabIndex        =   8
      Top             =   3900
      Width           =   2565
      Begin VB.OptionButton optMasculino 
         BackColor       =   &H00404040&
         Caption         =   "Masculino"
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
         Left            =   210
         TabIndex        =   3
         Top             =   960
         Width           =   1185
      End
      Begin VB.OptionButton optFemenino 
         BackColor       =   &H00404040&
         Caption         =   "Femenino"
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
         Left            =   210
         TabIndex        =   2
         Top             =   600
         Width           =   1125
      End
      Begin VB.OptionButton optAmbos 
         BackColor       =   &H00404040&
         Caption         =   "Ambos"
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
         Left            =   210
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Provincias"
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
      Height          =   2535
      Left            =   60
      TabIndex        =   6
      Top             =   3900
      Width           =   4395
      Begin VB.TextBox txtDistrito 
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
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   4155
      End
      Begin VB.ListBox lstDistritos 
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
         Height          =   1710
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   4155
      End
   End
End
Attribute VB_Name = "frmListados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkBloques_Click()
Dim i As Integer
If chkBloques.Value = vbChecked Then
    lstBloquesSeleccionados.Clear
    ActualizarBloques
    If lstBloques.ListCount > 0 Then
        For i = 0 To lstBloques.ListCount - 1
            lstBloquesSeleccionados.AddItem lstBloques.List(i)
        Next i
        lstBloques.Clear
    Else
        MsgBox "No tiene más bloques para seleccionar!"
    End If
    cmdMover.Enabled = False
    cmdRemover.Enabled = False
    cmdLimpiar.Enabled = False
Else
    lstBloques.Clear
    cmdMover.Enabled = True
    cmdRemover.Enabled = True
    cmdLimpiar.Enabled = True
    lstBloquesSeleccionados.Clear
    ActualizarBloques
End If
End Sub

Private Sub chkProvincias_Click()
If chkProvincias.Value = vbChecked Then
    lstDistritos.ListIndex = -1
    lstDistritos.Enabled = False
    txtDistrito.Enabled = False
Else
    txtDistrito.Enabled = True
    lstDistritos.Enabled = True
End If
End Sub

Private Sub cmdImprimir_Click()
Dim X As New rptFotos
Dim X2 As New rptListadoLegisladores
Dim Rs As ADODB.Recordset
Dim Sexo As String
Dim Consulta As String
Dim cTotal As Integer
Dim i As Integer
If lstBloquesSeleccionados.ListCount > 0 Then
    ActualizarActivosExtra
    frmOpt.Enabled = False
    cmdImprimir.Caption = "Cargando..."
    DoEvents
    If optConFotos.Value = True Then
        Consulta = "SELECT Legisladores.id AS idDiputado2, Legisladores.id AS idDiputado, Legisladores.codigo_persona AS idDiputado,Legisladores.PICTURE,Legisladores.apellido,Legisladores.nombre,Legisladores.bloque_politico,legisladores_activos.DeskId, distritos.distrito AS NDistrito, '" & Now() & "' As fecha FROM Legisladores INNER JOIN legisladores_activos ON legisladores_activos.id = Legisladores.id INNER JOIN distritos ON Legisladores.distrito = distritos.id_distrito"
    Else
        'Consulta = "SELECT legisladores_para_actualizar.DeskId AS idDiputado2,legisladores_para_actualizar.id AS idDB, Legisladores.codigo_persona AS idDiputado,Legisladores.apellido,Legisladores.nombre,Legisladores.bloque_politico,Legisladores.grupo_politico,distritos.distrito AS NDistrito, '" & Now() & "' As fecha FROM Legisladores INNER JOIN distritos ON Legisladores.distrito = distritos.id_distrito INNER JOIN legisladores_para_actualizar ON legisladores_para_actualizar.id = Legisladores.id"
        Consulta = "SELECT legisladores_para_actualizar.DeskId AS idDiputado2,legisladores_para_actualizar.id AS idDB, Legisladores.codigo_persona AS idDiputado,Legisladores.apellido,Legisladores.nombre,Legisladores.bloque_politico,Legisladores.grupo_politico,distritos.distrito AS NDistrito, '" & Now() & "' As fecha FROM Legisladores INNER JOIN legisladores_activos ON legisladores_activos.id = legisladores.id LEFT OUTER JOIN distritos ON Legisladores.distrito = distritos.id_distrito LEFT OUTER JOIN legisladores_para_actualizar ON legisladores_para_actualizar.id = Legisladores.id"
    End If
    If optAmbos.Value = True Then
        Sexo = ""
    ElseIf optMasculino.Value = True Then
        Sexo = "1"
    Else
        Sexo = "0"
    End If
    Consulta = Consulta & " WHERE Legisladores.sexo LIKE '" & Sexo & "%' "
    If lstBloquesSeleccionados.ListCount > 0 Then
        i = 0
        Consulta = Consulta & " AND (Legisladores.bloque_politico = '" & lstBloquesSeleccionados.List(i) & "'"
        Consulta = Consulta & " OR Legisladores.bloque_politico IS NULL "
        For i = 1 To lstBloquesSeleccionados.ListCount - 1
            Consulta = Consulta & " OR Legisladores.bloque_politico = '" & lstBloquesSeleccionados.List(i) & "'"
        Next i
        Consulta = Consulta & ")"
    End If
    If lstDistritos.ListIndex <> -1 Then
        Dim t As Integer
        t = Distrito_GetIdByName(lstDistritos.List(lstDistritos.ListIndex))
        If t <> -1 Then
            Consulta = Consulta & " AND Legisladores.distrito = " & t
        Else
            MsgBox "Error de integridad de datos en distritos", vbCritical
        End If
    End If
    If optOrdenAlfabetico.Value = True Then
        Consulta = Consulta & " ORDER BY Legisladores.apellido, Legisladores.nombre"
    ElseIf optPorID.Value = True Then
        Consulta = Consulta & " ORDER BY idDiputado2"
    ElseIf optIDInterno.Value = True Then
        Consulta = Consulta & " ORDER BY idDiputado"
    End If
    Set rsTemp = New ADODB.Recordset
    Set Rs = New ADODB.Recordset
    SetearRs Consulta, Rs
    cTotal = Rs.RecordCount
    If cTotal > 0 Then
        If optConFotos.Value = True Then
            X.DataControl1.Recordset = Rs
            X.Run False
            For i = 0 To X.Pages.Count - 1
                X.Pages(i).Width = 300
            Next i
            X.PrintReport True
        Else
            If chkIDInterno.Value = vbUnchecked Then
                X2.Label16.Left = X2.Label8.Left
                X2.Field10.Left = X2.Field3.Left
                X2.Label8.Visible = False
                X2.Field3.Visible = False
            End If
            X2.lblVotacion = X2.lblVotacion & " al " & Format(Now(), "dd/mm/yyyy")
            X2.lblCantidad = cTotal
            X2.DataControl1.Recordset = Rs
            X2.Run False
            For i = 0 To X2.Pages.Count - 1
                X2.Pages(i).Width = 300
            Next i
            X2.PrintReport True
        End If
        Rs.Close
        Set Rs = Nothing
        Set X = Nothing
        Set X2 = Nothing
    Else
        MsgBox "No se encontraron resultados", vbInformation
    End If
    cmdImprimir.Caption = "&Imprimir Listado"
    frmOpt.Enabled = True
Else
    MsgBox "Debe seleccionar al menos un bloque", vbInformation
End If
End Sub
Private Sub cmdLimpiar_Click()
lstBloquesSeleccionados.Clear
ActualizarBloques
End Sub
Private Sub cmdMover_Click()
If lstBloques.ListIndex <> -1 Then
    lstBloquesSeleccionados.AddItem lstBloques.List(lstBloques.ListIndex)
    lstBloques.RemoveItem (lstBloques.ListIndex)
Else
    Call MsgBox("Debe tener un bloque seleccionado", vbInformation)
End If
End Sub
Private Sub cmdPasarTodos_Click()

End Sub
Private Sub cmdRemover_Click()
If lstBloquesSeleccionados.ListIndex <> -1 Then
    lstBloques.AddItem lstBloquesSeleccionados.List(lstBloquesSeleccionados.ListIndex)
    lstBloquesSeleccionados.RemoveItem (lstBloquesSeleccionados.ListIndex)
Else
    Call MsgBox("Debe tener un bloque seleccionado", vbInformation)
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Form_Load()
ActualizarBloques
ActualizarDistritos
lstDistritos.ListIndex = -1
lstDistritos.Enabled = False
txtDistrito.Enabled = False
End Sub
Private Sub ActualizarBloques()
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
lstBloques.Clear
SetearRs "SELECT * FROM Bloques ORDER BY Bloque_Político", rsTemp
While Not rsTemp.EOF
    lstBloques.AddItem rsTemp.Fields("Bloque_Político")
    rsTemp.MoveNext
Wend
rsTemp.Close
Set rsTemp = Nothing
End Sub
Private Sub ActualizarDistritos()
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
lstDistritos.Clear
SetearRs "SELECT * FROM distritos WHERE distrito LIKE '" & txtDistrito.Text & "%' ORDER BY distrito", rsTemp
While Not rsTemp.EOF
    lstDistritos.AddItem rsTemp.Fields("distrito")
    rsTemp.MoveNext
Wend
rsTemp.Close
Set rsTemp = Nothing
End Sub
Private Function Bloque_GetIdByName(bloque As String) As Integer
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
SetearRs "SELECT id FROM Bloques WHERE LTrim(RTrim(Bloque_Político)) = '" & Trim(bloque) & "'", rsTemp
If rsTemp.EOF Then
    Bloque_GetIdByName = -1
Else
    Bloque_GetIdByName = Val(rsTemp.Fields(0))
End If
rsTemp.Close
Set rsTemp = Nothing
End Function
Private Function Distrito_GetIdByName(distrito As String) As Integer
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
SetearRs "SELECT id_distrito FROM distritos WHERE LTrim(RTrim(distrito)) = '" & Trim(distrito) & "'", rsTemp
If rsTemp.EOF Then
    Distrito_GetIdByName = -1
Else
    Distrito_GetIdByName = Val(rsTemp.Fields(0))
End If
rsTemp.Close
Set rsTemp = Nothing
End Function
Private Sub txtBloque_Change()
ActualizarBloques
End Sub

Private Sub optConFotos_Click()
optOrdenAlfabetico.Value = True
optOrdenAlfabetico.Enabled = False
optPorID.Enabled = False
optIDInterno.Enabled = False
chkIDInterno.Enabled = False
End Sub

Private Sub optRegular_Click()
optOrdenAlfabetico.Value = True
optOrdenAlfabetico.Enabled = True
optPorID.Enabled = True
optIDInterno.Enabled = True
chkIDInterno.Enabled = True
End Sub

Private Sub txtDistrito_Change()
ActualizarDistritos
End Sub
