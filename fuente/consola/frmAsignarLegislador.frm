VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmAsignarLegislador 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ASIGNAR LEGISLADOR A BANCA"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10275
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleMode       =   0  'User
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   10080
      Top             =   4620
   End
   Begin Proyecto1.ButtonOffice cmdAsignar 
      Height          =   705
      Left            =   60
      TabIndex        =   7
      Top             =   4620
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1244
      BackColor       =   16744576
      Caption         =   "&Asignar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Foto"
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
      Height          =   5280
      Left            =   5940
      TabIndex        =   6
      Top             =   60
      Width           =   4245
      Begin VB.Label lblDataExtra 
         BackStyle       =   0  'Transparent
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
         Height          =   495
         Left            =   180
         TabIndex        =   10
         Top             =   4740
         Width           =   3915
      End
      Begin VB.Image picDiputado 
         Height          =   4395
         Left            =   180
         Stretch         =   -1  'True
         Top             =   300
         Width           =   3915
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Búsqueda por aproximación"
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
      Height          =   1335
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   5805
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1020
         TabIndex        =   1
         Top             =   840
         Width           =   4665
      End
      Begin VB.TextBox txtApellido 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1020
         TabIndex        =   0
         Top             =   360
         Width           =   4665
      End
      Begin VB.Label lblNombre 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
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
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   2220
         TabIndex        =   4
         Top             =   360
         Width           =   555
      End
   End
   Begin MSDataListLib.DataList dlLegisladores 
      Height          =   3060
      Left            =   60
      TabIndex        =   2
      Top             =   1440
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   5398
      _Version        =   393216
      Appearance      =   0
      BackColor       =   8421504
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Proyecto1.ButtonOffice cmdCancelar 
      Height          =   705
      Left            =   3840
      TabIndex        =   8
      Top             =   4620
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1244
      BackColor       =   16744576
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
End
Attribute VB_Name = "frmAsignarLegislador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstLista As New ADODB.Recordset
Private mBanca As Long
Public BancaP As Integer
Dim IdDefault As Integer
Dim PVez As Boolean
Private lastId As String
Private Sub cmdAplicar_Click()

End Sub
Private Sub cmdAsignar_Click()
Dim nTick As Long
    If (gTipoUsuario <> 1) And (gTipoUsuario <> 4) Then
        If InStr(LCase(cmdAsignar.Caption), "orador") > 0 Then
            nTick = GetTickCount
            cmdAsignar.Enabled = False
            'frmABMLegisladores.lblid.Caption = dlLegisladores.BoundText
            cambiarOrador (dlLegisladores.BoundText)
        Else
            If (Trim(dlLegisladores.BoundText) <> Trim(IdDefault)) And IdDefault <> 0 Then
                Dim RTA As Integer
                RTA = MsgBox("Esta no es la banca donde normamente se sienta el diputado." & vbCrLf & _
                "¿Está seguro de querer asignarlo?", vbYesNo)
                If RTA = vbYes Then
                    MensajesSQV.CambiarIdBanca mBanca, dlLegisladores.BoundText
                End If
            Else
                MensajesSQV.CambiarIdBanca mBanca, dlLegisladores.BoundText
            End If
            
            Call LogSAUTOD(Trim(Str(mBanca)), dlLegisladores.BoundText)
        End If
    Else
        MsgBox "Error de permisos!"
    End If
    cmdCancelar_Click
End Sub

Private Sub cmdAsignar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub

Private Sub dlLegisladores_Click()
On Error Resume Next
Dim pic As New ADODB.Stream
Dim Rinfo As New ADODB.Recordset
If Trim(dlLegisladores.BoundText) <> "" Then
    picDiputado.Picture = LoadPicture(GetFoto(Trim(dlLegisladores.BoundText)))
    Call updateExtraData
End If
End Sub

Private Sub updateExtraData()
Dim s As String
s = "SELECT bloque_politico, ISNULL(BancasProbables.banca, '') AS banca FROM legisladores_activos LEFT JOIN BancasProbables ON " & _
" BancasProbables.id_legislador = legisladores_activos.id " & _
" WHERE legisladores_activos.id = " & dlLegisladores.BoundText
Dim Rs As New Recordset
SetearRs s, Rs
If Not Rs.EOF Then
    lblDataExtra.Caption = "Bloque: " & Rs.Fields(0) & vbCrLf
    lblDataExtra.Caption = lblDataExtra.Caption & "Banca Probable: " & Rs.Fields(1)
    lastId = dlLegisladores.BoundText
Else
    lblDataExtra.Caption = "No hay datos."
End If
Rs.Close
End Sub

Private Sub dlLegisladores_DblClick()
'    If (gTipoUsuario <> 1) And (gTipoUsuario <> 4) Then
'        MensajesSQV.CambiarIdBanca mBanca, dlLegisladores.BoundText
'        cmdCancelar_Click
'    Else
'        MsgBox "Ud. no dispone de permisos para realizar esta acción.", vbInformation + vbOKOnly
'    End If
End Sub
Public Sub mostrarLegisladores(pBanca As Integer)
    
    Dim strSql         As String
    Dim rstLista       As ADODB.Recordset
    Dim strCadena      As String
    Dim strWhere       As String
    Dim strVectWhere() As String
    Dim X              As Long
    Dim strCriterio    As String
    BancaP = pBanca
    Set rstLista = New ADODB.Recordset
    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Levantar Vector Identificación
    ' ----------------------------------------------------------------------------------------------------------------------------------
    strSql = "SELECT vector_identificacion FROM vector"
    Datos.SetearRs strSql, rstLista
    strCadena = Trim(rstLista.Fields(0).Value)
    rstLista.Close
    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Detectar todos los legisladores que estan sentados en el recinto
    ' ----------------------------------------------------------------------------------------------------------------------------------
    strWhere = strCadena
    'strWhere = Trim(Replace(strWhere, ";0;0", ""))
    'strWhere = Trim(Replace(strWhere, "0;", ""))
    'strWhere = Trim(Replace(strWhere, "; ;", ""))
    'If strWhere = "0" Then
    '    strWhere = ""
    'End If
    strCriterio = "WHERE "
    If strWhere <> "" Then
        strVectWhere = Split(strWhere, ";")
        For X = 0 To UBound(strVectWhere)
            If strVectWhere(X) <> "0" Then
                If strCriterio <> "WHERE " And X > 0 Then
                    strCriterio = strCriterio & " and "
                End If
                strCriterio = strCriterio & " Legisladores.ID <> " & strVectWhere(X)
            End If
        Next X
    End If
    'evitar que se seleccione al vice gobernador
    If strCriterio <> "WHERE " Then
        strCriterio = strCriterio & " and "
        strCadena = ""
        Dim i As Integer
        strCadena = mVectorIdentificacion(0) & ";"
        For i = 1 To 255
            strCadena = strCadena & "0;"
        Next i
        strCadena = strCadena & "0"
    End If
    strCriterio = strCriterio & " Legisladores.Es_Legislador >= 1"

    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Listar legisladores activos
    ' ----------------------------------------------------------------------------------------------------------------------------------
    If InStr(LCase(cmdAsignar.Caption), "orador") > 0 Then
        strCriterio = " AND Legisladores.ID <> " & mVectorIdentificacion(0) & " "
    End If
    If Trim(frmConsolaOperacion.lblTituloActa.Caption) <> "MANTENIMIENTO DEL SISTEMA SQV" Then
        strSql = "SELECT Legisladores.id, Legisladores.apellido + ' ' + dbo.Legisladores.nombre AS Legislador, " _
           & "legisladores_activos.ordenpresidente AS OPresidencia FROM Legisladores INNER JOIN legisladores_activos ON " _
           & "Legisladores.ID = legisladores_activos.ID " _
           & strCriterio _
           & " AND tipo = 1 " & CreaFiltroIdentificados & " AND legisladores_activos.descripcion <> 'Activo sin incorporar' ORDER BY Legislador" ' OPresidencia"
    Else
        strSql = "SELECT Legisladores.id, Legisladores.apellido + ' ' + dbo.Legisladores.nombre AS Legislador, " _
           & "'1' AS OPresidencia FROM Legisladores " _
           & strCriterio _
           & " AND tipo = 0 ORDER BY Legislador" ' OPresidencia"
    End If
    ' MsgBox strSql
    Datos.SetearRs strSql, rstLista
    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Cargar en la grilla contenido del recordset
    ' ----------------------------------------------------------------------------------------------------------------------------------
    Dim pId As String
    pId = ""
    If rstLista.EOF = False Then
        rstLista.MoveFirst
        Do While Not (rstLista.EOF)
            If InStr(1, strCadena, Trim(rstLista!id)) = 0 Then
                With dlLegisladores
                    Set .RowSource = rstLista
                    .ListField = "Legislador"
                    If pId = "" Then
                        pId = rstLista.Fields("id")
                    End If
                    .BoundColumn = "id"
                End With
            End If
            rstLista.MoveNext
        Loop
    Else
        MsgBox "No hay legisladores disponibles para presidir el recinto.", vbInformation, "Consola SQV"
        Unload Me
    End If
    mBanca = pBanca
    Dim bId As Boolean
    If PVez = True And mVectorIdentificacion(pBanca) <> IdDefault And pBanca <> 0 And frmConsolaOperacion.IDRepetida(Trim(Str(IdDefault))) = False Then
        dlLegisladores.BoundText = Trim(Str(IdDefault))
        dlLegisladores_Click
        PVez = False
    Else 'Seteo la lista en el primer diputado
        dlLegisladores.BoundText = pId
        dlLegisladores_Click
    End If
End Sub

Private Sub dlLegisladores_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub

Private Sub Form_Activate()
If InStr(LCase(cmdAsignar.Caption), "orador") > 0 Then
    'Nadap
Else
    txtNombre.Visible = False
    lblNombre.Visible = False
    Frame1.Height = 855
    dlLegisladores.Height = 3500
    dlLegisladores.Top = 1000
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
Dim RsTemp As ADODB.Recordset
LimpiaOrador = True
PVez = True
Set RsTemp = New ADODB.Recordset
IdDefault = 0
dlLegisladores.Enabled = False
SetearRs "SELECT id_legislador,BancasProbables.banca,legisladores_activos.nombre,legisladores_activos.apellido FROM BancasProbables INNER JOIN legisladores_activos ON BancasProbables.id_legislador = legisladores_activos.id WHERE (BancasProbables.banca = " & Str(BancaP) & ")", RsTemp
dlLegisladores.Enabled = True
If RsTemp.EOF Then
'    chkDiputadoDefault.Value = vbUnchecked
'    chkDiputadoDefault.Enabled = False
'    dlLegisladores.Enabled = True
'    lblDiputado.Caption = "Sin asignación"
'    lblDiputado.Enabled = False
Else
    Dim i As Integer
    Dim YaIdentificado As Boolean
    IdDefault = RsTemp.Fields("id_legislador")
    YaIdentificado = frmConsolaOperacion.IDRepetida(Str(IdDefault))
    If YaIdentificado = True Then
'        chkDiputadoDefault.Value = vbUnchecked
'        chkDiputadoDefault.Enabled = False
'        dlLegisladores.Enabled = True
'        lblDiputado.Caption = "Sin asignación"
'        lblDiputado.Enabled = False
    Else
'        lblDiputado.Caption = RsTemp.Fields("apellido") & ", " & RsTemp.Fields("nombre")
'        ActualizarPIC (Str(IdDefault))
    End If
End If
RsTemp.Close
Set RsTemp = Nothing
End Sub
Private Sub ActualizarPIC(l_id As String)
On Error Resume Next
Dim pic As New ADODB.Stream
Dim Rinfo As New ADODB.Recordset
picDiputado.Picture = LoadPicture(GetFoto(l_id))
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rstLista.State = adStateOpen Then
        rstLista.Close
    End If
    Set rstLista = Nothing
End Sub
Public Sub mostrarDiputadosFiltrados(pBanca As Integer, pPorApellido As Boolean)
    Dim strSql         As String
    Dim rstLista       As ADODB.Recordset
    Dim strCadena      As String
    Dim strWhere       As String
    Dim strVectWhere() As String
    Dim X              As Long
    Dim strCriterio    As String
    Dim PrimerID      As Integer
    PrimerID = 0
    BancaP = pBanca
    Set rstLista = New ADODB.Recordset
    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Levantar Vector Identificación
    ' ----------------------------------------------------------------------------------------------------------------------------------
    strSql = "SELECT vector_identificacion FROM vector"
    Datos.SetearRs strSql, rstLista
    strCadena = Trim(rstLista.Fields(0).Value)
    rstLista.Close
    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Detectar todos los legisladores que estan sentados en el recinto
    ' ----------------------------------------------------------------------------------------------------------------------------------
    strWhere = strCadena
    'strWhere = Trim(Replace(strWhere, ";0;0", ""))
    'strWhere = Trim(Replace(strWhere, "0;", ""))
    'strWhere = Trim(Replace(strWhere, "; ;", ""))
    'If strWhere = "0" Then
    '    strWhere = ""
    'End If
    strCriterio = "WHERE "
    If strWhere <> "" Then
        strVectWhere = Split(strWhere, ";")
        For X = 0 To UBound(strVectWhere)
            If strVectWhere(X) <> "0" Then
                If strCriterio <> "WHERE " And X > 0 Then
                    strCriterio = strCriterio & " and "
                End If
                strCriterio = strCriterio & " Legisladores.ID <> " & strVectWhere(X)
            End If
        Next X
    End If
    'evitar que se seleccione al vice gobernador
    If strCriterio <> "WHERE " Then
        strCriterio = strCriterio & " and "
    End If
    strCriterio = strCriterio & " Legisladores.Es_Legislador >= 1"


    If InStr(LCase(cmdAsignar.Caption), "orador") > 0 Then
        strCriterio = " AND Legisladores.ID <> " & mVectorIdentificacion(0) & " "
    End If
    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Listar legisladores activos
    ' ----------------------------------------------------------------------------------------------------------------------------------
    If Trim(frmConsolaOperacion.lblTituloActa.Caption) <> "MANTENIMIENTO DEL SISTEMA SQV" Then
        strSql = "SELECT Legisladores.id, Legisladores.apellido + ' ' + dbo.Legisladores.nombre AS Legislador, " _
           & "legisladores_activos.ordenpresidente AS OPresidencia FROM Legisladores INNER JOIN legisladores_activos ON " _
           & "Legisladores.ID = legisladores_activos.ID " _
           & strCriterio
            If pPorApellido = True Then
                strSql = strSql & " AND " & NormalizaCampo("Legisladores.apellido") & " LIKE '" & NormalizarNombre(txtApellido.Text) & "%' AND legisladores_activos.descripcion <> 'Activo sin incorporar' ORDER BY Legislador"  ' OPresidencia"
            Else
                strSql = strSql & " AND " & NormalizaCampo("Legisladores.nombre") & " LIKE '" & NormalizarNombre(txtNombre.Text) & "%' AND legisladores_activos.descripcion <> 'Activo sin incorporar' ORDER BY Legislador"  ' OPresidencia"
            End If
    Else
        strSql = "SELECT Legisladores.id, Legisladores.apellido + ' ' + dbo.Legisladores.nombre AS Legislador, " _
           & "'1' AS OPresidencia FROM Legisladores " _
           & strCriterio _
           & " AND Legisladores.apellido LIKE '" & txtApellido.Text & "%' AND tipo = 0 ORDER BY Legislador" ' OPresidencia"
    End If
    ' MsgBox strSql
    Datos.SetearRs strSql, rstLista
    ' ----------------------------------------------------------------------------------------------------------------------------------
    ' Cargar en la grilla contenido del recordset
    ' ----------------------------------------------------------------------------------------------------------------------------------
    If rstLista.EOF = False Then
        rstLista.MoveFirst
        Do While Not (rstLista.EOF)
            If PrimerID = 0 Then
                PrimerID = rstLista!id
            End If
            If InStr(1, strCadena, Trim(rstLista!id)) = 0 Then
                With dlLegisladores
                    Set .RowSource = rstLista
                    .ListField = "Legislador"
                    .BoundColumn = "id"
                End With
            End If
            rstLista.MoveNext
        Loop
        dlLegisladores.BoundText = PrimerID
    Else
        'Set dlLegisladores.RowSource = rstLista
        'dlLegisladores.Refresh
    End If
    mBanca = pBanca
End Sub

Private Function NormalizaCampo(pField As String)
Dim ret As String
ret = pField
ret = ReplaceDinamico(ret, "á", "a")
ret = ReplaceDinamico(ret, "é", "e")
ret = ReplaceDinamico(ret, "í", "i")
ret = ReplaceDinamico(ret, "ó", "o")
ret = ReplaceDinamico(ret, "ú", "u")
NormalizaCampo = ret
End Function

Private Function ReplaceDinamico(pText As String, pPattern As String, pReplace As String) As String
Dim ret As String
ret = "REPLACE(" & pText & ",'" & pPattern & "','" & pReplace & "')"
ReplaceDinamico = ret
End Function

Private Function NormalizarNombre(pTexto As String) As String
Dim ret As String
ret = pTexto
ret = LCase(Trim(ret))
ret = Replace(ret, "á", "a")
ret = Replace(ret, "é", "e")
ret = Replace(ret, "í", "i")
ret = Replace(ret, "ó", "o")
ret = Replace(ret, "ú", "u")
NormalizarNombre = ret
End Function

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim pic As New ADODB.Stream
Dim Rinfo As New ADODB.Recordset
If Trim(dlLegisladores.BoundText) <> "" Then
    picDiputado.Picture = LoadPicture(GetFoto(Trim(dlLegisladores.BoundText)))
    If lastId <> dlLegisladores.BoundText Then
        Call updateExtraData
    End If
End If
End Sub
Private Sub txtApellido_Change()
If txtApellido.Text <> "" Then
    Call mostrarDiputadosFiltrados(BancaP, True)
Else
    mostrarLegisladores (BancaP)
End If
End Sub
Private Sub txtApellido_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'End If
End Sub

Private Sub txtNombre_Change()
If txtNombre.Text <> "" Then
    Call mostrarDiputadosFiltrados(BancaP, False)
Else
    mostrarLegisladores (BancaP)
End If
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAplicar_Click
End If
End Sub
Public Function CreaFiltroIdentificados() As String
Dim i As Integer
Dim Buff As String
Buff = ""
If InStr(LCase(cmdAsignar.Caption), "orador") = 0 Then
    For i = 0 To 256
        If mVectorIdentificacion(i) <> "0" Then
            Buff = Buff & " AND Legisladores.id <> " & mVectorIdentificacion(i) & " "
        End If
    Next i
End If
CreaFiltroIdentificados = Buff
End Function
