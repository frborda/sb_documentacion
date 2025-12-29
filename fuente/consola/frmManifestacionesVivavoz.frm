VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmManifestacionesVivavoz 
   BackColor       =   &H00404040&
   Caption         =   "Manifestaciones a viva voz"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   12375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Manifestaciones actuales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   12255
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   675
         Left            =   11160
         TabIndex        =   6
         Top             =   5460
         Width           =   975
      End
      Begin VB.TextBox txtManifestacion 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5115
         Left            =   4380
         MaxLength       =   1024
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   7755
      End
      Begin VB.ListBox lstDiputados 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5100
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   4215
      End
      Begin Proyecto1.ButtonOffice cmdVivavoz 
         Height          =   555
         Left            =   120
         TabIndex        =   3
         Top             =   5460
         Width           =   2050
         _ExtentX        =   3625
         _ExtentY        =   979
         BackColor       =   33023
         Caption         =   "Agregar diputado"
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
      Begin Proyecto1.ButtonOffice cmdBorrar 
         Height          =   555
         Left            =   2220
         TabIndex        =   4
         Top             =   5460
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   979
         BackColor       =   33023
         Caption         =   "Borrar diputado"
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
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Haga clic en un diputado para modificar su manifestación."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   4440
         TabIndex        =   5
         Top             =   5580
         Width           =   7695
      End
   End
End
Attribute VB_Name = "frmManifestacionesVivavoz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mPeriodo As String
Public mSesion As String
Public mActa As String
Public mVersion As String
Public defaultText As String
Public mChanged As Boolean
Public mSinDiputados As Boolean
Public mCargarOriginales As Boolean
Dim ids() As String
Dim focusIndex As Integer
Dim Manifestaciones() As CManifestacion
Dim fTime As Boolean

Private Sub cmdBorrar_Click()
If (lstDiputados.ListIndex > -1) Then
    Dim r As Integer
    r = MsgBox("¿Está seguro de que desea eliminar las manifestaciones del diputado " & lstDiputados.List(lstDiputados.ListIndex) & "?", vbYesNo)
    If (r = vbYes) Then
        Call borrarDiputado(ids(lstDiputados.ListIndex))
        Call CargaDiputados
    End If
Else
    MsgBox "Debe seleccionar un diputado"
End If
End Sub

Private Sub cmdVivavoz_Click()
Dim f As New frmSeleccionDiputado
f.mPeriodo = mPeriodo
f.mSesion = mSesion
f.mActa = mActa
f.mVersion = 0 'En teoría sería siempre la última versión
f.Show vbModal, Me
If (f.Result <> 0) Then
    Call agregarDiputado(f.Result, "")
    Call CargaDiputados
End If
End Sub

Private Sub borrarDiputado(mId As String)
Dim s As String
s = "DELETE FROM manifestaciones_vivavoz WHERE " & _
" manifestaciones_vivavoz.id_diputado = " & mId & _
" AND manifestaciones_vivavoz.periodo = '" & mPeriodo & "'" & _
" AND manifestaciones_vivavoz.sesion = " & mSesion & _
" AND manifestaciones_vivavoz.nro_acta = " & mActa & _
" AND manifestaciones_vivavoz.version_acta = -1"
Call InsertSQL(s)
End Sub

Private Sub borrarDiputados()
Dim s As String
s = "DELETE FROM manifestaciones_vivavoz WHERE " & _
" manifestaciones_vivavoz.periodo = '" & mPeriodo & "'" & _
" AND manifestaciones_vivavoz.sesion = " & mSesion & _
" AND manifestaciones_vivavoz.nro_acta = " & mActa & _
" AND manifestaciones_vivavoz.version_acta = -1"
Call InsertSQL(s)
End Sub

Private Sub agregarDiputado(mId As String, comentario As String)
Dim s As String
s = "INSERT INTO manifestaciones_vivavoz(periodo, sesion, nro_acta, version_acta, id_diputado, comentario, ultima_edicion) " & _
" VALUES('" & mPeriodo & "'," & mSesion & "," & mActa & ", -1," & mId & ",'" & comentario & "', { fn NOW() })"
Call InsertSQL(s)
End Sub

Private Sub CargaDiputados()
Dim rs As New Recordset
Dim s As String
lblStatus.Caption = defaultText
txtManifestacion.Text = ""
txtManifestacion.Enabled = False
focusIndex = -1
lstDiputados.Clear
ReDim ids(0 To 0)
s = "SELECT Legisladores.id, Legisladores.apellido + ', ' + Legisladores.nombre AS diputado, manifestaciones_vivavoz.comentario FROM manifestaciones_vivavoz " & _
"INNER JOIN Legisladores ON Legisladores.id = manifestaciones_vivavoz.id_diputado " & _
" WHERE manifestaciones_vivavoz.periodo = '" & mPeriodo & "'" & _
" AND manifestaciones_vivavoz.sesion = " & mSesion & _
" AND manifestaciones_vivavoz.nro_acta = " & mActa & _
" AND manifestaciones_vivavoz.version_acta = -1" & _
" ORDER BY Legisladores.apellido, Legisladores.nombre"
SetearRs s, rs
If rs.EOF Then
    Exit Sub
End If
While Not rs.EOF
    lstDiputados.AddItem rs.Fields("diputado")
    ReDim Preserve ids(0 To (lstDiputados.ListCount - 1))
    ids(lstDiputados.ListCount - 1) = rs.Fields("id")
    rs.MoveNext
Wend
End Sub

Private Sub Command1_Click()
Dim rpt As New rptManifestaciones
Dim rs As New ADODB.Recordset
Dim s As String
s = "SELECT id_diputado, Legisladores.apellido + ', ' + Legisladores.nombre as diputado, comentario " & _
" FROM manifestaciones_vivavoz INNER JOIN Legisladores ON Legisladores.id = manifestaciones_vivavoz.id_diputado " & _
" ORDER BY diputado"
SetearRs s, rs
If rs.EOF Then
    MsgBox "Sin data!"
    Exit Sub
End If
rpt.DataControl1.Recordset = rs
rpt.PrintReport True
End Sub

Private Sub Form_Load()
mSinDiputados = True
mChanged = False
fTime = True
focusIndex = -1
defaultText = "Haga clic en un diputado para modificar su manifestación."
If Not (mCargarOriginales) Then
    Call PreparaDiputados(False)
    Call CargaDiputados
Else
    Call PreparaDiputados(True)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If (lstDiputados.ListIndex > -1) Then
    If (focusIndex <> -1) Then
        Call SaveData(txtManifestacion.Text)
    End If
End If
'Reviso si cambió alguno
Call CheckForChanges
End Sub

Private Sub lstDiputados_Click()
If (lstDiputados.ListIndex > -1) Then
    txtManifestacion.Enabled = True
    Call CargaManifestacion(ids(lstDiputados.ListIndex))
    lblStatus.Caption = "Actualmente editando: " & lstDiputados.List(lstDiputados.ListIndex)
    txtManifestacion.SetFocus
Else
    txtManifestacion.Enabled = False
    txtManifestacion.Text = ""
    lblStatus.Caption = defaultText
End If
End Sub

Private Sub txtManifestacion_GotFocus()
focusIndex = lstDiputados.ListIndex
End Sub

Private Sub txtManifestacion_LostFocus()
Call SaveData(txtManifestacion.Text)
lblStatus.Caption = defaultText
focusIndex = -1
End Sub

Private Sub CheckForChanges()
Dim s As String
Dim rs As New ADODB.Recordset
s = "SELECT Legisladores.id, Legisladores.apellido + ', ' + Legisladores.nombre AS diputado, manifestaciones_vivavoz.comentario FROM manifestaciones_vivavoz " & _
" INNER JOIN Legisladores ON Legisladores.id = manifestaciones_vivavoz.id_diputado " & _
" WHERE manifestaciones_vivavoz.periodo = '" & mPeriodo & "'" & _
" AND manifestaciones_vivavoz.sesion = " & mSesion & _
" AND manifestaciones_vivavoz.nro_acta = " & mActa & _
" AND manifestaciones_vivavoz.version_acta = -1" & _
" ORDER BY Legisladores.apellido, Legisladores.nombre"
SetearRs s, rs
If rs.EOF Then
    'Borro diputados con -1 aunque en este caso no sería necesario
    borrarDiputados
    Exit Sub
End If
If rs.RecordCount > 0 And Me.mSinDiputados = True Then
    Me.mChanged = True
    Exit Sub
End If
If rs.RecordCount <> (UBound(Manifestaciones) + 1) Then
    'Cambiaron! No hago nada, que queden los -1 en la tabla
    Me.mChanged = True
    Exit Sub
End If
'Si tiene la misma cantidad, recorro a ver si alguno es distinto
Dim i As Integer
i = -1
While Not rs.EOF
    i = i + 1
    If Not (Manifestaciones(i).idDiputado = rs.Fields("id") And Manifestaciones(i).manifestacion = rs.Fields("comentario")) Then
        'Cambio por lo menos uno! No hago nada, que queden los -1 en la tabla
        Me.mChanged = True
        Exit Sub
    End If
    rs.MoveNext
Wend
'Borro los diputados con -1, no hubieron cambios
borrarDiputados
End Sub

Private Sub CargaManifestacion(mId As String)
Dim rs As New Recordset
Dim s As String
s = "SELECT manifestaciones_vivavoz.comentario FROM manifestaciones_vivavoz " & _
" WHERE manifestaciones_vivavoz.periodo = '" & mPeriodo & "'" & _
" AND manifestaciones_vivavoz.sesion = " & mSesion & _
" AND manifestaciones_vivavoz.nro_acta = " & mActa & _
" AND manifestaciones_vivavoz.version_acta = -1" & _
" AND manifestaciones_vivavoz.id_diputado = " & mId
SetearRs s, rs
If rs.EOF Then
    MsgBox "Error fatal: no se pudo encontrar al diputado."
    Unload Me
    Exit Sub
End If
txtManifestacion.Text = rs.Fields("comentario")
If (Len(txtManifestacion.Text) > 0) Then
    txtManifestacion.SelStart = Len(txtManifestacion.Text)
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub PreparaDiputados(mOriginales As Boolean)
Dim rs As New Recordset
Dim s As String
lblStatus.Caption = defaultText
txtManifestacion.Text = ""
txtManifestacion.Enabled = False
focusIndex = -1
lstDiputados.Clear
ReDim ids(0 To 0)
s = "SELECT Legisladores.id, Legisladores.apellido + ', ' + Legisladores.nombre AS diputado, manifestaciones_vivavoz.comentario FROM manifestaciones_vivavoz " & _
" INNER JOIN Legisladores ON Legisladores.id = manifestaciones_vivavoz.id_diputado " & _
" WHERE manifestaciones_vivavoz.periodo = '" & mPeriodo & "'" & _
" AND manifestaciones_vivavoz.sesion = " & mSesion & _
" AND manifestaciones_vivavoz.nro_acta = " & mActa & _
" AND manifestaciones_vivavoz.version_acta = " & mVersion & _
" ORDER BY Legisladores.apellido, Legisladores.nombre"
SetearRs s, rs
If rs.EOF Then
    Exit Sub
End If
Dim i As Integer
i = -1
While Not rs.EOF
    i = i + 1
    mSinDiputados = False
    If (mOriginales = True) Then
        lstDiputados.AddItem rs.Fields("diputado")
    End If
    ReDim Preserve ids(0 To i)
    ids(i) = rs.Fields("id")
    ReDim Preserve Manifestaciones(0 To i)
    Manifestaciones(i).idDiputado = rs.Fields("id")
    Manifestaciones(i).manifestacion = rs.Fields("comentario")
    If (mOriginales = True) Then
        Call agregarDiputado(rs.Fields("id"), rs.Fields("comentario"))
    End If
    rs.MoveNext
Wend
End Sub

Private Sub SaveData(mComentario As String)
Dim s As String
Dim mId As String
mId = ids(focusIndex)
s = "UPDATE manifestaciones_vivavoz SET comentario = '" & mComentario & "'" & _
" , ultima_edicion = { fn NOW() } " & _
" WHERE manifestaciones_vivavoz.id_diputado = " & mId & _
" AND manifestaciones_vivavoz.periodo = '" & mPeriodo & "'" & _
" AND manifestaciones_vivavoz.sesion = " & mSesion & _
" AND manifestaciones_vivavoz.nro_acta = " & mActa & _
" AND manifestaciones_vivavoz.version_acta = -1"
Call InsertSQL(s)
End Sub
