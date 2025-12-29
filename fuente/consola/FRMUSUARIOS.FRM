VERSION 5.00
Begin VB.Form frmUsuarios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Usuarios"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   915
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   6225
      TabIndex        =   19
      Top             =   120
      Width           =   6280
      Begin VB.CommandButton Buscar 
         Caption         =   "&Buscar"
         Height          =   855
         Left            =   4980
         Picture         =   "frmUsuarios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Borrar 
         Caption         =   "Eliminar"
         Height          =   855
         Left            =   2490
         Picture         =   "frmUsuarios.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Grabar 
         Caption         =   "&Grabar"
         Height          =   855
         Left            =   1245
         Picture         =   "frmUsuarios.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Nuevo 
         Caption         =   "&Nuevo"
         Height          =   855
         Left            =   0
         Picture         =   "frmUsuarios.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdPassword 
         Caption         =   "Password"
         Height          =   855
         Left            =   3735
         Picture         =   "frmUsuarios.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   915
      Left            =   6480
      ScaleHeight     =   855
      ScaleWidth      =   1245
      TabIndex        =   15
      Top             =   120
      Width           =   1300
      Begin VB.CommandButton Salir 
         Caption         =   "&Salir"
         Height          =   855
         Left            =   0
         Picture         =   "frmUsuarios.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4005
      Left            =   120
      TabIndex        =   12
      Top             =   1140
      Width           =   7665
      Begin VB.CheckBox chkActualizarSB 
         Caption         =   "Actualizar datos de Servidor de Banca"
         Height          =   195
         Left            =   840
         TabIndex        =   34
         Top             =   2256
         Width           =   3165
      End
      Begin VB.CheckBox chkHabilitaBotonesConsola 
         Caption         =   "Habilitar botones de la Consola"
         Height          =   195
         Left            =   840
         TabIndex        =   33
         Top             =   2592
         Width           =   2565
      End
      Begin VB.CheckBox chkConsultaABMLegisladores 
         Caption         =   "Consulta ABM Legisladores"
         Height          =   195
         Left            =   840
         TabIndex        =   32
         Top             =   3600
         Width           =   2325
      End
      Begin VB.CheckBox chkDefinirOrdenPresidente 
         Caption         =   "Definir Orden de Presidente"
         Height          =   195
         Left            =   840
         TabIndex        =   31
         Top             =   2928
         Width           =   2325
      End
      Begin VB.CheckBox chkABMUsuarios 
         Caption         =   "ABM Usuarios"
         Height          =   195
         Left            =   840
         TabIndex        =   30
         Top             =   3264
         Width           =   1395
      End
      Begin VB.CheckBox chkUsuarioMantenimiento 
         Caption         =   "Usuario Mantenimiento"
         Height          =   195
         Left            =   4500
         TabIndex        =   29
         Top             =   2246
         Width           =   1965
      End
      Begin VB.CheckBox chkABMLegisladores 
         Caption         =   "ABM Legisladores"
         Height          =   195
         Left            =   840
         TabIndex        =   28
         Top             =   1920
         Width           =   1635
      End
      Begin VB.CheckBox chkImprimirActas 
         Caption         =   "Imprimir Actas"
         Height          =   195
         Left            =   4500
         TabIndex        =   27
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox chkModificaActas 
         Caption         =   "Modificar Actas"
         Height          =   195
         Left            =   4500
         TabIndex        =   4
         Top             =   2898
         Width           =   1485
      End
      Begin VB.CheckBox chkExportaActas 
         Caption         =   "Exportar titulos de actas"
         Height          =   195
         Left            =   4500
         TabIndex        =   5
         Top             =   3224
         Width           =   2055
      End
      Begin VB.CheckBox chkConsultaActas 
         Caption         =   "Consultar Actas"
         Height          =   195
         Left            =   4500
         TabIndex        =   3
         Top             =   2572
         Width           =   1455
      End
      Begin VB.TextBox txtUsuario 
         Height          =   285
         Left            =   1650
         TabIndex        =   2
         Top             =   1156
         Width           =   5385
      End
      Begin VB.TextBox txtApellido 
         Height          =   285
         Left            =   1650
         TabIndex        =   0
         Top             =   330
         Width           =   5385
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1650
         TabIndex        =   1
         Top             =   743
         Width           =   5385
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Permisos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   450
         TabIndex        =   24
         Top             =   1530
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre de Usuario : "
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   1186
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Apellido : "
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre : "
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   773
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   390
      Left            =   120
      ScaleHeight     =   330
      ScaleWidth      =   690
      TabIndex        =   9
      Top             =   5265
      Width           =   750
      Begin VB.CommandButton cmdPrimero 
         Height          =   325
         Left            =   0
         Picture         =   "frmUsuarios.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdAnterior 
         Height          =   325
         Left            =   345
         Picture         =   "frmUsuarios.frx":079E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   390
      Left            =   7080
      ScaleHeight     =   330
      ScaleWidth      =   690
      TabIndex        =   6
      Top             =   5280
      Width           =   750
      Begin VB.CommandButton cmdUltimo 
         Height          =   325
         Left            =   345
         Picture         =   "frmUsuarios.frx":0930
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdSiguiente 
         Height          =   325
         Left            =   0
         Picture         =   "frmUsuarios.frx":0AC2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
   End
   Begin VB.Label lblId 
      Caption         =   "nothing"
      Height          =   375
      Left            =   7560
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblRecordSet 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0/0 Usuarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1027
      TabIndex        =   17
      Top             =   5265
      Width           =   5895
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blEsNuevo     As Boolean
Private WithEvents Rs As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1

Private Sub MostrarRegistro()
    On Error GoTo TrapError
    Dim xPos As Long
    Dim xMax As Long
    With Rs
        xPos = .AbsolutePosition ' Registro Actual
        xMax = .RecordCount      ' Registros totales
        If xMax > 0 Then         ' Si hay registros para mostrar
            txtApellido.Text = "" + .Fields("Apellido").Value
            txtNombre.Text = "" + .Fields("Nombre").Value
            txtUsuario.Text = "" + .Fields("Login").Value
            chkConsultaActas.Value = 0 + .Fields("ConsultasActas").Value
            chkABMLegisladores.Value = 0 + .Fields("ABMLegisladores").Value
            chkModificaActas.Value = 0 + .Fields("ModificaActas").Value
            chkUsuarioMantenimiento.Value = 0 + .Fields("ControlaMantenimiento").Value
            chkExportaActas.Value = 0 + .Fields("ImportaTitulos").Value
            chkDefinirOrdenPresidente.Value = 0 + .Fields("OrdenPresidente").Value
            chkABMUsuarios.Value = 0 + .Fields("ABMUsuarios").Value
            chkImprimirActas.Value = 0 + .Fields("ImprimeActas").Value
            chkConsultaABMLegisladores.Value = 0 + .Fields("ConsultaLegisladores").Value
            chkHabilitaBotonesConsola.Value = 0 + .Fields("HabilitaBotonesConsola").Value
            chkActualizarSB.Value = 0 + .Fields("EnvioDatosSB").Value
            lblRecordSet.Caption = Trim(Str(xPos)) & "/" & Trim(Str(xMax)) & " Usuarios"
        Else                     ' Si no hay registros para mostrar
            Call Limpiar
            lblRecordSet.Caption = "0/0 Usuarios"
        End If
    End With
Exit Sub
TrapError:
    Select Case Err.Number
    
        Case Else
            MsgBox "Error Nº " & Err.Number & Chr(10) & Err.Description & Chr(10) & "Originado en " & Err.Source
            Resume Next
    End Select
Return
End Sub
Private Sub Borrar_Click()
    On Error GoTo TrapError
    Dim xRespuesta As Integer
    If Rs.RecordCount > 0 Then
        xRespuesta = MsgBox("¿Seguro de eliminar al usuario " & UCase(Trim(txtUsuario.Text)) & " de la base de datos?", vbQuestion + vbYesNo)
        If xRespuesta = vbYes Then
            Rs.Delete
            Call Limpiar
            If Rs.RecordCount > 0 Then
                Rs.MoveFirst
                Call MostrarRegistro
            End If
        End If
    End If
Exit Sub
TrapError:
    Select Case Err.Number
        Case Else
            MsgBox "Error Nº " & Err.Number & Chr(10) & Err.Description & Chr(10) & "Originado en " & Err.Source
            Resume
    End Select
Return
End Sub

Private Sub Buscar_Click()
    Dim strId       As String
    Dim blCondicion As Boolean
    
    frmSelUsuarios.Show vbModal, Me
    strId = Trim(lblId.Caption)
    blCondicion = True
    If LCase(strId) <> "nothing" Then
        With Rs
            If .RecordCount > 0 Then
                .MoveFirst
                While blCondicion
                    If .Fields("id").Value = strId Then
                        Call MostrarRegistro
                        blCondicion = False
                    Else
                        .MoveNext
                    End If
                Wend
            End If
        End With
    End If
    
End Sub

Private Sub cmdPassword_Click()
    frmPassword.id_Usuario = Rs.Fields("Id").Value
    frmPassword.NombreUsuario = txtUsuario.Text
    frmPassword.Show vbModal
End Sub

Private Sub Grabar_Click()
    On Error GoTo TrapError
    Dim strApellido             As String
    Dim strNombre               As String
    Dim strUsuario              As String
    Dim strPassword             As String
    Dim xConsultaActas          As Long
    Dim xABMLegisladores        As Long
    Dim xModificaActas          As Long
    Dim xUsuarioMantenimiento   As Long
    Dim xExportaActas           As Long
    Dim xDefinirOrdenPresidente As Long
    Dim xABMUsuarios            As Long
    Dim xImprimirActas          As Long
    
    Dim xConsultaABMLegisla     As Long
    Dim xHabilitaBotonesConsola As Long
    Dim xActualizarSB           As Long
    
    ' ------------------------------------------------------------------------------------
    ' Validar entrada de usuario
    ' ------------------------------------------------------------------------------------
    ' Apellido
    If Trim(txtApellido.Text) = "" Then
        MsgBox "El apellido es un dato obligatorio", vbInformation + vbOKOnly
        txtApellido.SetFocus
        Exit Sub
    End If
    ' Nombre
    If Trim(txtNombre.Text) = "" Then
        MsgBox "El nombre del usuario es un dato obligatorio", vbInformation + vbOKOnly
        txtNombre.SetFocus
        Exit Sub
    End If
    ' Nombre de Usuario
    If Trim(txtUsuario.Text) = "" Then
        MsgBox "El nombre de usuario es un dato obligatorio", vbInformation + vbOKOnly
        txtUsuario.SetFocus
        Exit Sub
    End If
    ' Asegurarse que el nombre de usuario sea unico
       ' Falta hacer
    ' Leer datos ingresados por el usuario
    strApellido = txtApellido.Text
    strNombre = txtNombre.Text
    strUsuario = txtUsuario.Text
    xConsultaActas = chkConsultaActas.Value
    xABMLegisladores = chkABMLegisladores.Value
    xModificaActas = chkModificaActas.Value
    xUsuarioMantenimiento = chkUsuarioMantenimiento.Value
    xExportaActas = chkExportaActas.Value
    xDefinirOrdenPresidente = chkDefinirOrdenPresidente.Value
    xABMUsuarios = chkABMUsuarios.Value
    xImprimirActas = chkImprimirActas.Value
    xConsultaABMLegisla = chkConsultaABMLegisladores.Value
    xHabilitaBotonesConsola = chkHabilitaBotonesConsola.Value
    xActualizarSB = chkActualizarSB.Value
    ' ------------------------------------------------------------------------------------
    ' Grabar en base de datos
    ' ------------------------------------------------------------------------------------
    With Rs
        If blEsNuevo Then
            strPassword = InputBox("Ingrese la palabra secreta de usuario", "Palabra secreta de usuario", strApellido & strNombre)
            If Trim(strPassword) = "" Then
                MsgBox "El la palabra clave de usuario es un dato obligatorio", vbInformation + vbOKOnly
                Exit Sub
            Else
                .AddNew
            End If
        End If
        .Fields("Apellido").Value = strApellido
        .Fields("Nombre").Value = strNombre
        .Fields("Login").Value = strUsuario
        If blEsNuevo Then
            .Fields("Clave").Value = strPassword
        End If
        .Fields("ConsultasActas").Value = xConsultaActas
        .Fields("ABMLegisladores").Value = xABMLegisladores
        .Fields("ModificaActas").Value = xModificaActas
        .Fields("ControlaMantenimiento").Value = xUsuarioMantenimiento
        .Fields("ImportaTitulos").Value = xExportaActas
        .Fields("OrdenPresidente").Value = xDefinirOrdenPresidente
        .Fields("ABMUsuarios").Value = xABMUsuarios
        .Fields("ImprimeActas").Value = xImprimirActas
        .Fields("ConsultaLegisladores").Value = xConsultaABMLegisla
        .Fields("HabilitaBotonesConsola").Value = xHabilitaBotonesConsola
        .Fields("EnvioDatosSB").Value = xActualizarSB
        .Update
        If blEsNuevo Then
            .MoveLast
            blEsNuevo = False
        End If
        Call MostrarRegistro
    End With
    Call ActivarBotones(True)
Exit Sub
TrapError:
    Select Case Err.Number
        Case Else
            MsgBox "Error Nº " & Err.Number & Chr(10) & Err.Description & Chr(10) & "Originado en " & Err.Source
            Resume
    End Select
Return
End Sub
Private Sub SetearRs()
    Dim strSql As String
    Set Rs = New ADODB.Recordset
    strSql = "SELECT * FROM UsuarioConsola ORDER BY Apellido"
    Datos.SetearRsW strSql, Rs
    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
    End If
End Sub
Private Sub Limpiar()
    txtApellido.Text = ""
    txtNombre.Text = ""
    txtUsuario.Text = ""
    chkConsultaActas.Value = 0
    chkABMLegisladores.Value = 0
    chkModificaActas.Value = 0
    chkUsuarioMantenimiento.Value = 0
    chkExportaActas.Value = 0
    chkDefinirOrdenPresidente.Value = 0
    chkABMUsuarios.Value = 0
End Sub
Private Sub ActivarBotones(blEstado As Boolean)
    Nuevo.Enabled = blEstado
    Grabar.Enabled = True
    Borrar.Enabled = blEstado
    
    cmdAnterior.Enabled = blEstado
    cmdSiguiente.Enabled = blEstado
    cmdPrimero.Enabled = blEstado
    cmdUltimo.Enabled = blEstado
    lblRecordSet.Enabled = blEstado
    cmdPassword.Enabled = blEstado
    If blEstado Then
        Salir.Caption = "&Salir"
    Else
        Salir.Caption = "&Cancelar"
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Nuevo_Click()
    Call ActivarBotones(False)
    Call Limpiar
    blEsNuevo = True
    txtApellido.SetFocus
End Sub

Private Sub cmdAnterior_Click()
    If Rs.AbsolutePosition > 1 Then
        Rs.MovePrevious
    Else
        Rs.MoveFirst
    End If
    Call MostrarRegistro
End Sub
Private Sub cmdSiguiente_Click()
    If Rs.AbsolutePosition < Rs.RecordCount Then
        Rs.MoveNext
    Else
        Rs.MoveLast
    End If
    Call MostrarRegistro
End Sub
Private Sub cmdPrimero_Click()
    Rs.MoveFirst
    Call MostrarRegistro
End Sub

Private Sub cmdUltimo_Click()
    Rs.MoveLast
    Call MostrarRegistro
End Sub
Private Sub Form_Load()
        blEsNuevo = False
        Call SetearRs
        Call MostrarRegistro
End Sub
Private Sub Salir_Click()
    If Salir.Caption = "&Salir" Then
        Unload Me
    Else
        blEsNuevo = False
        Call ActivarBotones(True)
        Call MostrarRegistro
    End If
End Sub


