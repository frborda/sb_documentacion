VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inicio de sesión en Sistema"
   ClientHeight    =   2880
   ClientLeft      =   2835
   ClientTop       =   3435
   ClientWidth     =   6465
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1701.6
   ScaleMode       =   0  'User
   ScaleWidth      =   6070.285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkModoPrueba 
      BackColor       =   &H00404040&
      Caption         =   "Modo &Prueba"
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
      Height          =   225
      Left            =   90
      TabIndex        =   4
      Top             =   2490
      Width           =   1635
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   3000
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1515
      Width           =   3315
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1905
      Width           =   3315
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   -60
      TabIndex        =   7
      Top             =   1350
      Width           =   6975
      Begin VB.Label lblLabels 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "&Contraseña:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1950
         TabIndex        =   9
         Top             =   600
         Width           =   1440
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   2250
         TabIndex        =   8
         Top             =   210
         Width           =   780
      End
      Begin VB.Image picLogo 
         Appearance      =   0  'Flat
         Height          =   795
         Left            =   0
         Picture         =   "frmLogin.frx":030A
         Top             =   0
         Width           =   1980
      End
   End
   Begin Proyecto1.ButtonOffice cmdOK 
      Height          =   435
      Left            =   2850
      TabIndex        =   2
      Top             =   2370
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   767
      BackColor       =   12230304
      Caption         =   "Aceptar"
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
   Begin Proyecto1.ButtonOffice cmdCancel 
      Height          =   435
      Left            =   4620
      TabIndex        =   3
      Top             =   2370
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   767
      BackColor       =   12230304
      Caption         =   "Cancelar"
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
   Begin VB.Label lblVersion 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
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
      Left            =   60
      TabIndex        =   10
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "de la Nación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   3
      Left            =   1560
      TabIndex        =   6
      Top             =   510
      Width           =   3180
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Honorable Cámara De Diputados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   2
      Left            =   330
      TabIndex        =   5
      Top             =   60
      Width           =   5775
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AMenu As Boolean
Option Explicit

Private Sub chkModoPrueba_Click()
    Dim strSql As String
    
    If chkModoPrueba.Value = 1 Then
        strconexion = strConexionConfig
        AbrirDB
        FlagBasePrueba = True
        strSql = "UPDATE Configuracion SET Valor = '" & strBasePrueba & "' WHERE Variable = 'base_vigente'"
        SenteciaSQl strSql
        strconexion = strBasePrueba
        Modo_Prueba_Seleccionado = True
    Else
        strconexion = strConexionConfig
        AbrirDB
        FlagBasePrueba = False
        strSql = "UPDATE Configuracion SET Valor = '" & strBaseProduccion & "' WHERE Variable = 'base_vigente'"
        SenteciaSQl strSql
        strconexion = strBaseProduccion
        Modo_Prueba_Seleccionado = False
    End If
    AbrirDB
End Sub

Private Sub cmdCancel_Click()
    'establecer la variable global a false
    'para indicar un inicio de sesión fallido
    gLoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    AMenu = True
    AbrirDB
    chkModoPrueba_Click
    'comprobar si la contraseña es correcta
    If validarControles = True Then
        If esClaveValida = True Then
            gLoginSucceeded = True
            EntroAMenu = True
            Unload Me
        Else
            txtPassword.SetFocus
    '        SendKeys "{Home}+{End}"
        End If
    End If
End Sub
Private Function esClaveValida() As Boolean
    Dim rstClave As New ADODB.Recordset
    Dim strSql As String
    strSql = "SELECT * FROM UsuarioConsola WHERE (Login='" & txtUserName.Text & "') AND (Clave='" & txtPassword.Text & "')"
    SetearRs strSql, rstClave
    If rstClave.EOF = False Then
        'If IsNull(rstClave!TipoUsuario) = False Then
        '    gTipoUsuario = rstClave!TipoUsuario
        'Else
        '    gTipoUsuario = 0
        'End If
        With PermisosTotales
            .id_persona = rstClave!id
            .ABMLegisladores = rstClave!ABMLegisladores
            .ABMUsuarios = rstClave!ABMUsuarios
            .ConsultaActas = rstClave!ConsultasActas
            .DefinirOrdenPresidente = rstClave!OrdenPresidente
            .ExportaActas = rstClave!ImportaTitulos
            .ModificaActas = rstClave!ModificaActas
            .UsuarioMantenimiento = rstClave!ControlaMantenimiento
            .ImprimeActas = rstClave!ImprimeActas
            .ConsultaABMLegislador = rstClave!ConsultaLegisladores
            .HabilitaBotonesConsola = rstClave!HabilitaBotonesConsola
            .ActualizaaSB = rstClave!EnvioDatosSB
        End With
        
        esClaveValida = True
    Else
        MsgBox "Los datos ingresados no correponden a un usuario registrado en el Sistema.", vbExclamation + vbOKOnly
        txtUserName.SetFocus
        esClaveValida = False
    End If
    If rstClave.State = adStateOpen Then
        rstClave.Close
    End If
    Set rstClave = Nothing
End Function
Private Function validarControles() As Boolean
    If (txtUserName.Text = "") Or (txtPassword.Text = "") Then
        MsgBox "Debe ingresar una combinación válida de nombre de usuario y clave antes de continuar.", vbInformation + vbOKOnly
        validarControles = False
    Else
        validarControles = True
    End If
End Function

Private Sub SetVersion()
 lblVersion.Caption = "Consola " & Consola_Version
 frmLogin.Caption = "Inicio de Sesión"
End Sub

Private Sub Form_Activate()
    Dim strMensaje As String
    Dim strSql     As String
    Dim Rs1        As ADODB.Recordset
    AutoCaptura = False
    html = ""
    Modo_Prueba_Seleccionado = False
    VistaPrevia = False
    Call SetVersion
    Set Rs1 = New ADODB.Recordset
    
    If App.PrevInstance = True Then
        End
    End If
    ' ----------------------------------------------------------------------
    ' Conexion con la base de datos de configuracion
    ' ----------------------------------------------------------------------
    strconexion = strConexionConfig
    AbrirDB
    WSData.createConfig
End Sub
Public Sub SetDefault()
Dim strSql As String
strconexion = strConexionConfig
AbrirDB
FlagBasePrueba = False
strSql = "UPDATE Configuracion SET Valor = '" & strBaseProduccion & "' WHERE Variable = 'base_vigente'"
SenteciaSQl strSql
strconexion = strBaseProduccion
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCaracter(KeyAscii)
End Sub
Private Sub Form_Load()
AMenu = False
EntroAMenu = False
EntroAConsola = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
If EntroAMenu = False Then
    SetDefault
End If
If AMenu = False Then
    'MsgBox "Hace Unload"
End If
End Sub
Private Sub txtPassword_Change()
    ' Funciones.seleccionadoTxt txtPassword
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOK_Click
    End If
End Sub
Private Sub txtUserName_GotFocus()
    Funciones.seleccionadoTxt txtUserName
End Sub
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPassword.SetFocus
    End If
End Sub
