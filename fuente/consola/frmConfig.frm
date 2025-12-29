VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurar Archivo de Conexión a datos"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Borrar Configuración"
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      ToolTipText     =   "Elimina actual archivo SQV.DAT"
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4935
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1785
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtUsuario 
         Height          =   285
         Left            =   1785
         TabIndex        =   9
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtBase 
         Height          =   285
         Left            =   1785
         TabIndex        =   8
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   1785
         TabIndex        =   7
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Password : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Usuario : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Base de Datos : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Servidor SQL : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      ToolTipText     =   "Cancelar operación y salir"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdGenerarSQVdat 
      Caption         =   "Guardar Configuración"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      ToolTipText     =   "Generar archivo de configuración inicial"
      Top             =   780
      Width           =   1935
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strArchivo As String
Private strCadenaConexionGlobal  As String

Private Sub InicializarValores()
    strArchivo = App.Path & "\Consola.dat"
   ' strCadenaConexionGlobal = "Provider=SQLOLEDB.1;Password=;Persist Security Info=True;" _
              & "User ID=sa;Initial Catalog=sqv;Data Source=ADV-PRO1"
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdGenerarSQVdat_Click()

    On Error GoTo Trap_Error

    Dim strServer   As String
    Dim strBase     As String
    Dim strUsuario  As String
    Dim strPassword As String
    Dim Cn          As ADODB.Connection
    
    strServer = txtServer.Text
    strBase = txtBase.Text
    strUsuario = txtUsuario.Text
    strPassword = txtPassword.Text
    ' ------------------------------------------------------------------------------------------
    ' Generar string de conexion
    ' ------------------------------------------------------------------------------------------
    strCadenaConexionGlobal = "Provider=SQLOLEDB.1;Password=" & strPassword & ";Persist Security Info=True;" _
              & "User ID=" & strUsuario & ";Initial Catalog=" & strBase & ";Data Source=" & strServer
    'strCadenaConexionGlobal = "Provider=SQLOLEDB.1;Password=" & strPassword & ";Persist Security Info=True;" _
              & "User ID=" & strUsuario & ";Initial Catalog=" & strBase & ";Data Source=" & strServer & _
              ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=ELCO02;Use Encryption for Data=False;Tag with column collation when possible=False"
    ' ------------------------------------------------------------------------------------------
    ' Probar string de conexion
    ' ------------------------------------------------------------------------------------------
    Set Cn = New ADODB.Connection
    With Cn
        .ConnectionString = strCadenaConexionGlobal
        .CursorLocation = adUseServer
        .ConnectionTimeout = 10
        .Open
    End With
    ' ------------------------------------------------------------------------------------------
    ' Si no hubo errores en la conexion, se guardan datos en archivos SQV.DAT
    ' ------------------------------------------------------------------------------------------
    If Err.Number = 0 Then
        Call GuardarArchivo
    End If
    Unload Me
    
Exit Sub
Trap_Error:
    MsgBox "Los valores de configuración no son correctos. Modifique los parámetros de acceso al servidor de base de datos y reintente la operación"
    Exit Sub
End Sub
Private Sub GuardarArchivo()
    
    Dim xFile As Long
    Dim strTemp As String
    
    xFile = FreeFile
    Datos.CadenaConexion = strCadenaConexionGlobal
    'strTemp = Encode(strCadenaConexionGlobal)
    strTemp = Encripta.EncryptString(strCadenaConexionGlobal)
    Open strArchivo For Binary As #xFile
        Put #xFile, , strTemp
    Close #xFile
End Sub
Private Sub Command1_Click()
    Dim xRespuesta As Long
    On Error GoTo Trap_Error
    
    xRespuesta = MsgBox("¿Esta seguro que desea eliminar el actual archivo de configuración Consola.DAT?", vbQuestion + vbYesNo, "Eliminar archivo Consola.DAT")
    If xRespuesta = vbYes Then
        Kill strArchivo
    End If
Exit Sub
Trap_Error:
    Select Case Err.Number
        Case 53
            MsgBox "No existe el archivo Consola.DAT en el directorio " & App.Path
            Resume Next
        Case Else
            MsgBox "Error " & Err.Number & Chr(10) & Err.Description & Chr(10) & "Originado en " & Err.Source
            Resume
    End Select
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    Call InicializarValores
End Sub

