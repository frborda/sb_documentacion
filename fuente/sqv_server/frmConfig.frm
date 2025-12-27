VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SQV - SERVER: Configurar Archivo de Inicio"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Borrar SQV.DAT"
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      ToolTipText     =   "Elimina actual archivo SQV.DAT"
      Top             =   120
      Width           =   1695
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
         Text            =   "sa"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtBase 
         Height          =   285
         Left            =   1785
         TabIndex        =   8
         Text            =   "sqv"
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   1785
         TabIndex        =   7
         Text            =   "ADV-PRO1"
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Alignment       =   1  'Right Justify
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
         Alignment       =   1  'Right Justify
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
         Alignment       =   1  'Right Justify
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
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenerarSQVdat 
      Caption         =   "Generar SQV.DAT"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      ToolTipText     =   "Generar archivo de configuración inicial"
      Top             =   780
      Width           =   1695
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
    strArchivo = App.Path & "\sqv.dat"
    strCadenaConexionGlobal = "Provider=SQLOLEDB.1;Password=;Persist Security Info=True;" _
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
    If err.Number = 0 Then
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
    strConexion = strCadenaConexionGlobal
    strTemp = Encripta.EncryptString(strConexion)
    Open strArchivo For Binary As #xFile
        Put #xFile, , strTemp
    Close #xFile
End Sub
Private Sub Command1_Click()
    Dim xRespuesta As Long
    On Error GoTo Trap_Error
    
    xRespuesta = MsgBox("¿Esta seguro que desea eliminar el actual archivo de configuración SQV.DAT?", vbQuestion + vbYesNo, "Eliminar archivo SQV.DAT")
    If xRespuesta = vbYes Then
        Kill strArchivo
    End If
Exit Sub
Trap_Error:
    Select Case err.Number
        Case 53
            MsgBox "No existe el archivo SQV.DAT en el directorio " & App.Path
            Resume Next
        Case Else
            MsgBox "Error " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            Resume
    End Select
    
End Sub
Private Sub Form_Load()
    Call InicializarValores
End Sub
