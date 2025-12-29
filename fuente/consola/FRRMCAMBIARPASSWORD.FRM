VERSION 5.00
Begin VB.Form frrmCambiarPassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualizar Palabra Secreta"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   7455
      Begin VB.TextBox txtUsuario 
         BackColor       =   &H00C0FFFF&
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
         IMEMode         =   3  'DISABLE
         Left            =   3825
         TabIndex        =   15
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H00C0FFFF&
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
         IMEMode         =   3  'DISABLE
         Left            =   3825
         TabIndex        =   14
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtPassword2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3825
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtPassword1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3825
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtActualPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3825
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Usuario : "
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   645
         Width           =   3645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre : "
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   3645
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   7200
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   7200
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Reingrese nuevamente su nueva palabra secreta : "
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1950
         Width           =   3645
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ingrese su nueva palabra secreta : "
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1605
         Width           =   3645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ingrese su actual palabra secreta : "
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1230
         Width           =   3645
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   915
      Left            =   6360
      ScaleHeight     =   855
      ScaleWidth      =   1245
      TabIndex        =   6
      Top             =   120
      Width           =   1300
      Begin VB.CommandButton Salir 
         Caption         =   "&Salir"
         Height          =   855
         Left            =   0
         Picture         =   "frrmCambiarPassword.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   915
      Left            =   240
      ScaleHeight     =   855
      ScaleWidth      =   1245
      TabIndex        =   0
      Top             =   120
      Width           =   1300
      Begin VB.CommandButton Grabar 
         Caption         =   "&Grabar"
         Height          =   855
         Left            =   0
         Picture         =   "frrmCambiarPassword.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Mantenimiento de palabra secreta de usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1620
      TabIndex        =   7
      Top             =   225
      Width           =   4275
   End
End
Attribute VB_Name = "frrmCambiarPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsUsr As ADODB.Recordset

Private Sub BuscarUsuario()
    Dim strSql As String
    Set RsUsr = New ADODB.Recordset
    strSql = "SELECT * FROM UsuarioConsola WHERE (Id = " & PermisosTotales.id_persona & ")"
    SetearRsW strSql, RsUsr
    With RsUsr
        txtNombre.Text = .Fields("Apellido").Value & ", " & .Fields("Nombre").Value
        txtUsuario.Text = .Fields("Login").Value
        txtNombre.Locked = True
        txtUsuario.Locked = True
    End With
'    Text1.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call BuscarUsuario
End Sub

Private Sub Grabar_Click()
    Dim strActualPassword As String
    Dim strPassword1      As String
    Dim strPassword2      As String
    
    If Trim(txtActualPassword.Text) = "" Then
        MsgBox "Debe ingresar su actual palabra secreta", vbInformation + vbOKOnly
        txtActualPassword.SetFocus
        Exit Sub
    Else
        strActualPassword = txtActualPassword.Text
    End If
    If Trim(txtPassword1.Text) = "" Then
        MsgBox "Debe ingresar su nueva palabra secreta", vbInformation + vbOKOnly
        txtPassword1.SetFocus
        Exit Sub
    Else
        strPassword1 = txtPassword1.Text
    End If
    If Trim(txtPassword2.Text) = "" Then
        MsgBox "Debe ingresar su nueva palabra secreta", vbInformation + vbOKOnly
        txtPassword2.SetFocus
        Exit Sub
    Else
        strPassword2 = txtPassword2.Text
    End If
    If strPassword2 <> strPassword1 Then
        MsgBox "Error en la definición de la nueva palabrra secreta. Reintente la operación", vbInformation + vbOKOnly
        txtPassword1.Text = ""
        txtPassword2.Text = ""
        txtPassword1.SetFocus
        Exit Sub
    End If
    
    With RsUsr
        If .Fields("Clave") <> strActualPassword Then
            MsgBox "Error al ingresar su actual palabra secreta", vbInformation + vbOKOnly
            txtActualPassword.SetFocus
            Exit Sub
        Else
            .Fields("Clave").Value = strPassword1
            .Update
            MsgBox "Su palabra clave se ha actualizado con exito"
        End If
    End With
End Sub
Private Sub Salir_Click()
    Unload Me
End Sub
