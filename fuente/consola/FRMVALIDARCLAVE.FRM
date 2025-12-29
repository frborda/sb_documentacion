VERSION 5.00
Begin VB.Form frmValidarClave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Validación de clave de usuario"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2580
      TabIndex        =   1
      Top             =   960
      Width           =   915
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3540
      TabIndex        =   2
      Top             =   960
      Width           =   915
   End
   Begin VB.TextBox txtClave 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1500
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Por favor, vuelva a ingresar la nueva contraseña del usuario."
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   4440
   End
   Begin VB.Label Label3 
      Caption         =   "Contraseña"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   1380
   End
End
Attribute VB_Name = "frmValidarClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mContrasegna As String

Public Property Let ContrasegnaIngresada(ByVal vNewValue As String)
    mContrasegna = vNewValue
End Property

Private Sub cmdAceptar_Click()
    If validarDatos = True Then
        If txtClave.Text = mContrasegna Then
            FrmGestionarUsuarios.ContrasegnaValida = True
            Unload Me
        Else
            MsgBox "La contraseña ingresada no corresponde con la nueva contraseña del usuario.", vbExclamation + vbOKOnly, "Vuelva a intentar"
            FrmGestionarUsuarios.ContrasegnaValida = False
        End If
    End If
End Sub

Private Function validarDatos() As Boolean
    If txtClave.Text = "" Then
        MsgBox "Debe ingresar una comprobación de contraseña válida.", vbInformation + vbOKOnly
        validarDatos = False
        txtClave.SetFocus
    Else
        validarDatos = True
    End If
End Function
Private Sub cmdCancelar_Click()
    FrmGestionarUsuarios.ContrasegnaValida = False
    Unload Me
End Sub

Private Sub Form_Load()
    FrmGestionarUsuarios.ContrasegnaValida = False
End Sub
