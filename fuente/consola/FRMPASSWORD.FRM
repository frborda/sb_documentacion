VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password de Usuario"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   4155
      Begin VB.PictureBox Picture1 
         Height          =   915
         Left            =   180
         ScaleHeight     =   855
         ScaleWidth      =   1245
         TabIndex        =   9
         Top             =   240
         Width           =   1305
         Begin VB.CommandButton Grabar 
            Caption         =   "&Grabar"
            Height          =   855
            Left            =   0
            Picture         =   "frmPassword.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   0
            Width           =   1245
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   915
         Left            =   2670
         ScaleHeight     =   855
         ScaleWidth      =   1245
         TabIndex        =   7
         Top             =   240
         Width           =   1305
         Begin VB.CommandButton Salir 
            Caption         =   "&Salir"
            Height          =   855
            Left            =   0
            Picture         =   "frmPassword.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   0
            Width           =   1245
         End
      End
      Begin VB.TextBox txtPassword2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2010
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1965
         Width           =   1905
      End
      Begin VB.TextBox txtPassword1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2010
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1605
         Width           =   1905
      End
      Begin VB.Label lblUsuario 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1140
         TabIndex        =   6
         Top             =   1275
         Width           =   75
      End
      Begin VB.Label Label2 
         Caption         =   "Reingrese Password : "
         Height          =   285
         Left            =   270
         TabIndex        =   4
         Top             =   1965
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Left            =   270
         TabIndex        =   2
         Top             =   1275
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese Password : "
         Height          =   285
         Left            =   270
         TabIndex        =   1
         Top             =   1605
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private xIdUsuario       As Long
Private strNombreUsuario As String
Private strSql           As String
Private Rs               As ADODB.Recordset


Private Sub BuscarUsuario()
    Set Rs = New ADODB.Recordset
    SetearRsW strSql, Rs
    With Rs
        If .RecordCount = 1 Then
            If .Fields("Clave").Value <> "" Then
                txtPassword1.Text = .Fields("Clave").Value
                txtPassword2.Text = .Fields("Clave").Value
            Else
                txtPassword1.Text = ""
                txtPassword2.Text = ""
            End If
        End If
    End With
End Sub


Public Property Let id_Usuario(ByVal vNewValue As Variant)
    xIdUsuario = vNewValue
    strSql = "SELECT Clave FROM UsuarioConsola WHERE id = " & Trim(Str(xIdUsuario))
    Call BuscarUsuario
End Property


Public Property Let NombreUsuario(ByVal vNewValue As Variant)
    strNombreUsuario = vNewValue
    lblUsuario.Caption = strNombreUsuario
End Property
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Grabar_Click()
    If Trim(txtPassword1.Text) = "" Then
        MsgBox "Debe ingresar el password del usuario", vbInformation + vbOKOnly
        txtPassword1.SetFocus
        Exit Sub
    End If
    If txtPassword1.Text <> txtPassword2.Text Then
        MsgBox "Error en el ingreso del password. Reintente la operacion", vbInformation + vbOKOnly
        txtPassword1.SetFocus
        Exit Sub
    End If
    Rs.Fields("Clave").Value = txtPassword1.Text
    Rs.Update
    Unload Me
End Sub

Private Sub Salir_Click()
    Unload Me
End Sub
Private Sub txtPassword1_GotFocus()
    With txtPassword1
        .SelStart = 0
        .SelLength = Len(txtPassword1.Text)
    End With
End Sub
Private Sub txtPassword2_GotFocus()
    With txtPassword2
        .SelStart = 0
        .SelLength = Len(txtPassword1.Text)
    End With
End Sub
