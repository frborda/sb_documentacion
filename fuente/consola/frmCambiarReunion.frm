VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmCambiarReunion 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambio de Reunión"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdAceptar 
      Height          =   645
      Left            =   90
      TabIndex        =   1
      Top             =   750
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1138
      BackColor       =   12230304
      Caption         =   "&Aceptar"
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
   Begin VB.TextBox txtReunion 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   60
      TabIndex        =   0
      Text            =   "5"
      Top             =   60
      Width           =   4185
   End
   Begin Proyecto1.ButtonOffice cmdCancelar 
      Height          =   645
      Left            =   2310
      TabIndex        =   2
      Top             =   750
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1138
      BackColor       =   12230304
      Caption         =   "Cancelar"
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
End
Attribute VB_Name = "frmCambiarReunion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
'cambiarReunion (txtNumeroReunion.Text)
cambiarReunion (txtReunion.Text)
Unload Me
End Sub
Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If (txtReunion.Text <> "") Then
    txtReunion.SelStart = Len(txtReunion.Text)
    txtReunion.SelLength = Len(txtReunion)
End If
End Sub

Private Sub txtReunion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAceptar_Click
End If
End Sub
