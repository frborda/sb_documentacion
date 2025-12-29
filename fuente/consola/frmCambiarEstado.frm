VERSION 5.00
Begin VB.Form frmCambiarEstado 
   BorderStyle     =   0  'None
   Caption         =   "Seleccione un estado"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOrdenPresidente 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Text            =   "99"
      Top             =   2220
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observaciones"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   2835
      Begin VB.TextBox txtObservaciones 
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2595
      End
   End
   Begin VB.CommandButton cmbCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   2700
      Width           =   1455
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "Aplicar"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2700
      Width           =   1455
   End
   Begin VB.ComboBox cmbEstados 
      Height          =   315
      ItemData        =   "frmCambiarEstado.frx":0000
      Left            =   120
      List            =   "frmCambiarEstado.frx":0002
      TabIndex        =   0
      Text            =   "-Seleccione un estado-"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Orden Presidente:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   4
      Height          =   3105
      Left            =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmCambiarEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbCancelar_Click()
Unload Me
End Sub
Private Sub cmdAplicar_Click()
If cmbEstados.ListIndex = -1 Then
    MsgBox "Debe seleccionar un estado!", vbCritical
Else
    frmHistorico.EstadoACambiar = cmbEstados.Text
    frmHistorico.ObservacionesACambiar = txtObservaciones.Text
    frmHistorico.OrdenPresidente = txtOrdenPresidente.Text
    Unload Me
End If
End Sub
Private Sub Form_Load()
frmHistorico.LlenaCombo cmbEstados, "descripcion", "estados", "orden"
End Sub
