VERSION 5.00
Begin VB.Form frmTiempoVotacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tiempo de votación"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4815
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTiempo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2070
      TabIndex        =   0
      Top             =   1110
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2190
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3510
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblTiempoActual 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2070
      TabIndex        =   7
      Top             =   750
      Width           =   1125
   End
   Begin VB.Label Label4 
      Caption         =   "Tiempo de votación actual:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   6
      Top             =   750
      Width           =   1995
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "segundos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3210
      TabIndex        =   5
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tiempo de votación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   1140
      Width           =   1845
   End
   Begin VB.Label Label1 
      Caption         =   "Indique el tiempo de votación en segundos, o selecciónelo entre las opciones disponibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4530
   End
End
Attribute VB_Name = "frmTiempoVotacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cargo As Boolean
Private Sub cmdAceptar_Click()
    Dim tiempo As String
    If txtTiempo.Text = "" Then
        MsgBox "Ingrese un valor de tiempo o cancele la operación."
        Exit Sub
    Else
        If IsNumeric(txtTiempo.Text) Then
            Dim num As Integer
            num = txtTiempo.Text
            If (num < 5) Then
                MsgBox "El tiempo no puede ser menor a 5."
                Exit Sub
            End If
        Else
            MsgBox "El valor es inválido."
        End If
    End If
    If (gTipoUsuario = 0) Or (gTipoUsuario = 2) Then
        'Modificacion de AP. Volvi a poner txtTiempo.Text, no se veia la informacion
        'If MsgBox("Está Ud. seguro de establecer el tiempo de votación a " & tiempo & " (MM:SS)", vbQuestion + vbYesNo, "Confirma la operación?") = vbYes Then
        If MsgBox("Está Ud. seguro de establecer el tiempo de votación a " & txtTiempo.Text & " segundos ?", vbQuestion + vbYesNo, "Confirma la operación?") = vbYes Then
            MensajesSQV.tiempo txtTiempo.Text
            Unload Me
        End If
    Else
        MsgBox "Ud. no dispone de permisos para realizar esta acción.", vbInformation + vbOKOnly
    End If
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
Cargo = False
SetearRs "SELECT Tiempo_de_votación FROM vector", rsTemp
If rsTemp.EOF Then
    Call MsgBox("Error al obtener el tiempo de votación", vbCritical)
    Unload Me
End If
If IsNull(rsTemp.Fields(0)) Then
    lblTiempoActual.Caption = "0"
Else
    lblTiempoActual.Caption = rsTemp.Fields(0) & " segundos"
End If
'Select Case txtTiempo.Text
'    Case "15"
'        OpcionTiempo(15).Value = True
'    Case "10"
'        OpcionTiempo(10).Value = True
'    Case "5"
'        OpcionTiempo(5).Value = True
'End Select
txtTiempo.Text = "15"
txtTiempo.SelStart = 0
txtTiempo.SelLength = "2"
End Sub
Private Sub txtTiempo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar_Click
    Else
        If ((KeyAscii < vbKey0) Or (KeyAscii > vbKey9)) And (KeyAscii <> vbKeyBack) Then
            KeyAscii = 0
        End If
    End If
End Sub
