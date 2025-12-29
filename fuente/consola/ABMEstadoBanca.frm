VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmABMEstadoBanca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estados de las Bancas"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   4635
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "..."
      Height          =   285
      Left            =   4005
      TabIndex        =   16
      Top             =   900
      Width           =   285
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6030
      TabIndex        =   15
      Top             =   4365
      Width           =   1275
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3150
      TabIndex        =   14
      Top             =   4365
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4590
      TabIndex        =   13
      Top             =   4365
      Width           =   1275
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   6030
      TabIndex        =   12
      Top             =   1530
      Width           =   1275
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   6030
      TabIndex        =   11
      Top             =   1080
      Width           =   1275
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   6030
      TabIndex        =   10
      Top             =   630
      Width           =   1275
   End
   Begin VB.TextBox txtNombreColor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2115
      TabIndex        =   4
      Top             =   1530
      Width           =   1815
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2115
      TabIndex        =   3
      Top             =   1215
      Width           =   3435
   End
   Begin VB.TextBox txtColor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2115
      TabIndex        =   2
      Top             =   900
      Width           =   1815
   End
   Begin VB.TextBox txtEstado 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2115
      TabIndex        =   1
      Top             =   585
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid flxEstados 
      Height          =   1950
      Left            =   225
      TabIndex        =   0
      Top             =   2205
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   3440
      _Version        =   393216
      BackColor       =   16777215
      Appearance      =   0
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2115
      TabIndex        =   18
      Top             =   270
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      Height          =   195
      Index           =   4
      Left            =   270
      TabIndex        =   17
      Top             =   270
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Listado de Tipos de Estado Registrados"
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
      Index           =   3
      Left            =   270
      TabIndex        =   9
      Top             =   1935
      Width           =   3435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Color"
      Height          =   195
      Index           =   2
      Left            =   270
      TabIndex        =   8
      Top             =   1530
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion del Estado"
      Height          =   195
      Index           =   1
      Left            =   270
      TabIndex        =   7
      Top             =   1215
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Relacionado"
      Height          =   195
      Left            =   270
      TabIndex        =   6
      Top             =   900
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Estado"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   5
      Top             =   585
      Width           =   1365
   End
End
Attribute VB_Name = "frmABMEstadoBanca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objE As New estadoBanca

Private Sub cmdAceptar_Click()
   If txtEstado.Text = "" Then
      MsgBox "Ingrese el nombre del estado.", vbExclamation
      If txtEstado.Enabled Then txtEstado.SetFocus
      Exit Sub
   End If
   If txtColor.Text = "" Then
      MsgBox "Ingrese el color del estado.", vbExclamation
      If txtColor.Enabled Then txtColor.SetFocus
      Exit Sub
   End If
   If txtDescripcion.Text = "" Then
      MsgBox "Ingrese la descripcion del estado.", vbExclamation
      If txtDescripcion.Enabled Then txtDescripcion.SetFocus
      Exit Sub
   End If
   If txtNombreColor.Text = "" Then
      MsgBox "Ingrese el nombre del color.", vbExclamation
      If txtNombreColor.Enabled Then txtNombreColor.SetFocus
      Exit Sub
   End If
   Screen.MousePointer = vbArrowHourglass
   Set objE = New estadoBanca
   objE.Codigo = Val(lblCodigo.Caption)
   objE.Nombre = Trim(txtEstado.Text)
   objE.CodigoColor = Trim(txtColor.Text)
   objE.Descripcion = Trim(txtDescripcion.Text)
   objE.NombreColor = Trim(txtNombreColor.Text)
   If objE.Existe Then
      objE.Actualizar
   Else
      objE.Guardar
   End If
   Set objE = Nothing
   CargarGrilla
   Screen.MousePointer = vbDefault
   Me.Caption = "Estados de las Bancas"
   limpiarCampos
   InhabilitarCampos
   cmdNUevo.Enabled = True
   cmdModificar.Enabled = False
   cmdEliminar.Enabled = False
   cmdAceptar.Enabled = False
   cmdCancelar.Enabled = False
   cmdSalir.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
   limpiarCampos
   InhabilitarCampos
   cmdNUevo.Enabled = True
   cmdModificar.Enabled = False
   cmdEliminar.Enabled = False
   cmdAceptar.Enabled = False
   cmdCancelar.Enabled = False
   cmdSalir.Enabled = True
   Me.Caption = "Estados de las Bancas"
End Sub

Private Sub cmdColor_Click()
   CommonDialog.ShowColor
   txtColor.Text = Hex(CommonDialog.Color)
End Sub

Private Sub cmdEliminar_Click()
Dim RTA As Byte
   RTA = MsgBox("Desea eliminar un tipo de estado?", vbCritical + vbYesNo + vbDefaultButton2)
   If RTA = vbNo Then Exit Sub
   Set objE = New estadoBanca
   If Not objE.Eliminar(Val(lblCodigo.Caption)) Then
      MsgBox "Imposible eliminar un estado asignado a una Banca.", vbExclamation
      Set objE = Nothing
      Exit Sub
   End If
   CargarGrilla
   limpiarCampos
   Set objE = Nothing
End Sub

Private Sub cmdModificar_Click()
   Me.Caption = "Estados de las Bancas [Edicion]"
   HabilitarCampos
   cmdNUevo.Enabled = False
   cmdModificar.Enabled = False
   cmdEliminar.Enabled = False
   cmdAceptar.Enabled = True
   cmdCancelar.Enabled = True
   cmdSalir.Enabled = False
End Sub

Private Sub cmdNuevo_Click()
   Me.Caption = "Estados de las Bancas [Nuevo]"
   HabilitarCampos
   limpiarCampos
   cmdNUevo.Enabled = False
   cmdAceptar.Enabled = True
   cmdCancelar.Enabled = True
   cmdSalir.Enabled = False
   txtEstado.SetFocus
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub flxEstados_DblClick()
   cmdNUevo.Enabled = False
   cmdModificar.Enabled = True
   cmdEliminar.Enabled = True
   cmdAceptar.Enabled = False
   cmdCancelar.Enabled = True
   cmdSalir.Enabled = False
   InhabilitarCampos
   With flxEstados
      lblCodigo.Caption = .TextMatrix(.Row, 0)
      txtEstado.Text = .TextMatrix(.Row, 1)
      txtColor.Text = .TextMatrix(.Row, 2)
      txtDescripcion.Text = .TextMatrix(.Row, 3)
      txtNombreColor.Text = .TextMatrix(.Row, 4)
   End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
'cGRIS = &HC0C0C0
'cBLANCO = &HFFFFFF
'cAMARILLO = &HFFFF&
'cROJO = &HFF&
'cCELESTE = &HFFFF00
'cNARANJA = &H80FF&
'cVERDE = &HFF00&
'cNEGRO = &H0&
   flxEstados.FormatString = "Id|Nombre del Estado|Color Asignado|Descripción|Nombre Color"
   CargarGrilla
   cmdAceptar.Enabled = False
   cmdCancelar.Enabled = False
   cmdModificar.Enabled = False
   cmdEliminar.Enabled = False
   InhabilitarCampos
End Sub

Private Sub CargarGrilla()
Dim rst As New ADODB.Recordset
   flxEstados.Cols = 5
   flxEstados.Rows = 1
   If SetearRs(objE.GenerarSql, rst) Then
      While Not rst.EOF
         flxEstados.AddItem rst.Fields("id") & vbTab & rst.Fields("estado") & vbTab & rst.Fields("color") & vbTab & rst.Fields("descripcion") & vbTab & rst.Fields("nombrecolor")
         flxEstados.RowData(flxEstados.Rows - 1) = rst.Fields("id")
         rst.MoveNext
      Wend
   End If
   Set rst = Nothing
   Set objE = Nothing
End Sub

Private Sub InhabilitarCampos()
   txtEstado.Enabled = False
   txtColor.Enabled = False
   txtDescripcion.Enabled = False
   txtNombreColor.Enabled = False
   cmdColor.Enabled = False
End Sub

Private Sub HabilitarCampos()
   txtEstado.Enabled = True
   txtColor.Enabled = True
   txtDescripcion.Enabled = True
   txtNombreColor.Enabled = True
   cmdColor.Enabled = True
End Sub

Private Sub limpiarCampos()
   lblCodigo.Caption = ""
   txtEstado.Text = ""
   txtColor.Text = ""
   txtDescripcion.Text = ""
   txtNombreColor.Text = ""
End Sub
