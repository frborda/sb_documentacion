VERSION 5.00
Begin VB.Form frmABMMandatos 
   Caption         =   "Mandatos"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   390
      Left            =   6000
      ScaleHeight     =   330
      ScaleWidth      =   690
      TabIndex        =   13
      Top             =   1935
      Width           =   750
      Begin VB.CommandButton cmdSiguiente 
         Height          =   325
         Left            =   0
         Picture         =   "frmABMMandatos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdUltimo 
         Height          =   325
         Left            =   345
         Picture         =   "frmABMMandatos.frx":0192
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   390
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   690
      TabIndex        =   10
      Top             =   1935
      Width           =   750
      Begin VB.CommandButton cmdAnterior 
         Height          =   325
         Left            =   345
         Picture         =   "frmABMMandatos.frx":0324
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrimero 
         Height          =   325
         Left            =   0
         Picture         =   "frmABMMandatos.frx":04B6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   30
      TabIndex        =   7
      Top             =   1035
      Width           =   6705
      Begin VB.TextBox txtMandato 
         Height          =   285
         Left            =   1470
         TabIndex        =   8
         Top             =   300
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mandato"
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1300
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   915
      Left            =   5460
      ScaleHeight     =   855
      ScaleWidth      =   1245
      TabIndex        =   5
      Top             =   60
      Width           =   1300
      Begin VB.CommandButton Salir 
         Caption         =   "&Salir"
         Height          =   855
         Left            =   0
         Picture         =   "frmABMMandatos.frx":0648
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   915
      Left            =   60
      ScaleHeight     =   855
      ScaleWidth      =   4980
      TabIndex        =   0
      Top             =   60
      Width           =   5040
      Begin VB.CommandButton Borrar 
         Caption         =   "Eliminar"
         Height          =   855
         Left            =   2490
         Picture         =   "frmABMMandatos.frx":074A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Grabar 
         Caption         =   "&Grabar"
         Height          =   855
         Left            =   1245
         Picture         =   "frmABMMandatos.frx":084C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Nuevo 
         Caption         =   "&Nuevo"
         Height          =   855
         Left            =   0
         Picture         =   "frmABMMandatos.frx":094E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Buscar 
         Caption         =   "&Buscar"
         Height          =   855
         Left            =   3735
         Picture         =   "frmABMMandatos.frx":0A50
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.Label lblid 
      Caption         =   "nothing"
      Height          =   255
      Left            =   5160
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblRecordSet 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0/0 Mandatos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   780
      TabIndex        =   16
      Top             =   1935
      Width           =   5190
   End
End
Attribute VB_Name = "frmABMMandatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blEsNuevo     As Boolean
Private WithEvents Rs As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1

Private Sub MostrarRegistro()

    On Error GoTo TrapError
    Dim xPos As Long
    Dim xMax As Long
    With Rs
        xPos = .AbsolutePosition ' Registro Actual
        xMax = .RecordCount      ' Registros totales
        If xMax > 0 Then         ' Si hay registros para mostrar
            txtMandato.Text = .Fields("fecha_mandato").Value
            lblRecordSet.Caption = Trim(Str(xPos)) & "/" & Trim(Str(xMax)) & " Mandatos"
        Else                     ' Si no hay registros para mostrar
            Call Limpiar
            lblRecordSet.Caption = "0/0 Mandatos"
        End If
    End With
Exit Sub
TrapError:
    Select Case Err.Number
    
        Case Else
            MsgBox "Error Nº " & Err.Number & Chr(10) & Err.Description & Chr(10) & "Originado en " & Err.Source
            Resume Next
    End Select
Return

End Sub

Private Sub Borrar_Click()
    On Error GoTo TrapError
    Dim xRespuesta As Integer
    If Rs.RecordCount > 0 Then
        xRespuesta = MsgBox("¿Seguro de eliminar el Mandato " & UCase(Trim(txtMandato.Text)) & " de la base de datos?", vbQuestion + vbYesNo)
        If xRespuesta = vbYes Then
            Rs.Delete
            Call Limpiar
            If Rs.RecordCount > 0 Then
                Rs.MoveFirst
                Call MostrarRegistro
            End If
        End If
    End If
Exit Sub
TrapError:
    Select Case Err.Number
    
        Case Else
            MsgBox "Error Nº " & Err.Number & Chr(10) & Err.Description & Chr(10) & "Originado en " & Err.Source
            Resume
    End Select
Return
End Sub
Private Sub Buscar_Click()
    Dim strId       As String
    Dim blCondicion As Boolean
    
    frmSelMandato.Show vbModal, Me
    strId = Trim(lblid.Caption)
    blCondicion = True
    If LCase(strId) <> "nothing" Then
        With Rs
            If .RecordCount > 0 Then
                .MoveFirst
                While blCondicion
                    If Trim(.Fields("fecha_mandato").Value) = Trim(strId) Then
                        Call MostrarRegistro
                        blCondicion = False
                    Else
                        .MoveNext
                    End If
                Wend
            End If
        End With
    End If
End Sub
Private Sub cmdAnterior_Click()
    If Rs.AbsolutePosition > 1 Then
        Rs.MovePrevious
    Else
        Rs.MoveFirst
    End If
    Call MostrarRegistro
End Sub
Private Sub cmdSiguiente_Click()
    If Rs.AbsolutePosition < Rs.RecordCount Then
        Rs.MoveNext
    Else
        Rs.MoveLast
    End If
    Call MostrarRegistro
End Sub
Private Sub cmdPrimero_Click()
    Rs.MoveFirst
    Call MostrarRegistro
End Sub

Private Sub cmdUltimo_Click()
    Rs.MoveLast
    Call MostrarRegistro
End Sub
Private Sub Form_Load()
    blEsNuevo = False
    Call SetearRs
    Call MostrarRegistro
End Sub
Private Sub Grabar_Click()
    On Error GoTo TrapError
    Dim strMandato As String
    ' ------------------------------------------------------------------------------------
    ' Validar entrada de usuario
    ' ------------------------------------------------------------------------------------
    ' Partido
    If Trim(txtMandato.Text) = "" Then
        MsgBox "El Mandato es un dato obligatorio", vbInformation + vbOKOnly
        txtMandato.SetFocus
        Exit Sub
    Else
        strMandato = txtMandato.Text
    End If
    ' ------------------------------------------------------------------------------------
    ' Grabar en base de datos
    ' ------------------------------------------------------------------------------------
    With Rs
        If blEsNuevo Then
            .AddNew
        End If
        .Fields("fecha_mandato").Value = strMandato
        .Update
        ' .Resync adAffectGroup
        If blEsNuevo Then
            .MoveLast
            blEsNuevo = False
        End If
        Call MostrarRegistro
    End With
    Call ActivarBotones(True)
Exit Sub
TrapError:
    Select Case Err.Number
    
        Case Else
            MsgBox "Error Nº " & Err.Number & Chr(10) & Err.Description & Chr(10) & "Originado en " & Err.Source
            Resume
    End Select
Return
End Sub

Private Sub Salir_Click()
    If Salir.Caption = "&Salir" Then
        Unload Me
    Else
        blEsNuevo = False
        Call ActivarBotones(True)
        Call MostrarRegistro
    End If
End Sub
Private Sub SetearRs()
    Dim strSql As String
    Set Rs = New ADODB.Recordset
    strSql = "SELECT fecha_mandato FROM mandatos ORDER BY fecha_mandato"
    Datos.SetearRsW strSql, Rs
    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
    End If
End Sub
Private Sub Limpiar()
    txtMandato.Text = ""
    
End Sub
Private Sub ActivarBotones(blEstado As Boolean)
    Nuevo.Enabled = blEstado
    Grabar.Enabled = True
    Borrar.Enabled = blEstado
    Buscar.Enabled = blEstado
    cmdAnterior.Enabled = blEstado
    cmdSiguiente.Enabled = blEstado
    cmdPrimero.Enabled = blEstado
    cmdUltimo.Enabled = blEstado
    lblRecordSet.Enabled = blEstado
    
    If blEstado Then
        Salir.Caption = "&Salir"
    Else
        Salir.Caption = "&Cancelar"
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Nuevo_Click()
    Call ActivarBotones(False)
    Call Limpiar
    blEsNuevo = True
    txtMandato.SetFocus
End Sub




