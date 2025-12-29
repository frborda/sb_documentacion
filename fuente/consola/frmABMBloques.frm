VERSION 5.00
Begin VB.Form frmABMBloques 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ABM Bloques Políticos"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Height          =   390
      Left            =   6090
      ScaleHeight     =   330
      ScaleWidth      =   690
      TabIndex        =   21
      Top             =   2340
      Width           =   750
      Begin VB.CommandButton cmdSiguiente 
         Height          =   325
         Left            =   0
         Picture         =   "frmABMBloques.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdUltimo 
         Height          =   325
         Left            =   345
         Picture         =   "frmABMBloques.frx":0192
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   390
      Left            =   90
      ScaleHeight     =   330
      ScaleWidth      =   690
      TabIndex        =   19
      Top             =   2340
      Width           =   750
      Begin VB.CommandButton cmdAnterior 
         Height          =   325
         Left            =   345
         Picture         =   "frmABMBloques.frx":0324
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrimero 
         Height          =   325
         Left            =   0
         Picture         =   "frmABMBloques.frx":04B6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   14
      Top             =   1020
      Width           =   6705
      Begin VB.TextBox txtBancaMaxima 
         Height          =   285
         Left            =   5220
         TabIndex        =   7
         Top             =   690
         Width           =   825
      End
      Begin VB.TextBox txtBancaMinima 
         Height          =   285
         Left            =   1470
         TabIndex        =   6
         Top             =   690
         Width           =   825
      End
      Begin VB.TextBox txtClave 
         Height          =   285
         Left            =   5220
         TabIndex        =   5
         Top             =   270
         Width           =   825
      End
      Begin VB.TextBox txtBloquePolitico 
         Height          =   285
         Left            =   1470
         TabIndex        =   4
         Top             =   300
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Banca Máxima : "
         Height          =   285
         Left            =   3840
         TabIndex        =   18
         Top             =   780
         Width           =   1305
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Banca Mínima : "
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   750
         Width           =   1305
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Clave : "
         Height          =   285
         Left            =   3840
         TabIndex        =   16
         Top             =   330
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bloque Político :"
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1300
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   915
      Left            =   5550
      ScaleHeight     =   855
      ScaleWidth      =   1245
      TabIndex        =   13
      Top             =   50
      Width           =   1300
      Begin VB.CommandButton Salir 
         Caption         =   "&Salir"
         Height          =   855
         Left            =   0
         Picture         =   "frmABMBloques.frx":0648
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   915
      Left            =   150
      ScaleHeight     =   855
      ScaleWidth      =   4980
      TabIndex        =   0
      Top             =   50
      Width           =   5040
      Begin VB.CommandButton Buscar 
         Caption         =   "&Buscar"
         Height          =   855
         Left            =   3735
         Picture         =   "frmABMBloques.frx":074A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Borrar 
         Caption         =   "Eliminar"
         Height          =   855
         Left            =   2490
         Picture         =   "frmABMBloques.frx":084C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Grabar 
         Caption         =   "&Grabar"
         Height          =   855
         Left            =   1245
         Picture         =   "frmABMBloques.frx":094E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Nuevo 
         Caption         =   "&Nuevo"
         Height          =   855
         Left            =   0
         Picture         =   "frmABMBloques.frx":0A50
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.Label lblClave 
      Caption         =   "nothing"
      Height          =   375
      Left            =   7080
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblRecordSet 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0/0 Bloques Políticos"
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
      Left            =   870
      TabIndex        =   20
      Top             =   2340
      Width           =   5190
   End
End
Attribute VB_Name = "frmABMBloques"
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
            txtBloquePolitico.Text = NullCadena(.Fields("Bloque_Político").Value)
            txtClave.Text = NullCadena(.Fields("Clave").Value)
            txtBancaMinima.Text = NullCadena(.Fields("BancaMinima").Value)
            txtBancaMaxima.Text = NullCadena(.Fields("BancaMaxima").Value)
            lblRecordSet.Caption = Trim(Str(xPos)) & "/" & Trim(Str(xMax)) & " Bloques Políticos"
        Else                     ' Si no hay registros para mostrar
            Call Limpiar
            lblRecordSet.Caption = "0/0 Bloques Políticos"
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
        xRespuesta = MsgBox("¿Seguro de eliminar al bloque " & UCase(Trim(txtBloquePolitico.Text)) & " de la base de datos?", vbQuestion + vbYesNo)
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
    
    frmSelBloquePolitico.Show vbModal, Me
    strId = Trim(lblClave.Caption)
    blCondicion = True
    If LCase(strId) <> "nothing" Then
        With Rs
            If .RecordCount > 0 Then
                .MoveFirst
                While blCondicion
                    If .Fields("clave").Value = strId Then
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
    Dim strBloque    As String
    Dim strClave     As String
    Dim xBancaMinima As Long
    Dim xBancaMaxima As Long
    ' ------------------------------------------------------------------------------------
    ' Validar entrada de usuario
    ' ------------------------------------------------------------------------------------
    ' Bloque
    If Trim(txtBloquePolitico.Text) = "" Then
        MsgBox "El bloque político es un dato obligatorio", vbInformation + vbOKOnly
        txtBloquePolitico.SetFocus
        Exit Sub
    Else
        strBloque = txtBloquePolitico.Text
    End If
    ' Clave
    If Trim(txtClave.Text) = "" Then
        MsgBox "La clave del  bloque político es un dato obligatorio", vbInformation + vbOKOnly
        txtClave.SetFocus
        Exit Sub
    Else
        strClave = txtClave.Text
    End If
    ' Banca Minima
    If Trim(txtBancaMinima.Text) = "" Then
        xBancaMinima = 0
    Else
        If Not IsNumeric(txtBancaMinima.Text) Then
            MsgBox "En caso de registrar el valor Banca Minima, debe hacerlo con un numero entero positivo", vbInformation + vbOKOnly
            txtBancaMinima.SetFocus
            Exit Sub
        Else
            xBancaMinima = Int(txtBancaMinima.Text)
        End If
    End If
    ' Banca Maxima
    If Trim(txtBancaMaxima.Text) = "" Then
        xBancaMaxima = 0
    Else
        If Not IsNumeric(txtBancaMaxima.Text) Then
            MsgBox "En caso de registrar el valor Banca Máxima, debe hacerlo con un numero entero positivo", vbInformation + vbOKOnly
            txtBancaMaxima.SetFocus
            Exit Sub
        Else
            xBancaMaxima = Int(txtBancaMaxima.Text)
        End If
    End If
    ' ------------------------------------------------------------------------------------
    ' Grabar en base de datos
    ' ------------------------------------------------------------------------------------
    With Rs
        If blEsNuevo Then
            .AddNew
        End If
        .Fields("Bloque_Político").Value = strBloque
        .Fields("Clave").Value = strClave
        .Fields("BancaMinima").Value = xBancaMaxima
        .Fields("BancaMaxima").Value = xBancaMinima
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
    strSql = "SELECT Bloque_Político, Clave, BancaMinima, BancaMaxima FROM Bloques ORDER BY Bloque_Político"
    Datos.SetearRsW strSql, Rs
    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
    End If
End Sub
Private Sub Limpiar()
    txtBloquePolitico.Text = ""
    txtClave.Text = ""
    txtBancaMinima.Text = ""
    txtBancaMaxima.Text = ""
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
    txtBloquePolitico.SetFocus
End Sub



