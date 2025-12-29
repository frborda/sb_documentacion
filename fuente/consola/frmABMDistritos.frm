VERSION 5.00
Begin VB.Form frmABMDistritos 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   390
      Left            =   6000
      ScaleHeight     =   330
      ScaleWidth      =   690
      TabIndex        =   13
      Top             =   2655
      Width           =   750
      Begin VB.CommandButton cmdSiguiente 
         Height          =   325
         Left            =   0
         Picture         =   "frmABMDistritos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdUltimo 
         Height          =   325
         Left            =   345
         Picture         =   "frmABMDistritos.frx":0192
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
      Top             =   2655
      Width           =   750
      Begin VB.CommandButton cmdAnterior 
         Height          =   325
         Left            =   345
         Picture         =   "frmABMDistritos.frx":0324
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrimero 
         Height          =   325
         Left            =   0
         Picture         =   "frmABMDistritos.frx":04B6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   30
      TabIndex        =   7
      Top             =   1035
      Width           =   6705
      Begin VB.ComboBox cmbSeccion 
         Height          =   315
         ItemData        =   "frmABMDistritos.frx":0648
         Left            =   1470
         List            =   "frmABMDistritos.frx":064F
         TabIndex        =   18
         Top             =   360
         Width           =   3500
      End
      Begin VB.TextBox txtDistrito 
         Height          =   285
         Left            =   1470
         TabIndex        =   8
         Top             =   900
         Width           =   5055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Distrito Electoral"
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   375
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Distrito Electoral"
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1305
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
         Picture         =   "frmABMDistritos.frx":0658
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
         Picture         =   "frmABMDistritos.frx":075A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Grabar 
         Caption         =   "&Grabar"
         Height          =   855
         Left            =   1245
         Picture         =   "frmABMDistritos.frx":085C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Nuevo 
         Caption         =   "&Nuevo"
         Height          =   855
         Left            =   0
         Picture         =   "frmABMDistritos.frx":095E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Buscar 
         Caption         =   "&Buscar"
         Height          =   855
         Left            =   3735
         Picture         =   "frmABMDistritos.frx":0A60
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Distrito Electoral"
      Height          =   285
      Left            =   240
      TabIndex        =   19
      Top             =   1560
      Width           =   1305
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
      Caption         =   "0/0 Distritos Electorales"
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
      Top             =   2655
      Width           =   5190
   End
End
Attribute VB_Name = "frmABMDistritos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blEsNuevo     As Boolean
Private WithEvents Rs As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1
Public arrSecciones As String
Private Function buscarIndexSeccion(strSeccion)
    Dim arrSec, sec
    
    Dim nPointer As Integer
    buscarIndexSeccion = 1
    arrSec = Split(arrSecciones, "||")
    'MsgBox (uboun(arrSec))
    For nPointer = 0 To UBound(arrSec) - 1
        sec = Split(arrSec(nPointer), "|")
        If sec(0) = strSeccion Then
            buscarIndexSeccion = sec(1)
            Exit Function
        End If
    Next
End Function

Private Sub MostrarRegistro()
    On Error GoTo TrapError
    Dim xPos As Long
    Dim xMax As Long
    
    With Rs
        xPos = .AbsolutePosition ' Registro Actual
        xMax = .RecordCount      ' Registros totales
        If xMax > 0 Then         ' Si hay registros para mostrar
            txtDistrito.Text = .Fields("distrito").Value
            If Len(Rs.Fields("seccionstr").Value) > 0 Then
                cmbSeccion.Text = .Fields("seccionstr").Value
                
                cmbSeccion.ListIndex = buscarIndexSeccion(Rs.Fields("seccionstr").Value) - 1
                
                
            Else
                cmbSeccion.ListIndex = 1
            End If
            
            lblRecordSet.Caption = Trim(Str(xPos)) & "/" & Trim(Str(xMax)) & " Distritos Electorales"
        Else                     ' Si no hay registros para mostrar
            Call Limpiar
            lblRecordSet.Caption = "0/0 Distritos Electorales"
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
    Dim strSql As String
    Dim Cn As ADODB.Connection
    Set Cn = New ADODB.Connection
    
    
    If Rs.RecordCount > 0 Then
        xRespuesta = MsgBox("¿Seguro de eliminar el Distrito " & UCase(Trim(txtDistrito.Text)) & " de la base de datos?", vbQuestion + vbYesNo)
        If xRespuesta = vbYes Then
            'Rs.Delete
            With Cn
            .ConnectionString = strconexion
            .CommandTimeout = 15
            .CursorLocation = adUseClient
            .Open
            End With
            Cn.BeginTrans
                strSql = "DELETE FROM distritos where id_distrito=" & Rs.Fields("id_distrito").Value
                Cn.Execute (strSql)
            Cn.CommitTrans
    
    
            Call Limpiar
            If Rs.RecordCount > 0 Then
                Rs.MoveFirst
                
                Call LlenarCombos
                Call SetearRs
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
    
    frmSelDistrito.Show vbModal, Me
    strId = Trim(lblid.Caption)
    blCondicion = True
    If LCase(strId) <> "nothing" Then
        With Rs
            If .RecordCount > 0 Then
                .MoveFirst
                While blCondicion
                    If Trim(.Fields("distrito").Value) = Trim(strId) Then
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
    Call LlenarCombos
    Call MostrarRegistro
End Sub
Private Sub cmdSiguiente_Click()
    If Rs.AbsolutePosition < Rs.RecordCount Then
        Rs.MoveNext
    Else
        Rs.MoveLast
    End If
    Call LlenarCombos
    Call MostrarRegistro
End Sub
Private Sub cmdPrimero_Click()
    Call SetearRs
    Rs.MoveFirst
    Call MostrarRegistro
End Sub

Private Sub cmdUltimo_Click()
    Call SetearRs
    Rs.MoveLast
    Call MostrarRegistro
End Sub
Private Sub Form_Load()
    blEsNuevo = False
    Call LlenarCombos
    Call SetearRs
    Call MostrarRegistro
End Sub

Private Sub Grabar_Click()
    On Error GoTo TrapError
    Dim strDistrito As String
    Dim nSeccion, n As Integer
    Dim Cn As ADODB.Connection
    Dim strSql As String
     Dim rsActual As Integer
    Set Cn = New ADODB.Connection
    
       
    ' ------------------------------------------------------------------------------------
    ' Validar entrada de usuario
    ' ------------------------------------------------------------------------------------
    ' Partido
    If Trim(txtDistrito.Text) = "" Then
        MsgBox "El Distrito Electoral es un dato obligatorio", vbInformation + vbOKOnly
        txtDistrito.SetFocus
        Exit Sub
    Else
        strDistrito = txtDistrito.Text
        nSeccion = cmbSeccion.ItemData(cmbSeccion.ListIndex)
        'MsgBox (cmbSeccion.ItemData(cmbSeccion.ListIndex))
    End If
    ' ------------------------------------------------------------------------------------
    ' Grabar en base de datos
    ' ------------------------------------------------------------------------------------
    
        If blEsNuevo Then
            With Rs
                .AddNew
                .Fields("distrito").Value = strDistrito
                .Fields("seccion").Value = nSeccion
                .Update
                .MoveLast
                blEsNuevo = False
            End With
        Else
            With Cn
                .ConnectionString = strconexion
                .CommandTimeout = 15
                .CursorLocation = adUseClient
                .Open
            End With
            Cn.BeginTrans
                strSql = "UPDATE distritos set distrito='" & strDistrito & "',seccion=" & nSeccion & "  where id_distrito=" & Rs.Fields("id_distrito").Value
                'MsgBox strSql
                Cn.Execute (strSql)
                
            Cn.CommitTrans
        End If
        
       
        
        
        '
        ' .Resync adAffectGroup
        With Rs
            rsActual = Rs.Bookmark - 1
            Call SetearRs
            Rs.Move (rsActual)

            Call MostrarRegistro
            'Rs.Bookmark = rsActual
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
    strSql = "SELECT d.id_distrito,d.distrito,d.seccion,s.seccion as seccionstr FROM distritos d INNER JOIN secciones s ON s.id_seccion = d.seccion ORDER BY distrito"
    Datos.SetearRsW strSql, Rs
    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
    End If
End Sub
Private Sub Limpiar()
    txtDistrito.Text = ""
    
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
    txtDistrito.SetFocus
End Sub

Private Sub LlenarCombos()
    Dim strSql As String
    Dim RsTemp As ADODB.Recordset
    Dim nIndex As Integer
    nIndex = 1
     
    Set RsTemp = New ADODB.Recordset
    
    ' -------------------------------------------------------
    ' Combo Seccion
    ' -------------------------------------------------------
    arrSecciones = ""
    strSql = "SELECT seccion,id_seccion FROM secciones ORDER BY id_seccion"
    Datos.SetearRsW strSql, RsTemp
    With cmbSeccion
        .Clear
        If RsTemp.RecordCount > 0 Then
            RsTemp.MoveFirst
            While Not RsTemp.EOF
                .AddItem RsTemp.Fields("seccion").Value '& " / " & RsTemp.Fields("id_seccion").Value
                .ItemData(cmbSeccion.NewIndex) = RsTemp.Fields("id_seccion").Value
                arrSecciones = arrSecciones & RsTemp.Fields("seccion").Value & "|" & nIndex & "||"
                RsTemp.MoveNext
                nIndex = nIndex + 1
            Wend
        End If
        RsTemp.Close
    End With
    
    Set RsTemp = Nothing
End Sub

