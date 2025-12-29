VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmTituloActa 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestión de órdenes del día"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10050
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdNUevo 
      Height          =   495
      Left            =   180
      TabIndex        =   3
      Top             =   6930
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   873
      BackColor       =   16744576
      Caption         =   "&Nuevo"
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
   Begin MSFlexGridLib.MSFlexGrid vsGrilla 
      Height          =   5355
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   9446
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtTitulo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   930
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5490
      Width           =   8865
   End
   Begin Proyecto1.ButtonOffice cmdModificar 
      Height          =   495
      Left            =   1350
      TabIndex        =   4
      Top             =   6930
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   873
      BackColor       =   16744576
      Caption         =   "&Modificar"
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
   Begin Proyecto1.ButtonOffice cmdEliminar 
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   6930
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   873
      BackColor       =   16744576
      Caption         =   "&Eliminar"
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
   Begin Proyecto1.ButtonOffice cmdEliminarTodos 
      Height          =   495
      Left            =   3690
      TabIndex        =   6
      Top             =   6930
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   873
      BackColor       =   16744576
      Caption         =   "Eliminar &Todos"
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
   Begin Proyecto1.ButtonOffice cmdSeleccionar 
      Height          =   495
      Left            =   4860
      TabIndex        =   7
      Top             =   6930
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BackColor       =   16744576
      Caption         =   "Se&leccionar"
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
   Begin Proyecto1.ButtonOffice cmdAceptar 
      Height          =   495
      Left            =   6090
      TabIndex        =   8
      Top             =   6930
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BackColor       =   16744576
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
   Begin Proyecto1.ButtonOffice cmdCancelar 
      Height          =   495
      Left            =   7320
      TabIndex        =   9
      Top             =   6930
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BackColor       =   16744576
      Caption         =   "&Cancelar"
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
   Begin Proyecto1.ButtonOffice cmdSAlir 
      Height          =   495
      Left            =   8550
      TabIndex        =   10
      Top             =   6930
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BackColor       =   16744576
      Caption         =   "&Salir"
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Título"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   5490
      Width           =   1275
   End
End
Attribute VB_Name = "frmTituloActa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstTitulo As New ADODB.Recordset
Private rstMaxTitulo As New ADODB.Recordset
Private rstPeriodo As New ADODB.Recordset
Private strSql As String
Private strSqlMax As String
Private mPeriodoLegislativo As String
Private mIdSesion As String
Private mCodigoActual As Long

Private Sub limpiarControles()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
         Select Case TypeName(ctrl)
            Case "TextBox"
                ctrl.Text = ""
            Case "DTPicker"
                ctrl.Value = Date
            Case "DataCombo"
                ctrl.BoundText = ""
        End Select
    Next
    mCodigoActual = -1
End Sub

Private Property Let ControlesHabilitados(ByVal pModo As Variant)
    Dim ctrl As Control
    For Each ctrl In Me.Controls
         Select Case TypeName(ctrl)
            Case "TextBox", "DTPicker", "DataCombo", "UpDown"
                ctrl.Enabled = pModo
        End Select
    Next
    If pModo = True Then
        cmdAceptar.Enabled = True
        cmdCancelar.Enabled = True
        cmdSAlir.Enabled = False
        cmdNUevo.Enabled = False
    Else
        cmdAceptar.Enabled = False
        cmdCancelar.Enabled = False
        cmdSAlir.Enabled = True
        cmdNUevo.Enabled = True
    End If
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdEliminarTodos.Enabled = False
    cmdSeleccionar.Enabled = False
    
End Property

Private Sub cmdAceptar_Click()
    Dim strSentencia As String
    If MsgBox("Está Ud. seguro de registrar las modificaciones realizadas?", vbQuestion + vbYesNo, "Confirma la operación?") = vbYes Then
        If validarDatos = True Then
            On Error GoTo ErrorDatos
            Datos.IniciarTransaccion
            verificarOrden 'compruebo el orden por las dudas
            If mCodigoActual = -1 Then
                'nuevo
                strSentencia = "INSERT INTO TitulosActas (Titulo) " _
                   & " VALUES ('" & txtTitulo.Text & "')"
            Else
                'update
                strSentencia = "UPDATE TitulosActas SET Titulo='" & txtTitulo.Text & "' WHERE id=" & mCodigoActual
            End If
            Datos.SenteciaSQl strSentencia
            Datos.FinalizarTransaccion True
            CargarGrilla
            cmdCancelar_Click
            
        End If
    End If
Exit Sub
ErrorDatos:
    Datos.FinalizarTransaccion False
End Sub

Private Sub verificarOrden()
    'If Val(txtORden.Text) <> upDown.Max Then
    '    Datos.SenteciaSQl "UPDATE TitulosActas set orden=orden+1 WHERE (Periodo_Legislativo='" & mPeriodoLegislativo & "') AND (orden>=" & txtORden.Text & ")"
    'End If
End Sub
Private Function validarDatos() As Boolean
    'If txtTitulo.Text = "" Then
    '    txtTitulo.Text = txtExtracto.Text
    'End If
    validarDatos = True
End Function
Private Sub cmdCancelar_Click()
    ControlesHabilitados = False
    limpiarControles
End Sub

Private Sub TitulosGRilla()
    With vsGrilla
        .Cols = 3
        .TextMatrix(0, 1) = "Titulo"
        .ColWidth(0) = 1
        .ColWidth(1) = vsGrilla.Width
        .ColWidth(2) = 1
    End With
End Sub
Private Sub CargarGrilla()
    'strSql = "SELECT Periodo_Legislativo, Fecha, Tipo, CodigoSesion, Titulo, Destino, " _
        & " Origen, Extracto, Orden, id, ordenDiaN FROM TitulosActas WHERE (Periodo_Legislativo='" & mPeriodoLegislativo & "') AND (CodigoSesion=" & mIdSesion & ") ORDER BY Fecha, orden"
    strSql = "SELECT id,Titulo FROM TitulosActas ORDER BY id ASC"
    'strSqlMax = "SELECT max(orden) as ord FROM TitulosActas WHERE (Periodo_Legislativo='" & mPeriodoLegislativo & "') AND (CodigoSesion=" & mIdSesion & ")"
     strSqlMax = "SELECT max(orden) as ord FROM TitulosActas "
    
    SetearRs strSql, rstTitulo
    SetearRs strSqlMax, rstMaxTitulo
    vsGrilla.Rows = 1
    'cargo datos en la grilla
    Do While Not (rstTitulo.EOF)
        vsGrilla.AddItem vbTab & rstTitulo!Titulo & vbTab & rstTitulo!id
        rstTitulo.MoveNext
    Loop
    'cierro rs
    If rstTitulo.State = adStateOpen Then
       rstTitulo.Close
    End If
    If rstMaxTitulo.State = adStateOpen Then
       rstMaxTitulo.Close
    End If
    TitulosGRilla
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Está ud seguro de eliminar el Título seleccionado?", vbQuestion + vbYesNo, "Confirma la operación?") = vbYes Then
        Eliminar (mCodigoActual)
        limpiarControles
        MsgBox "El registro se ha eliminado con éxito.", vbInformation + vbOKOnly
        ControlesHabilitados = False
        CargarGrilla
    End If
End Sub

Private Sub Eliminar(pCodigo As Long)
    If pCodigo < 0 Then
        Datos.SenteciaSQl ("TRUNCATE TABLE TitulosActas")
    Else
        Datos.SenteciaSQl ("DELETE FROM TitulosActas WHERE id=" & pCodigo)
    End If
End Sub

Private Sub cmdEliminarTodos_Click()
If MsgBox("Está ud seguro de todos los Títulos?", vbQuestion + vbYesNo, "Confirma la operación?") = vbYes Then
        Eliminar (-2)
        limpiarControles
        MsgBox "Todos los registros han sido eliminados con éxito.", vbInformation + vbOKOnly
        ControlesHabilitados = False
        CargarGrilla
    End If
End Sub

Private Sub cmdModificar_Click()
    ControlesHabilitados = True
End Sub

Private Sub cmdNuevo_Click()
    ControlesHabilitados = True
    limpiarControles
    'txtORden.Text = upDown.Max
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
    If mCodigoActual <> -1 Then
        MensajesSQV.CambioTituloActa mCodigoActual, vsGrilla.TextMatrix(vsGrilla.Row, 1)
        frmConsolaOperacion.txtTituloTemp.Text = vsGrilla.TextMatrix(vsGrilla.Row, 1)
        cmdSalir_Click
    Else
        MsgBox "Debe seleccionar un título para continuar.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Form_Activate()
vsGrilla.SelectionMode = flexSelectionByRow
vsGrilla.FocusRect = flexFocusHeavy
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCaracter(KeyAscii)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rstTitulo.State = adStateOpen Then
        rstTitulo.Close
    End If
    Set rstTitulo = Nothing
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = validarNumero(KeyAscii)
End Sub

Public Sub MostrarDatos(pPeriodo As String, Optional pSesion As String)
    mPeriodoLegislativo = pPeriodo
    mIdSesion = pSesion
    CargarGrilla
    limpiarControles
    ControlesHabilitados = False
End Sub

Private Sub Label7_Click()

End Sub

Private Sub txtORden_KeyPress(KeyAscii As Integer)
    KeyAscii = validarNumero(KeyAscii)
End Sub

Private Sub txtORden_Validate(Cancel As Boolean)
    'If Val(txtORden.Text) > UpDown.Max Then
    '    MsgBox "El máximo número de orden permitido es " & UpDown.Max, vbInformation + vbOKOnly
    '    txtORden.Text = UpDown.Max
    'End If
End Sub

Private Sub txtOrdenDia_KeyPress(KeyAscii As Integer)
    KeyAscii = Funciones.validarNumero(KeyAscii)
End Sub

Private Sub vsGrilla_Click()
    Dim Row As Integer
    Dim X As Integer
    If (vsGrilla.Row > 0) Then
        Row = vsGrilla.Row
        If Row > 0 Then
            With vsGrilla
                txtTitulo.Text = .TextMatrix(Row, 1)
                mCodigoActual = Val(.TextMatrix(Row, 2))
                ControlesHabilitados = False
                cmdModificar.Enabled = True
                cmdEliminar.Enabled = True
                cmdEliminarTodos.Enabled = True
                cmdSeleccionar.Enabled = True
            End With
        End If
    End If
    vsGrilla_KeyPress 0
End Sub

Private Sub vsGrilla_DblClick()
    cmdSeleccionar_Click
End Sub

Private Sub vsGrilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        vsGrilla_DblClick
    End If
End Sub
