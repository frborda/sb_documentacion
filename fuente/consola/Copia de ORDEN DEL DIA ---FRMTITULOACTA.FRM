VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTituloActa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestión de órdenes del día"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7995
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOrdenDia 
      Height          =   315
      Left            =   1500
      MaxLength       =   8
      TabIndex        =   5
      Top             =   3900
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid vsGrilla 
      Height          =   2355
      Left            =   120
      TabIndex        =   23
      Top             =   60
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   4154
      _Version        =   393216
      Rows            =   1
      Cols            =   11
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "Se&leccionar "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   22
      Top             =   6120
      Width           =   2235
   End
   Begin VB.TextBox txtOrigen 
      Height          =   315
      Left            =   1500
      MaxLength       =   200
      TabIndex        =   2
      Top             =   3180
      Width           =   6375
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   1020
      TabIndex        =   9
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdSAlir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7020
      TabIndex        =   13
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdNUevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5220
      TabIndex        =   11
      Top             =   6120
      Width           =   855
   End
   Begin MSComCtl2.UpDown upDown 
      Height          =   315
      Left            =   7635
      TabIndex        =   20
      Top             =   3540
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtORden"
      BuddyDispid     =   196618
      OrigLeft        =   7620
      OrigTop         =   3540
      OrigRight       =   7860
      OrigBottom      =   3855
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtFecha 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   3540
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   23527425
      CurrentDate     =   37985
   End
   Begin VB.TextBox txtTitulo 
      Height          =   855
      Left            =   1500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   5160
      Width           =   6375
   End
   Begin VB.TextBox txtORden 
      Height          =   315
      Left            =   6000
      TabIndex        =   4
      Top             =   3540
      Width           =   1635
   End
   Begin VB.TextBox txtExtracto 
      Height          =   855
      Left            =   1500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4260
      Width           =   6375
   End
   Begin VB.TextBox txtDestino 
      Height          =   315
      Left            =   1500
      MaxLength       =   200
      TabIndex        =   1
      Top             =   2820
      Width           =   6375
   End
   Begin VB.TextBox txtTipo 
      Height          =   315
      Left            =   1500
      MaxLength       =   200
      TabIndex        =   0
      Top             =   2460
      Width           =   6375
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Orden del día Nº"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   3960
      Width           =   1275
   End
   Begin VB.Label Label5 
      Caption         =   "Origen"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   1275
   End
   Begin VB.Label Label9 
      Caption         =   "Fecha"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Label Label8 
      Caption         =   "Título"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   5220
      Width           =   1275
   End
   Begin VB.Label Label7 
      Caption         =   "Orden en lista"
      Height          =   195
      Left            =   4620
      TabIndex        =   17
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Label Label6 
      Caption         =   "Extracto"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Destino"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   2520
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
        cmdSalir.Enabled = False
        cmdNUevo.Enabled = False
    Else
        cmdAceptar.Enabled = False
        cmdCancelar.Enabled = False
        cmdSalir.Enabled = True
        cmdNUevo.Enabled = True
    End If
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
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
                strSentencia = "INSERT INTO TitulosActas (OrigenDatos,Tipo,CodigoSesion,Destino,Fecha,Periodo_Legislativo,Origen,Extracto,Orden,Titulo, OrdenDiaN) " _
                   & " VALUES ('SQV','" & txtTipo.Text & "','" & mIdSesion & "','" & txtDestino.Text & "','" & dtFecha.Value & "','" _
                   & mPeriodoLegislativo & " ','" & txtOrigen.Text & "','" & txtExtracto.Text & "'," & txtORden.Text & ",'" & txtTitulo.Text & "'," & txtOrdenDia.Text & ")"
            Else
                'update
                strSentencia = "UPDATE TitulosActas SET Origen='" & txtOrigen.Text & "',Tipo='" & txtTipo.Text & "',Destino='" & txtDestino.Text & "'," _
                & "Fecha='" & dtFecha.Value & "',Periodo_Legislativo='" & mPeriodoLegislativo & "',Extracto='" & txtExtracto.Text & "',Orden=" & txtORden.Text & ",Titulo='" & txtTitulo.Text & "', ORdenDiaN=" & txtOrdenDia.Text _
                & " WHERE id=" & mCodigoActual
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
    If Val(txtORden.Text) <> upDown.Max Then
        Datos.SenteciaSQl "UPDATE TitulosActas set orden=orden+1 WHERE (Periodo_Legislativo='" & mPeriodoLegislativo & "') AND (orden>=" & txtORden.Text & ")"
    End If
End Sub
Private Function validarDatos() As Boolean
    If txtTitulo.Text = "" Then
        txtTitulo.Text = txtExtracto.Text
    End If
    If txtOrdenDia.Text = "" Then
        txtOrdenDia.Text = Val(txtOrdenDia.Text)
    End If
    validarDatos = True
End Function
Private Sub cmdCancelar_Click()
    ControlesHabilitados = False
    limpiarControles
End Sub

Private Sub TitulosGRilla()
    With vsGrilla
        .Cols = 12
        .TextMatrix(0, 1) = "Periodo"
        .TextMatrix(0, 2) = "Fecha"
        .TextMatrix(0, 3) = "Tipo"
        .TextMatrix(0, 4) = "Sesión"
        .TextMatrix(0, 5) = "Título"
        .ColWidth(0) = 250
        .ColWidth(1) = 800
        .ColWidth(2) = 1200
        .ColWidth(3) = 1500
        .ColWidth(4) = 900
        .ColWidth(5) = 2750
        .ColWidth(6) = 0 'destino
        .ColWidth(7) = 0 'origen
        .ColWidth(8) = 0 'extracto
        .ColWidth(9) = 0 'Orden
        .ColWidth(10) = 0  'id
        .ColWidth(11) = 0  'Orden dia Nº
    End With
End Sub
Private Sub CargarGrilla()
    strSql = "SELECT Periodo_Legislativo, Fecha, Tipo, CodigoSesion, Titulo, Destino, " _
        & " Origen, Extracto, Orden, id, ordenDiaN FROM TitulosActas WHERE (Periodo_Legislativo='" & mPeriodoLegislativo & "') AND (CodigoSesion=" & mIdSesion & ") ORDER BY Fecha, orden"
    strSqlMax = "SELECT max(orden) as ord FROM TitulosActas WHERE (Periodo_Legislativo='" & mPeriodoLegislativo & "') AND (CodigoSesion=" & mIdSesion & ")"
    
    SetearRs strSql, rstTitulo
    SetearRs strSqlMax, rstMaxTitulo
    vsGrilla.Rows = 1
    'cargo datos en la grilla
    Do While Not (rstTitulo.EOF)
        vsGrilla.AddItem vbTab & rstTitulo!periodo_Legislativo & vbTab & rstTitulo!Fecha & vbTab & rstTitulo!Tipo & vbTab & rstTitulo!CodigoSesion & vbTab & rstTitulo!titulo & vbTab & rstTitulo!Destino & vbTab _
           & rstTitulo!Origen & vbTab & rstTitulo!Extracto & vbTab & rstTitulo!orden & vbTab & rstTitulo!Id & vbTab & rstTitulo!ordenDiaN
        rstTitulo.MoveNext
    Loop
    If IsNull(rstMaxTitulo!ord) = False Then
        upDown.Max = Val(rstMaxTitulo!ord) + 1
    Else
        upDown.Max = 1
    End If
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
    Datos.SenteciaSQl ("DELETE FROM TitulosActas WHERE id=" & pCodigo)
End Sub

Private Sub cmdModificar_Click()
    ControlesHabilitados = True
End Sub

Private Sub cmdNuevo_Click()
    ControlesHabilitados = True
    limpiarControles
    txtORden.Text = upDown.Max
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
    If mCodigoActual <> -1 Then
        MensajesSQV.CambioTituloActa mCodigoActual, vsGrilla.TextMatrix(vsGrilla.row, 5)
        cmdSalir_Click
    Else
        MsgBox "Debe seleccionar un título para continuar.", vbInformation + vbOKOnly
    End If
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

Private Sub txtORden_KeyPress(KeyAscii As Integer)
    KeyAscii = validarNumero(KeyAscii)
End Sub

Private Sub txtORden_Validate(Cancel As Boolean)
    If Val(txtORden.Text) > upDown.Max Then
        MsgBox "El máximo número de orden permitido es " & upDown.Max, vbInformation + vbOKOnly
        txtORden.Text = upDown.Max
    End If
End Sub

Private Sub txtOrdenDia_KeyPress(KeyAscii As Integer)
    KeyAscii = Funciones.validarNumero(KeyAscii)
End Sub

Private Sub vsGrilla_Click()
    Dim row As Integer
    If (vsGrilla.row > 0) Then
        row = vsGrilla.row
        With vsGrilla
            txtTipo.Text = .TextMatrix(row, 3)
            txtDestino.Text = .TextMatrix(row, 6)
            dtFecha.Value = .TextMatrix(row, 2)
            txtORden.Text = .TextMatrix(row, 9)
            txtExtracto.Text = .TextMatrix(row, 8)
            txtTitulo.Text = .TextMatrix(row, 5)
            txtOrdenDia.Text = .TextMatrix(row, 11)
            mCodigoActual = Val(.TextMatrix(row, 10))
            ControlesHabilitados = False
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = True
            cmdSeleccionar.Enabled = True
        End With
    End If
End Sub

Private Sub vsGrilla_DblClick()
    cmdSeleccionar_Click
End Sub

Private Sub vsGrilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        vsGrilla_DblClick
    End If
End Sub
