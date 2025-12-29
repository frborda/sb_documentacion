VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAltaPeriodo 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administración de Períodos legislativos"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6240
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbSesion 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3960
      Width           =   2175
   End
   Begin VB.ComboBox cmbPeriodo 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CheckBox chkHistorico 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404040&
      Caption         =   "Histórico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   780
      TabIndex        =   6
      Top             =   4740
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtActual 
      Height          =   315
      Left            =   2280
      MaxLength       =   200
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid vsGrilla 
      Height          =   2355
      Left            =   120
      TabIndex        =   17
      Top             =   60
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   4154
      _Version        =   393216
      Rows            =   1
      Cols            =   11
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   5220
      Width           =   855
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   1260
      TabIndex        =   8
      Top             =   5220
      Width           =   855
   End
   Begin VB.CommandButton cmdSAlir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   5220
      Width           =   855
   End
   Begin VB.CommandButton cmdNUevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5220
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   5220
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtFecha 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   3600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Format          =   50331649
      CurrentDate     =   37985
   End
   Begin VB.TextBox txtNumero 
      Height          =   315
      Left            =   2280
      MaxLength       =   200
      TabIndex        =   1
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtPeriodo 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      MaxLength       =   200
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4020
      TabIndex        =   11
      Top             =   5220
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      Caption         =   "Nº sesión actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   780
      TabIndex        =   19
      Top             =   4380
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "Tipo de sesión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   780
      TabIndex        =   18
      Top             =   4020
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Tipo período"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   780
      TabIndex        =   16
      Top             =   3300
      Width           =   1275
   End
   Begin VB.Label Label9 
      BackColor       =   &H00404040&
      Caption         =   "Fecha de comienzo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   780
      TabIndex        =   15
      Top             =   3660
      Width           =   1395
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   780
      TabIndex        =   14
      Top             =   2940
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   780
      TabIndex        =   13
      Top             =   2580
      Width           =   1275
   End
End
Attribute VB_Name = "frmAltaPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstTitulo As New ADODB.Recordset
Private strSql As String
Private mCodigoActual As String
Private Const CodigoVacio As String = "N_U_L_O"

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
            Case "ComboBox"
                ctrl.ListIndex = -1
        End Select
    Next
    mCodigoActual = CodigoVacio
End Sub

Private Property Let ControlesHabilitados(ByVal pModo As Variant)
    Dim ctrl As Control
    For Each ctrl In Me.Controls
         Select Case TypeName(ctrl)
            Case "TextBox", "DTPicker", "DataCombo", "ComboBox"
                If ctrl.Name <> "txtPeriodo" Then
                    ctrl.Enabled = pModo
                End If
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
    cmbSesion.Enabled = False
End Property

Private Sub cmdAceptar_Click()
    Dim strSentencia As String
    Dim fecha As String
        
    If MsgBox("Está Ud. seguro de registrar las modificaciones realizadas?", vbQuestion + vbYesNo, "Confirma la operación?") = vbYes Then
        If validarDatos = True Then
            On Error GoTo ErrorDatos
            Datos.IniciarTransaccion
            If IsDate(dtFecha.Value) Then
                fecha = dtFecha.Value
            Else
                fecha = "01/01/1900"
            End If
        
            If mCodigoActual = CodigoVacio Then
                'nuevo
                Dim rs As New ADODB.Recordset
                SetearRs "SELECT id, leyenda_para_actas FROM tipo_sesion", rs
                Dim strTp As String
                Select Case cmbPeriodo.Text
                    Case Is = "Ordinario"
                        strTp = "O"
                    Case Is = "Extraordinario"
                        strTp = "E"
                    Case Is = "Legislativo"
                        strTp = "L"
                    Case Is = "Prórroga"
                        strTp = "P"
                End Select
                If Not rs.EOF Then
                    While Not rs.EOF
                        strSentencia = "INSERT INTO perparl (Período_Legislativo,Nro_de_Período_Legislativo,Tipo_de_período_sesión,Fecha_de_comienzo,Tipo_de_Sesión,Nro_de_Sesion_actual,Histórico,Ultima_Reunion) " _
                           & " VALUES ('" & txtNumero.Text & LCase(strTp) & LCase(rs!id) & "'," & txtNumero.Text & ",'" & cmbPeriodo.Text & "','" & fecha & "','" & rs!leyenda_para_actas & "','" & txtActual.Text & "'," & chkHistorico.Value & ",1)"
                        Datos.SenteciaSQl strSentencia
                        rs.MoveNext
                    Wend
                Else
                    MsgBox "Faltan cargar los tipos de sesion en la base de datos"
                End If
'                strSentencia = "INSERT INTO perparl (Período_Legislativo,Nro_de_Período_Legislativo,Tipo_de_período_sesión,Fecha_de_comienzo,Tipo_de_Sesión,Nro_de_Sesion_actual,Histórico,Ultima_Reunion) " _
'                   & " VALUES ('" & txtPeriodo.Text & "'," & txtNumero.Text & ",'" & cmbPeriodo.Text & "','" & fecha & "','" & cmbSesion.Text & "','" & txtActual.Text & "'," & chkHistorico.Value & ",3)"
            Else
                'update
                    
                strSentencia = "UPDATE perparl  SET Período_Legislativo='" & txtPeriodo.Text _
                & "', Nro_de_Período_Legislativo= " & txtNumero.Text _
                & ",Tipo_de_período_sesión='" & cmbPeriodo.Text _
                & "',Fecha_de_comienzo='" & fecha _
                & "',Tipo_de_Sesión='" & cmbSesion.Text _
                & "',Nro_de_Sesion_actual=" & txtActual.Text _
                & ",Histórico=" & chkHistorico.Value _
                & " WHERE Período_Legislativo='" & txtPeriodo.Text & "'"
                Datos.SenteciaSQl strSentencia
            End If
            Datos.FinalizarTransaccion True
            CargarGrilla
            cmdCancelar_Click
        End If
    End If
Exit Sub
ErrorDatos:
    Datos.FinalizarTransaccion False
End Sub
Private Function validarDatos() As Boolean
    validarDatos = True
    If cmbPeriodo.ListIndex = -1 Then
        cmbPeriodo.ListIndex = 0
        'cmbPeriodo.Text = "Ordinario"
    End If
    If (txtNumero.Text = "") Or (cmbPeriodo.ListIndex = -1) Or (cmbSesion.Locked = -1) Or _
        (txtActual.Text = "") Then
        MsgBox "Debe completar todos los campos antes de continuar.", vbInformation + vbOKOnly
        validarDatos = False
        Exit Function
    End If
    Dim strTp As String
    Dim strTs As String
    'cadena periodo
    Select Case cmbPeriodo.Text
        Case Is = "Ordinario"
            strTp = "O"
        Case Is = "Extraordinario"
            strTp = "E"
        Case Is = "Legislativo"
            strTp = "L"
        Case Is = "Prórroga"
            strTp = "P"
        Case Else
            validarDatos = False
            MsgBox "Seleccione un tipo de período válido de la lista.", vbInformation + vbOKOnly
            Exit Function
    End Select
    'cadena tipo sesion
'    Select Case cmbSesion.Text 'Tablas, Especial, Ordinaria, Extraordinaria, Asamblea Legislativa
'        Case Is = "Tablas" ' No disponible en SBA2009
'            strTs = "T"
'        Case Is = "Especial"
'            strTs = "E"
'        Case Is = "Preparatoria"
'            strTs = "P"
'        Case Is = "Informativa"
'            strTs = "I"
'        Case Is = "Homenaje"
'            strTs = "H"
'        Case Else
'            validarDatos = False
'            MsgBox "Seleccione un tipo de sesión válido de la lista.", vbInformation + vbOKOnly
'            Exit Function
'    End Select
    txtPeriodo.Text = LCase(Trim(txtNumero.Text) & strTp & strTs)
End Function
Private Sub cmdCancelar_Click()
    ControlesHabilitados = False
    limpiarControles
End Sub

Private Sub TitulosGRilla()
    With vsGrilla
        .Cols = 8
        .TextMatrix(0, 1) = "Periodo"
        .TextMatrix(0, 2) = "Número"
        .TextMatrix(0, 3) = "Tipo período"
        .TextMatrix(0, 5) = "Fecha comienzo"
        .TextMatrix(0, 4) = "Tipo sesión"
        .TextMatrix(0, 6) = "Próx. Sesión"
        .TextMatrix(0, 7) = "Histórico"
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 1300
        .ColWidth(3) = 1300
        .ColWidth(5) = 1800
        .ColWidth(4) = 1300
        .ColWidth(6) = 0
        .ColWidth(7) = 0
    End With
End Sub
Private Sub CargarGrilla()
    strSql = "SELECT Período_Legislativo, Nro_de_Período_Legislativo, Tipo_de_período_sesión, Fecha_de_comienzo, Tipo_de_Sesión, Nro_de_Sesion_actual, Histórico " _
     & " FROM perparl " _
     & " ORDER BY Período_Legislativo DESC"
    
    SetearRs strSql, rstTitulo
    vsGrilla.Rows = 1
    'cargo datos en la grilla
    Do While Not (rstTitulo.EOF)
        vsGrilla.AddItem vbTab & rstTitulo!Período_Legislativo & vbTab & rstTitulo!Nro_de_Período_Legislativo & vbTab & rstTitulo!Tipo_de_período_sesión & vbTab & rstTitulo!Tipo_de_Sesión & vbTab & rstTitulo!Fecha_de_comienzo & vbTab & rstTitulo!Nro_de_Sesion_actual & vbTab _
           & rstTitulo!Histórico
        rstTitulo.MoveNext
    Loop

    'cierro rs
    If rstTitulo.State = adStateOpen Then
       rstTitulo.Close
    End If
    TitulosGRilla
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Está ud seguro de eliminar el Período Legislativo seleccionado?", vbQuestion + vbYesNo, "Confirma la operación?") = vbYes Then
        Eliminar (mCodigoActual)
        limpiarControles
        ControlesHabilitados = False
        CargarGrilla
    End If
End Sub

Private Sub Eliminar(pCodigo As String)
    Dim rstVerificar As New ADODB.Recordset

    SetearRs "SELECT * FROM sesion WHERE rtrim(Período_Legislativo)='" & Trim(pCodigo) & "'", rstVerificar
    If rstVerificar.EOF = False Then
        MsgBox "No se puede eliminar el período legislativo seleccionado." & Chr(13) & "Este período posee sesiones asociadas.", vbInformation + vbOKOnly
    Else
        Datos.SenteciaSQl ("DELETE FROM perparl WHERE Período_Legislativo='" & pCodigo & "'")
        MsgBox "El registro se ha eliminado con éxito.", vbInformation + vbOKOnly
    End If
    rstVerificar.Close
    Set rstVerificar = Nothing
End Sub

Private Sub cmdModificar_Click()
    ControlesHabilitados = True
End Sub

Private Sub cmdNuevo_Click()
    ControlesHabilitados = True
    limpiarControles
    txtActual.Text = "1"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCaracter(KeyAscii)
End Sub

Private Sub Form_Load()
    'Cargo periodos legs en combo Ordinario y Extraordinario.
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    SetearRs "SELECT leyenda_para_formulario FROM tipo_periodo ORDER BY id", RsTemp
    While Not RsTemp.EOF
        cmbPeriodo.AddItem RsTemp.Fields(0)
        RsTemp.MoveNext
    Wend
    RsTemp.Close
    Set RsTemp = Nothing
    cmbPeriodo.Visible = True 'Se guardan como ordinarios todos los periodos y no se muestra la etiqueta en ninguna salida del sistema.
    Label5.Visible = True 'Se guardan como ordinarios todos los periodos y no se muestra la etiqueta en ninguna salida del sistema.
    'Cargo tipos de sesiones Especial, Ordinaria, Extraordinaria, Asamblea Legislativa, Preparatoria
    Set RsTemp = New ADODB.Recordset
    SetearRs "SELECT leyenda_para_formulario FROM tipo_sesion ORDER BY id", RsTemp
    While Not RsTemp.EOF
        cmbSesion.AddItem RsTemp.Fields(0)
        RsTemp.MoveNext
    Wend
    RsTemp.Close
    Set RsTemp = Nothing
    CargarGrilla
    cmdCancelar_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rstTitulo.State = adStateOpen Then
        rstTitulo.Close
    End If
    Set rstTitulo = Nothing
End Sub

Private Sub txtActual_KeyPress(KeyAscii As Integer)
    KeyAscii = validarNumero(KeyAscii)
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = validarNumero(KeyAscii)
End Sub

Private Sub vsGrilla_Click()
    Dim Row As Integer
    If (vsGrilla.Row > 0) Then
        Row = vsGrilla.Row
        With vsGrilla
            txtPeriodo.Text = .TextMatrix(Row, 1)
            txtNumero.Text = .TextMatrix(Row, 2)
            cmbPeriodo.ListIndex = Funciones.determinarListindex(.TextMatrix(Row, 3), cmbPeriodo)
            dtFecha.Value = .TextMatrix(Row, 5)
            cmbSesion.ListIndex = Funciones.determinarListindex(.TextMatrix(Row, 5), cmbSesion)
            txtActual.Text = .TextMatrix(Row, 6)
            If UCase(.TextMatrix(Row, 7)) = "VERDADERO" Then
                chkHistorico.Value = 1
            Else
                chkHistorico.Value = 0
            End If
            mCodigoActual = .TextMatrix(Row, 1)
            ControlesHabilitados = False
            cmdModificar.Enabled = False 'Cambiado HCDN 2011
            cmdEliminar.Enabled = False 'Cambiado HCDN 2011
        End With
    End If
End Sub
