VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCrearSesion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crear nueva sesión Sesión"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1740
      TabIndex        =   4
      Top             =   1500
      Width           =   1335
   End
   Begin VB.CheckBox chkProrroga 
      Alignment       =   1  'Right Justify
      Caption         =   "Prórroga"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1140
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1500
      Width           =   1335
   End
   Begin VB.TextBox txtProximo 
      Height          =   315
      Left            =   1740
      TabIndex        =   2
      Text            =   "1"
      Top             =   780
      Width           =   2955
   End
   Begin VB.TextBox txtSesion 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1740
      TabIndex        =   0
      Top             =   60
      Width           =   2955
   End
   Begin MSComCtl2.DTPicker dtFecha 
      Height          =   315
      Left            =   1740
      TabIndex        =   1
      Top             =   420
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Format          =   68878337
      CurrentDate     =   37985
   End
   Begin VB.Label Label9 
      Caption         =   "Fecha"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Próximo Nº de Acta"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Sesión"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "frmCrearSesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mPeriodoLegislativo
Private mNuevaSesion As Long

Private Sub cmdAceptar_Click()
Call Acepta
End Sub
Public Sub Acepta()
    Dim strSql As String
    If Not IsNumeric(txtProximo.Text) Then
        MsgBox "El próximo acta debe ser un número entero", vbInformation + vbOKOnly, "Error de usuario"
        txtProximo.SelStart = 0
        txtProximo.SelLength = Len(txtProximo.Text)
        txtProximo.SetFocus
        Exit Sub
    End If
    If Int(txtProximo.Text) <= 0 Then
        MsgBox "El próximo debe ser un valor mayor a cero", vbInformation + vbOKOnly, "Error de usuario"
        txtProximo.SelStart = 0
        txtProximo.SelLength = Len(txtProximo.Text)
        txtProximo.SetFocus
        Exit Sub
    End If
    
    If MsgBox("¿Está Ud. seguro de registrar la nueva sesión?", vbQuestion + vbYesNo, "Confirmar operación?") = vbYes Then
        If (gTipoUsuario <> 1) And (gTipoUsuario <> 4) Then
            If ExisteRegistro = False Then
                strSql = "INSERT INTO sesion (Período_Legislativo, Sesión,Fecha_de_inicio, Próximo_Acta, Estado_sesión, Prorroga) " _
                   & " VALUES ('" & mPeriodoLegislativo & "','" & txtSesion.Text & "','" & dtFecha.Value & "','" & txtProximo.Text & "','nueva'," & chkProrroga.Value & ")"
                Datos.SenteciaSQl strSql
                EjecutarSQL ("UPDATE perparl SET Nro_de_Sesion_actual = " & Val(txtSesion.Text) & " WHERE Período_Legislativo = '" & mPeriodoLegislativo & "'")
                MsgBox "Los datos se han registrado con éxito.", vbInformation + vbOKOnly
                MensajesSQV.cambiosesion txtSesion.Text
                cmdCancelar_Click
            Else
                MsgBox "Ya existe una sesión con este número en el período legislativo seleccionado.", vbExclamation + vbOKOnly
            End If
        Else
            MsgBox "Ud. no dispone de permisos para realizar esta acción.", vbInformation + vbOKOnly
        End If
    End If
End Sub
Private Function ExisteRegistro() As Boolean
    Dim rstExiste As New ADODB.Recordset
    SetearRs "SELECT * FROM sesion WHERE (Período_Legislativo='" & mPeriodoLegislativo & "') AND (Sesión='" & txtSesion.Text & "')", rstExiste
    If rstExiste.EOF = False Then
        ExisteRegistro = True
    Else
        ExisteRegistro = False
    End If
    If rstExiste.State = adStateOpen Then
        rstExiste.Close
    End If
    Set rstExiste = Nothing
End Function
Private Sub dtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
        Unload Me
End If
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dtFecha.Value = Date
End Sub

Private Sub txtProximo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
    KeyAscii = Funciones.validarNumero(KeyAscii)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtSesion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
    KeyAscii = Funciones.validarNumero(KeyAscii)
End Sub

Public Sub AgregarDatos(pPeriodo As String)
    mPeriodoLegislativo = pPeriodo
    mNuevaSesion = nuevoNumeroSesion
    txtSesion.Text = mNuevaSesion
End Sub
Private Function nuevoNumeroSesion() As Long
    Dim rstNumero As New ADODB.Recordset
    SetearRs "SELECT max(Sesión) as maximo FROM sesion WHERE (Sesión <> 9999)", rstNumero
    If rstNumero.EOF = False Then
        If IsNull(rstNumero!maximo) = False Then
            nuevoNumeroSesion = rstNumero!maximo + 1
        Else
            nuevoNumeroSesion = 1
        End If
    Else
        nuevoNumeroSesion = 1
    End If
    If rstNumero.State = adStateOpen Then
        rstNumero.Close
    End If
    Set rstNumero = Nothing
End Function
