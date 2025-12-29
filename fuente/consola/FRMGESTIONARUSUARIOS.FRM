VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmGestionarUsuarios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administración de datos de usuarios del Sistema"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5460
      TabIndex        =   13
      Top             =   4200
      Width           =   915
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4500
      TabIndex        =   12
      Top             =   4200
      Width           =   915
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3540
      TabIndex        =   11
      Top             =   4200
      Width           =   915
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   1980
      TabIndex        =   10
      Top             =   4200
      Width           =   915
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   1020
      TabIndex        =   9
      Top             =   4200
      Width           =   915
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   60
      TabIndex        =   8
      Top             =   4200
      Width           =   915
   End
   Begin VB.ComboBox cmbNivelUsuario 
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   3660
      Width           =   2955
   End
   Begin VB.TextBox txtClave 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1620
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3240
      Width           =   2955
   End
   Begin VB.TextBox txtLogin 
      Height          =   315
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   1
      Top             =   2820
      Width           =   2955
   End
   Begin MSFlexGridLib.MSFlexGrid vsGrilla 
      Height          =   2355
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4154
      _Version        =   393216
      Rows            =   1
      Cols            =   11
      SelectionMode   =   1
   End
   Begin VB.Label Label4 
      Caption         =   "Nivel de usuario"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3660
      Width           =   1380
   End
   Begin VB.Label Label3 
      Caption         =   "Contraseña"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione un usuario de la lista para ver sus datos"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   4275
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre de usuario:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2820
      Width           =   1380
   End
End
Attribute VB_Name = "FrmGestionarUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mUsuarioActual As Integer
Public ContrasegnaValida As Boolean
Private mCambioPassword As Boolean

Private Sub cmdAceptar_Click()
    On Error GoTo ErrorDatos
    If validarControles = True Then
        Dim strSql As String
        If mUsuarioActual = -1 Then
            strSql = "INSERT INTO UsuarioConsola (Login, Clave, TipoUsuario) VALUES ('" & Trim(txtLogin.Text) & "','" & Trim(txtClave.Text) & "'," & cmbNivelUsuario.ListIndex & ")"
        Else
            strSql = "UPDATE UsuarioConsola SET Login='" & Trim(txtLogin.Text) & "', Clave='" & Trim(txtClave.Text) & "', TipoUsuario= " & cmbNivelUsuario.ListIndex & " WHERE id=" & mUsuarioActual
        End If
        Datos.SenteciaSQl strSql
        MsgBox "Los datos se han grabado de manera exitosa.", vbInformation + vbOKOnly
        cmdCancelar_Click
    End If
Exit Sub
ErrorDatos:
    MsgBox "Ha ocurrido un error al registrar los datos.", vbExclamation + vbOKOnly
End Sub
Private Function validarControles() As Boolean
    If (txtLogin.Text = "") Or (txtClave.Text = "") Or (cmbNivelUsuario.ListIndex = -1) Then
        validarControles = False
        MsgBox "Debe completar todos los campos antes de continuar.", vbInformation + vbOKOnly
    Else
        If mCambioPassword = True Then
            Dim valida As New frmValidarClave
            valida.ContrasegnaIngresada = txtClave.Text
            valida.Show vbModal
            If ContrasegnaValida = False Then
                validarControles = False
            Else
                validarControles = True
            End If
        Else
            validarControles = True
        End If
    End If
End Function
Private Sub cmdCancelar_Click()
    ControlesHabilitados = False
    BotonesEdicion = False
    limpiarCampos
    cargarGrillaUsuarios
End Sub

Private Sub cmdEliminar_Click()
    If mUsuarioActual <> -1 Then
        If MsgBox("Está Ud. seguro de eliminar el usuario seleccionado?", vbQuestion + vbYesNo, "COnfirma la operación?") = vbYes Then
            Datos.SenteciaSQl "DELETE FROM UsuarioConsola WHERE id=" & mUsuarioActual
            cmdCancelar_Click
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    ControlesHabilitados = True
    BotonesEdicion = True
End Sub

Private Sub cmdNuevo_Click()
    limpiarCampos
    ControlesHabilitados = True
    BotonesEdicion = True
End Sub
Private Sub limpiarCampos()
    mUsuarioActual = -1
    txtLogin.Text = ""
    txtClave.Text = ""
    cmbNivelUsuario.ListIndex = -1
    mCambioPassword = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCaracter(KeyAscii)
End Sub

Private Sub Form_Load()
    cargarComboNiveles
    cargarGrillaUsuarios
End Sub

Private Sub cargarComboNiveles()
    cmbNivelUsuario.AddItem "Administrador", 0
    cmbNivelUsuario.AddItem "Administrador bancas", 1
    cmbNivelUsuario.AddItem "Operador Avanzado", 2
    cmbNivelUsuario.AddItem "Operador Básico", 3
    cmbNivelUsuario.AddItem "Operador Consulta", 4
End Sub

Private Sub cargarGrillaUsuarios()
    Dim strSql As String
    Dim rstLista As New ADODB.Recordset
    strSql = "SELECT *, case TipoUsuario " _
    & " when 0 then 'Administrador' when 1 then 'Administrador bancas' when 2 then 'Operador Avanzado' when 3 then 'Operador Básico' when 4 then 'Operador Consulta' end as DescripcionTipoUsuario " _
    & " from UsuarioConsola ORDER by login"
    Datos.SetearRs strSql, rstLista
    TitulosGRilla
    vsGrilla.Rows = 1
    If rstLista.EOF = False Then
        Do While Not (rstLista.EOF)
            vsGrilla.AddItem vbTab & rstLista!Login & vbTab & rstLista!clave & _
            vbTab & rstLista!TipoUsuario & vbTab & rstLista!id & vbTab & rstLista!DescripcionTipoUsuario
            rstLista.MoveNext
        Loop
    Else
        MsgBox "No se han encontrado sesiones asociadas a este Período Legislativo.", vbInformation + vbOKOnly
    End If
    If rstLista.State = adStateOpen Then
        rstLista.Close
    End If
    Set rstLista = Nothing
End Sub

Private Sub TitulosGRilla()
    With vsGrilla
        .Cols = 6
        .ColWidth(0) = 100
        .ColWidth(1) = 3000
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 3000
        .TextMatrix(0, 1) = "Nombre de usuario"
        .TextMatrix(0, 2) = "Clave"
        .TextMatrix(0, 3) = "Codigo Tipo de usuario"
        .TextMatrix(0, 4) = "Id"
        .TextMatrix(0, 5) = "Tipo de usuario"
    End With
End Sub

Private Sub txtClave_Change()
    mCambioPassword = True
End Sub

Private Sub vsGrilla_Click()
    If vsGrilla.row > 0 Then
        txtLogin.Text = vsGrilla.TextMatrix(vsGrilla.row, 1)
        txtClave.Text = vsGrilla.TextMatrix(vsGrilla.row, 2)
        cmbNivelUsuario.ListIndex = Val(vsGrilla.TextMatrix(vsGrilla.row, 3))
        mUsuarioActual = Val(vsGrilla.TextMatrix(vsGrilla.row, 4))
        ControlesHabilitados = False
    End If
End Sub

Private Property Let ControlesHabilitados(ByVal vNewValue As Variant)
    txtLogin.Enabled = vNewValue
    txtClave.Enabled = vNewValue
    cmbNivelUsuario.Enabled = vNewValue
    vsGrilla.Enabled = Not (vNewValue)
End Property

Public Property Let BotonesEdicion(ByVal vNewValue As Variant)
    cmdModificar.Enabled = Not (vNewValue)
    cmdEliminar.Enabled = Not (vNewValue)
    cmdAceptar.Enabled = vNewValue
    cmdCancelar.Enabled = vNewValue
End Property
