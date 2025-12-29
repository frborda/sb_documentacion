VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmCambiarSesion 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambiar la sesión"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5895
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdActualizarGrilla 
      Height          =   645
      Left            =   4800
      TabIndex        =   8
      Top             =   60
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1138
      BackColor       =   16744576
      Caption         =   "Actualizar &Grilla"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdActas 
      Height          =   765
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1349
      BackColor       =   16744576
      Caption         =   "&Aceptar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin VB.CheckBox chkPasLis 
      BackColor       =   &H00404040&
      Caption         =   "Pase de Lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2010
      TabIndex        =   5
      Top             =   450
      Width           =   2145
   End
   Begin VB.CheckBox chkVotNom 
      BackColor       =   &H00404040&
      Caption         =   "Votación Nominal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   450
      Width           =   1995
   End
   Begin VB.CheckBox chkVotNum 
      BackColor       =   &H00404040&
      Caption         =   "Votación Numérica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2010
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CheckBox chkActas 
      BackColor       =   &H00404040&
      Caption         =   "&Todas las actas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   1725
   End
   Begin MSFlexGridLib.MSFlexGrid dgSesion 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   1110
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   6
      SelectionMode   =   1
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
   Begin Proyecto1.ButtonOffice cmdCancelar 
      Height          =   765
      Left            =   3090
      TabIndex        =   7
      Top             =   3840
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1349
      BackColor       =   16744576
      Caption         =   "&Volver"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doble click sobre la sesión seleccionada para elegir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   750
      Width           =   4965
   End
End
Attribute VB_Name = "frmCambiarSesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstLista         As New ADODB.Recordset
Private mActualizarDatos As Boolean
Private mPeriodo         As String
Private strTipo_Filtro   As String

Public Property Let ActualizarDatos(ByVal vNewValue As Boolean)
    mActualizarDatos = vNewValue
End Property
Private Sub chkActas_Click()
    If chkActas.Value = 1 Then
        chkVotNum.Value = 0
        chkVotNom.Value = 0
        chkPasLis.Value = 0
    End If
End Sub
Private Sub chkPasLis_Click()
    If chkVotNum.Value = 1 And chkVotNom.Value = 1 And chkPasLis.Value = 1 Then
        chkVotNum.Value = 0
        chkVotNom.Value = 0
        chkPasLis.Value = 0
        chkActas.Value = 1
    Else
        chkActas.Value = 0
    End If
End Sub
Private Sub chkVotNom_Click()
    If chkVotNum.Value = 1 And chkVotNom.Value = 1 And chkPasLis.Value = 1 Then
        chkVotNum.Value = 0
        chkVotNom.Value = 0
        chkPasLis.Value = 0
        chkActas.Value = 1
    Else
        chkActas.Value = 0
    End If
End Sub
Private Sub chkVotNum_Click()
    If chkVotNum.Value = 1 And chkVotNom.Value = 1 And chkPasLis.Value = 1 Then
        chkVotNum.Value = 0
        chkVotNom.Value = 0
        chkPasLis.Value = 0
        chkActas.Value = 1
        Call CargarGrilla
    Else
        chkActas.Value = 0
    End If
End Sub
Private Sub cmdActas_Click()
    If dgSesion.Row > 0 Then
        mostrarActas dgSesion.TextMatrix(dgSesion.Row, 1)
    End If
End Sub


Private Sub cmdActualizarGrilla_Click()
        
    If chkVotNum.Value = 0 And chkVotNom.Value = 0 And chkPasLis.Value = 0 Then
        chkActas.Value = 1
        Exit Sub
    Else
        Call CargarGrilla
    End If
End Sub

Private Sub cmdActualizarGrilla2_Click()

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Function CargarGrilla() As Boolean
    Dim strSql       As String
    Dim strCondicion As String
    
    If chkVotNom.Value = 0 And chkVotNum.Value = 0 And chkPasLis.Value = 1 Then  ' Solo pase de lista
        strCondicion = "AND (actas.Tipo_de_operación = 'paslis') "
        strTipo_Filtro = "=paslis"
    ElseIf chkVotNom.Value = 1 And chkVotNum.Value = 0 And chkPasLis.Value = 0 Then ' Solo votacion nominal
        strCondicion = "AND (actas.Tipo_de_operación = 'votnom') "
        strTipo_Filtro = "=votnom"
    ElseIf chkVotNom.Value = 1 And chkVotNum.Value = 0 And chkPasLis.Value = 1 Then ' votacion nominal y pase de lista
        strCondicion = "AND (actas.Tipo_de_operación <> 'votnum') "
        strTipo_Filtro = "<>votnum"
    ElseIf chkVotNom.Value = 0 And chkVotNum.Value = 1 And chkPasLis.Value = 0 Then ' Solo votacion numerica
        strCondicion = "AND (actas.Tipo_de_operación = 'votnum') "
        strTipo_Filtro = "=votnum"
    ElseIf chkVotNom.Value = 0 And chkVotNum.Value = 1 And chkPasLis.Value = 1 Then ' Votacion Numerica y Pase de Lista
        strCondicion = "AND (actas.Tipo_de_operación <> 'votnom') "
        strTipo_Filtro = "<>votnom"
    ElseIf chkVotNom.Value = 1 And chkVotNum.Value = 1 And chkPasLis.Value = 0 Then  ' Votacion Numerica y Nominal
        strCondicion = "AND (actas.Tipo_de_operación <> 'paslis') "
        strTipo_Filtro = "<>paslis"
    End If
    
    If chkActas.Value = 1 Then
        strSql = "SELECT *, (SELECT Ultima_Reunion FROM perparl WHERE Período_Legislativo = '" & mPeriodo & "') AS ultimaReunion from sesion where ((rtrim(LOWER(sesion.Estado_sesión))='abierta')or rtrim(LOWER(sesion.Estado_sesión))='usoint' or (rtrim(LOWER(sesion.Estado_sesión))='nueva')) AND (Período_Legislativo='" & mPeriodo & "') order by Sesión desc "
        strTipo_Filtro = ""
    Else
        If Trim(strCondicion) = "" Then
            Exit Function
        End If
        strSql = "SELECT DISTINCT sesion.Sesión,(SELECT Ultima_Reunion FROM perparl WHERE Período_Legislativo = '" & mPeriodo & "') AS ultimaReunion, sesion.Fecha_de_inicio, sesion.Próximo_acta, sesion.Período_legislativo, actas.Tipo_de_operación " _
               & "FROM sesion INNER JOIN  actas ON sesion.Sesión = actas.Sesión AND sesion.Período_Legislativo = actas.Período_Legislativo " _
               & "WHERE (RTRIM(LOWER(dbo.sesion.Estado_sesión)) = 'abierta' OR " _
               & "RTRIM(LOWER(dbo.sesion.Estado_sesión)) = 'nueva' OR RTRIM(LOWER(dbo.sesion.Estado_sesión)) = 'usoint') AND (dbo.sesion.Período_Legislativo = '" & mPeriodo & "') " _
               & strCondicion _
               & "ORDER BY dbo.sesion.Sesión DESC"
    End If
    Datos.SetearRs strSql, rstLista
    With dgSesion
        .ColWidth(0) = 100
        .ColWidth(1) = 1000
        .ColWidth(2) = 2000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 0
        .TextMatrix(0, 1) = "Sesión"
        .TextMatrix(0, 2) = "Fecha"
        .TextMatrix(0, 3) = "Próximo acta"
        .TextMatrix(0, 4) = "Reunión"
        .TextMatrix(0, 5) = "PL"
        .Rows = 1
    End With
    If rstLista.EOF = False Then
        CargarGrilla = True
        Do While Not (rstLista.EOF)
            dgSesion.AddItem vbTab & rstLista.Fields("Sesión") & vbTab & Format(rstLista.Fields("Fecha_de_inicio"), "dd/mm/yyyy") _
            & vbTab & rstLista.Fields("Próximo_acta") & vbTab & rstLista.Fields("ultimaReunion") & vbTab & rstLista.Fields("Período_legislativo")
            rstLista.MoveNext
        Loop
        'If dgSesion.Rows > 2 Then
            'dgSesion.RemoveItem (1)
        'End If
    Else
        CargarGrilla = False
        MsgBox "No se han encontrado sesiones asociadas a este Período Legislativo.", vbInformation + vbOKOnly
    End If
End Function
Private Sub dgSesion_DblClick()
    If dgSesion.Row > 0 Then
        If mActualizarDatos = True Then
            If (gTipoUsuario <> 1) And (gTipoUsuario <> 4) Then
                MensajesSQV.cambiosesion dgSesion.TextMatrix(dgSesion.Row, 1)
                cmdCancelar_Click
            Else
                MsgBox "Ud. no dispone de permisos para realizar esta acción.", vbInformation + vbOKOnly
            End If
        Else
            mostrarActas dgSesion.TextMatrix(dgSesion.Row, 1)
        End If
    End If
End Sub
Private Sub mostrarActas(pSesion As Integer)
    Dim mTodas As Boolean
    If chkActas.Value = vbChecked Then
        mTodas = True
    Else
        mTodas = False
    End If
    Dim acta As New frmListarActas
    If acta.MostrarDatos(mPeriodo, pSesion, mTodas, strTipo_Filtro) = True Then
        acta.Show vbModal
    End If
    Set acta = Nothing
End Sub
Public Function MostrarDatos(pPeriodo As String) As Boolean
    mPeriodo = pPeriodo
    MostrarDatos = CargarGrilla
End Function
Private Sub dgSesion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dgSesion_DblClick
    End If
End Sub
Private Sub Form_Activate()
    If mActualizarDatos = True Then
        chkActas.Visible = False
    Else
        chkActas.Visible = True
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If rstLista.State = adStateOpen Then
        rstLista.Close
    End If
    Set rstLista = Nothing
End Sub
