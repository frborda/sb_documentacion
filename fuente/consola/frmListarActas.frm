VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmListarActas 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de actas registradas"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7665
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdImprRapida 
      Height          =   525
      Left            =   3570
      TabIndex        =   4
      Top             =   3090
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   926
      BackColor       =   12230304
      Caption         =   "Imprimir Acta"
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
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   6480
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
      Begin VB.CommandButton cmdImprimir 
         Height          =   550
         Left            =   0
         Picture         =   "frmListarActas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1030
      End
   End
   Begin MSFlexGridLib.MSFlexGrid dgSesion 
      Height          =   2655
      Left            =   90
      TabIndex        =   1
      Top             =   390
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   6
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
   Begin Proyecto1.ButtonOffice cmdCancelar 
      Height          =   525
      Left            =   5610
      TabIndex        =   5
      Top             =   3090
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   926
      BackColor       =   12230304
      Caption         =   "&Volver"
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
   Begin Proyecto1.ButtonOffice cmdModificarActa 
      Height          =   525
      Left            =   1530
      TabIndex        =   6
      Top             =   3090
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   926
      BackColor       =   33023
      Caption         =   "&Modificar Acta"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doble click sobre el acta seleccionada para ver su detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   5505
   End
End
Attribute VB_Name = "frmListarActas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstLista As New ADODB.Recordset
Private mPeriodo As String
Private mSesion  As Integer
Private mTodas   As Boolean
Private mFiltro  As String


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Function CargarGrilla() As Boolean
    Dim strSql     As String
    Dim strFiltro2 As String
    
    Select Case mFiltro
        Case "=paslis"
            strFiltro2 = " AND actas.Tipo_de_Operación = 'paslis' "
        Case "=votnum"
            strFiltro2 = " AND actas.Tipo_de_Operación = 'votnum' "
        Case "=votnom"
            strFiltro2 = " AND actas.Tipo_de_Operación = 'votnom' "
        Case "<>paslis"
            strFiltro2 = " AND actas.Tipo_de_Operación <> 'paslis' "
        Case "<>votnum"
            strFiltro2 = " AND actas.Tipo_de_Operación <> 'votnum' "
        Case "<>votnom"
            strFiltro2 = " AND actas.Tipo_de_Operación <> 'votnom' "
    End Select
    
    
    strSql = "SELECT tipoop.Tipo_de_operación, actas.Sesión, actas.Número_de_Acta, case when Ultima_Versión_Acta = 0 then 'Original' " _
        & " when Versión_Acta=0 then 'Ult. Mod. ver. ' + cast(Ultima_Versión_Acta + 1 as varchar) " _
        & " Else 'Ver. ' + cast(Ultima_Versión_Acta + 1 as varchar) end as Modificaciones, Versión_Acta as version " _
        & " FROM actas LEFT OUTER JOIN tipoop ON rtrim(actas.Tipo_de_operación) = rtrim(tipoop.identificador_en_mensajes) " _
        & " WHERE (actas.Período_Legislativo='" & mPeriodo & "') AND (Sesión=" & mSesion & ") "
    
    strSql = "SELECT tipoop.Tipo_de_operación, actas.Sesión, actas.Número_de_Acta, case when Ultima_Versión_Acta = 0 then 'Original' " _
        & " when Versión_Acta=0 then 'Ult. Mod. ver. ' + cast(Ultima_Versión_Acta + 1 as varchar) " _
        & " Else 'Ver. ' + cast(Ultima_Versión_Acta + 1 as varchar) end as Modificaciones, Versión_Acta as version " _
        & " FROM actas LEFT OUTER JOIN tipoop ON rtrim(actas.Tipo_de_operación) = rtrim(tipoop.identificador_en_mensajes) " _
        & " WHERE (actas.Período_Legislativo='" & mPeriodo & "') AND (Sesión=" & mSesion & ") AND Actas.Versión_Acta = 0 "
    If mTodas = False Then
        strSql = strSql & " AND (Versión_Acta=0) " & strFiltro2
    End If
    strSql = strSql & " ORDER BY sesión DESC, actas.Número_de_Acta desc"
    
    Datos.SetearRs strSql, rstLista
    With dgSesion
        .ColWidth(0) = 100
        .ColWidth(1) = 2000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1000
        .ColWidth(4) = 2700
        .ColWidth(5) = 0
        .TextMatrix(0, 1) = "Operación"
        .TextMatrix(0, 2) = "Sesión"
        .TextMatrix(0, 3) = "Nro. acta"
        .TextMatrix(0, 4) = "Modificaciones"
        .TextMatrix(0, 5) = "Version"
    End With
    If rstLista.EOF = False Then
        CargarGrilla = True
        Do While Not (rstLista.EOF)
            dgSesion.AddItem vbTab & Trim(rstLista!Tipo_de_operación) & vbTab & rstLista!Sesión _
            & vbTab & rstLista!Número_de_Acta & vbTab & rstLista.Fields!Modificaciones & vbTab & rstLista!Version
            rstLista.MoveNext
        Loop
        dgSesion.RemoveItem (1)
    Else
        CargarGrilla = False
        MsgBox "No se han encontrado actas asociadas a esta Sesión.", vbInformation + vbOKOnly, "Consola SQV"
    End If
End Function

Private Sub cmdImprimir_Click()
    Dim strLista As String
    Dim xS       As Long
    Dim X As Long
    With dgSesion
        If .Rows > 1 Then
            For X = 1 To .Rows - 1
                strLista = strLista & .TextMatrix(X, 3) & "; "
                xS = .TextMatrix(X, 2)
            Next X
            
        End If
    End With
    frmImprimirTotalActas.SesionActual = xS
    frmImprimirTotalActas.ListaActas = strLista
    frmImprimirTotalActas.Show vbModal
End Sub

Private Sub cmdImprRapida_Click()
Dim verActa As New frmConsultarActa

If PermisosTotales.ConsultaActas = 1 Then
    frmConsultarActa.MostrarDatos dgSesion.TextMatrix(dgSesion.Row, 3), mPeriodo, mSesion, dgSesion.TextMatrix(dgSesion.Row, 5)
    Tipo_PreActa = "consulta"
    frmConsultarActa.cmdReporte_Click
    Tipo_PreActa = ""
    Set verActa = Nothing
Else
    MsgBox "El usuario no tiene permisos para realizar esta operacion", vbInformation + vbOKOnly
End If
End Sub

Private Sub cmdModificarActa_Click()
cmdModificarActa.Enabled = False
If dgSesion.Row > 0 And EntroAConsola = False Then
    cmdImprRapida.Enabled = False
    cmdCancelar.Enabled = False
    DoEvents
    mostrarActa dgSesion.TextMatrix(dgSesion.Row, 3), mPeriodo, mSesion, dgSesion.TextMatrix(dgSesion.Row, 5)
    cmdImprRapida.Enabled = True
    cmdCancelar.Enabled = True
Else
    If dgSesion.Row <= 0 Then
        MsgBox "Seleccione un acta!", vbInformation
    End If
End If
cmdModificarActa.Enabled = True
End Sub
Private Sub dgSesion_DblClick()
    If dgSesion.Row > 0 And EntroAConsola = False Then
        cmdImprRapida.Enabled = False
        cmdCancelar.Enabled = False
        DoEvents
        mostrarActa dgSesion.TextMatrix(dgSesion.Row, 3), mPeriodo, mSesion, dgSesion.TextMatrix(dgSesion.Row, 5)
        cmdImprRapida.Enabled = True
        cmdCancelar.Enabled = True
    End If
End Sub
Private Sub mostrarActa(pActa As Integer, pPeriodo As String, pSesion As Integer, pVersion As Integer)
    Dim verActa As New frmConsultarActa
    If PermisosTotales.ConsultaActas = 1 Then
        verActa.MostrarDatos pActa, pPeriodo, pSesion, pVersion
        verActa.Show vbModal
        Set verActa = Nothing
    Else
        MsgBox "El usuario no tiene permisos para realizar esta operacion", vbInformation + vbOKOnly
    End If
End Sub
Public Function MostrarDatos(pPeriodo As String, pSesion As Integer, pTodas As Boolean, pFiltro As String) As Boolean
    mPeriodo = pPeriodo
    mSesion = pSesion
    mTodas = pTodas
    mFiltro = pFiltro
    MostrarDatos = CargarGrilla
End Function

Private Sub dgSesion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dgSesion_DblClick
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
