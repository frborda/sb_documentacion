VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmElegirPeriodo 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccionar Período Legislativo"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7305
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFiltro 
      BackColor       =   &H00404040&
      Caption         =   "Mostrar únicamente el período actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   3960
      Width           =   5415
   End
   Begin Proyecto1.ButtonOffice cmdCancelar 
      Height          =   795
      Left            =   3840
      TabIndex        =   2
      Top             =   4350
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1402
      BackColor       =   12230304
      Caption         =   "V&olver"
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
   Begin MSFlexGridLib.MSFlexGrid vsgrilla 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Proyecto1.ButtonOffice cmdSesiones 
      Height          =   795
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1402
      BackColor       =   12230304
      Caption         =   "&Ver sesiones"
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
      Caption         =   "Doble click sobre el período seleccionado para elegir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   5115
   End
End
Attribute VB_Name = "frmElegirPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstGrilla As New ADODB.Recordset
Private mActualizarDatos As Boolean
Private mFiltro As String
Private CodSes As String

Private Sub chkFiltro_Click()
If chkFiltro.Value = vbChecked Then
    mFiltro = " WHERE Período_Legislativo LIKE '" & Mid(CodSes, 1, 3) & "%' "
Else
    mFiltro = ""
End If
armarGrilla
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdSesiones_Click()
    If vsgrilla.Row > 0 Then
        buscarSesiones vsgrilla.TextMatrix(vsgrilla.Row, 1)
    Else
        MsgBox "Debe seleccionar un período legislativo para ver las sesiones relacionadas.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Form_Activate()
    If mActualizarDatos = False Then
        Label1.Caption = "Períodos legislativos registrados"
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    SetearRs "SELECT Período_Legislativo FROM vector", RsTemp
    CodSes = Trim(RsTemp.Fields(0))
    RsTemp.Close
    Set RsTemp = Nothing
    armarGrilla
    vsgrilla.ColWidth(6) = 0
    mFiltro = ""
    chkFiltro.Value = vbChecked
    chkFiltro_Click
End Sub

Private Sub armarGrilla()
    Dim RsTemp As ADODB.Recordset
    Dim strSql As String
    Dim j      As Long
    Dim blFind As Boolean
    Dim xFila  As Long
    Dim vColor As Variant
    
    Set RsTemp = New ADODB.Recordset
    blFind = False
    xFila = 1
    vsgrilla.Clear
    vsgrilla.Rows = 1
    ' Buscar todos los periodos que no tienen actas definidas
    strSql = "SELECT período_legislativo From perparl WHERE período_legislativo not in (SELECT DISTINCT Período_legislativo FROM actas) AND período_legislativo = '" & CodSes & "'"
    SetearRs strSql, RsTemp
    SetearRs "SELECT * FROM perparl Nro_de_Período_Legislativo " & mFiltro & " order by Nro_de_Período_Legislativo DESC, Tipo_de_período_sesión, Tipo_de_Sesión", rstGrilla
    Do While Not (rstGrilla.EOF)
        vsgrilla.AddItem vbTab & rstGrilla!Período_Legislativo & vbTab & rstGrilla!Nro_de_Período_Legislativo _
        & vbTab & rstGrilla!Tipo_de_período_sesión & vbTab & rstGrilla!Tipo_de_Sesión & vbTab & rstGrilla!Fecha_de_comienzo _
        & vbTab & rstGrilla!Nro_de_Sesion_actual
        rstGrilla.MoveNext
    Loop
    TitulosGRilla
    ' Recorrer la grilla..  si se trata de un periodo sin actas, pintarla de gris
    For xFila = 1 To vsgrilla.Rows - 1
        vsgrilla.Row = xFila
        With RsTemp
            If .RecordCount > 0 Then
                .MoveFirst
                blFind = False
                While Not .EOF
                    If Trim(LCase(vsgrilla.TextMatrix(xFila, 1))) = Trim(LCase(.Fields(0).Value)) Then
                        blFind = True
                        .MoveLast
                    End If
                    .MoveNext
                Wend
            End If
            If blFind Then
                blFind = False
                For j = 1 To vsgrilla.Cols - 1
                    vsgrilla.Col = j
                    'vsGrilla.CellBackColor = &H8000000F
                Next j
            End If
        End With
    Next xFila
    Set RsTemp = Nothing
End Sub
Private Sub TitulosGRilla()
    With vsgrilla
        .TextMatrix(0, 1) = "Período"
        .TextMatrix(0, 2) = "Número"
        .TextMatrix(0, 3) = "Tipo"
'        .TextMatrix(0, 4) = "Comienzo"
'        .TextMatrix(0, 5) = "Tipo sesión"
        .TextMatrix(0, 5) = "Comienzo"
        .TextMatrix(0, 4) = "Tipo sesión"
        .TextMatrix(0, 6) = "Sesión actual"
        .ColWidth(0) = 250
        .ColWidth(1) = 0
        .ColWidth(2) = 800
        .ColWidth(3) = 1500
'        .ColWidth(4) = 1300
'        .ColWidth(5) = 3000
        If .Rows < 13 Then
            .ColWidth(5) = 1500
        Else
            .ColWidth(5) = 1300
        End If
        .ColWidth(4) = 3000
        .ColWidth(6) = 1200
    End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If rstGrilla.State = adStateOpen Then
        rstGrilla.Close
    End If
    Set rstGrilla = Nothing
End Sub

Private Sub vsGrilla_DblClick()
    If mActualizarDatos = True Then
        If (gTipoUsuario <> 1) And (gTipoUsuario <> 4) Then
            If vsgrilla.Row > 0 Then
                MensajesSQV.cambioPeriodo vsgrilla.TextMatrix(vsgrilla.Row, 1)
                cmdCancelar_Click
            End If
        Else
            MsgBox "Ud. no dispone de permisos para realizar esta acción.", vbInformation + vbOKOnly
        End If
    Else
        buscarSesiones vsgrilla.TextMatrix(vsgrilla.Row, 1)
        'Dim f As New frmReuniones
        'f.periodoSesion = vsGrilla.TextMatrix(vsGrilla.Row, 1)
        'f.Show vbModal, Me
    End If
End Sub

Public Property Let ActualizarDatos(ByVal vNewValue As Boolean)
    mActualizarDatos = vNewValue
End Property

Private Sub vsGrilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        vsGrilla_DblClick
    End If
End Sub

Private Sub buscarSesiones(pPeriodo As String)
    Dim sesiones As New frmCambiarSesion
    sesiones.ActualizarDatos = False
    If sesiones.MostrarDatos(pPeriodo) = True Then
        sesiones.Show vbModal
    End If
    Set sesiones = Nothing
End Sub
