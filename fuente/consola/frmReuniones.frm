VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReuniones 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reuniones"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid gridView 
      Height          =   7395
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   13044
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doble clic para seleccionar"
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
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   8160
      Width           =   2655
   End
   Begin VB.Label lblPeriodo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "XXXXX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   240
      Width           =   7275
   End
End
Attribute VB_Name = "frmReuniones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public periodoSesion As String
Private periodoCodigo As String
Private sesionCodigo As String
Private periodo As String
Private sesion As String
Private firstTime As Boolean

Private Sub Form_Activate()
If firstTime = True Then
    firstTime = False
    CargaGrilla
End If
End Sub

Private Sub Form_Load()
firstTime = True
If (periodoSesion = "") Then
    MsgBox "Esta ventana debe ser usada con un parametro."
    Unload Me
Else
    periodoCodigo = mId(periodoSesion, 4, 1)
    sesionCodigo = mId(periodoSesion, 5, 1)
    'Obtengo el tipo periodo
    Dim rsPeriodo As New ADODB.Recordset
    SetearRs "SELECT leyenda_para_actas FROM tipo_periodo WHERE id = '" & periodoCodigo & "'", rsPeriodo
    If (rsPeriodo.EOF) Then
        MsgBox "Ocurrio un error tratando de obtener el nombe del tipo de periodo"
        Unload Me
    Else
        periodo = rsPeriodo!leyenda_para_actas
    End If
    rsPeriodo.Close
    Set rsPeriodo = Nothing
    'Obtengo el tipo de sesion
    sesion = ""
End If
IniciaGrilla
End Sub

Private Function ObtenerTipoSesion(pTipoSesion As String) As String
Dim rsSesion As New ADODB.Recordset
Dim ret As String
SetearRs "SELECT leyenda_para_actas FROM tipo_sesion WHERE id = '" & pTipoSesion & "'", rsSesion
If (rsSesion.EOF) Then
    MsgBox "Ocurrio un error tratando de obtener el nombre del tipo de sesion"
    Unload Me
Else
    ret = rsSesion!leyenda_para_actas
End If
rsSesion.Close
Set rsSesion = Nothing
ObtenerTipoSesion = ret
End Function

Private Sub IniciaGrilla()
lblPeriodo.Caption = "Período " & Left(periodoSesion, 3) & " - " & periodo
gridView.ColWidth(5) = 0
gridView.ColWidth(2) = 1500
gridView.TextMatrix(0, 0) = "Reunion"
gridView.TextMatrix(0, 1) = "Sesion"
gridView.TextMatrix(0, 2) = "Tipo"
gridView.TextMatrix(0, 3) = "Fecha"
gridView.TextMatrix(0, 4) = "Actas"
End Sub

Private Sub CargaGrilla()
Dim rs As New Recordset
Dim consulta As String
consulta = "SELECT     Reunion, Sesión AS Sesion, Período_Legislativo AS perleg, " & _
" (SELECT     TOP (1) CAST(DAY(Fecha) AS varchar(2)) + '/' + CAST(MONTH(Fecha) AS varchar(2)) + '/' + CAST(YEAR(Fecha) " & _
"                      AS varchar(4)) AS fecha " & _
" FROM          actas AS a1 " & _
" WHERE      (Período_Legislativo LIKE '" & mId(periodoSesion, 1, 4) & "%') AND (Versión_Acta = 0) AND (Sesión = a2.Sesión) " & _
" ORDER BY Fecha DESC) AS fecha, COUNT(Número_de_Acta) AS Cantidad " & _
"FROM actas AS a2 " & _
"WHERE (Período_Legislativo LIKE '" & mId(periodoSesion, 1, 4) & "%') AND (Versión_Acta = 0) " & _
"GROUP BY Reunion, Período_Legislativo, Sesión " & _
"ORDER BY Reunion DESC"
SetearRs consulta, rs
If (rs.EOF) Then
    MsgBox "No se encontraron sesiones para este periodo"
    Unload Me
Else
    While Not rs.EOF
        gridView.AddItem rs!Reunion & vbTab & IIf(rs!sesion = -1, 0, rs!sesion) & vbTab & ObtenerTipoSesion(mId(rs!perleg, 5, 1)) & vbTab & rs!fecha & vbTab & rs!Cantidad & vbTab & rs!perleg
        rs.MoveNext
    Wend
End If
rs.Close
Set rs = Nothing
End Sub
Private Sub gridView_DblClick()
Dim nForm As New frmMostrarActas
nForm.periodo = gridView.TextMatrix(gridView.Row, 5)
nForm.sesion = IIf(gridView.TextMatrix(gridView.Row, 1) = 0, -1, gridView.TextMatrix(gridView.Row, 1))
nForm.Show vbModal
End Sub
