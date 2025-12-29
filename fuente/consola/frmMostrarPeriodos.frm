VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMostrarPeriodos 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Períodos"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid gridView 
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   1
      Cols            =   4
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
      Left            =   900
      TabIndex        =   1
      Top             =   5580
      Width           =   2655
   End
End
Attribute VB_Name = "frmMostrarPeriodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo mError
Dim consulta As String
consulta = "SELECT     DISTINCT Nro_de_Período_Legislativo AS periodo, Tipo_de_período_sesión AS tipoPeriodo, CAST(DAY(Fecha_de_comienzo) AS varchar(2)) " & _
                      " + '/' + CAST(MONTH(Fecha_de_comienzo) AS varchar(2)) + '/' + CAST(YEAR(Fecha_de_comienzo) AS varchar(4)) AS fechaInicio, Fecha_de_comienzo " & _
" From perparl " & _
" ORDER BY periodo DESC, tipoPeriodo"
Dim rs As New Recordset
SetearRs consulta, rs
If (rs.EOF) Then
    MsgBox "No se encontraron los periodos. Error fatal."
    End
End If
gridView.Cols = 4
gridView.ColWidth(3) = 0
gridView.ColWidth(1) = 2000
gridView.TextMatrix(0, 0) = "Período"
gridView.TextMatrix(0, 1) = "Tipo"
gridView.TextMatrix(0, 2) = "Inicio"
Dim arr() As String
ReDim arr(0) As String
Dim find As Integer
find = 0
While Not rs.EOF
    Dim yaExiste As Boolean
    yaExiste = False
    For i = LBound(arr) To UBound(arr)
        If (arr(i) = (rs!periodo & vbTab & rs!tipOperiodo)) Then
            yaExiste = True
        End If
    Next i
    find = find + 1
    ReDim Preserve arr(find)
    arr(find) = rs!periodo & vbTab & rs!tipOperiodo
    If Not (yaExiste) Then
        gridView.AddItem rs!periodo & vbTab & rs!tipOperiodo & vbTab & rs!fechaInicio & vbTab & rs!periodo & Mid(rs!tipOperiodo, 1, 1) & "X"
    End If
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
Exit Sub
mError:
MsgBox err.Description
End Sub

Private Sub gridView_DblClick()
Dim f As New frmReuniones
f.periodoSesion = gridView.TextMatrix(gridView.Row, 3)
f.Show vbModal, Me
End Sub
