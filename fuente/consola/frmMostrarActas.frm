VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMostrarActas 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actas"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   12315
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid gridView 
      Height          =   4875
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   8599
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
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
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
      Left            =   60
      TabIndex        =   1
      Top             =   5040
      Width           =   12195
   End
End
Attribute VB_Name = "frmMostrarActas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
     Public periodo As String
Public sesion As String
Private sinActas As Boolean

Private Sub Form_Activate()
On Error GoTo errlab
If sinActas Then
    Dim RTA As Integer
    RTA = MsgBox("No hay actas para mostrar. ¿Desea imprimir una hoja de prueba?", vbYesNo)
    If RTA = vbYes Then
        Dim rp As New rptPruebaImpresion
        rp.lblTiempo = Format(Now(), "dd/mm/yyyy hh:mm:ss")
        rp.PrintReport False
    End If
    Unload Me
End If
Exit Sub
errlab: MsgBox "Error: " & err.Description
Unload Me
End Sub

Private Sub Form_Load()
Dim consulta As String
consulta = "SELECT     CASE Tipo_de_operación WHEN 'votnum' THEN Ultima_Versión_Acta ELSE '0' END AS version, Número_de_Acta AS nroActa, Nombre_del_Acta AS nombreActa, CAST(DAY(Fecha) AS varchar(2)) + '/' + CAST(MONTH(Fecha) AS varchar(2)) + '/' + CAST(YEAR(Fecha) " & _
"                      AS varchar(4)) AS fecha " & _
" From actas " & _
" WHERE     (Período_Legislativo = '" & periodo & "') AND (Versión_Acta = 0) AND (Sesión = " & sesion & ") " & _
" ORDER BY nroActa DESC "
Dim rs As New Recordset
SetearRs consulta, rs
If (rs.EOF) Then
    sinActas = True
End If
While Not rs.EOF
    Dim s As String
    s = Replace(rs!nombreActa, vbTab, "")
    s = Replace(s, vbCrLf, "")
    s = LTrim(RTrim(s))
    gridView.AddItem rs!nroActa & vbTab & s & vbTab & rs!fecha & vbTab & Trim(rs.Fields("version"))
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
gridView.ColWidth(1) = 9000
gridView.ColWidth(3) = 0
gridView.TextMatrix(0, 0) = "Acta"
gridView.TextMatrix(0, 1) = "Titulo"
gridView.TextMatrix(0, 2) = "Fecha"
End Sub

Private Sub gridView_DblClick()
If (gridView.Rows > 1) Then
    gridView.Enabled = False
    lblEstado.Caption = "Cargando, espere por favor..."
    lblEstado.ForeColor = vbRed
    mostrarActa gridView.TextMatrix(gridView.Row, 0), periodo, Val(sesion), gridView.TextMatrix(gridView.Row, 3)
    lblEstado.Caption = "Doble clic para seleccionar"
    lblEstado.ForeColor = vbWhite
    gridView.Enabled = True
End If
End Sub

Private Sub mostrarActa(pActa As Integer, pPeriodo As String, pSesion As Integer, pVersion As Integer)
    Dim verActa As New frmConsultarActa
    If PermisosTotales.ConsultaActas = 1 Then
        verActa.MostrarDatos pActa, pPeriodo, pSesion, pVersion
        If ImpresionDeConsola Then
            verActa.cmdReporte_Click
            Unload verActa
        Else
            verActa.Show vbModal
        End If
        
        Set verActa = Nothing
    Else
        MsgBox "El usuario no tiene permisos para realizar esta operacion", vbInformation + vbOKOnly
    End If
End Sub

