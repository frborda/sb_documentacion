VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHistoricoCompleto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Histórico Completo"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   780
      Width           =   8775
   End
   Begin MSFlexGridLib.MSFlexGrid flexDatos 
      Height          =   5595
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9869
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtro por banca"
      Height          =   735
      Left            =   2880
      TabIndex        =   2
      Top             =   0
      Width           =   2835
      Begin VB.ComboBox cmbBancas 
         Height          =   315
         ItemData        =   "frmHistoricoCompleto.frx":0000
         Left            =   120
         List            =   "frmHistoricoCompleto.frx":0002
         TabIndex        =   3
         Text            =   " - Seleccione una banca -"
         Top             =   240
         Width           =   2595
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro por causa"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2835
      Begin VB.ComboBox cmbCausas 
         Height          =   315
         ItemData        =   "frmHistoricoCompleto.frx":0004
         Left            =   120
         List            =   "frmHistoricoCompleto.frx":0006
         TabIndex        =   1
         Text            =   " - Seleccione una causa -"
         Top             =   240
         Width           =   2595
      End
   End
End
Attribute VB_Name = "frmHistoricoCompleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActualizar_Click()
Dim severidad As Integer
Dim conta As Integer
Dim rsActualizar As ADODB.Recordset
Dim Consulta As String
Set rsActualizar = New ADODB.Recordset
Select Case cmbCausas.List(cmbCausas.ListIndex)
Case ""
    severidad = -1 'Nada seleccionado
Case "Error de Switch"
    severidad = 2 'Switch
Case "Duplicidad de ID"
    severidad = 5 'Id duplicada
Case "Identificación Manual"
    severidad = 3 'SAUTOD
Case Else
    severidad = -2
End Select
Limpiar
conta = 0
If severidad <> -2 And severidad <> -1 Then
    If cmbBancas.ListIndex <= 0 Then
        Consulta = "SELECT TOP 100 * FROM log_general WHERE severidad = " & Trim(Str(severidad)) & " ORDER BY id DESC"
        SetearRs Consulta, rsActualizar
    Else
        Consulta = "SELECT TOP 100 * FROM log_general WHERE severidad = " & Trim(Str(severidad)) & " AND LTrim(objeto ) = " & cmbBancas.List(cmbBancas.ListIndex)
        SetearRs Consulta, rsActualizar
    End If
Else
    If cmbBancas.ListIndex <= 0 Then
        Consulta = "SELECT TOP 100 * FROM log_general ORDER BY id DESC"
        SetearRs Consulta, rsActualizar
    Else
        Consulta = "SELECT TOP 100 * FROM log_general WHERE LTrim(objeto) = " & cmbBancas.List(cmbBancas.ListIndex)
        SetearRs Consulta, rsActualizar
    End If
End If
While Not rsActualizar.EOF
    flexDatos.AddItem ""
    conta = conta + 1
    flexDatos.TextMatrix(conta, 0) = rsActualizar.Fields("origen")
    flexDatos.TextMatrix(conta, 1) = rsActualizar.Fields("fecha")
    flexDatos.TextMatrix(conta, 2) = rsActualizar.Fields("descripcion")
    If Trim(rsActualizar.Fields("objeto")) = "" Then
        flexDatos.TextMatrix(conta, 3) = "Masivo"
    ElseIf InStr(rsActualizar.Fields("objeto"), ";") > 0 Then
        flexDatos.TextMatrix(conta, 3) = "Semi-Masivo"
    Else
        flexDatos.TextMatrix(conta, 3) = rsActualizar.Fields("objeto")
    End If
    rsActualizar.MoveNext
Wend
rsActualizar.Close
Set rsActualizar = Nothing
End Sub
Private Sub flexDatos_DblClick()
If flexDatos.TextMatrix(flexDatos.Row, 2) <> "" Then
    MsgBox flexDatos.TextMatrix(flexDatos.Row, 2), vbInformation, "Descripción"
End If
End Sub
Private Sub Form_Load()
Dim i As Integer
Limpiar
cmbBancas.AddItem "Todas"
For i = 1 To 257
    cmbBancas.AddItem Trim(Str(i))
Next i
cmbCausas.AddItem "Duplicidad de ID"
cmbCausas.AddItem "Otros"
End Sub
Private Sub Limpiar()
flexDatos.Clear
flexDatos.Cols = 4
flexDatos.TextMatrix(0, 0) = "Resumen"
flexDatos.ColWidth(0) = 4000
flexDatos.TextMatrix(0, 1) = "Fecha"
flexDatos.ColWidth(1) = 2000
flexDatos.TextMatrix(0, 2) = "Descripción"
flexDatos.ColWidth(2) = 0
flexDatos.TextMatrix(0, 3) = "Objeto/Banca"
flexDatos.ColWidth(3) = 2000
flexDatos.Rows = 1
End Sub
