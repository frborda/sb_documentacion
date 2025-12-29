VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHistoricoBanca 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Histórico "
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   14445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDuplicidades 
      Caption         =   "Ver solo duplicidades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1620
      TabIndex        =   2
      Top             =   60
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid dgLogBanca 
      Height          =   4875
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   8599
      _Version        =   393216
      Cols            =   4
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Escape para salir"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmHistoricoBanca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strBanca       As String
Private RsBanca        As ADODB.Recordset
Private strCaptionForm As String

Private Sub chkDuplicidades_Click()
If chkDuplicidades.Value = vbChecked Then
    armarGrilla
    Dim strSql        As String
    Dim xId           As Long
    Dim strOrigen     As String
    Dim dFecha        As Date
    Dim strDescripcion As String
    Set RsBanca = New ADODB.Recordset
    strSql = "SELECT * FROM Log_General WHERE (Rtrim(Objeto) = " & strBanca & ") AND severidad = 5 ORDER BY id DESC"
    SetearRs strSql, RsBanca
    With RsBanca
        If .RecordCount > 0 Then
            While Not .EOF
                xId = .Fields("Id").Value
                strOrigen = .Fields("Origen").Value
                dFecha = .Fields("Fecha").Value
                strDescripcion = .Fields("Descripcion").Value
                dgLogBanca.AddItem xId & vbTab & strOrigen & vbTab & dFecha & vbTab & strDescripcion
                .MoveNext
            Wend
        End If
    End With
Else
    armarGrilla
    CargarGrilla
End If

End Sub

Private Sub Form_Load()
    Call armarGrilla
    Call CargarGrilla
End Sub
Private Sub CargarGrilla()
    On Error GoTo TrapError
    Dim strSql        As String
    Dim xId           As Long
    Dim strOrigen     As String
    Dim dFecha        As Date
    Dim strDescripcion As String
    
    Set RsBanca = New ADODB.Recordset
    strSql = "SELECT * FROM Log_General WHERE (Rtrim(Objeto) = '') Or (Rtrim(Objeto) = " & strBanca & ") ORDER BY id DESC"
    SetearRs strSql, RsBanca
    With RsBanca
        If .RecordCount > 0 Then
            While Not .EOF
                xId = .Fields("Id").Value
                strOrigen = .Fields("Origen").Value
                dFecha = .Fields("Fecha").Value
                strDescripcion = .Fields("Descripcion").Value
                dgLogBanca.AddItem xId & vbTab & strOrigen & vbTab & dFecha & vbTab & strDescripcion
                .MoveNext
            Wend
        End If
    End With
Exit Sub
TrapError:
    Select Case Err.Number
        Case Else
            MsgBox "Error Nº " & Err.Number & Chr(10) & Err.Description & Chr(10) & "Originado en " & Err.Source
            Resume
    End Select
End Sub
Private Sub armarGrilla()
    With dgLogBanca
        .Clear
        .Cols = 4
        .Rows = 2
        .FixedCols = 1
        .ColWidth(0) = 800
        .ColWidth(1) = 3000
        .ColWidth(2) = 2000
        .ColWidth(3) = 8000
        .TextMatrix(0, 0) = "Id"
        .TextMatrix(0, 1) = "Origen"
        .TextMatrix(0, 2) = "Fecha"
        .TextMatrix(0, 3) = "Descripción"
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Public Property Let Banca(ByVal vNewValue As Variant)
    strBanca = vNewValue
    strBanca = Trim(strBanca)
    strCaptionForm = "Histórico Banca " & strBanca
    Me.Caption = strCaptionForm
End Property
