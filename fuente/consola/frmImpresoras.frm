VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmImpresoras 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Impresoras del Sistema"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Salir 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar Impresora"
      Height          =   855
      Left            =   2040
      Picture         =   "frmImpresoras.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nueva  Impresora"
      Height          =   855
      Left            =   120
      Picture         =   "frmImpresoras.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7223
      _Version        =   393216
      FixedCols       =   0
   End
End
Attribute VB_Name = "frmImpresoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strconexion As String
Private Rs As New ADODB.Recordset


Private Sub Form_Load()
    Call SetearGrilla
    Call LlenarGrilla
End Sub
Private Sub LlenarGrilla()
    
    Dim strSql         As String
    Dim strImpresora   As String
    Dim strDescripcion As String
    Dim strPredetermin As String
    Dim strAccesoDirec As String
    Dim xFila          As Long
    
    strSql = "SELECT * FROM impres"
    SetearRs strSql, Rs
    xFila = 1
    With Rs
        If .RecordCount > 0 Then
            Grilla.Rows = .RecordCount + 1
            While Not .EOF
                strImpresora = .Fields("Impresora").Value
                strDescripcion = .Fields("Descripción").Value
                If .Fields("Predeterminada").Value = 1 Then
                    strPredetermin = "Sí"
                Else
                    strPredetermin = "No"
                End If
                strAccesoDirec = .Fields("Acceso_Directo").Value
                Grilla.TextMatrix(xFila, 0) = strImpresora
                Grilla.TextMatrix(xFila, 1) = strDescripcion
                Grilla.TextMatrix(xFila, 2) = strPredetermin
                Grilla.TextMatrix(xFila, 3) = strAccesoDirec
                xFila = xFila + 1
                .MoveNext
            Wend
        End If
    End With
    Rs.Close
    Set Rs = Nothing
End Sub
Private Sub SetearGrilla()
    With Grilla
        .Cols = 4
        .TextMatrix(0, 0) = "Impresora"
        .TextMatrix(0, 1) = "Descripción"
        .TextMatrix(0, 2) = "Pred."
        .TextMatrix(0, 3) = "Acceso Directo"
        .ColWidth(0) = 2000
        .ColWidth(1) = 2500
        .ColWidth(2) = 1200
        .ColWidth(3) = 3000
    End With
End Sub

Private Sub Salir_Click()
    Unload Me
End Sub
