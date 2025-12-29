VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmElegirPresidenteActas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccionar Presidente para el acta"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid dgPresidente 
      Height          =   3195
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5636
      _Version        =   393216
      Cols            =   5
      SelectionMode   =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Buscar"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Doble click sobre el nombre seleccionado para elegir"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmElegirPresidenteActas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As Form

Public Function MostrarDatos(Grilla As MSFlexGrid, pColumnaCodigo As Integer, pColumnaNombre As Integer, pColumnaVoto As Integer, pForm As Form) As Boolean
    Dim i As Integer
    Dim agregado As Boolean
    Set mForm = pForm
    agregado = False
    For i = 1 To Grilla.Rows - 1
        If Trim(UCase(Grilla.TextMatrix(i, pColumnaVoto))) = "AUSENTE" Then
            dgPresidente.AddItem vbTab & Grilla.TextMatrix(i, pColumnaNombre) & vbTab & Grilla.TextMatrix(i, pColumnaCodigo) & vbTab & i
            agregado = True
        End If
    Next i
    If agregado = False Then
        MsgBox "No se han encontrado legisladores activos no asignados.", vbInformation + vbOKOnly
        MostrarDatos = False
    Else
        MostrarDatos = True
        dgPresidente.RemoveItem (1)
    End If
End Function
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub dgPresidente_DblClick()
    If dgPresidente.Row > 0 Then
        If dgPresidente.CellBackColor <> &H8000000F Then
            mForm.RealizarCambioPresidente dgPresidente.TextMatrix(dgPresidente.Row, 2), dgPresidente.TextMatrix(dgPresidente.Row, 1), dgPresidente.TextMatrix(dgPresidente.Row, 3)
            cmdCancelar_Click
        End If
    End If
End Sub

Private Sub dgPresidente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dgPresidente_DblClick
    End If
End Sub

Private Sub Form_Load()
    TitulosGRilla
End Sub
Private Sub TitulosGRilla()
    With dgPresidente
        .Cols = 4
        .ColWidth(0) = 100
        .ColWidth(1) = 3000
        .ColWidth(2) = 0 'codigo
        .ColWidth(3) = 0 'nro de fila
        .TextMatrix(0, 1) = "Legislador"
    End With
End Sub

Private Sub txtBuscar_GotFocus()
    seleccionadoTxt txtBuscar
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim Col As Integer
        Dim Row As Integer
        Funciones.BuscarEnGrilla dgPresidente, 3, txtBuscar.Text, Col, Row
        If (Col <> -1) And (Row <> -1) Then
            dgPresidente.SetFocus
            dgPresidente.Row = Row
            dgPresidente.RowSel = Row
            dgPresidente.ColSel = 2
        Else
            MsgBox "No se ha encontrado el texto deseado." & Chr(13) & "Intente con otra búsqueda.", vbInformation + vbOKOnly
        End If
    End If
End Sub
