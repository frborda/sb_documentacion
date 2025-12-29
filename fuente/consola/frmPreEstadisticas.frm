VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmPreEstadisticas 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selección de Estadísticas"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdEstadísticasIndividuales 
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "Estadísticas Individuales"
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
   Begin Proyecto1.ButtonOffice cmdEstadisticasGenerales 
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "Estadísticas Generales"
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
   Begin Proyecto1.ButtonOffice cmdCancelar 
      Height          =   495
      Left            =   60
      TabIndex        =   4
      Top             =   1440
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   873
      BackColor       =   12230304
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Muestra las estadísticas agrupadas por bloque, distrito, etcétera."
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
      Height          =   555
      Left            =   2820
      TabIndex        =   3
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Muestra las estadísticas de un diputado en particular"
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
      Left            =   2820
      TabIndex        =   1
      Top             =   180
      Width           =   5295
   End
End
Attribute VB_Name = "frmPreEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdCancelar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
Private Sub cmdEstadisticasGenerales_Click()
'Dim rsTemp As ADODB.Recordset
'Dim rpt As New rptEstadisticasGenerales
'Dim i As Integer
'Set rsTemp = New ADODB.Recordset
'SetearRs "SELECT tbEstadisticas.*,Legisladores.apellido + ', ' + Legisladores.nombre AS CDiputado FROM tbEstadisticas INNER JOIN Legisladores ON Legisladores.id = tbEstadisticas.id INNER JOIN distritos on distritos.id_distrito = Legisladores.distrito ORDER BY Legisladores.apellido,Legisladores.nombre", rsTemp
'If Not rsTemp.EOF Then
'    Set rpt.DataControl1.Recordset = rsTemp
'    rpt.Run False
'    For i = 0 To (rpt.Pages.Count - 1)
'        rpt.Pages(i).Width = 300
'    Next i
'    rpt.PrintReport True
'End If
'rsTemp.Close
'Set rsTemp = Nothing
frmEstadisticasGenerales.Show vbModal, Me
End Sub
Private Sub cmdEstadisticasGenerales_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
Private Sub cmdEstadísticasIndividuales_Click()
frmEstadisticas.Show vbModal, Me
End Sub
Private Sub cmdEstadísticasIndividuales_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

