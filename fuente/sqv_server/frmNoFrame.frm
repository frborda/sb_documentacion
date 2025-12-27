VERSION 5.00
Begin VB.Form frmNoFrame 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Nuevo"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   10680
      Top             =   6720
   End
   Begin VB.Shape shpQuorum 
      BorderColor     =   &H000000C0&
      Height          =   915
      Left            =   8100
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   7035
   End
   Begin VB.Label lblHora 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   4560
      TabIndex        =   3
      Top             =   60
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   4380
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   3840
      X2              =   4200
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   3615
   End
   Begin VB.Label lblFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   240
      TabIndex        =   2
      Top             =   60
      Width           =   3315
   End
   Begin VB.Label lblQuorum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NO HAY QUORUM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   38.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1035
      Left            =   8100
      TabIndex        =   1
      Top             =   0
      Width           =   7035
   End
   Begin VB.Label lblVotacion 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   795
      Left            =   7140
      TabIndex        =   0
      Top             =   7920
      Visible         =   0   'False
      Width           =   4215
   End
End
Attribute VB_Name = "frmNoFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblQuorum_Change()
If lblQuorum.Caption = "QUORUM" Then
    lblQuorum.ForeColor = &HFFFF&
    shpQuorum.BorderColor = &HFFFF&
Else
    lblQuorum.ForeColor = &HC0&
    shpQuorum.BorderColor = &HC0&
End If
End Sub
Private Sub Timer1_Timer()
lblFecha.Caption = Format(Now(), "dd/mm/yyyy")
If lblHora.Caption <> Format(Now(), "hh:mm") Then
    lblHora.Caption = Format(Now(), "hh:mm")
End If
lblQuorum.Caption = UCase(Trim(CartelActual.LeyendaQuorum))
End Sub
Private Function NumeroPeriodo() As String
Dim Temp As String
Temp = Mid(EstadoActual.PeriodoLegislativo, 1, 3)
NumeroPeriodo = Temp
End Function
Private Function ObtenerTipoPeriodo() As String
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
frmMain.SetearRsAux "SELECT leyenda_para_cartel FROM tipo_periodo WHERE id = '" & Mid(EstadoActual.PeriodoLegislativo, 4, 1) & "'", rsTemp
If rsTemp.EOF Then
    ObtenerTipoPeriodo = "Invalido"
Else
    ObtenerTipoPeriodo = rsTemp.Fields(0)
End If
rsTemp.Close
Set rsTemp = Nothing
End Function
Private Function ObtenerTipoSesion() As String
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
frmMain.SetearRsAux "SELECT leyenda_para_cartel FROM tipo_sesion WHERE id = '" & Mid(EstadoActual.PeriodoLegislativo, 5, 1) & "'", rsTemp
If rsTemp.EOF Then
    ObtenerTipoSesion = "Invalido"
Else
    ObtenerTipoSesion = rsTemp.Fields(0)
End If
rsTemp.Close
Set rsTemp = Nothing
End Function
