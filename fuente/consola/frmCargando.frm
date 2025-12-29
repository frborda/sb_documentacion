VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmCargando 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Cargando"
   ClientHeight    =   1755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdCancelar 
      Height          =   345
      Left            =   1080
      TabIndex        =   2
      Top             =   420
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      BackColor       =   8421631
      Caption         =   "Cancelar"
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
   Begin ComctlLib.ProgressBar prgBar 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   1380
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   476
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Timer tmOK2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3750
      Top             =   420
   End
   Begin VB.Timer tmOK 
      Interval        =   2000
      Left            =   3360
      Top             =   240
   End
   Begin Proyecto1.ButtonOffice cmdContinuar 
      Height          =   285
      Left            =   660
      TabIndex        =   3
      Top             =   1020
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   503
      BackColor       =   33023
      Caption         =   "Forzar ingreso a la Consola"
      Enabled         =   0   'False
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
      State           =   3
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   420
      X2              =   3600
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label lblState 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando, espere"
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
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3915
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PVez As Boolean
Private Sub cmdCancelar_Click()
tmOK.Enabled = False
tmOK2.Enabled = False
Error_Carga = True
Unload Me
End Sub

Private Sub cmdContinuar_Click()
Unload Me
End Sub
Private Sub Form_Load()
Error_Carga = False
PVez = True
prgBar.max = 200
prgBar.Min = 0
prgBar.Value = 0
nTick = GetTickCount
End Sub
Private Sub tmOK_Timer()
Dim nTick As Long
Dim Rs As ADODB.Recordset
If PVez = True Then
    lblState.Caption = "Esperando carga de presidente..."
    PVez = False
Else
    Set Rs = New ADODB.Recordset
    SetearRs "SELECT * FROM vector", Rs
    If Not Rs.EOF Then
        hacerSplitVector Trim(Rs!vector_identificacion), mVectorIdentificacion
    End If
    If mVectorIdentificacion(0) <> "0" Then
        lblState.Caption = "Presidente OK."
        cmdContinuar.Enabled = True
        DoEvents
        prgBar.Value = 1
        nTick = GetTickCount
        While GetTickCount - nTick < 1000
            DoEvents
        Wend
        tmOK2.Enabled = True
        PVez = True
        tmOK.Enabled = False
    End If
End If
End Sub
Private Sub tmOK2_Timer()
Dim nTick As Long
Dim rsTemp As ADODB.Recordset
Dim Cont As Integer
If PVez = True Then
    lblState.Caption = "Esperando respuesta de bancas..."
    PVez = False
Else
    Set rsTemp = New ADODB.Recordset
    SetearRs "SELECT * FROM vector", rsTemp
    Cont = CuentaX(rsTemp.Fields("vector_presencia"))
    lblState.Caption = "Bancas pendientes: " & Trim(Str(Cont))
    If Cont <= 57 Then
        prgBar.Value = 200
        lblState.Caption = "Bancas OK..."
        nTick = GetTickCount
        While GetTickCount - nTick < 2000
            DoEvents
        Wend
        Unload Me
    Else
        prgBar.Value = 257 - Cont
        DoEvents
    End If
    rsTemp.Close
    Set rsTemp = Nothing
End If
End Sub
Private Function CuentaX(cad As String) As Integer
Dim i As Integer
Dim Buff As String
Buff = cad
i = 0
While InStr(Buff, "X")
    Buff = Replace(Buff, "X", "", , 1)
    i = i + 1
Wend
CuentaX = i
End Function
Private Sub hacerSplitVector(ByVal pCadena As String, ByRef pVector() As String)
    pVector = Split(pCadena, ";")
End Sub
