VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmNuevoInfo 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORMACIÓN DE BANCA"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmScan 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4500
      Top             =   2520
   End
   Begin VB.Frame frmFoto 
      BackColor       =   &H00404040&
      Height          =   4995
      Left            =   4860
      TabIndex        =   0
      Top             =   60
      Width           =   3855
      Begin VB.Image imgDiputado 
         Height          =   4695
         Left            =   120
         Stretch         =   -1  'True
         Top             =   180
         Width           =   3615
      End
   End
   Begin Proyecto1.ButtonOffice cmdLimpiarID 
      Height          =   600
      Left            =   2400
      TabIndex        =   5
      Top             =   3480
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   1058
      BackColor       =   12230304
      Caption         =   "Limpiar Identificación"
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
   Begin Proyecto1.ButtonOffice cmdReiniciarBanca 
      Height          =   420
      Left            =   180
      TabIndex        =   6
      Top             =   4140
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   741
      BackColor       =   192
      Caption         =   "Reiniciar"
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
   Begin Proyecto1.ButtonOffice cmdPruebaScan 
      Height          =   600
      Left            =   180
      TabIndex        =   7
      Top             =   3480
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   1058
      BackColor       =   12230304
      Caption         =   "Prueba de Scan"
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
   Begin Proyecto1.ButtonOffice cmdDeshabilitarBanca 
      Height          =   420
      Left            =   2400
      TabIndex        =   8
      Top             =   4140
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   741
      BackColor       =   192
      Caption         =   "Deshabilitar"
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
   Begin Proyecto1.ButtonOffice cmdAsignarVoto 
      Height          =   420
      Left            =   180
      TabIndex        =   9
      Top             =   4620
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   741
      BackColor       =   12230304
      Caption         =   "Asignar Voto"
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
      Height          =   420
      Left            =   2400
      TabIndex        =   10
      Top             =   4620
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   741
      BackColor       =   12230304
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Banca"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   1035
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   4620
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblBanca 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   1200
      TabIndex        =   11
      Top             =   0
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4680
      Y1              =   3300
      Y2              =   3300
   End
   Begin VB.Label lblApellido 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4515
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   4620
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Label lblBloque 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Top             =   1860
      Width           =   4575
   End
   Begin VB.Label lblProvincia 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2820
      Width           =   4575
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1140
      Width           =   4515
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4680
      Y1              =   2700
      Y2              =   2700
   End
End
Attribute VB_Name = "frmNuevoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BancaID As Integer
Private Sub cmdAsignarVoto_Click()
Dim voto As New frmDefinirVoto
voto.Banca = BancaID
voto.Show vbModal
Set voto = Nothing
Unload Me
End Sub

Private Sub cmdAsignarVoto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdDeshabilitarBanca_Click()
Datos.GrabarMensaje "banca?deshabilitar", Trim(Str(BancaID)), , True
Unload Me
End Sub
Private Sub cmdDeshabilitarBanca_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub

Private Sub cmdLimpiarID_Click()
Datos.GrabarMensaje "limpieza_individual", Trim(Str(BancaID)), , True
frmConsolaOperacion.MensajeEsperado = MensajeVacio
Unload Me
End Sub
Private Sub cmdLimpiarID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub

Private Sub cmdPruebaScan_Click()
If tmScan.Enabled = False Then
    Datos.GrabarMensaje "scan?prueba", Trim(Str(BancaID)), , True
    cmdPruebaScan.Caption = "Prueba Iniciada"
    cmdDeshabilitarBanca.Enabled = False
    cmdAsignarVoto.Enabled = False
    cmdLimpiarID.Enabled = False
    cmdPruebaScan.Enabled = False
    cmdReiniciarBanca.Enabled = False
    tmScan.Enabled = True
End If
End Sub

Private Sub cmdPruebaScan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub

Private Sub cmdReiniciarBanca_Click()
MensajesSQV.reiniciarBanca Trim(Str(BancaID))
Unload Me
End Sub
Private Sub cmdReiniciarBanca_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub

Private Sub Form_Load()
Dim RsTemp As ADODB.Recordset
lblBanca.Caption = Trim(Str(BancaID))
Set RsTemp = New ADODB.Recordset
SetearRs "SELECT habilitada FROM BancasDeshabilitadas WHERE banca = " & BancaID, RsTemp
If RsTemp.EOF Then
    cmdDeshabilitarBanca.Caption = "Error"
Else
    If RsTemp.Fields(0) = 0 Then
        cmdDeshabilitarBanca.Caption = "Habilitar Banca"
    Else
        cmdDeshabilitarBanca.Caption = "Deshabilitar Banca"
    End If
End If
RsTemp.Close
Set RsTemp = Nothing
cmdPruebaScan.Enabled = False
cmdAsignarVoto.Enabled = False
If BancaID <> -1 Then
    If mVectorIdentificacion(BancaID) <> "0" Then
        ActualizaInfo (mVectorIdentificacion(BancaID))
    End If
End If
If mVectorIdentificacion(BancaID) <> "0" Then
    cmdPruebaScan.Enabled = False
End If
If frmConsolaOperacion.dcTipoOperacion.BoundText = "paslis" Or frmConsolaOperacion.dcTipoOperacion.BoundText = "votnom" Then
    If mVectorIdentificacion(BancaID) <> "0" Then
        cmdAsignarVoto.Enabled = True
    End If
ElseIf frmConsolaOperacion.dcTipoOperacion.BoundText = "votnum" Then
    cmdAsignarVoto.Enabled = True
End If
If Not mModo_Ident_Nom And frmConsolaOperacion.dcTipoOperacion.BoundText = "quorum" And mVectorIdentificacion(BancaID) = "0" Then
    cmdPruebaScan.Enabled = True
End If
End Sub
Private Sub ActualizaInfo(id As String)
Dim RsTemp As ADODB.Recordset
Dim pic As ADODB.Stream
Set RsTemp = New ADODB.Recordset
SetearRs "SELECT Legisladores.nombre,Legisladores.apellido, Legisladores.PICTURE, Legisladores.bloque_politico,distritos.distrito AS Provincia FROM Legisladores LEFT OUTER JOIN distritos ON Legisladores.distrito = distritos.id_distrito WHERE Legisladores.id = " & id, RsTemp
If RsTemp.EOF Then
    Call MsgBox("No se encontró al diputado", vbCritical)
    Unload Me
End If
lblNombre.Caption = RsTemp.Fields("nombre")
lblApellido.Caption = RsTemp.Fields("apellido")
lblBloque.Caption = IIf(IsNull(RsTemp.Fields("bloque_politico")), "-", RsTemp.Fields("bloque_politico"))
lblProvincia.Caption = IIf(IsNull(RsTemp.Fields("Provincia")), "-", RsTemp.Fields("Provincia"))
If Not IsNull(RsTemp.Fields("PICTURE")) Then
    Set pic = New ADODB.Stream
    pic.Type = adTypeBinary
    pic.Open
    pic.Write RsTemp.Fields("PICTURE")
    pic.SaveToFile App.Path & "\temp.jpg", adSaveCreateOverWrite
    imgDiputado.Picture = LoadPicture(App.Path & "\temp.jpg")
    pic.Close
    Set pic = Nothing
Else
    Set imgDiputado.Picture = Nothing
End If
RsTemp.Close
Set RsTemp = Nothing
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If tmScan.Enabled = True Then
    Datos.GrabarMensaje "scan?finprueba", Trim(Str(BancaID)), , True
End If
End Sub

Private Sub frmFoto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmConsolaOperacion.pctInfo.Visible = True Then
    frmConsolaOperacion.pctInfo.Visible = False
End If
End Sub
Private Sub tmScan_Timer()
Dim RsTemp As ADODB.Recordset
Dim nTick As Long
Dim MiVector() As String
If lblNombre.Caption = "-" Then
    Set RsTemp = New ADODB.Recordset
    SetearRs "SELECT * FROM vector", RsTemp
    If Not RsTemp.EOF Then
        frmConsolaOperacion.hacerSplitVector Trim(RsTemp!vector_identificacion), MiVector
        If MiVector(BancaID) <> "0" Then
            ActualizaInfo (MiVector(BancaID))
            cmdCancelar.Caption = "Cerrar Ventana"
        End If
    End If
    RsTemp.Close
    Set RsTemp = Nothing
End If
End Sub
