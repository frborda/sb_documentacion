VERSION 5.00
Object = "{429F6260-B945-11D3-9A1F-9E6707138531}#1.0#0"; "Vsflex7N.ocx"
Begin VB.Form frmrSelLegislador 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SELECCIÓN DE ORADOR"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleMode       =   0  'User
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Foto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   4560
      Left            =   6030
      TabIndex        =   4
      Top             =   60
      Width           =   4095
      Begin VB.Image picDiputado 
         Height          =   4095
         Left            =   150
         Stretch         =   -1  'True
         Top             =   300
         Width           =   3795
      End
   End
   Begin VB.TextBox txtApellido 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   450
      Width           =   4605
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Búsqueda por aproximación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1065
      Left            =   90
      TabIndex        =   2
      Top             =   60
      Width           =   5775
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   150
         TabIndex        =   3
         Top             =   420
         Width           =   795
      End
   End
   Begin VSFlex7NCtl.VSFlexGrid vsGrilla 
      Height          =   3390
      Left            =   90
      TabIndex        =   0
      Top             =   1230
      Width           =   5775
      _cx             =   10186
      _cy             =   5980
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmrSelLegislador.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   0   'False
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmrSelLegislador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql   As String
Dim strOrder As String
Public Seleccionado As Boolean
Private Sub CargarGrilla()
    Dim RsTemp As ADODB.Recordset
    Dim xFila  As Long
    Set RsTemp = New ADODB.Recordset
    SetearRs strSql & " " & strOrder, RsTemp
    xFila = 1
    vsGrilla.Rows = 1
    With RsTemp
        If .RecordCount > 0 Then
            .MoveFirst
            vsGrilla.Rows = .RecordCount + 1
            While Not .EOF
                vsGrilla.TextMatrix(xFila, 0) = .Fields("id").Value
                vsGrilla.TextMatrix(xFila, 1) = .Fields("apellido").Value & ", " & .Fields("nombre").Value
                xFila = xFila + 1
                .MoveNext
            Wend
        End If
        .Close
    End With
    Set RsTemp = Nothing
End Sub
Private Sub Ok()
    frmABMLegisladores.lblId.Caption = vsGrilla.TextMatrix(vsGrilla.Row, 0)
    LimpiaOrador = False
    Seleccionado = True
    Unload Me
End Sub
Private Sub Cancelar()
    frmABMLegisladores.lblId.Caption = "nothing"
    Unload Me
End Sub

Private Sub chkOrdenarPorApellido_Click()

End Sub

Private Sub cmdAplicar_Click()

End Sub

Private Sub Command1_Click()
    Call Ok
End Sub
Private Sub Command2_Click()
    Call Cancelar
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    LimpiaOrador = True
    'strSql = "SELECT * FROM Legisladores WHERE Es_Legislador " & IIf(mModoMantenimiento, ">=0", "=1")
    strSql = "SELECT * FROM legisladores_activos"
    strOrder = "ORDER BY apellido"
    Call CargarGrilla
    vsGrilla.ColWidth(0) = 0
    vsGrilla.ColWidth(1) = 5500
    vsGrilla_EnterCell
    Seleccionado = False
    vsGrilla_Click
End Sub

Private Sub Option1_Click()
    strOrder = "ORDER BY Apellido"
    Call CargarGrilla
End Sub
Private Sub Option2_Click()
    strOrder = "ORDER BY Id"
    Call CargarGrilla
End Sub

Private Sub optNombre_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Seleccionado = False Then

End If
End Sub

Private Sub txtApellido_Change()
'If txtApellido.Text <> "" Then
'    strSql = "SELECT * FROM legisladores_activos WHERE apellido LIKE '" & _
'            txtApellido.Text & "%'"
'    strOrder = " ORDER BY apellido "
'    Call CargarGrilla
'    Set picDiputado.Picture = Nothing
'    vsGrilla.SetFocus
'    vsGrilla.RowSel = 1
'    vsGrilla.ColSel = 1
'    vsGrilla_EnterCell
'    txtApellido.SetFocus
'Else
'    strSql = "SELECT * FROM legisladores_activos"
'    strOrder = " ORDER BY apellido "
'    Call CargarGrilla
'    Set picDiputado.Picture = Nothing
'    vsGrilla.SetFocus
'    vsGrilla.RowSel = 1
'    vsGrilla.ColSel = 1
'    vsGrilla_EnterCell
'    txtApellido.SetFocus
'End If
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAplicar_Click
End If
End Sub
Private Sub vsGrilla_Click()
If vsGrilla.Rows = 2 Or (vsGrilla.RowSel = 1 And vsGrilla.TextMatrix(vsGrilla.Row, 0) <> "") Then
    vsGrilla_EnterCell
End If
End Sub

Private Sub vsGrilla_DblClick()
    Call Ok
End Sub

Private Sub vsGrilla_EnterCell()
On Error Resume Next
Dim pic As New ADODB.Stream
Dim Rinfo As New ADODB.Recordset
If vsGrilla.TextMatrix(vsGrilla.Row, 0) <> "" Then
    SetearRs "SELECT PICTURE FROM legisladores WHERE id = " & vsGrilla.TextMatrix(vsGrilla.Row, 0), Rinfo
    If Not Rinfo.EOF Then
        If Not IsNull(Rinfo.Fields(0)) Then
            Set pic = New ADODB.Stream
            pic.Type = adTypeBinary
            pic.Open
            pic.Write Rinfo.Fields(0)
            pic.SaveToFile App.Path & "\temp.jpg", adSaveCreateOverWrite
            picDiputado.Picture = LoadPicture(App.Path & "\temp.jpg")
        Else
            Set picDiputado.Picture = Nothing
        End If
    Else
        Set picDiputado.Picture = Nothing
    End If
End If
End Sub

Private Sub vsGrilla_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
MsgBox ("asd")
End Sub

Private Sub vsGrilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Ok
    End If
End Sub
