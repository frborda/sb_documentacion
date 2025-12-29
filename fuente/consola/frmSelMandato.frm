VERSION 5.00
Object = "{429F6260-B945-11D3-9A1F-9E6707138531}#1.0#0"; "Vsflex7N.ocx"
Begin VB.Form frmSelMandato 
   Caption         =   "Seleccionar Mandato"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   550
      Left            =   4005
      ScaleHeight     =   495
      ScaleWidth      =   1710
      TabIndex        =   0
      Top             =   0
      Width           =   1770
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   855
         TabIndex        =   2
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   855
      End
   End
   Begin VSFlex7NCtl.VSFlexGrid vsGrilla 
      Height          =   6135
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   5775
      _cx             =   10186
      _cy             =   10821
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSelMandato.frx":0000
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
      ShowComboButton =   -1  'True
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
Attribute VB_Name = "frmSelMandato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql   As String
Dim strOrder As String

Private Sub CargarGrilla()
    Dim RsTemp As ADODB.Recordset
    Dim xFila  As Long
    Set RsTemp = New ADODB.Recordset
    SetearRs strSql + strOrder, RsTemp
    xFila = 1
    vsGrilla.Rows = 1
    With RsTemp
        If .RecordCount > 0 Then
            .MoveFirst
            vsGrilla.Rows = .RecordCount + 1
            While Not .EOF
                vsGrilla.TextMatrix(xFila, 0) = .Fields("fecha_mandato").Value
                xFila = xFila + 1
                .MoveNext
            Wend
        End If
        .Close
    End With
    Set RsTemp = Nothing
End Sub
Private Sub Ok()
    frmABMMandatos.lblid.Caption = vsGrilla.TextMatrix(vsGrilla.row, 0)
     
    Unload Me
End Sub
Private Sub Cancelar()
    frmABMMandatos.lblid.Caption = "nothing"
    Unload Me
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
    strSql = "SELECT * FROM mandatos "
    strOrder = "ORDER BY fecha_mandato"
    Call CargarGrilla
End Sub
Private Sub vsGrilla_DblClick()
    Call Ok
End Sub
Private Sub vsGrilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Ok
    End If
End Sub


