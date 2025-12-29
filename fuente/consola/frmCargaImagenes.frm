VERSION 5.00
Begin VB.Form frmCargaImagenes 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cargando imágenes de diputados..."
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmImgs 
      Interval        =   500
      Left            =   2700
      Top             =   300
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descargando imágenes de diputados..."
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
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmCargaImagenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dirWs As Boolean

Private Sub CreaCarpeta()
On Error Resume Next
dirWs = WSExiste
If Not dirWs Then
    MsgBox "No se copiarán las fotos al Web Service ya que el directorio " & WSData.getWSFolder & " no existe."
Else
    Kill WSData.getWSFolder & "\*.*"
End If
MkDir (App.Path & "\FotosDiputados")
MkDir (App.Path & "\FotosDiputadosA")
Kill App.Path & "\FotosDiputados\*.*"
Kill App.Path & "\FotosDiputadosA\*.*"
End Sub

Private Function WSExiste() As Boolean
Dim b As Boolean
b = False
If (dir(WSData.getWSFolder, vbDirectory) <> "") Then
    b = True
End If
WSExiste = b
End Function

Private Sub tmImgs_Timer()
Dim rs As New ADODB.Recordset
Dim consulta As String
CreaCarpeta
consulta = "SELECT legisladores_activos.id, legisladores_activos.id_a, legisladores.PICTURE FROM legisladores_activos_full legisladores_activos INNER JOIN "
consulta = consulta & "legisladores ON legisladores.id = legisladores_activos.id"
SetearRs consulta, rs
While Not rs.EOF
    'NicoNicoNico
    If Not IsNull(rs.Fields("PICTURE")) Then
        Dim p As New ADODB.Stream
        p.Type = adTypeBinary
        p.Open
        p.Write rs.Fields("PICTURE")
        p.SaveToFile App.Path & "\FotosDiputados\" & rs.Fields("id") & ".jpg"
        p.SaveToFile App.Path & "\FotosDiputadosA\" & rs.Fields("id_a") & ".jpg"
        If dirWs Then
            p.SaveToFile WSData.getWSFolder & "\" & rs.Fields("id") & ".jpg"
        End If
        p.Close
    Else
        Call FileCopy(App.Path & "\sinfoto.jpg", App.Path & "\FotosDiputados\" & rs.Fields("id") & ".jpg")
    End If
    rs.MoveNext
    DoEvents
Wend
rs.Close
Unload Me
End Sub
