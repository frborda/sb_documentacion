VERSION 5.00
Begin VB.Form frmNuevoScan 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Escaneando..."
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmId 
      Interval        =   100
      Left            =   3960
      Top             =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   4620
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label lblApellido 
      BackStyle       =   0  'Transparent
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
      Left            =   60
      TabIndex        =   4
      Top             =   840
      Width           =   4515
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   4620
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblBloque 
      BackStyle       =   0  'Transparent
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
      Left            =   60
      TabIndex        =   3
      Top             =   2040
      Width           =   4575
   End
   Begin VB.Label lblProvincia 
      BackStyle       =   0  'Transparent
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
      Left            =   60
      TabIndex        =   2
      Top             =   3000
      Width           =   4575
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
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
      Left            =   60
      TabIndex        =   1
      Top             =   1260
      Width           =   4515
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   4620
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Image imgFoto 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3795
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2955
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Prueba de Scan de Banca"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4635
   End
End
Attribute VB_Name = "frmNuevoScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Banca As String
Public Titulo As String
Private Sub Form_Load()
lblTitulo.Caption = lblTitulo.Caption & " " & Banca
Me.Caption = Titulo & Banca
lblNombre.Caption = "Escaneando..."
End Sub
Private Sub Form_Unload(Cancel As Integer)
Datos.GrabarMensaje "scan?finprueba", Trim(Banca), , True
End Sub
Private Sub tmId_Timer()
Dim Rinfo As ADODB.Recordset
Dim pic As ADODB.Stream
If Trim(mVectorIdentificacion(Val(Banca))) <> "0" Then
    tmId.Enabled = False
    Set Rinfo = New ADODB.Recordset
    SetearRs "SELECT nombre,apellido,bloque_politico,PICTURE,distritos.distrito AS Provincia FROM Legisladores LEFT OUTER JOIN distritos ON Legisladores.distrito = distritos.id_distrito WHERE id='" & mVectorIdentificacion(Val(Banca)) & "'", Rinfo
    If Rinfo.EOF Then
        Call MsgBox("Error de integridad", vbCritical)
        Unload Me
    Else
        If (IsNull(Rinfo.Fields("PICTURE")) = False) Then
            Set pic = New ADODB.Stream
            pic.Type = adTypeBinary
            pic.Open
            pic.Write Rinfo.Fields("PICTURE")
            pic.SaveToFile App.Path & "\temp.jpg", adSaveCreateOverWrite
            imgFoto.Picture = LoadPicture(App.Path & "\temp.jpg")
        Else
            Set imgFoto.Picture = Nothing
        End If
        lblNombre.Caption = Rinfo.Fields("nombre")
        lblApellido.Caption = Rinfo.Fields("apellido")
        lblBloque.Caption = IIf(IsNull(Rinfo.Fields("bloque_politico")), "Bloque Vacio", Rinfo.Fields("bloque_politico"))
        lblProvincia.Caption = IIf(IsNull(Rinfo.Fields("Provincia")), "Provincia Vacio", Rinfo.Fields("Provincia"))
    End If
    Rinfo.Close
    Set Rinfo = Nothing
End If
End Sub
