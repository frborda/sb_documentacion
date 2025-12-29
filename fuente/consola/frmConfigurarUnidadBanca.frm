VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmConfigurarUnidadBanca 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración de Bancas"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdActualizarIPs 
      Height          =   525
      Left            =   150
      TabIndex        =   5
      Top             =   7170
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   926
      BackColor       =   12230304
      Caption         =   "Actualizar IPs en el Servidor de Bancas"
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
   Begin MSFlexGridLib.MSFlexGrid Grilla 
      Height          =   5685
      Left            =   120
      TabIndex        =   1
      Top             =   1470
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10028
      _Version        =   393216
      Cols            =   5
      Appearance      =   0
   End
   Begin VB.Label versionsqv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión actual de datos SQV: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   810
      Width           =   2580
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<ESC> para salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   9210
      TabIndex        =   3
      Top             =   1185
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doble click para editar valores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   1185
      Width           =   2625
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de configuración que no requieren mantenimiento. Su modificación puede alterar el normal funcionamiento del sistema SQV."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1305
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10545
   End
End
Attribute VB_Name = "frmConfigurarUnidadBanca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Rs As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1

Private Sub CargarDatosGrilla()
    
    Dim strSql As String
    Dim xFila  As Long
    Dim i As Integer
    Set Rs = New ADODB.Recordset
    strSql = "SELECT * FROM BancasIP ORDER BY BancaNumero"
    SetearRs strSql, Rs
    xFila = 1
    With Rs
        If .RecordCount > 0 Then
            Grilla.Rows = .RecordCount + 1
            .MoveFirst
            While Not .EOF
                Grilla.TextMatrix(xFila, 0) = Trim(.Fields("BancaNumero").Value) ' "Banca"
                Grilla.TextMatrix(xFila, 1) = Trim(.Fields("Ip").Value)          ' "IP"
                Grilla.TextMatrix(xFila, 2) = Trim(.Fields("Puerto").Value)      ' "Puerto"
                Grilla.TextMatrix(xFila, 3) = "" & Trim(.Fields("Comentario").Value)  ' "Comentario"
                'Grilla.TextMatrix(xFila, 4) = Trim(.Fields("IdString").Value)    ' "Id String"
                Grilla.TextMatrix(xFila, 4) = Trim(.Fields("Version").Value)    ' "Id String"
                Grilla.TextMatrix(xFila, 5) = Trim(.Fields("version_datos_banca").Value)    ' "Id String"
                Grilla.TextMatrix(xFila, 6) = Trim(.Fields("version_datos_sqv").Value)    ' "Id String"
                xFila = xFila + 1
                .MoveNext
            Wend
            .Close
        End If
    End With
    'For i = 1 To Grilla.Rows
    '    MsgBox (Grilla.TextMatrix(i, 1))
    'Next i
    'MsgBox (Grilla.TextMatrix(Grilla.Rows - 1, 1))
    Set Rs = Nothing
End Sub

Private Sub SincronizarTodasLasBancas()
    MensajesSQV.SincronizarBancas ("brc")
    Unload Me
End Sub
Private Sub cmdSincronizar_Click()
    Call SincronizarTodasLasBancas
End Sub

Private Sub cmdActualizarIPs_Click()
Datos.GrabarMensaje "actualizarips", "", "", True
End Sub
Private Sub Form_Load()
    'versionsqv.Caption = versionsqv.Caption & " " & strVersion_datos_sqv
    Call TitulosGRilla
    Call CargarDatosGrilla
End Sub
Private Sub TitulosGRilla()
    Call SetVersion_datos_sqv
    versionsqv.Caption = versionsqv.Caption & " " & strVersion_datos_sqv
    With Grilla
        .Cols = 7
        .ColWidth(0) = 1000 ' banca
        .ColWidth(1) = 1400 ' ip
        .ColWidth(2) = 800 ' Puerto
        .ColWidth(3) = 1900 ' Comentario
        .ColWidth(4) = 1700 ' version
        .ColWidth(5) = 1700 ' version datos banca
        .ColWidth(6) = 1700 ' version sqv
        .TextMatrix(0, 0) = "Banca"
        .TextMatrix(0, 1) = "IP"
        .TextMatrix(0, 2) = "Puerto"
        .TextMatrix(0, 3) = "Comentario"
        '.TextMatrix(0, 4) = "Id String"
        .TextMatrix(0, 4) = "Versión última sinc. "
        .TextMatrix(0, 5) = "Versión Datos Banca"
        .TextMatrix(0, 6) = "Versión Datos SQV"
    End With
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Grilla_DblClick()
    Grilla.Col = 0
    frmEditarDatosUnidadBanca.Banca = Trim(Grilla.Text)
    frmEditarDatosUnidadBanca.Show vbModal
    Grilla.Clear
    Call TitulosGRilla
    Call CargarDatosGrilla
End Sub

Private Sub Label4_Click()

End Sub

