VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConfigurarUnidadBanca 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración de Bancas"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSincronizar 
      Caption         =   "Sincronizar datos de bancas con SQV"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   7560
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla 
      Height          =   6855
      Left            =   120
      TabIndex        =   1
      Top             =   690
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   12091
      _Version        =   393216
      Cols            =   5
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "<ESC> para salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9180
      TabIndex        =   3
      Top             =   500
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   500
      Width           =   2625
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Datos de configuración que no requieren mantenimiento. Su modificación puede alterar el normal funcionamiento del sistema SQV."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10485
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
                Grilla.TextMatrix(xFila, 4) = Trim(.Fields("IdString").Value)    ' "Id String"
                xFila = xFila + 1
                .MoveNext
            Wend
            .Close
        End If
    End With
    Set Rs = Nothing
End Sub

Private Sub SincronizarTodasLasBancas()
    MensajesSQV.SincronizarBancas ("brc")
    Unload Me
End Sub
Private Sub cmdSincronizar_Click()
    Call SincronizarTodasLasBancas
End Sub

Private Sub Form_Load()
    Call TitulosGRilla
    Call CargarDatosGrilla
End Sub
Private Sub TitulosGRilla()
    With Grilla
        .Cols = 5
        .ColWidth(0) = 1000 ' banca
        .ColWidth(1) = 2000 ' ip
        .ColWidth(2) = 1000 ' Puerto
        .ColWidth(3) = 3000 ' Comentario
        .ColWidth(4) = 3200 ' Id String
        .TextMatrix(0, 0) = "Banca"
        .TextMatrix(0, 1) = "IP"
        .TextMatrix(0, 2) = "Puerto"
        .TextMatrix(0, 3) = "Comentario"
        .TextMatrix(0, 4) = "Id String"
    End With
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Grilla_DblClick()
    Grilla.col = 0
    frmEditarDatosUnidadBanca.Banca = Trim(Grilla.Text)
    frmEditarDatosUnidadBanca.Show vbModal
End Sub
