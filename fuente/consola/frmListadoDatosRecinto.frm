VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmListadoDatosRecinto 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de datos de Recinto"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Orden"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   120
      TabIndex        =   4
      Top             =   780
      Width           =   6255
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Apellido y nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   6255
      Begin VB.ComboBox cmbCalidadMaxima 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmListadoDatosRecinto.frx":0000
         Left            =   5460
         List            =   "frmListadoDatosRecinto.frx":0002
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cmbCalidadMinima 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmListadoDatosRecinto.frx":0004
         Left            =   4440
         List            =   "frmListadoDatosRecinto.frx":0006
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "y"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5280
         TabIndex        =   6
         Top             =   300
         Width           =   195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Solo listar las huellas que tengan una calidad entre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   4215
      End
   End
   Begin Proyecto1.ButtonOffice cmdGenerarListado 
      Height          =   465
      Left            =   4380
      TabIndex        =   2
      Top             =   1620
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   820
      BackColor       =   12230304
      Caption         =   "Generar listado"
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
   Begin Proyecto1.ButtonOffice cmdVolver 
      Height          =   465
      Left            =   120
      TabIndex        =   7
      Top             =   1620
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   820
      BackColor       =   12230304
      Caption         =   "Volver"
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
End
Attribute VB_Name = "frmListadoDatosRecinto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGenerarListado_Click()
Dim rpt As New rptHuellas
Dim rs As New Recordset
If (IsNumeric(cmbCalidadMinima.Text) And IsNumeric(cmbCalidadMaxima.Text)) Then
    SetearRs "SELECT GETDATE() as fecha_actual, apellido + ', ' + nombre AS apellidoNombre, banca AS bancaProbable, IsNull(dedo,'-') AS dedo, IsNull(CAST(calidadMinima AS varchar(3)) + '/' + CAST(calidadMaxima AS varchar(3)),'-') AS calidad FROM VistaDedos WHERE IsNull(calidadMinima,0) >= " & cmbCalidadMinima.Text & " AND IsNull(calidadMaxima,0) <= " & cmbCalidadMaxima.Text & " ORDER BY apellido,nombre", rs
    rpt.DataControl1.Recordset = rs
    rpt.PrintReport True
Else
    MsgBox "Los datos de calida deben ser numéricos"
End If
End Sub

Private Sub cmdVolver_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
Dim i As Integer
cmbCalidadMinima.Clear
cmbCalidadMaxima.Clear
For i = 0 To 100
    cmbCalidadMinima.AddItem Trim(Str(i))
    cmbCalidadMaxima.AddItem Trim(Str(i))
Next i
cmbCalidadMinima.ListIndex = 0
cmbCalidadMaxima.ListIndex = cmbCalidadMaxima.ListCount - 1
End Sub
