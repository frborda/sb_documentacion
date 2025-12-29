VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmEditarPartidos 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edición de Partidos Políticos"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBusqueda 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin VB.ListBox lstPartidos 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   90
      TabIndex        =   8
      Top             =   600
      Width           =   8475
   End
   Begin VB.Frame frmOpt 
      BackColor       =   &H00404040&
      Caption         =   "Edición"
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
      Height          =   1425
      Left            =   90
      TabIndex        =   3
      Top             =   4260
      Width           =   8385
      Begin VB.TextBox txtPartido 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   930
         MaxLength       =   80
         TabIndex        =   5
         Top             =   360
         Width           =   7335
      End
      Begin Proyecto1.ButtonOffice cmdGrabar 
         Height          =   585
         Left            =   6210
         TabIndex        =   4
         Top             =   750
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1032
         BackColor       =   16744576
         Caption         =   "&Guardar Cambios"
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
      Begin Proyecto1.ButtonOffice cmdLimpiar 
         Height          =   585
         Left            =   4110
         TabIndex        =   6
         Top             =   750
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1032
         BackColor       =   16744576
         Caption         =   "&Limpiar"
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
         Caption         =   "Nombre"
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
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   390
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Otros"
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
      Height          =   795
      Left            =   90
      TabIndex        =   1
      Top             =   5700
      Width           =   8385
      Begin VB.TextBox txtAddPartido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   300
         Width           =   5535
      End
      Begin Proyecto1.ButtonOffice cmdAgregarPartido 
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   714
         BackColor       =   16744576
         Caption         =   "&Agregar Partido Político"
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
   Begin Proyecto1.ButtonOffice cmdBuscar 
      Height          =   435
      Left            =   6510
      TabIndex        =   9
      Top             =   60
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   767
      BackColor       =   33023
      Caption         =   "&Buscar"
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
   Begin Proyecto1.ButtonOffice cmdUnload 
      Height          =   405
      Left            =   5940
      TabIndex        =   10
      Top             =   6570
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   714
      BackColor       =   33023
      Caption         =   "Cerrar Ventana"
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
Attribute VB_Name = "frmEditarPartidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lista() As String
Dim ListaIds() As String

Private Sub cmdAgregarPartido_Click()
Dim X As String
Dim Consulta As String
Dim rsTemp As ADODB.Recordset
X = txtAddPartido.Text
If Trim(X) <> "" Then
    Set rsTemp = New ADODB.Recordset
    SetearRs "SELECT Agrupación_Política FROM Grupos", rsTemp
    While Not rsTemp.EOF
        If Trim(rsTemp.Fields(0)) = Trim(X) Then
            MsgBox "Este partido ya existe, no es posible agregarlo", vbCritical
            rsTemp.Close
            Set rsTemp = Nothing
            Exit Sub
        End If
        rsTemp.MoveNext
    Wend
    rsTemp.Close
    Set rsTemp = Nothing
    Consulta = "INSERT INTO Grupos(Agrupación_Política) VALUES ('" & X & "')"
    EjecutarSQL (Consulta)
    MsgBox "Agregado con éxito!", vbInformation
    txtBusqueda.Text = ""
    cmdLimpiar_Click
    ArmaListas
Else
    MsgBox "Nombre de partido inválido", vbCritical
End If
txtAddPartido.Text = ""
End Sub
Private Sub cmdBuscar_Click()
ArmaListas
End Sub
Private Sub cmdGrabar_Click()
Dim i As Integer
Dim Consulta As String
If txtPartido.Text <> "" Then
    If lstPartidos.ListIndex <> -1 Then
        i = MsgBox("¿Confirma que desea cambiar el partido '" & lstPartidos.List(lstPartidos.ListIndex) & "' por '" & txtPartido.Text & "'?", vbYesNo)
        If i = vbYes Then
            Consulta = "UPDATE Grupos SET Agrupación_Política = '" & txtPartido.Text & "' WHERE id_grupos = " & ListaIds(lstPartidos.ListIndex)
            EjecutarSQL (Consulta)
            MsgBox "Cambiado con éxito!", vbInformation
            txtBusqueda.Text = ""
            cmdLimpiar_Click
            ArmaListas
        End If
    Else
        Call MsgBox("No tiene ningún Bloque seleccionado", vbCritical)
    End If
Else
    MsgBox "Nombre del partido en blanco", vbCritical
    txtPartido.SetFocus
End If
End Sub
Private Sub cmdLimpiar_Click()
txtPartido.Text = ""
End Sub
Private Sub cmdUnload_Click()
Unload Me
End Sub
Private Sub Form_Load()
ArmaListas
End Sub
Private Sub ArmaListas()
Dim rsTemp As ADODB.Recordset
Dim i As Integer
i = -1
Set rsTemp = New ADODB.Recordset
SetearRs "SELECT id_grupos,Agrupación_Política FROM Grupos WHERE Agrupación_Política LIKE '" & txtBusqueda.Text & "%' ORDER BY Agrupación_Política", rsTemp
If rsTemp.RecordCount > 0 Then
    ReDim Lista(0 To rsTemp.RecordCount - 1)
    ReDim ListaIds(0 To rsTemp.RecordCount - 1)
    While Not rsTemp.EOF
        i = i + 1
        ListaIds(i) = rsTemp.Fields("id_grupos")
        Lista(i) = rsTemp.Fields("Agrupación_Política")
        rsTemp.MoveNext
    Wend
    rsTemp.Close
    Set rsTemp = Nothing
    lstPartidos.Clear
    For i = LBound(Lista) To UBound(Lista)
        lstPartidos.AddItem Lista(i)
    Next i
    If lstPartidos.Enabled = False Then
        lstPartidos.Enabled = True
    End If
    If frmOpt.Enabled = False Then
        frmOpt.Enabled = True
    End If
    txtPartido.Text = ""
Else
    txtPartido.Text = ""
    lstPartidos.Clear
    lstPartidos.Enabled = False
    frmOpt.Enabled = False
End If
End Sub
Private Sub lstPartidos_Click()
If lstPartidos.ListIndex <> -1 Then
    txtPartido.Text = lstPartidos.List(lstPartidos.ListIndex)
End If
End Sub
Private Sub txtBusqueda_Change()
ArmaListas
End Sub
