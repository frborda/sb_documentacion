VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAltaLegislador 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enrolar Legislador"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   7575
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         TabIndex        =   32
         Top             =   3240
         Width           =   2200
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   31
         Top             =   2760
         Width           =   2200
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   29
         Top             =   2280
         Width           =   2200
      End
      Begin VB.ComboBox cmbBloque 
         Height          =   315
         Left            =   1845
         TabIndex        =   23
         Top             =   3795
         Width           =   2200
      End
      Begin VB.Frame Frame2 
         Caption         =   "Banca Asignada : "
         Height          =   1280
         Left            =   4080
         TabIndex        =   19
         Top             =   2040
         Width           =   3255
         Begin VB.TextBox txtBanca 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   8
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton cmdAsignar 
            Caption         =   "Asignar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1920
            TabIndex        =   9
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label Label9 
            Caption         =   "Banca : "
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.CheckBox chkEsLegislador 
         Caption         =   "Es Legislador"
         Height          =   255
         Left            =   4200
         TabIndex        =   7
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CheckBox chkPersonalMantenimiento 
         Caption         =   "Personal de Mantenimiento"
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtDepartamento 
         Height          =   285
         Left            =   1725
         TabIndex        =   4
         Top             =   4200
         Width           =   2200
      End
      Begin MSComCtl2.DTPicker dtFechaNacimiento 
         Height          =   315
         Left            =   5760
         TabIndex        =   5
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52166657
         CurrentDate     =   38027
      End
      Begin VB.TextBox txtApellido 
         Height          =   285
         Left            =   1725
         TabIndex        =   1
         Top             =   675
         Width           =   2200
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1725
         TabIndex        =   2
         Top             =   255
         Width           =   2200
      End
      Begin VB.ComboBox cmbSexo 
         Height          =   315
         Left            =   1725
         TabIndex        =   3
         Top             =   1170
         Width           =   2200
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52166657
         CurrentDate     =   38027
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Fin de mandato : "
         Height          =   405
         Left            =   240
         TabIndex        =   30
         Top             =   2760
         Width           =   1500
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio de mandato : "
         Height          =   405
         Left            =   240
         TabIndex        =   27
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha de inicio de actividad : "
         Height          =   405
         Left            =   480
         TabIndex        =   24
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Agrupación Política : "
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   3285
         Width           =   1500
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Departamento : "
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   4200
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha de Nacimiento : "
         Height          =   285
         Left            =   3915
         TabIndex        =   17
         Top             =   360
         Width           =   1785
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Bloque Político : "
         Height          =   285
         Left            =   195
         TabIndex        =   16
         Top             =   3795
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Apellido  : "
         Height          =   285
         Left            =   75
         TabIndex        =   15
         Top             =   675
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre : "
         Height          =   285
         Left            =   75
         TabIndex        =   14
         Top             =   285
         Width           =   1500
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Sexo : "
         Height          =   285
         Left            =   75
         TabIndex        =   13
         Top             =   1200
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   920
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   5445
      TabIndex        =   0
      Top             =   120
      Width           =   5505
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   855
         Left            =   4200
         TabIndex        =   11
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   855
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha de inicio de actividades : "
      Height          =   405
      Left            =   600
      TabIndex        =   26
      Top             =   3480
      Width           =   1140
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha de inicio de actividades : "
      Height          =   405
      Left            =   600
      TabIndex        =   25
      Top             =   3480
      Width           =   1140
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Asistente para enrolamiento de personas vinculadas a bancas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6090
      TabIndex        =   21
      Top             =   600
      Width           =   5025
   End
End
Attribute VB_Name = "frmAltaLegislador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents RsAltaLegislador As ADODB.Recordset
Attribute RsAltaLegislador.VB_VarHelpID = -1
Private WithEvents RsLegxBanca      As ADODB.Recordset
Attribute RsLegxBanca.VB_VarHelpID = -1
Private strIdLegislador             As String
Private strOperacion                As String
Private strApellido                 As String
Private strNombre                   As String
Private strHuella                   As String

Private Sexo                        As Integer ' 0 es Nena, 1 Es Nene
Private strBolquePolitico           As String
Private strAgrupacionPolitica       As String
Private strDepartamento             As String
Private strFechaNacimiento          As Date
Private blPersonalMantenimiento     As Integer  ' 1 es legislador, 0 es informatico de mantenimiento
Private blEsLegislador              As Integer  ' 1 es legislador, 0 es vicegobernador
Private strDni                      As String
Private xBanca                      As Long

Private strSql                      As String

Public Property Let Id(ByVal vNewValue As Variant)
    strIdLegislador = vNewValue
    strIdLegislador = Trim(strIdLegislador)
    txtID.Text = strIdLegislador
    Call BuscarLegisladorEnBD
    Call BuscarBancaAsignada
End Property
Public Property Let Apellido(ByVal vNewValue As Variant)
    strApellido = vNewValue
    txtApellido.Text = strApellido
End Property
Public Property Let Nombre(ByVal vNewValue As Variant)
    strNombre = vNewValue
    txtNombre.Text = strNombre
End Property
Public Property Let Huella(ByVal vNewValue As Variant)
    strHuella = vNewValue
End Property
Public Property Let Operacion(ByVal vNewValue As Variant)
    strOperacion = vNewValue
    strOperacion = LCase(Trim(strOperacion))
    If strOperacion <> "update" And strOperacion <> "addnew" Then
        MsgBox "Error de operación en proceso de enrolamiento", vbCritical + vbOKOnly
    End If
End Property

Private Sub CargarCombo()
    
    Dim strSql  As String
    Dim RsCombo As ADODB.Recordset
    Set RsCombo = New ADODB.Recordset
    ' ------------------------------------------------------------------------
    ' Combo sexo
    ' ------------------------------------------------------------------------
    cmbSexo.Clear
    cmbSexo.AddItem "Femenino"
    cmbSexo.AddItem "Masculino"
    ' ------------------------------------------------------------------------
    ' Combo Bloque Politico
    ' ------------------------------------------------------------------------
    strSql = "SELECT Bloque_Político FROM Bloques ORDER BY Bloque_Político"
    SetearRs strSql, RsCombo
    With RsCombo
        If .RecordCount > 0 Then
            cmbBloque.Clear
            .MoveFirst
            While Not .EOF
                cmbBloque.AddItem .Fields(0).Value
                .MoveNext
            Wend
            .Close
        End If
    End With
    Set RsCombo = Nothing
End Sub

Private Sub cmdAsignar_Click()
    
    Dim xBancaAsignada As Long
    Dim strSql         As String
    Dim xRespuesta     As Integer
    Dim RsLegxBanca    As ADODB.Recordset
    
    Set RsLegxBanca = New ADODB.Recordset
    xRespuesta = 0
    
    If Trim(txtBanca.Text) <> "" Then
        If Not IsNumeric(txtBanca.Text) Then
            MsgBox "La banca asignada debe ser un número entero entre las bancas disponibles", vbInformation + vbOKOnly, "Error de usuario"
            txtBanca.SelStart = 0
            txtBanca.SelLength = Len(txtBanca.Text)
            txtBanca.SetFocus
            Exit Sub
        Else
            xBancaAsignada = Int(txtBanca.Text)
        End If
    Else
        xBancaAsignada = -1 ' NULL: es decir, se desasigna la banca actual
    End If
    ' ------------------------------------------------------------------
    ' Grabar en base de datos
    ' ------------------------------------------------------------------
    strSql = "SELECT id, deskid FROM legisladores_activos WHERE id = '" & strIdLegislador & "'"
    SetearRsW strSql, RsLegxBanca
    With RsLegxBanca
        If (.RecordCount = 0) Or (.EOF = True) Then
            xRespuesta = MsgBox("¿Seguro que quiere modificar la asignación de bancas?" & Chr(10), vbQuestion + vbYesNo)
            If xRespuesta = vbYes Then
                .AddNew
                .Fields("id").Value = strIdLegislador
                .Fields("deskid").Value = xBancaAsignada
                .Update
            End If
        ElseIf .RecordCount = 1 Then
            If Not (xBancaAsignada = -1) Then
                xRespuesta = MsgBox("¿Seguro que quiere modificar la asignación de bancas?" & Chr(10), vbQuestion + vbYesNo)
                If xRespuesta = vbYes Then
                    .Fields("deskid").Value = xBancaAsignada
                End If
            Else
                xRespuesta = MsgBox("¿Seguro que quiere dejar al legislador " & UCase(Trim(txtApellido.Text)) & " sin banca asignada?" & Chr(10), vbQuestion + vbYesNo)
                If xRespuesta = vbYes Then
                    .Fields("deskid").Value = Null
                End If
            End If
            .Update
        Else
            MsgBox "El legislador tiene mas de una banca asignada.", vbInformation + vbOKOnly, "Error de coherencia de datos en base SQV"
        End If
        ' Mostrar en pantalla banca asignada
        If xBancaAsignada >= 0 Then
            txtBanca.Text = Trim(Str(xBancaAsignada))
        ElseIf xBancaAsignada = -1 Then
            txtBanca.Text = ""
        End If
        .Close
    End With
    Set RsLegxBanca = Nothing
End Sub

Private Sub cmdGrabar_Click()
    Dim RsLegis As ADODB.Recordset
    Dim strSql As String
    ' ---------------------------------------------------------------------------
    ' Validar informacion ingresada por usuario
    ' ---------------------------------------------------------------------------
    If Trim(txtID.Text) = "" Then
        MsgBox "El Id del legislador es un dato obligatorio", vbInformation, "Error de usuario en carga de datos"
        txtID.SetFocus
        Exit Sub
    End If
    If Trim(txtApellido.Text) = "" Then
        MsgBox "El apellido del legislador es un dato obligatorio", vbInformation, "Error de usuario en carga de datos"
        txtApellido.SetFocus
        Exit Sub
    End If
    If Trim(txtNombre.Text) = "" Then
        MsgBox "El nombre del legislador es un dato obligatorio", vbInformation, "Error de usuario en carga de datos"
        txtNombre.SetFocus
        Exit Sub
    End If
    ' ---------------------------------------------------------------------------
    ' Leer valores ingresados por usuario
    ' ---------------------------------------------------------------------------
    strIdLegislador = txtID.Text
    strApellido = txtApellido.Text
    strNombre = txtNombre.Text
    If Trim(LCase(cmbSexo.Text)) = "masculino" Then
        Sexo = 1
    ElseIf Trim(LCase(cmbSexo.Text)) = "femenino" Then
        Sexo = 0
    End If
    strBolquePolitico = Trim(cmbBloque.Text)
    strAgrupacionPolitica = Trim(txtAgrupacionPolitica.Text)
    strDepartamento = Trim(txtDepartamento.Text)
    strFechaNacimiento = dtFechaNacimiento.Value
    ' Tipo: 1 es Legislador, 0 es Mantenimiento
    If chkPersonalMantenimiento.Value = 0 Then
        blPersonalMantenimiento = 1
    ElseIf chkPersonalMantenimiento.Value = 1 Then
        blPersonalMantenimiento = 0
    End If
    If chkEsLegislador.Value = 0 Then
        blEsLegislador = 0
    ElseIf chkEsLegislador.Value = 1 Then
        blEsLegislador = 1
    End If
    strDni = Trim(txtDni.Text)
    'xBanca = Int(txtBanca.Text)
    
    ' ---------------------------------------------------------------------------
    ' Grabar en Base de datos
    ' ---------------------------------------------------------------------------
    strSql = "SELECT * FROM Legisladores WHERE Id = '" & strIdLegislador & "'"
    Set RsLegis = New ADODB.Recordset
    SetearRsW strSql, RsLegis
    With RsLegis
        If strOperacion = "addnew" Then
            .AddNew
            'strSql = "sp_insert_Legisladores('" & strIdLegislador & "', '" & strNombre & "', '" & strApellido & "', " & Sexo & ", '" & strFechaNacimiento & "', '" & strBolquePolitico & "', '" & strAgrupacionPolitica & "', '" & strDepartamento & "', '" & strHuella & "', " & blPersonalMantenimiento & ", " & blEsLegislador & ", '" & strDni & "')"
            ' SenteciaSQl strSql
        End If
        .Fields("id").Value = strIdLegislador
        .Fields("apellido").Value = strApellido
        .Fields("Nombre").Value = strNombre
        .Fields("Sexo").Value = Sexo
        .Fields("Fecha_Nacimiento").Value = strFechaNacimiento
        .Fields("Bloque_Politico").Value = strBolquePolitico
        .Fields("Grupo_Politico").Value = strAgrupacionPolitica
        .Fields("Departamento").Value = strDepartamento
        .Fields("Tipo").Value = blPersonalMantenimiento
        .Fields("Es_Legislador").Value = blEsLegislador
        .Fields("Dni").Value = strDni
        .Fields("Template1").Value = strHuella
        .Update
        .Close
    End With
    strOperacion = ""
    Set RsLegis = Nothing
    cmdAsignar.Enabled = True
End Sub

Private Sub BuscarLegisladorEnBD()
    'Call AbrirDB
    strSql = "SELECT * FROM Legisladores WHERE Id = '" & strIdLegislador & "'"
    SetearRs strSql, RsAltaLegislador
    With RsAltaLegislador
        If .RecordCount = 1 Then
            txtID.Text = .Fields("id").Value
            txtApellido.Text = .Fields("Apellido").Value
            txtNombre.Text = .Fields("Nombre").Value
            If IsNull(.Fields("Sexo").Value) Then
                cmbSexo.Text = ""
            ElseIf .Fields("Sexo").Value = 1 Then
                cmbSexo.Text = "Masculino"
            ElseIf .Fields("Sexo").Value = 0 Then
                cmbSexo.Text = "Femenino"
            End If
            If IsNull(.Fields("Bloque_Politico").Value) Then
                cmbBloque.Text = " "
            Else
                cmbBloque.Text = .Fields("Bloque_Politico").Value
            End If
            If IsNull(.Fields("Grupo_Politico").Value) Then
                txtAgrupacionPolitica.Text = ""
            Else
                txtAgrupacionPolitica.Text = .Fields("Grupo_Politico").Value
            End If
            If IsNull(.Fields("Departamento").Value) Then
                txtDepartamento.Text = ""
            Else
                txtDepartamento.Text = .Fields("Departamento").Value
            End If
            
            If Not IsNull(.Fields("Fecha_Nacimiento").Value) Then
                dtFechaNacimiento.Value = .Fields("fecha_nacimiento").Value
            Else
                dtFechaNacimiento.Value = "1/1/1900"
            End If
            If .Fields("Fecha_Nacimiento").Value = 0 Then
                chkPersonalMantenimiento.Value = 0
            Else
                chkPersonalMantenimiento.Value = 1
            End If
            If .Fields("Es_Legislador").Value = 1 Then
                chkEsLegislador.Value = 1
            Else
                chkEsLegislador.Value = 0
            End If
            If Not IsNull(.Fields("Dni").Value) Then
                txtDni = .Fields("Dni").Value
            Else
                txtDni = ""
            End If

            ' txtBanca.Text = 0
        End If
    End With
End Sub

Private Sub BuscarBancaAsignada()
    Dim strSql      As String
    
    Set RsLegxBanca = New ADODB.Recordset
    strSql = "SELECT id, deskid FROM legisladores_activos WHERE id = '" & strIdLegislador & "'"
    SetearRsW strSql, RsLegxBanca
    With RsLegxBanca
        If (.RecordCount = 0) Or (.EOF = True) Then
            txtBanca.Text = ""
            Exit Sub
        Else
            txtBanca.Text = .Fields("deskid").Value
        End If
    End With
    RsLegxBanca.Close
    Set RsLegxBanca = Nothing

End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set RsAltaLegislador = New ADODB.Recordset
    Call CargarCombo
    Call AbrirDB
End Sub

Private Sub txtAgrupacionPolitica_Change()

End Sub
