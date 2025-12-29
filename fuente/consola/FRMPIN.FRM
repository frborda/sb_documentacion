VERSION 5.00
Begin VB.Form frmPIN 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PIN de Legislador"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6045
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1605
      Left            =   90
      TabIndex        =   4
      Top             =   1020
      Width           =   5865
      Begin VB.TextBox txtNuevoPin2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1050
         Width           =   1470
      End
      Begin VB.CommandButton Command1 
         Caption         =   "D"
         Height          =   285
         Left            =   3300
         TabIndex        =   9
         ToolTipText     =   "Restaurar valor por defecto (DNI de legislador)"
         Top             =   690
         Width           =   315
      End
      Begin VB.TextBox txtNuevoPin 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   720
         Width           =   1470
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Reingrese PIN : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "PIN :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1410
      End
      Begin VB.Label lblLegislador 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   4005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Legislador :"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   1410
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   915
      Left            =   4650
      ScaleHeight     =   855
      ScaleWidth      =   1245
      TabIndex        =   2
      Top             =   50
      Width           =   1305
      Begin VB.CommandButton Salir 
         Caption         =   "&Salir"
         Height          =   855
         Left            =   0
         Picture         =   "frmPIN.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   915
      Left            =   90
      ScaleHeight     =   855
      ScaleWidth      =   1245
      TabIndex        =   0
      Top             =   50
      Width           =   1305
      Begin VB.CommandButton Grabar 
         Caption         =   "&Grabar"
         Height          =   855
         Left            =   0
         Picture         =   "frmPIN.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmPIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strIdLegislador  As String
Private strDniLegislador As String
Private WithEvents RsPin As ADODB.Recordset
Attribute RsPin.VB_VarHelpID = -1
Private Sub BuscarLegislador()
    Dim strSql As String
    Set RsPin = New ADODB.Recordset
    strSql = "SELECT * FROM LEgisladores WHERE Id= '" & strIdLegislador & "'"
    SetearRsW strSql, RsPin
    With RsPin
        If .RecordCount = 0 Then
            MsgBox "Error de datos. No se encontró al legislador Id " & strIdLegislador, vbInformation + vbOKOnly
            .Close
            Unload Me
        End If
        If .RecordCount > 1 Then
            MsgBox "Error de datos. Existen " & Trim(Str(RsPin.RecordCount)) & " Legisladores con el ID  " & strIdLegislador & ". Verifique los datos de legisladores registrados.", vbInformation + vbOKOnly
            .Close
            Unload Me
        End If
        If .RecordCount = 1 Then
            lblLegislador.Caption = .Fields("Apellido").Value & ", " & .Fields("Nombre").Value
            If IsNull(.Fields("Pin").Value) And Not IsNull(.Fields("Dni").Value) Then
                .Fields("Pin").Value = Encripta.EncryptString(.Fields("Dni").Value)
                .Update
            End If
        End If
        If Trim(.Fields("pin")) = "" Then
            MsgBox "El legislador actual no tiene pin asignado", vbInformation + vbOKOnly
        End If
        If Trim(.Fields("pin")) = "" Then
            MsgBox "El legislador actual no tiene DNI asignado", vbInformation + vbOKOnly
        End If
    End With
End Sub
Private Sub Command1_Click()
    txtNuevoPin.Text = strDniLegislador
    txtNuevoPin2.Text = strDniLegislador
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Public Property Let ID_Legislador(ByVal vNewValue As Variant)
    strIdLegislador = vNewValue
End Property
Private Sub Form_Load()
    Call BuscarLegislador
    txtNuevoPin.Text = Encripta.EncryptString(RsPin.Fields("pin").Value)
End Sub
Private Sub Grabar_Click()
    Dim strSql        As String
    Dim RsTemp        As ADODB.Recordset
    Dim BlPinNoValido As Boolean
    Dim strTemp       As String
    
    BlPinNoValido = False
    ' Validar entrada de usuario
    If Trim(Len(txtNuevoPin.Text)) < 5 Then
        MsgBox "El PIN de legislador debe tener mas de 5 dígitos", vbInformation + vbOKCancel
        txtNuevoPin.SelStart = 0
        txtNuevoPin.SelLength = Len(txtNuevoPin.Text)
        txtNuevoPin.SetFocus
        Exit Sub
    End If
    If txtNuevoPin.Text <> txtNuevoPin2.Text Then
        MsgBox "Error en el ingreso de PIN. Reintente la operacion", vbInformation + vbOKOnly
        txtNuevoPin.Text = ""
        txtNuevoPin2.Text = ""
        txtNuevoPin.SetFocus
        Exit Sub
    End If
    'rellena con ceros a izquierda
    txtNuevoPin.Text = CerosIzquierda(txtNuevoPin.Text, 8)
    txtNuevoPin2.Text = txtNuevoPin.Text
    ' Verificar que no se duplique el PIN
    Set RsTemp = New ADODB.Recordset
    strSql = "SELECT Pin FROM Legisladores WHERE  id <> '" & Trim(RsPin.Fields("id").Value) & "'"
    SetearRs strSql, RsTemp
    With RsTemp
        If .RecordCount > 0 Then
            .MoveFirst
            While Not .EOF
                 If Not IsNull(.Fields("PIN").Value) Then
                     strTemp = Trim(Encripta.EncryptString(.Fields("PIN").Value))
                     If LCase(Trim(strTemp)) = LCase(Trim(txtNuevoPin.Text)) Then
                         BlPinNoValido = True
                         .MoveLast
                     End If
                 End If
                .MoveNext
            Wend
        End If
    End With
    
    If BlPinNoValido Then
        MsgBox "El PIN ingresado no es válido. Reintente la operacion", vbInformation + vbOKOnly
        RsTemp.Close
        Set RsTemp = Nothing
        Exit Sub
    Else
        RsTemp.Close
        Set RsTemp = Nothing
    End If
    
    With RsPin
        .Fields("pin").Value = Encripta.EncryptString(txtNuevoPin.Text)
        .Update
    End With
    MsgBox "El PIN del legislador se ha actualizado con éxito"
End Sub
Private Function CerosIzquierda(strText As String, nLong As Long) As String
    If nLong > Len(strText) Then
        CerosIzquierda = Left(String(nLong - Len(strText), "0") & strText, nLong)
    Else
        CerosIzquierda = Right(strText, nLong)
    End If
End Function
Private Sub Salir_Click()
    Unload Me
End Sub
Public Property Let DNI(ByVal vNewValue As Variant)
    strDniLegislador = vNewValue
End Property

Private Sub txtNuevoPin_GotFocus()
    With txtNuevoPin
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNuevoPin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Exit Sub
    End If
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub
