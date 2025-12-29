VERSION 5.00
Begin VB.Form frmSetearConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración General SQV"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSegundosScreens 
      Height          =   285
      Left            =   4650
      TabIndex        =   40
      Top             =   5640
      Width           =   2205
   End
   Begin VB.CommandButton cmdSegundosScreens 
      Caption         =   "D"
      Height          =   300
      Left            =   6930
      TabIndex        =   39
      ToolTipText     =   "Restaurar valor por defecto"
      Top             =   5640
      Width           =   285
   End
   Begin VB.CommandButton cmdPathScreensDefault 
      Caption         =   "D"
      Height          =   300
      Left            =   6930
      TabIndex        =   37
      ToolTipText     =   "Restaurar valor por defecto"
      Top             =   5280
      Width           =   285
   End
   Begin VB.TextBox txtPathScreens 
      Height          =   285
      Left            =   4650
      TabIndex        =   36
      Top             =   5280
      Width           =   2205
   End
   Begin VB.TextBox txtFileEnrolamiento 
      Height          =   285
      Left            =   4650
      TabIndex        =   34
      Top             =   4920
      Width           =   2205
   End
   Begin VB.CommandButton Command10 
      Caption         =   "D"
      Height          =   300
      Left            =   6930
      TabIndex        =   33
      ToolTipText     =   "Restaurar valor por defecto"
      Top             =   4920
      Width           =   285
   End
   Begin VB.TextBox txtDirEnrolamiento 
      Height          =   285
      Left            =   4650
      TabIndex        =   31
      Top             =   4560
      Width           =   2205
   End
   Begin VB.CommandButton Command9 
      Caption         =   "D"
      Height          =   300
      Left            =   6930
      TabIndex        =   30
      ToolTipText     =   "Restaurar valor por defecto"
      Top             =   4560
      Width           =   285
   End
   Begin VB.CommandButton Command8 
      Caption         =   "D"
      Height          =   300
      Left            =   6930
      TabIndex        =   29
      ToolTipText     =   "Restaurar valor por defecto"
      Top             =   4200
      Width           =   285
   End
   Begin VB.TextBox txtPathSb 
      Height          =   285
      Left            =   4650
      TabIndex        =   28
      Top             =   4200
      Width           =   2205
   End
   Begin VB.CommandButton Command7 
      Caption         =   "D"
      Height          =   300
      Left            =   6930
      TabIndex        =   27
      ToolTipText     =   "Restaurar valor por defecto"
      Top             =   3840
      Width           =   285
   End
   Begin VB.TextBox txtPathSqv 
      Height          =   285
      Left            =   4650
      TabIndex        =   26
      Top             =   3840
      Width           =   2205
   End
   Begin VB.CommandButton Command6 
      Caption         =   "D"
      Height          =   300
      Left            =   6930
      TabIndex        =   25
      ToolTipText     =   "Restaurar valor por defecto"
      Top             =   3480
      Width           =   285
   End
   Begin VB.TextBox txtBancasTotales 
      Height          =   285
      Left            =   4650
      TabIndex        =   24
      Top             =   3480
      Width           =   2205
   End
   Begin VB.CommandButton Command5 
      Caption         =   "D"
      Height          =   300
      Left            =   6930
      TabIndex        =   23
      ToolTipText     =   "Restaurar valor por defecto"
      Top             =   3120
      Width           =   285
   End
   Begin VB.TextBox txtFinOperacion 
      Height          =   285
      Left            =   4650
      TabIndex        =   22
      Top             =   3120
      Width           =   2205
   End
   Begin VB.CommandButton Command4 
      Caption         =   "D"
      Height          =   300
      Left            =   6930
      TabIndex        =   21
      ToolTipText     =   "Restaurar valor por defecto"
      Top             =   2760
      Width           =   285
   End
   Begin VB.TextBox txtInicioOperacion 
      Height          =   285
      Left            =   4650
      TabIndex        =   20
      Top             =   2760
      Width           =   2205
   End
   Begin VB.CommandButton Command3 
      Caption         =   "D"
      Height          =   300
      Left            =   6930
      TabIndex        =   19
      ToolTipText     =   "Restaurar valor por defecto"
      Top             =   2400
      Width           =   285
   End
   Begin VB.TextBox txtReinicioSQV 
      Height          =   285
      Left            =   4650
      TabIndex        =   18
      Top             =   2400
      Width           =   2205
   End
   Begin VB.CommandButton Command2 
      Caption         =   "D"
      Height          =   300
      Left            =   6930
      TabIndex        =   17
      ToolTipText     =   "Restaurar valor por defecto"
      Top             =   2040
      Width           =   285
   End
   Begin VB.TextBox txtTiempoPaseLista 
      Height          =   285
      Left            =   4650
      TabIndex        =   16
      Top             =   2040
      Width           =   2205
   End
   Begin VB.CommandButton Command1 
      Caption         =   "D"
      Height          =   300
      Left            =   6930
      TabIndex        =   15
      ToolTipText     =   "Restaurar valor por defecto"
      Top             =   1680
      Width           =   285
   End
   Begin VB.TextBox txtMiembrosTotales 
      Height          =   285
      Left            =   4650
      TabIndex        =   14
      Top             =   1680
      Width           =   2205
   End
   Begin VB.PictureBox Picture1 
      Height          =   915
      Left            =   180
      ScaleHeight     =   855
      ScaleWidth      =   1245
      TabIndex        =   2
      Top             =   50
      Width           =   1300
      Begin VB.CommandButton Grabar 
         Caption         =   "&Grabar"
         Height          =   855
         Left            =   0
         Picture         =   "frmSetearConfig.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   915
      Left            =   5880
      ScaleHeight     =   855
      ScaleWidth      =   1245
      TabIndex        =   0
      Top             =   60
      Width           =   1300
      Begin VB.CommandButton Salir 
         Caption         =   "&Salir"
         Height          =   855
         Left            =   0
         Picture         =   "frmSetearConfig.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Segundos entre cada captura de pantalla :"
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
      Left            =   80
      TabIndex        =   41
      Top             =   5640
      Width           =   4500
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Path carpeta de capturas de pantalla :"
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
      Left            =   80
      TabIndex        =   38
      Top             =   5280
      Width           =   4500
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Path archivo intercambio de enrolamiento : "
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
      Left            =   150
      TabIndex        =   35
      Top             =   4920
      Width           =   4500
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Path directorio Imágenes y datos de enrolamiento : "
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
      Left            =   150
      TabIndex        =   32
      Top             =   4560
      Width           =   4500
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Path ejecutable  Servidor de Bancas : "
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
      Left            =   150
      TabIndex        =   13
      Top             =   4200
      Width           =   4500
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Path ejecutable SQV Server : "
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
      Left            =   150
      TabIndex        =   12
      Top             =   3840
      Width           =   4500
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Cantidad de Bancas Totales : "
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
      Left            =   150
      TabIndex        =   11
      Top             =   3480
      Width           =   4500
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Tiempo requerido para fin de operación : "
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
      Left            =   150
      TabIndex        =   10
      Top             =   3120
      Width           =   4500
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Tiempo requerido para inicio de operación : "
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
      Left            =   150
      TabIndex        =   9
      Top             =   2760
      Width           =   4500
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Tiempo de Espera para reinicio de SQV Server : "
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
      Left            =   150
      TabIndex        =   8
      Top             =   2400
      Width           =   4500
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Tiempo de espera de pase de lista : "
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
      Left            =   150
      TabIndex        =   7
      Top             =   2040
      Width           =   4500
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Miembros Totales del Cuerpo : "
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
      Left            =   150
      TabIndex        =   6
      Top             =   1680
      Width           =   4500
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "La modificación incorrecta de los datos expuestos puede alterar el normal funcionamiento del sistema SQV."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   525
      Left            =   180
      TabIndex        =   5
      Top             =   1020
      Width           =   7005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Datos de configuración que no requieren mantenimiento. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   150
      Width           =   4275
   End
End
Attribute VB_Name = "frmSetearConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Rs As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1

Private Sub cmdPathScreensDefault_Click()
txtPathScreens.Text = "C:\screens\"
End Sub

Private Sub cmdSegundosScreens_Click()
txtSegundosScreens.Text = 5
End Sub
Private Sub Command1_Click()
    txtMiembrosTotales.Text = "70"
End Sub

Private Sub Command10_Click()
 txtFileEnrolamiento = "bdExportEnrolamiento.mdb"
End Sub

Private Sub Command2_Click()
    txtTiempoPaseLista.Text = "000010"
End Sub

Private Sub Command3_Click()
    txtReinicioSQV.Text = "000015"
End Sub

Private Sub Command4_Click()
    txtInicioOperacion.Text = "5"
End Sub

Private Sub Command5_Click()
    txtFinOperacion.Text = "3"
End Sub

Private Sub Command6_Click()
    txtBancasTotales.Text = "71"
End Sub

Private Sub Command7_Click()
    txtPathSqv.Text = "e:\exes_sqv\sqvservidor.exe"
End Sub

Private Sub Command8_Click()
    txtPathSb.Text = "e:\exes_sqv\servidorB1.exe"
End Sub

Private Sub Command9_Click()
 txtDirEnrolamiento = "c:\temp\"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call SetearRs
    Call MostrarDatos
End Sub

Private Sub MostrarDatos()
    With Rs
        txtMiembrosTotales.Text = .Fields("Cantidad_de_Legisladores").Value
        txtTiempoPaseLista.Text = .Fields("Tiempo_Espera_Pase_de_Lista").Value
        txtReinicioSQV.Text = .Fields("Tiempo_Espera_Reinicio_Server").Value
        txtInicioOperacion.Text = .Fields("Segundos_de_Inicio_Operacion").Value
        txtFinOperacion.Text = .Fields("Segundos_de_fin_Operacion").Value
        txtBancasTotales.Text = .Fields("Cantidad_de_Bancas").Value
        txtPathSqv.Text = .Fields("Ejecutable_SQV").Value
        txtPathSb.Text = .Fields("Ejecutable_SB").Value
        txtDirEnrolamiento.Text = .Fields("directorio_enrolamiento").Value
        txtFileEnrolamiento.Text = .Fields("archivo_enrolamiento").Value
        If (IsNull(.Fields("Carpeta_Screens")) = False) Then
            txtPathScreens.Text = .Fields("Carpeta_Screens")
        Else
            txtPathScreens.Text = ""
        End If
        If (IsNull(.Fields("Segundos_Screens")) = False) Then
            txtSegundosScreens.Text = .Fields("Segundos_Screens")
        Else
            txtSegundosScreens.Text = ""
        End If
    End With
End Sub
Private Sub SetearRs()
    Dim strSql As String
    'Public strDirectorioFotos As String
    strSql = "SELECT * FROM Config"
    Set Rs = New ADODB.Recordset
    SetearRsW strSql, Rs
    Rs.MoveFirst
End Sub

Private Sub Grabar_Click()

    Dim xValor As Long
    
    ' ------------------------------------------------------------
    ' Validar datos ingresados por usuario
    ' ------------------------------------------------------------
    ' Miembros totales
    If Not IsNumeric(txtMiembrosTotales.Text) Or Trim(txtMiembrosTotales.Text) = "" Or Trim(txtMiembrosTotales.Text) = "0" Then
        MsgBox "Valor de miembros totales incorrecto", vbCritical + vbOKOnly, "ERROR FATAL COMETIDO POR USUARIO"
        txtMiembrosTotales.SetFocus
        Exit Sub
    Else
        xValor = Int(txtMiembrosTotales.Text)
    End If
    ' Tiempo de pase de lista
    If Not IsNumeric(txtTiempoPaseLista.Text) Or Trim(txtTiempoPaseLista.Text) = "" Or Trim(txtTiempoPaseLista.Text) = "0" Then
        MsgBox "Valor de tiempo de espera de pase de lista incorrecto", vbCritical + vbOKOnly, "ERROR FATAL COMETIDO POR USUARIO"
        txtTiempoPaseLista.SetFocus
        Exit Sub
    Else
        xValor = Int(txtTiempoPaseLista.Text)
        txtTiempoPaseLista.Text = Format(xValor, "000000")
    End If
    ' Tiempo de reinicio de Server
    If Not IsNumeric(txtReinicioSQV.Text) Or Trim(txtReinicioSQV.Text) = "" Or Trim(txtReinicioSQV.Text) = "0" Then
        MsgBox "Valor de tiempo de reinicio de server incorrecto", vbCritical + vbOKOnly, "ERROR FATAL COMETIDO POR USUARIO"
        txtReinicioSQV.SetFocus
        Exit Sub
    Else
        xValor = Int(txtReinicioSQV.Text)
        txtReinicioSQV.Text = Format(xValor, "000000")
    End If
    ' Segundos de reinicio de Operacion
    If Not IsNumeric(txtInicioOperacion.Text) Or Trim(txtInicioOperacion.Text) = "" Or Trim(txtInicioOperacion.Text) = "0" Then
        MsgBox "Valor de tiempo de reinicio de operación incorrecto", vbCritical + vbOKOnly, "ERROR FATAL COMETIDO POR USUARIO"
        txtInicioOperacion.SetFocus
        Exit Sub
    End If
    ' Cantidad de Bancas
    If Not IsNumeric(txtBancasTotales.Text) Or Trim(txtBancasTotales.Text) = "" Or Trim(txtBancasTotales.Text) = "0" Then
        MsgBox "Valor de cantidad de bancas incorrecto", vbCritical + vbOKOnly, "ERROR FATAL COMETIDO POR USUARIO"
        txtBancasTotales.SetFocus
        Exit Sub
    End If
    
    ' ------------------------------------------------------------
    ' Grabar en base de datos
    ' ------------------------------------------------------------
    With Rs
        .Fields("Cantidad_de_Legisladores").Value = txtMiembrosTotales.Text
        .Fields("Tiempo_Espera_Pase_de_Lista").Value = txtTiempoPaseLista.Text
        .Fields("Tiempo_Espera_Reinicio_Server").Value = txtReinicioSQV.Text
        .Fields("Segundos_de_Inicio_Operacion").Value = txtInicioOperacion.Text
        .Fields("Segundos_de_fin_Operacion").Value = txtFinOperacion.Text
        .Fields("Cantidad_de_Bancas").Value = txtBancasTotales.Text
        .Fields("Ejecutable_SQV").Value = txtPathSqv.Text
        .Fields("Ejecutable_SB").Value = txtPathSb.Text
        .Fields("directorio_enrolamiento").Value = txtDirEnrolamiento
        .Fields("archivo_enrolamiento").Value = txtFileEnrolamiento
        If (txtPathScreens.Text <> "") Then
            .Fields("Carpeta_Screens").Value = txtPathScreens.Text
        Else
            MsgBox ("La carpeta de screens no puede estar en blanco")
        End If
        If (txtSegundosScreens.Text <> "") Then
            .Fields("Segundos_Screens").Value = Int(txtSegundosScreens.Text)
        Else
            MsgBox ("La carpeta de screens no puede estar en blanco")
        End If
        .Update
    End With


End Sub

Private Sub Salir_Click()
    Unload Me
End Sub

