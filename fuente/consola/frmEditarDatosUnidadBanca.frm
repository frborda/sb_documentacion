VERSION 5.00
Begin VB.Form frmEditarDatosUnidadBanca 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Unidad de Banca"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   915
      Left            =   3240
      ScaleHeight     =   855
      ScaleWidth      =   1245
      TabIndex        =   10
      Top             =   50
      Width           =   1300
      Begin VB.CommandButton Salir 
         Caption         =   "&Salir"
         Height          =   855
         Left            =   0
         Picture         =   "frmEditarDatosUnidadBanca.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   915
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1245
      TabIndex        =   9
      Top             =   50
      Width           =   1305
      Begin VB.CommandButton Grabar 
         Caption         =   "&Grabar"
         Height          =   855
         Left            =   0
         Picture         =   "frmEditarDatosUnidadBanca.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.Frame frameBanca 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Left            =   60
      TabIndex        =   0
      Top             =   1020
      Width           =   4455
      Begin VB.TextBox txtIdString 
         Height          =   1365
         Left            =   1170
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1800
         Width           =   3105
      End
      Begin VB.TextBox txtComentario 
         Height          =   285
         Left            =   1170
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1350
         Width           =   3105
      End
      Begin VB.TextBox txtPuerto 
         Height          =   285
         Left            =   1170
         MaxLength       =   50
         TabIndex        =   6
         Top             =   930
         Width           =   3105
      End
      Begin VB.TextBox txtIp 
         Height          =   285
         Left            =   1170
         MaxLength       =   50
         TabIndex        =   5
         Top             =   540
         Width           =   3105
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Id String : "
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   1830
         Width           =   900
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Comentario : "
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   1460
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Puerto : "
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   1000
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "IP : "
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   540
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmEditarDatosUnidadBanca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strBancaActual As String
Private WithEvents Rs  As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1

Private Sub Grabar_Click()
    Dim i As Integer
    Dim Repetido As Boolean
    Dim BancaRep As String
    Dim Validado As Boolean
    Dim xIP() As String
    xIP = Split(txtIp.Text, ".")
    Validado = True
    If UBound(xIP) <> 3 Then
        Validado = False
        MsgBox "La IP no tiene tres puntos", vbInformation
    End If
    If Validado Then
        For i = LBound(xIP) To UBound(xIP)
            If IsNumeric(xIP(i)) Then
                If Val(xIP(i)) > 255 Or Val(xIP(i)) < 0 Then
                    Validado = False
                    MsgBox "La IP tiene un rango mayor a 255 o es menor a 0 (" & i + 1 & ")"
                End If
            Else
                Validado = False
                MsgBox "La IP no es numérica", vbCritical
            End If
        Next i
    End If
    If Validado Then
        Repetido = False
        For i = 1 To (frmConfigurarUnidadBanca.Grilla.Rows - 1)
            If Trim(frmConfigurarUnidadBanca.Grilla.TextMatrix(i, 1)) = Trim(txtIp.Text) Then
                BancaRep = frmConfigurarUnidadBanca.Grilla.TextMatrix(i, 0)
                Repetido = True
                i = frmConfigurarUnidadBanca.Grilla.Rows - 1
            End If
        Next i
        If Repetido = True Then
            MsgBox ("La IP ya existe en la banca " & BancaRep)
        Else
            ' ------------------------------------------------------------------------------------
            ' validar datos ingresados por usuarios
            ' ------------------------------------------------------------------------------------
            If Trim(txtIp.Text) = "" Then
                MsgBox "El número de IP de la banca no puede ser nulo", vbInformation + vbOKOnly
                txtIp.SetFocus
                Exit Sub
            End If
            If Trim(txtPuerto.Text) = "" Then
                MsgBox "El número de puerto de la banca no puede ser nulo", vbInformation + vbOKOnly
                txtPuerto.SetFocus
                Exit Sub
            End If
            If Trim(txtIdString.Text) = "" Then
                MsgBox "Id String de la banca no puede ser nulo", vbInformation + vbOKOnly
                txtIdString.SetFocus
                Exit Sub
            End If
            Dim xCantRows As Integer
            xCantRows = frmConfigurarUnidadBanca.Grilla.Rows
            'MsgBox xCantRows
            With Rs
                .Fields("ip").Value = Trim(txtIp.Text)
                .Fields("puerto").Value = Trim(txtPuerto.Text)
                .Fields("comentario").Value = Trim(txtComentario.Text)
                .Fields("idstring").Value = txtIdString.Text
                .Update
            End With
        End If
    End If
    Unload Me
End Sub
Private Sub MostrarDatos()
    With Rs
        txtIp.Text = .Fields("ip").Value
        txtPuerto.Text = .Fields("puerto").Value
        txtComentario.Text = "" & .Fields("comentario").Value
        txtIdString.Text = .Fields("idstring").Value
    End With
End Sub
Private Sub Form_Load()
    Call SetearRecordSet
    Call MostrarDatos
End Sub
Private Sub SetearRecordSet()
    Dim strSql As String
    Set Rs = New ADODB.Recordset
    strSql = "SELECT * FROM BancasIP WHERE (BancaNumero = '" & strBancaActual & "')"
    SetearRsW strSql, Rs
    If Rs.RecordCount = 0 Then
        MsgBox "No se localizo la banca en cuestion", vbInformation
        Unload Me
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Public Property Let Banca(ByVal vNewValue As Variant)
    strBancaActual = vNewValue
    strBancaActual = Trim(strBancaActual)
    frameBanca.Caption = "Banca " & strBancaActual
End Property

Private Sub Salir_Click()
    Unload Me
End Sub
