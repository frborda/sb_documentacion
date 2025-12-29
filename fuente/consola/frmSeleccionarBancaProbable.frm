VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmSeleccionarBancaProbable 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Selección de Banca Probable"
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdAplicar 
      Height          =   345
      Left            =   90
      TabIndex        =   3
      Top             =   900
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      BackColor       =   12230304
      Caption         =   "&Aplicar"
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
   Begin VB.ComboBox cmbBancasProbables 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   2
      Text            =   " - Seleccione una banca probable -"
      Top             =   480
      Width           =   3495
   End
   Begin Proyecto1.ButtonOffice cmdCancelar 
      Height          =   345
      Left            =   1950
      TabIndex        =   4
      Top             =   900
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      BackColor       =   12230304
      Caption         =   "&Cancelar"
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
   Begin VB.Label lblDiputado 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   285
      Left            =   900
      TabIndex        =   1
      Top             =   90
      Width           =   2715
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Edición:"
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
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   765
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   4
      Height          =   1395
      Left            =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmSeleccionarBancaProbable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbBancasProbables_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAplicar_Click
End If
End Sub
Private Sub cmdAplicar_Click()
Dim Asignado As Boolean
Asignado = False
If cmbBancasProbables.ListIndex <> -1 Then
    frmBancasProbables.BancaAAsignar = Val(cmbBancasProbables.List(cmbBancasProbables.ListIndex))
    Unload Me
Else
    If cmbBancasProbables.Text = " - Seleccione una banca probable -" Then
        MsgBox "Debe seleccionar una banca probable!", vbCritical, "Alerta"
    Else
        If Not IsNumeric(cmbBancasProbables.Text) Then
            MsgBox "Sólo se permiten ingresar números!", vbInformation
        Else
            Dim i As Integer
            For i = 0 To cmbBancasProbables.ListCount - 1
                If Trim(cmbBancasProbables.Text) = cmbBancasProbables.List(i) Then
                    Asignado = True
                    frmBancasProbables.BancaAAsignar = Val(cmbBancasProbables.List(i))
                    i = cmbBancasProbables.ListCount - 1
                End If
            Next i
            If Asignado = True Then
                Unload Me
            Else
                MsgBox "Esa banca ya está en uso", vbInformation
                cmbBancasProbables.Text = " - Seleccione una banca probable -"
                cmbBancasProbables.SetFocus
            End If
        End If
    End If
End If
End Sub
Private Sub cmdCancelar_Click()
frmBancasProbables.BancaAAsignar = -1
Unload Me
End Sub
