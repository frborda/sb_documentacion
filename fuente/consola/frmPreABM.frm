VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmPreABM 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modificación de Datos"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ButtonOffice cmdLegisladores 
      Height          =   495
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "Modificar / Ver &Estado de Diputados"
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
   Begin Proyecto1.ButtonOffice cmdBloques 
      Height          =   495
      Left            =   150
      TabIndex        =   1
      Top             =   750
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "Modificar / Ver &Bloques Políticos"
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
   Begin Proyecto1.ButtonOffice cmdEditarDistritos 
      Height          =   495
      Left            =   150
      TabIndex        =   2
      Top             =   1350
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "Modificar / Ver &Partidos Políticos"
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
Attribute VB_Name = "frmPreABM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonOffice1_Click()

End Sub
Private Sub cmdBloques_Click()
Me.Enabled = False
frmEditarBloques.Show vbModal, Me
Me.Enabled = True
End Sub

Private Sub cmdEditarDistritos_Click()
Me.Enabled = False
frmEditarPartidos.Show vbModal, Me
Me.Enabled = True
End Sub

Private Sub cmdLegisladores_Click()
    Me.Enabled = False
    If PermisosTotales.ABMLegisladores = 1 Then
        frmHistorico.Show vbModal
    Else
        MsgBox "El usuario no tiene permisos para realizar esta tarea", vbInformation + vbOKOnly
    End If
    Me.Enabled = True
End Sub
