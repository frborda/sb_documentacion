VERSION 5.00
Begin VB.Form frmDefinirVoto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambiar el voto de la banca nº"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbstener 
      Caption         =   "Abstener"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   6
      Top             =   1320
      Width           =   2235
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "Vota por &No"
      Height          =   435
      Left            =   180
      TabIndex        =   4
      Top             =   720
      Width           =   2235
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "&Volver"
      Height          =   435
      Left            =   180
      TabIndex        =   2
      Top             =   1920
      Width           =   2235
   End
   Begin VB.CommandButton cmdSi 
      Caption         =   "Vota por &SÍ"
      Height          =   435
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2235
   End
   Begin VB.Label Label4 
      Caption         =   "Abstiene al diputado"
      Height          =   435
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   3675
   End
   Begin VB.Label Label3 
      Caption         =   "Asigna un voto negativo a la banca seleccionada"
      Height          =   435
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   3675
   End
   Begin VB.Label Label2 
      Caption         =   "Cierra esta ventana sin asignar un voto a la banca."
      Height          =   435
      Left            =   2520
      TabIndex        =   3
      Top             =   1920
      Width           =   3675
   End
   Begin VB.Label Label1 
      Caption         =   "Asignar un voto positivo a la banca seleccionada."
      Height          =   435
      Left            =   2520
      TabIndex        =   1
      Top             =   180
      Width           =   3675
   End
End
Attribute VB_Name = "frmDefinirVoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mBanca As Integer

Private Sub cmdAbstener_Click()
    MensajesSQV.cambioVoto Str(mBanca), "A"
    Unload Me
End Sub

Private Sub cmdNo_Click()
    MensajesSQV.cambioVoto Str(mBanca), "N"
    Unload Me
End Sub

Private Sub cmdSi_Click()
    MensajesSQV.cambioVoto Str(mBanca), "S"
    Unload Me
End Sub

Private Sub cmdVolver_Click()
    Unload Me
End Sub

Public Property Let Banca(ByVal vNewValue As Integer)
    Me.Caption = "Cambiar el voto de la banca nº " & vNewValue
    mBanca = vNewValue
End Property

