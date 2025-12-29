VERSION 5.00
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmImpresion 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1095
   ClientLeft      =   8985
   ClientTop       =   12570
   ClientWidth     =   4095
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmPuntos 
      Interval        =   250
      Left            =   120
      Top             =   390
   End
   Begin Proyecto1.ButtonOffice cmdSinImprimir 
      Height          =   435
      Left            =   300
      TabIndex        =   1
      Top             =   570
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   767
      BackColor       =   12230304
      Caption         =   "No esperar (No recomendado)"
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
   Begin VB.Timer tmCheck 
      Interval        =   100
      Left            =   120
      Top             =   1050
   End
   Begin VB.Label lblPuntos 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   1740
      TabIndex        =   2
      Top             =   0
      Width           =   2325
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando el almacenamiento del acta..."
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
      Height          =   495
      Left            =   330
      TabIndex        =   0
      Top             =   120
      Width           =   4155
   End
End
Attribute VB_Name = "frmImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSinImprimir_Click()
Dim r As Integer
tmCheck.Enabled = False
r = MsgBox("¿Realmente desea cancelar la espera/impresión?" & vbCrLf & "Nota: esta acción no interfiere con la grabación del acta", vbYesNo)
If r = vbYes Then
    Unload Me
Else
    tmCheck.Enabled = True
End If
End Sub
Private Sub tmCheck_Timer()
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
If Ultimo_Periodo = "" Or Ultima_Sesion = "" Or Ultimo_Acta = "" Then
    MsgBox "Se cerró la Consola con una votacion en curso. Por favor, imprima el acta manualmente.", vbInformation
    Unload Me
    Exit Sub
End If
SetearRs "SELECT * FROM actas WHERE Período_Legislativo = '" & Ultimo_Periodo & "' AND Sesión = " & Ultima_Sesion & " AND Número_de_Acta = " & Ultimo_Acta, rsTemp
If Not rsTemp.EOF Then
    If ImpresionAutomaticaActivada = True Then
        frmConsolaOperacion.Impresion
    End If
    Unload Me
End If
rsTemp.Close
Set rsTemp = Nothing
End Sub

Private Sub tmPuntos_Timer()
If Len(lblPuntos.Caption) = 0 Then
    lblPuntos.Caption = "."
ElseIf Len(lblPuntos.Caption) = 1 Then
    lblPuntos.Caption = ".."
ElseIf Len(lblPuntos.Caption) = 2 Then
    lblPuntos.Caption = "..."
ElseIf Len(lblPuntos.Caption) = 3 Then
    lblPuntos.Caption = ""
End If
End Sub
