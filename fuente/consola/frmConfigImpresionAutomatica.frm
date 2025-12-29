VERSION 5.00
Begin VB.Form frmConfigImpresionAutomatica 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Configuracíon de Impresión Automática"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabarYSalir 
      Caption         =   "Grabar y Salir"
      Height          =   315
      Left            =   1920
      TabIndex        =   14
      Top             =   3900
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir sin grabar"
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   3900
      Width           =   1695
   End
   Begin VB.TextBox txtAusentes 
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtAbstenciones 
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtNegativos 
      Height          =   315
      Left            =   2040
      TabIndex        =   9
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtAfirmativos 
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdSubir 
      Caption         =   "Subir"
      Height          =   255
      Left            =   780
      TabIndex        =   5
      Top             =   1920
      Width           =   1275
   End
   Begin VB.CommandButton cmdBajar 
      Caption         =   "Bajar"
      Height          =   255
      Left            =   2100
      TabIndex        =   4
      Top             =   1920
      Width           =   1275
   End
   Begin VB.ListBox lstOrden 
      Height          =   1230
      Left            =   780
      TabIndex        =   3
      Top             =   660
      Width           =   2595
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   3600
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label7 
      Caption         =   "Copias AUSENTES:"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   3420
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Copias ABSTENCIONES:"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   2340
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Copias NEGATIVO:"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   3060
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Copias AFIRMATIVO:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   2700
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Orden:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   660
      Width           =   555
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   3540
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblInfo 
      Caption         =   "dentro del recinto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Valores de impresión para"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "frmConfigImpresionAutomatica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DentroRecinto As Boolean
Dim NombreOrden As String
Dim NombreCopias As String
Dim CAfirmativos As Integer 'En estas variables guardo el valor inicial de la cantidad
Dim CNegativos As Integer 'Para luego saber si hacer el update
Dim CAbstenciones As Integer
Dim CAusentes As Integer
Dim iCOMP As String 'String para ver si cambio el listbox, ver cmdGrabarYSalir
Dim rsDatos As Recordset
Private Sub cmdBajar_Click()
Dim bc As String
Dim aInd As Integer
If lstOrden.ListIndex = lstOrden.ListCount - 1 Then
    MsgBox ("No se puede bajar mas!")
Else
    bc = lstOrden.List(lstOrden.ListIndex)
    aInd = lstOrden.ListIndex
    lstOrden.List(aInd) = lstOrden.List(aInd + 1)
    lstOrden.List(aInd + 1) = bc
    lstOrden.ListIndex = aInd + 1
End If
End Sub
Private Sub cmdGrabarYSalir_Click()
On Error GoTo pError
Dim xTemp As String
Dim i As Integer
'Reviso el listbox
For i = 0 To lstOrden.ListCount - 1
    xTemp = xTemp & lstOrden.List(i)
Next i
If xTemp <> iCOMP Then
    'Cambió el listbox, cambió el orden
    For i = 0 To lstOrden.ListCount - 1
        EjecutarSQL ("UPDATE paramorden SET " & NombreOrden & " = " & i + 1 & " WHERE Resultado = '" & lstOrden.List(i) & "          '")
    Next i
Else
    'MsgBox ("No cambio el listbox!") PARA PRUEBAS
End If
'Reviso las cantidades
If txtAfirmativos.Text <> CAfirmativos Then
    EjecutarSQL ("UPDATE paramorden SET " & NombreCopias & " = " & txtAfirmativos.Text & " WHERE Resultado = 'AFIRMATIVO          '")
End If
If txtNegativos.Text <> CNegativos Then
    EjecutarSQL ("UPDATE paramorden SET " & NombreCopias & " = " & txtNegativos.Text & " WHERE Resultado = 'NEGATIVO          '")
End If
If txtAbstenciones.Text <> CAbstenciones Then
    EjecutarSQL ("UPDATE paramorden SET " & NombreCopias & " = " & txtAbstenciones.Text & " WHERE Resultado = 'ABSTENCION          '")
End If
If txtAusentes.Text <> CAusentes Then
    EjecutarSQL ("UPDATE paramorden SET " & NombreCopias & " = " & txtAusentes.Text & " WHERE Resultado = 'AUSENTE          '")
End If
Unload Me
Exit Sub
pError: MsgBox ("Se ha producido el siguiente error" & Err.Description)
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub cmdSubir_Click()
Dim bc As String
Dim aInd As Integer
If lstOrden.ListIndex = 0 Then
    MsgBox ("No se puede subir más!")
Else
    bc = lstOrden.List(lstOrden.ListIndex)
    aInd = lstOrden.ListIndex
    lstOrden.List(aInd) = lstOrden.List(aInd - 1)
    lstOrden.List(aInd - 1) = bc
    lstOrden.ListIndex = aInd - 1
End If
End Sub
Private Sub Form_Load()
Dim SQL_Imp As String
If DentroRecinto = True Then
    NombreOrden = "Orden_r"
    NombreCopias = "copias_r"
Else
    NombreOrden = "Orden"
    NombreCopias = "copias"
    lblInfo.Caption = "afuera del recinto"
End If
SQL_Imp = "SELECT Resultado," & NombreCopias & " FROM paramorden ORDER BY " & NombreOrden
Set rsDatos = New Recordset
SetearRs SQL_Imp, rsDatos
While rsDatos.EOF = False
    Select Case Trim(rsDatos.Fields("Resultado"))
    Case "AFIRMATIVO"
        lstOrden.AddItem "AFIRMATIVO"
        iCOMP = iCOMP & "AFIRMATIVO"
        txtAfirmativos.Text = rsDatos.Fields(NombreCopias)
        CAfirmativos = Int(txtAfirmativos.Text)
    Case "NEGATIVO"
        lstOrden.AddItem "NEGATIVO"
        iCOMP = iCOMP & "NEGATIVO"
        txtNegativos.Text = rsDatos.Fields(NombreCopias)
        CNegativos = Int(txtNegativos.Text)
    Case "ABSTENCION"
        lstOrden.AddItem "ABSTENCION"
        iCOMP = iCOMP & "ABSTENCION"
        txtAbstenciones.Text = rsDatos.Fields(NombreCopias)
        CAbstenciones = Int(txtAbstenciones.Text)
    Case "AUSENTE"
        lstOrden.AddItem "AUSENTE"
        iCOMP = iCOMP & "AUSENTE"
        txtAusentes.Text = rsDatos.Fields(NombreCopias)
        CAusentes = Int(txtAusentes.Text)
    End Select
    rsDatos.MoveNext
Wend
rsDatos.Close
Set rsDatos = Nothing
End Sub
