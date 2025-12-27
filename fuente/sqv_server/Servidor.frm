VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Servidor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imprime Cartel.."
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6240
   Icon            =   "Servidor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   840
      Top             =   720
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Servidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strLinea(11)      As String
Dim i                 As Long
Dim Orden(11)         As Long
Dim strLineaNueva(11) As String

Private Sub CargaOrden()
    ' Ingreso el Oden de Secuencia que se Muestran los Datos
    Orden(1) = 9
    Orden(2) = 6
    Orden(3) = 8
    Orden(4) = 2
    Orden(5) = 10
    Orden(6) = 11
    Orden(7) = 5
    Orden(8) = 4
    Orden(9) = 7
    Orden(10) = 3
    Orden(11) = 1
End Sub

Private Sub ConectarCom()
On Error GoTo err
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
Exit Sub
err:
End Sub

Private Sub datos()
    With sCartel
        .strAbtenciones = " "
        .strAfirmativos = " "
        .strAusentes = " "
        .strMayoria = " "
        .strNegativos = " "
        .strOrdenDia = " "
        .strPresentes = " "
        .strQuorum = " "
        .strResultado = " "
        .strSesion = " "
        .strTiempoVota = " "
        .strTipoVota = " "
        .strTitulo = " "
        .strLineaCartel10 = " "
        .strLineaCartel11 = " "
        .strAtributo03 = "^V^3" '"^L" & "^V^8"
        .strAtributo04 = "^V^3"
        .strAtributo05 = "^V^3"
        .strAtributo10 = "^V^3"
        .strAtributo11 = "^V^3"
    End With
End Sub
Private Sub ComDatosEnviar()
    Dim FinTex As String
    FinTex = "^_" '& Chr(94) & Chr(95)
    ' Definimos los datos para el Cartel
    With sCartel
        'Linea 1
        'strLineaNueva(1) = "^L " & "^V^6" & Date & " " & Format(Time, "hh") & ":" & Format(Time, "nn") & " " & .strQuorum & FinTex
        'linea 2
        'strLineaNueva(2) = "^L" & "^V^8" & "" & .strPresentes & Space(3 - Len(.strPresentes)) & " " & .strAusentes & Space(3 - Len(.strAusentes)) & FinTex
        'Linea 3
        'strLineaNueva(3) = .strAtributo03 & .strSesion & "    " & FinTex
        'Linea 4
        'strLineaNueva(4) = .strAtributo04 & .strOrdenDia & "    " & FinTex
        'Linea 5
        'strLineaNueva(5) = .strAtributo05 & .strTitulo & "    " & FinTex
        'Linea 6
        'strLineaNueva(6) = "^L" & "^V^8" & Left(.strTipoVota, 9) & "  " & Space(3 - Min(3, Len(Trim(.strTiempoVota)))) & Left(Trim(.strTiempoVota), 3) & FinTex
        ''Linea 7
        'strLineaNueva(7) = "^L" & "^V^8" & .strMayoria & FinTex
        'Linea 8
        'strLineaNueva(8) = "^L" & "^V^8" & .strAfirmativos & Space(3 - Len(.strAfirmativos)) & " " & .strNegativos & Space(3 - Len(.strNegativos)) & " " & .strAbtenciones & Space(3 - Len(.strAbtenciones)) & FinTex
        'Linea 9
        'strLineaNueva(9) = "^L" & "^V^8" & .strResultado & FinTex
        'Linea 10
        'strLineaNueva(10) = .strAtributo10 & .strLineaCartel10 & FinTex
        'Linea 11
        'strLineaNueva(11) = .strAtributo11 & .strLineaCartel11 & FinTex
    End With
    Timer.Enabled = True
End Sub
Private Function Min(n1 As Long, n2 As Long) As Long
    Min = IIf(n1 > n2, n2, n1)
End Function

Private Sub Form_Load()
   MSComm1.Settings = "4800,n,8,2"
   X = 1
   i = 1
   Call datos
   Call CargaOrden
   Call ConectarCom
   Timer.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MSComm1.PortOpen = False
End Sub

Private Sub Timer_Timer()
  If strLineaNueva(Orden(i)) <> strLinea(Orden(i)) Then
    Timer.Interval = 200
    strLinea(Orden(i)) = strLineaNueva(Orden(i))
    If MSComm1.PortOpen Then
        MSComm1.Output = Chr(128 + Orden(i)) & strLinea(Orden(i))
    End If
    DoEvents
  Else
    Timer.Interval = 40
  End If
  i = i + 1
  X = X + 1
  If i = 12 Then i = 1
  Label9.Caption = Time
  Call ComDatosEnviar
End Sub

'**************************************************
' Defino Procedimientos para SQV Server
'**************************************************

Public Property Let sAfirmativo(ByVal vNewValue As Variant)
    sCartel.strAfirmativos = vNewValue
End Property

Public Property Let sQuorum(ByVal vNewValue As Variant)
    sCartel.strQuorum = vNewValue
End Property

Public Property Let sNegativo(ByVal vNewValue As Variant)
    sCartel.strNegativos = vNewValue
End Property

Public Property Let sPresentes(ByVal vNewValue As Variant)
    sCartel.strPresentes = vNewValue
End Property

Public Property Let sAusentes(ByVal vNewValue As Variant)
    sCartel.strAusentes = vNewValue
End Property
Public Property Let sAbstenciones(ByVal vNewValue As Variant)
    sCartel.strAbtenciones = vNewValue
End Property
Public Property Let sResultado(ByVal vNewValue As Variant)
    sCartel.strResultado = vNewValue
End Property
Public Property Let sSesion(ByVal vNewValue As Variant)
    sCartel.strSesion = vNewValue
End Property
Public Property Let sOrdendia(ByVal vNewValue As Variant)
    sCartel.strOrdenDia = vNewValue
End Property
Public Property Let sTitulo(ByVal vNewValue As Variant)
    sCartel.strTitulo = vNewValue
End Property
Public Property Let sTipoVota(ByVal vNewValue As Variant)
    sCartel.strTipoVota = vNewValue
End Property
Public Property Let sTiempoVota(ByVal vNewValue As Variant)
    sCartel.strTiempoVota = vNewValue
End Property
Public Property Let sMayoria(ByVal vNewValue As Variant)
    sCartel.strMayoria = vNewValue
End Property
Public Property Let sLineaCartel10(ByVal vNewValue As Variant)
    sCartel.strLineaCartel10 = vNewValue
End Property
Public Property Let sLineaCartel11(ByVal vNewValue As Variant)
    sCartel.strLineaCartel11 = vNewValue
End Property




