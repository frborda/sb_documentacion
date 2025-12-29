VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{172CC8FF-7909-413F-9341-19B0B44AB0F8}#1.0#0"; "ocx-button-ofiice-xp-2003.ocx"
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estadísticas"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      Caption         =   "Impresora a utilizar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   795
      Left            =   60
      TabIndex        =   19
      Top             =   4980
      Width           =   4605
      Begin VB.ComboBox cmbImpresoras 
         Height          =   315
         ItemData        =   "frmEstadisticas.frx":0000
         Left            =   120
         List            =   "frmEstadisticas.frx":0002
         TabIndex        =   20
         Text            =   "- Seleccione una impresora -"
         Top             =   300
         Width           =   4335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   "Lista de Impresión"
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
      Height          =   6495
      Left            =   5820
      TabIndex        =   12
      Top             =   60
      Width           =   5535
      Begin VB.ListBox lstSeleccionados 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5700
         ItemData        =   "frmEstadisticas.frx":0004
         Left            =   60
         List            =   "frmEstadisticas.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   5415
      End
      Begin Proyecto1.ButtonOffice cmdLimpiar 
         Height          =   435
         Left            =   90
         TabIndex        =   14
         Top             =   6000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         BackColor       =   8421631
         Caption         =   "Limpiar lista"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         PicOpacity      =   0
      End
   End
   Begin Proyecto1.ButtonOffice cmdImprimir 
      Height          =   675
      Left            =   60
      TabIndex        =   10
      Top             =   5880
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   1191
      BackColor       =   12230304
      Caption         =   "Imprimir Estadísticas"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1635
      Left            =   60
      TabIndex        =   3
      Top             =   3240
      Width           =   4605
      Begin Proyecto1.ButtonOffice cmdCopyFecha 
         Height          =   315
         Left            =   3180
         TabIndex        =   9
         Top             =   720
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BackColor       =   12230304
         Caption         =   "&C"
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
      Begin VB.CheckBox chkFecha 
         BackColor       =   &H00404040&
         Caption         =   "Cualquier Fecha"
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
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1995
      End
      Begin MSComCtl2.DTPicker dtDesde 
         Height          =   315
         Left            =   780
         TabIndex        =   4
         Top             =   300
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57999361
         CurrentDate     =   40666
      End
      Begin MSComCtl2.DTPicker dtHasta 
         Height          =   315
         Left            =   780
         TabIndex        =   5
         Top             =   720
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57999361
         CurrentDate     =   40666
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Diputado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3105
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4575
      Begin VB.CheckBox chkSoloActivos 
         BackColor       =   &H00404040&
         Caption         =   "Listar sólo activos"
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
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2115
      End
      Begin VB.ListBox lstDiputados 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1020
         Width           =   4305
      End
      Begin VB.TextBox txtApellido 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   4305
      End
   End
   Begin Proyecto1.ButtonOffice cmdMoverDerecha 
      Height          =   435
      Left            =   4770
      TabIndex        =   15
      Top             =   180
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   767
      BackColor       =   12230304
      Caption         =   ">"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdMoverIzquierda 
      Height          =   435
      Left            =   4770
      TabIndex        =   16
      Top             =   720
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   767
      BackColor       =   12230304
      Caption         =   "<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicOpacity      =   0
   End
   Begin Proyecto1.ButtonOffice cmdPasarTodos 
      Height          =   645
      Left            =   4770
      TabIndex        =   17
      Top             =   1260
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1138
      BackColor       =   12230304
      Caption         =   "Pasar todos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicOpacity      =   0
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "Listo."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   18
      Top             =   6660
      Width           =   9075
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ids(0 To 256) As Integer
Dim diputadoActual As String
Private Sub ActualizarLista()
CargaDiputados
chkFecha.Value = vbChecked
End Sub
Private Sub chkFecha_Click()
If chkFecha.Value = vbChecked Then
    dtDesde.Enabled = False
    dtHasta.Enabled = False
    cmdCopyFecha.Enabled = False
Else
    dtDesde.Enabled = True
    cmdCopyFecha.Enabled = True
    dtHasta.Enabled = True
End If
End Sub
Private Sub cmbPeriodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
Private Sub chkSoloActivos_Click()
If chkSoloActivos.Value = vbChecked Then
    CargaDiputadosActivos
Else
    CargaDiputados
End If
RevisaListas
End Sub
Private Sub cmdCopyFecha_Click()
CopiarFechas
End Sub
Private Sub cmdImprimir_Click()
Dim RTA As Integer
Dim totalFile As String
Dim links As String
Dim table As String
Dim middle As String
Dim linkNum As Integer
linkNum = 1
totalFile = ""
links = ""
table = ""
middle = ""
totalFile = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & vbCrLf
totalFile = totalFile & "<TTArticles xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation=""opencms://system/modules/ar.gov.hcdn/schemas/hcdntablas.xsd"">" & vbCrLf
totalFile = totalFile & "  <TTArticle language=""en"">" & vbCrLf
totalFile = totalFile & "    <Title><![CDATA[Estadísticas Individuales Período PER_PARAM]]></Title>" & vbCrLf
totalFile = totalFile & "    <Paragraphs>" & vbCrLf
totalFile = totalFile & "      <Text name=""Text0"">" & vbCrLf
'Table
table = table & "<table width=""96%"" cellspacing=""1"" cellpadding=""1"" border=""1"" align=""center"">" & vbCrLf
table = table & "    <thead>" & vbCrLf
table = table & "        <tr>" & vbCrLf
table = table & "            <th scope=""col"" style=""text-align: center;"">LEGISLADOR</th>" & vbCrLf
table = table & "            <th scope=""col"" style=""text-align: center;"">BLOQUE POLÍTICO</th>" & vbCrLf
table = table & "            <th scope=""col"" style=""text-align: center;"">PROVINCIA</th>" & vbCrLf
table = table & "        </tr>" & vbCrLf
table = table & "    </thead>" & vbCrLf
table = table & "    <tbody>"
'Middle
middle = "        <content><![CDATA[<p>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style=""text-decoration: underline;""><span style=""font-weight: bold;"">Informe Detallado de Votaciones Nominales por Legislador Período N&ordm; PAR_PER</span></span></p>" & vbCrLf
middle = middle & "<p>&nbsp;</p>" & vbCrLf
middle = middle & "<p>Los datos presentados corresponden exclusivamente a los extraídos de las Actas de Votación Nominal emitidas en el Recinto de Sesiones en el Período N&ordm; PAR_PER comprendido entre el PAR_INICIO y PAR_FIN.<br />" & vbCrLf
middle = middle & "Se excluye del presente informe al  Sr. Presidente de esta Honorable Cámara Dr. DOMINGUEZ, Julián Andrés; que conforme a sus obligaciones reglamentarias no vota salvo casos excepcionales.</p>" & vbCrLf
middle = middle & "<p>&nbsp;</p>" & vbCrLf
middle = middle & "<table width=""500"" cellspacing=""1"" cellpadding=""1"" border=""1"" align=""center"">" & vbCrLf
middle = middle & "    <tbody>" & vbCrLf
middle = middle & "        <tr>" & vbCrLf
middle = middle & "            <td style=""text-align: center;font-size: 12pt;""><span style=""font-weight: bold;""><a href=""%(link0)"" target=""_self"" title=""Estadísticas Generales"">Informe Alfabético</a></span></td>" & vbCrLf
middle = middle & "            <td style=""text-align: center;font-size: 12pt;""><span style=""font-weight: bold;""><a href=""%(link1)"" target=""_self"" title=""Fechas Actas"">Fecha de Actas</a></span></td>" & vbCrLf
middle = middle & "        </tr>" & vbCrLf
middle = middle & "    </tbody>" & vbCrLf
middle = middle & "</table>" & vbCrLf
middle = middle & "<p><span style=""text-decoration: underline;""><span style=""font-weight: bold;"">Votaciones Individuales de cada Legislador</span></span>&nbsp;</p>"
'Process
frmEstadisticas.Enabled = False
If lstSeleccionados.ListCount = 0 Then
    MsgBox "No ha seleccionado a ningún diputado de la lista", vbCritical
Else
    RTA = MsgBox("Se van a imprimir " & lstSeleccionados.ListCount & " diputados. " & vbCrLf & "¿Desea continuar?", vbYesNo)
    If RTA = vbYes Then
        Dim i As Integer
        Dim pags As Integer
        Dim Buff() As String
        If lstSeleccionados.ListCount > 0 Then
            For i = 0 To lstSeleccionados.ListCount - 1
                EstadisticasTotalAfirmativos = 0
                EstadisticasTotalNegativos = 0
                EstadisticasTotalAusentes = 0
                EstadisticasTotalAbstenciones = 0
                EstadisticasTotalPresidencias = 0
                pags = 0
                HeaderEstadisticaIndividualImpreso = False
                Buff() = Split(lstSeleccionados.List(i), ";")
                diputadoActual = Buff(0)
                Buff(1) = Trim(Buff(1))
                pags = getPaginasEstadistica(Buff(1))
                EstadisticaIndividualTotal = pags
                Dim b As Boolean
                b = ImprimirEstadistica(Buff(1), Trim(Buff(2)), Trim(Buff(0)))
                If (b) Then
                    linkNum = linkNum + 1
                    If (links = "") Then
                       links = "        <links>" & vbCrLf
                       'Links estaticos
                       links = links & "          <link name=""link0"" internal=""false"" type=""A"">" & vbCrLf
                       links = links & "            <target><![CDATA[http://www1.hcdn.gov.ar/dependencias/dselectronicos/actas/PAR_INF/Estadística General.pdf]]></target>" & vbCrLf
                       links = links & "          </link>" & vbCrLf
                       links = links & "          <link name=""link1"" internal=""false"" type=""A"">" & vbCrLf
                       links = links & "            <target><![CDATA[http://www1.hcdn.gov.ar/dependencias/dselectronicos/actas/PAR_INF/Fechas Actas Periodo 132.pdf]]></target>" & vbCrLf
                       links = links & "          </link>" & vbCrLf
                       'Links Nuevos
                       links = links & "          <link name=""link" & Trim(CStr(linkNum)) & """ internal=""false"" type=""A"">" & vbCrLf
                       links = links & "            <target><![CDATA[   http://www1.hcdn.gov.ar/dependencias/dselectronicos/actas/PAR_INF/" & Trim(Buff(2)) & ".pdf   ]]></target>" & vbCrLf
                       links = links & "          </link>"
                       table = table & vbCrLf & "        <tr>" & vbCrLf
                       table = table & "            <td><a title=""Estadistica Individual pdf"" href=""%(link" & Trim(CStr(linkNum)) & ")"" target=""_self"">" & vbTab & Buff(0) & vbTab & "</a></td>" & vbCrLf
                       table = table & "            <td>" & Buff(4) & "</td>" & vbCrLf
                       table = table & "            <td>" & Buff(3) & "</td>" & vbCrLf
                       table = table & "        </tr>"
                    Else
                       links = links & vbCrLf & "          <link name=""link" & Trim(CStr(linkNum)) & """ internal=""false"" type=""A"">" & vbCrLf
                       links = links & "            <target><![CDATA[   http://www1.hcdn.gov.ar/dependencias/dselectronicos/actas/PAR_INF/" & Trim(Buff(2)) & ".pdf   ]]></target>" & vbCrLf
                       links = links & "          </link>"
                       table = table & vbCrLf & "        <tr>" & vbCrLf
                       table = table & "            <td><a title=""Estadistica Individual pdf"" href=""%(link" & Trim(CStr(linkNum)) & ")"" target=""_self"">" & vbTab & Buff(0) & vbTab & "</a></td>" & vbCrLf
                       table = table & "            <td>" & Buff(4) & "</td>" & vbCrLf
                       table = table & "            <td>" & Buff(3) & "</td>" & vbCrLf
                       table = table & "        </tr>"
                    End If
                End If
                DoEvents
            Next i
            links = links & vbCrLf & "        </links>"
            table = table & vbCrLf & "    </tbody>"
            table = table & vbCrLf & "</table>"
            totalFile = totalFile & links & vbCrLf & middle & vbCrLf & table
            totalFile = totalFile & vbCrLf & "<p><br />"
            totalFile = totalFile & vbCrLf & "&nbsp;</p>]]></content>"
            totalFile = totalFile & vbCrLf & "      </Text>"
            totalFile = totalFile & vbCrLf & "    </Paragraphs>"
            totalFile = totalFile & vbCrLf & "  </TTArticle>"
            totalFile = totalFile & vbCrLf & "</TTArticles>"
            MesDesde = AgregaCero(dtDesde.Month)
            DiaDesde = AgregaCero(dtDesde.Day)
            MesHasta = AgregaCero(dtHasta.Month)
            DiaHasta = AgregaCero(dtHasta.Day)
            totalFile = Replace(totalFile, "PAR_INICIO", AgregaCero(dtDesde.Day) & "/" & AgregaCero(dtDesde.Month) & "/" & Trim(CStr(dtDesde.Year)))
            totalFile = Replace(totalFile, "PAR_FIN", AgregaCero(dtHasta.Day) & "/" & AgregaCero(dtHasta.Month) & "/" & Trim(CStr(dtHasta.Year)))
            Dim sPer As String
            sPer = ""
            While (sPer = "")
                sPer = InputBox("Ingrese el número de período a imprimir", "Datos para OpenCMS", "")
            Wend
            totalFile = Replace(totalFile, "PAR_PER", sPer)
            Dim sInf As String
            sInf = ""
            While (sInf = "")
                sInf = InputBox("Ingrese el nombre de la carpeta de Informes", "Datos para OpenCMS", "")
            Wend
            totalFile = Replace(totalFile, "PAR_INF", sInf)
            Clipboard.Clear
            Clipboard.SetText (totalFile)
            MsgBox ("El nuevo HTML de estadísticas se ha copiado al Clipboard")
        Else
            MsgBox "No ha seleccionado a ningún diputado de la lista", vbCritical
        End If
    End If
End If
frmEstadisticas.Enabled = True
End Sub
Public Function getPaginasEstadistica(id As String) As Integer
Dim consulta As String
Dim periodo As String
Dim FiltroFechas As String
Dim DiaDesde As String
Dim MesDesde As String
Dim DiaHasta As String
Dim MesHasta As String
Dim fecha1 As String
Dim fecha2 As String
Dim Desempates_Negativos As Integer
Dim Desempates_Afirmativos As Integer
Dim rsDesempates As ADODB.Recordset
cmdImprimir.Enabled = False
cmdImprimir.Caption = "Cargando..."
DoEvents
Desempates_Negativos = 0
Desempates_Afirmativos = 0
MesDesde = AgregaCero(dtDesde.Month)
DiaDesde = AgregaCero(dtDesde.Day)
MesHasta = AgregaCero(dtHasta.Month)
DiaHasta = AgregaCero(dtHasta.Day)
fecha1 = DiaDesde & "/" & MesDesde & "/" & dtDesde.Year & " 00:00:00"
fecha2 = DiaHasta & "/" & MesHasta & "/" & dtHasta.Year & " 23:59:59"
If chkFecha.Value = vbUnchecked Then
    FiltroFechas = " BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' "
    FiltroFechas = " AND ((SELECT     actas.Fecha " & _
                              " From actas " & _
                              " WHERE     Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                    " actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) " & _
                    FiltroFechas & ")"
Else
    FiltroFechas = ""
End If
Dim RsTemp As ADODB.Recordset
periodo = "'%'"
Dim rpt As rptEstadisticaIndividual
If lstSeleccionados.ListCount > 0 Then
    Dim FechaDesempate As String
    If chkFecha.Value = vbChecked Then
        FechaDesempate = ""
    Else
        FechaDesempate = " AND Fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "'"
    End If
    'Id = "'" & Ids(lstDiputados.ListIndex) & "'"
    '*****CONSULTA PARA LOS DESEMPATES*****
    consulta = "SELECT COUNT(Desempate) FROM actas WHERE Desempate = 'Si' AND Período_Legislativo = " & periodo & " AND " & _
       " Votacion = 'AFIRMATIVO' AND actas.Presidente = " & id & FechaDesempate
    Set rsDesempates = New ADODB.Recordset
    SetearRs consulta, rsDesempates
    If rsDesempates.EOF Then
        Desempates_Afirmativos = 0
    Else
        Desempates_Afirmativos = rsDesempates.Fields(0)
    End If
    rsDesempates.Close
    Set rsDesempates = Nothing
    Set rsDesempates = New ADODB.Recordset
    consulta = "SELECT COUNT(Desempate) FROM actas WHERE Desempate = 'Si' AND Período_Legislativo = " & periodo & " AND " & _
       " Votacion = 'NEGATIVO' AND actas.Presidente = " & id & FechaDesempate
    SetearRs consulta, rsDesempates
    If rsDesempates.EOF Then
        Desempates_Negativos = 0
    Else
        Desempates_Negativos = rsDesempates.Fields(0)
    End If
    rsDesempates.Close
    Set rsDesempates = Nothing
    consulta = "SELECT '" & Desempates_Negativos & "' AS Desempates_Negativos, " & _
                          "'" & Desempates_Afirmativos & "' AS Desempates_Afirmativos, Legisladores.apellido + ', ' + Legisladores.nombre AS Diputado, " & _
                      " detalleactas.bloque_político, Legisladores.grupo_politico, distritos.distrito AS Provincia, " & _
                          "(SELECT     COUNT(detalleactas.Resultado) " & _
                            "From detalleactas " & _
                            "WHERE      (Versión_Acta = 0) " & FiltroFechas & " AND (Legislador_asignado = " & id & ") AND (Período_Legislativo = " & periodo & ") AND LTrim(RTrim(Resultado)) = 'AFIRMATIVO') " & _
                      "AS CantAfirm, " & _
                          "(SELECT     COUNT(detalleactas.Resultado) " & _
                            "From detalleactas " & _
                            "WHERE      (Versión_Acta = 0) " & FiltroFechas & " AND (Legislador_asignado = " & id & ") AND (Período_Legislativo = " & periodo & ") AND LTrim(RTrim(Resultado)) = 'NEGATIVO') " & _
                      "AS CantNeg, " & _
                          "(SELECT     COUNT(detalleactas.Resultado) " & _
                            "From detalleactas " & _
                            "WHERE      (Versión_Acta = 0) " & FiltroFechas & " AND (Legislador_asignado = " & id & ") AND (Período_Legislativo = " & periodo & ") AND LTrim(RTrim(Resultado)) = 'AUSENTE' AND " & _
                                                    "detalleactas.Legislador_asignado <> " & _
                                                       "(SELECT     actas.Presidente " & _
                                                         "From actas " & _
                                                         "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                                                "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0)) AS CantAus, "
       consulta = consulta & "SUBSTRING(detalleactas.Período_Legislativo, 1, 3) + ' - ' + CASE SUBSTRING(detalleactas.Período_Legislativo, 4, 1) " & _
                      "WHEN 'O' THEN 'Ordinario' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Prórroga' WHEN 'L' THEN 'Legislativo' END + ' - ' + CASE SUBSTRING(detalleactas.Período_Legislativo, " & _
                       "5, 2) " & _
                      "WHEN 'T' THEN 'Tablas' WHEN 'H' THEN 'Homenajes' WHEN 'E' THEN 'Especial' WHEN 'P' THEN 'Preparatoria ' WHEN 'I' THEN 'Informativa' END AS " & _
                       "Per,SUBSTRING(detalleactas.Período_Legislativo, 1, 3) AS NumPer, " & _
                          "(SELECT     actas.Nombre_del_Acta " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) AS Nombre_del_Acta, "
consulta = consulta & "(SELECT     actas.Número_de_Acta " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) AS NroActa, " & _
"(SELECT     actas.Fecha " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) AS Fecha, CASE " & _
                          "(SELECT     CASE actas.Presidente WHEN detalleactas.Legislador_asignado THEN 'Pte.' ELSE 'N' END " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) WHEN 'Pte.' THEN " & _
                          "(SELECT     CASE actas.resultado_voto_presidente WHEN 's' THEN 'AFIRMATIVO' WHEN 'n' THEN 'NEGATIVO' ELSE 'S/VOTO' END AS Ex " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) ELSE (Resultado) END AS Voto, " & _
                          "(SELECT     CASE actas.Presidente WHEN detalleactas.Legislador_asignado THEN 'Pte.' ELSE '' END " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) AS PresiDato, "
        consulta = consulta & "(SELECT    actas.Desempate " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) AS Acta_Desempate, "
        consulta = consulta & "(SELECT    actas.Votacion " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) AS Acta_Votacion, "
        consulta = consulta & "CASE detalleactas.Operación WHEN 'votnom' THEN 'Votación Nominal' WHEN 'paslis' THEN 'Pase de Lista' END AS TipoOp, " & _
                      "detalleactas.sesión " & _
"FROM         detalleactas INNER JOIN " & _
                      "Legisladores ON Legisladores.id = detalleactas.Legislador_asignado INNER JOIN " & _
                      "distritos ON Legisladores.distrito = distritos.id_distrito " & _
"WHERE (detalleactas.Versión_Acta = 0) " & FiltroFechas & " AND (detalleactas.Legislador_asignado = " & id & ")" & _
"ORDER BY Fecha ASC,detalleactas.Período_Legislativo"
    Set RsTemp = New ADODB.Recordset
    SetearRs consulta, RsTemp
    Set rpt = New rptEstadisticaIndividual
    If chkFecha.Value = vbUnchecked Then
        rpt.lblVotacion.Caption = rpt.lblVotacion.Caption & " desde el " & dtDesde.Day & "/" & dtDesde.Month & "/" & dtDesde.Year & " hasta el " & _
            dtHasta.Day & "/" & dtHasta.Month & "/" & dtHasta.Year
    End If
    If RsTemp.EOF Then
        'MsgBox ("No hay resultados para estas fechas!")
        lblEstado.Caption = "No hay resultados para " & diputadoActual
        RsTemp.Close
        Set RsTemp = Nothing
    Else
        Set rpt.DataControl1.Recordset = RsTemp
        rpt.Run False
        Dim i As Integer
        getPaginasEstadistica = rpt.Pages.Count
        RsTemp.Close
        Set RsTemp = Nothing
    End If
    Set rpt = Nothing
Else
    MsgBox "Seleccione a un diputado!", vbCritical
End If
cmdImprimir.Enabled = True
cmdImprimir.Caption = "Imprimir Estadísticas"
End Function
Private Function ImprimirEstadistica(id As String, idInterno As String, pNombre As String) As Boolean
Dim consulta As String
Dim periodo As String
Dim FiltroFechas As String
Dim DiaDesde As String
Dim MesDesde As String
Dim DiaHasta As String
Dim MesHasta As String
Dim fecha1 As String
Dim fecha2 As String
Dim Desempates_Negativos As Integer
Dim Desempates_Afirmativos As Integer
Dim rsDesempates As ADODB.Recordset
cmdImprimir.Enabled = False
cmdImprimir.Caption = "Cargando..."
DoEvents
Desempates_Negativos = 0
Desempates_Afirmativos = 0
MesDesde = AgregaCero(dtDesde.Month)
DiaDesde = AgregaCero(dtDesde.Day)
MesHasta = AgregaCero(dtHasta.Month)
DiaHasta = AgregaCero(dtHasta.Day)
fecha1 = DiaDesde & "/" & MesDesde & "/" & dtDesde.Year & " 00:00:00"
fecha2 = DiaHasta & "/" & MesHasta & "/" & dtHasta.Year & " 23:59:59"
If chkFecha.Value = vbUnchecked Then
    FiltroFechas = " BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' "
    FiltroFechas = " AND ((SELECT     actas.Fecha " & _
                              " From actas " & _
                              " WHERE     Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                    " actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) " & _
                    FiltroFechas & ")"
Else
    FiltroFechas = ""
End If
Dim RsTemp As ADODB.Recordset
periodo = "'%'"
Dim rpt As rptEstadisticaIndividual
If lstSeleccionados.ListCount > 0 Then
    Dim FechaDesempate As String
    If chkFecha.Value = vbChecked Then
        FechaDesempate = ""
    Else
        FechaDesempate = " AND Fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "'"
    End If
    'Id = "'" & Ids(lstDiputados.ListIndex) & "'"
    '*****CONSULTA PARA LOS DESEMPATES*****
    consulta = "SELECT COUNT(Desempate) FROM actas WHERE Desempate = 'Si' AND Período_Legislativo = " & periodo & " AND " & _
       " Votacion = 'AFIRMATIVO' AND actas.Presidente = " & id & FechaDesempate
    Set rsDesempates = New ADODB.Recordset
    SetearRs consulta, rsDesempates
    If rsDesempates.EOF Then
        Desempates_Afirmativos = 0
    Else
        Desempates_Afirmativos = rsDesempates.Fields(0)
    End If
    rsDesempates.Close
    Set rsDesempates = Nothing
    Set rsDesempates = New ADODB.Recordset
    consulta = "SELECT COUNT(Desempate) FROM actas WHERE Desempate = 'Si' AND Período_Legislativo = " & periodo & " AND " & _
       " Votacion = 'NEGATIVO' AND actas.Presidente = " & id & FechaDesempate
    SetearRs consulta, rsDesempates
    If rsDesempates.EOF Then
        Desempates_Negativos = 0
    Else
        Desempates_Negativos = rsDesempates.Fields(0)
    End If
    rsDesempates.Close
    Set rsDesempates = Nothing
    consulta = "SELECT '" & Desempates_Negativos & "' AS Desempates_Negativos, " & _
                          "'" & Desempates_Afirmativos & "' AS Desempates_Afirmativos, Legisladores.apellido + ', ' + Legisladores.nombre AS Diputado, " & _
                      "Legisladores.PICTURE, detalleactas.bloque_político,detalleactas.Nro_de_Acta AS NroActa, Legisladores.grupo_politico, distritos.distrito AS Provincia, " & _
                          "(SELECT     COUNT(detalleactas.Resultado) " & _
                            "From detalleactas " & _
                            "WHERE      (Versión_Acta = 0) " & FiltroFechas & " AND (Legislador_asignado = " & id & ") AND (Período_Legislativo = " & periodo & ") AND LTrim(RTrim(Resultado)) = 'AFIRMATIVO') " & _
                      "AS CantAfirm, " & _
                          "(SELECT     COUNT(detalleactas.Resultado) " & _
                            "From detalleactas " & _
                            "WHERE      (Versión_Acta = 0) " & FiltroFechas & " AND (Legislador_asignado = " & id & ") AND (Período_Legislativo = " & periodo & ") AND LTrim(RTrim(Resultado)) = 'NEGATIVO') " & _
                      "AS CantNeg, " & _
                          "(SELECT     COUNT(detalleactas.Resultado) " & _
                            "From detalleactas " & _
                            "WHERE      (Versión_Acta = 0) " & FiltroFechas & " AND (Legislador_asignado = " & id & ") AND (Período_Legislativo = " & periodo & ") AND LTrim(RTrim(Resultado)) = 'AUSENTE' AND " & _
                                                    "detalleactas.Legislador_asignado <> " & _
                                                       "(SELECT     actas.Presidente " & _
                                                         "From actas " & _
                                                         "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                                                "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0)) AS CantAus, "
       consulta = consulta & "SUBSTRING(detalleactas.Período_Legislativo, 1, 3) + ' - ' + CASE SUBSTRING(detalleactas.Período_Legislativo, 4, 1) " & _
                      "WHEN 'O' THEN 'Ordinario' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Prórroga' WHEN 'L' THEN 'Legislativo' END + ' - ' + CASE SUBSTRING(detalleactas.Período_Legislativo, " & _
                       "5, 2) " & _
                      "WHEN 'T' THEN 'Tablas' WHEN 'H' THEN 'Homenajes' WHEN 'E' THEN 'Especial' WHEN 'P' THEN 'Preparatoria ' WHEN 'I' THEN 'Informativa' END AS " & _
                       "Per,SUBSTRING(detalleactas.Período_Legislativo, 1, 3) AS NumPer, " & _
                          "(SELECT     actas.Nombre_del_Acta " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) AS Nombre_del_Acta, " & _
                          "(SELECT     actas.Fecha " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) AS Fecha, CASE " & _
                          "(SELECT     CASE actas.Presidente WHEN detalleactas.Legislador_asignado THEN 'Pte.' ELSE 'N' END " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) WHEN 'Pte.' THEN " & _
                          "(SELECT     CASE actas.resultado_voto_presidente WHEN 's' THEN 'AFIRMATIVO' WHEN 'n' THEN 'NEGATIVO' ELSE 'S/VOTO' END AS Ex " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) ELSE (Resultado) END AS Voto, " & _
                          "(SELECT     CASE actas.Presidente WHEN detalleactas.Legislador_asignado THEN 'Pte.' ELSE '' END " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) AS PresiDato, "
        consulta = consulta & "(SELECT    actas.Desempate " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) AS Acta_Desempate, "
        consulta = consulta & "(SELECT    actas.Votacion " & _
                            "From actas " & _
                            "WHERE      Período_Legislativo = detalleactas.Período_Legislativo AND actas.sesión = detalleactas.sesión AND " & _
                                                   "actas.Número_de_Acta = detalleactas.Nro_de_Acta AND actas.Versión_Acta = 0) AS Acta_Votacion, "
        consulta = consulta & "CASE detalleactas.Operación WHEN 'votnom' THEN 'Votación Nominal' WHEN 'paslis' THEN 'Pase de Lista' END AS TipoOp, " & _
                      "detalleactas.sesión " & _
"FROM         detalleactas INNER JOIN " & _
                      "Legisladores ON Legisladores.id = detalleactas.Legislador_asignado INNER JOIN " & _
                      "distritos ON Legisladores.distrito = distritos.id_distrito " & _
"WHERE (detalleactas.Versión_Acta = 0) " & FiltroFechas & " AND (detalleactas.Legislador_asignado = " & id & ")" & _
"ORDER BY Fecha ASC,detalleactas.Período_Legislativo"
    Set RsTemp = New ADODB.Recordset
    SetearRs consulta, RsTemp
    Set rpt = New rptEstadisticaIndividual
    If chkFecha.Value = vbUnchecked Then
        rpt.lblVotacion.Caption = rpt.lblVotacion.Caption & " desde el " & dtDesde.Day & "/" & dtDesde.Month & "/" & dtDesde.Year & " hasta el " & _
            dtHasta.Day & "/" & dtHasta.Month & "/" & dtHasta.Year
    End If
    If RsTemp.EOF Then
        'MsgBox ("No hay resultados para estas fechas!")
        lblEstado.Caption = "No hay resultados para " & diputadoActual
        RsTemp.Close
        Set RsTemp = Nothing
        ImprimirEstadistica = False
    Else
        rpt.documentName = idInterno
        Set rpt.DataControl1.Recordset = RsTemp
        rpt.Run False
        Dim i As Integer
        For i = 0 To rpt.Pages.Count - 1
            rpt.Pages(i).Width = 300
            rpt.Pages.Commit
        Next i
        If (cmbImpresoras.ListIndex > -1) Then
            rpt.Printer.DeviceName = cmbImpresoras.List(cmbImpresoras.ListIndex)
        End If
        rpt.PrintReport False
        RsTemp.Close
        Set RsTemp = Nothing
        ImprimirEstadistica = True
    End If
    Set rpt = Nothing
Else
    MsgBox "Seleccione a un diputado!", vbCritical
End If
cmdImprimir.Enabled = True
cmdImprimir.Caption = "Imprimir Estadísticas"
End Function
Private Sub cmdImprimir_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub cmdLimpiar_Click()
lstSeleccionados.Clear
If chkSoloActivos.Value = vbChecked Then
    CargaDiputadosActivos
Else
    CargaDiputados
End If
End Sub

Private Sub cmdMoverDerecha_Click()
Dim i As Integer
If lstDiputados.ListIndex > -1 Then
    i = lstDiputados.ListIndex
    lstSeleccionados.AddItem (lstDiputados.List(lstDiputados.ListIndex))
    lstDiputados.RemoveItem lstDiputados.ListIndex
    If i > lstDiputados.ListCount - 1 Then
        i = lstDiputados.ListCount - 1
    End If
    lstDiputados.ListIndex = i
    lstDiputados.SetFocus
End If
RevisaListas
End Sub

Private Sub cmdMoverIzquierda_Click()
Dim i As Integer
If lstSeleccionados.ListIndex > -1 Then
    lstDiputados.AddItem lstSeleccionados.List(lstSeleccionados.ListIndex)
    lstSeleccionados.RemoveItem lstSeleccionados.ListIndex
End If
RevisaListas
End Sub
Private Sub cmdPasarTodos_Click()
Dim i As Integer
RevisaListas
For i = 0 To lstDiputados.ListCount - 1
    lstSeleccionados.AddItem lstDiputados.List(i)
Next i
lstDiputados.Clear
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
Dim RsTemp As ADODB.Recordset
Dim consulta As String
Dim Indice As Integer
Indice = 0
ActualizarLista
dtDesde.Day = Format(Now(), "dd")
dtDesde.Month = Format(Now(), "mm")
dtDesde.Year = Format(Now(), "YYYY")
CopiarFechas
Dim X As Printer
cmbImpresoras.Clear
For Each X In Printers
    cmbImpresoras.AddItem X.DeviceName
Next X
End Sub
Private Sub CopiarFechas()
dtHasta.Day = dtDesde.Day
dtHasta.Month = dtDesde.Month
dtHasta.Year = dtDesde.Year
End Sub

Private Sub lstDiputados_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
Private Sub txtApellido_Change()
If chkSoloActivos.Value = vbChecked Then
    CargaDiputadosActivosFiltrados (txtApellido.Text)
Else
    CargaDiputadosFiltrados (txtApellido.Text)
End If
RevisaListas
End Sub
Private Sub txtApellido_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
Private Function AgregaCero(Param As String) As String
Dim X As String
If Len(Param) = 1 Then
    X = "0" & Param
Else
    X = Param
End If
AgregaCero = X
End Function
Private Sub CargaDiputados()
Dim RsTemp As ADODB.Recordset
Set RsTemp = New ADODB.Recordset
lstDiputados.Clear
SetearRs "SELECT Legisladores.apellido + ', ' + Legisladores.nombre AS DFull, Legisladores.id, Legisladores.codigo_persona, Legisladores.Provincia, Legisladores.bloque_politico FROM Legisladores WHERE Legisladores.tipo = 1 ORDER BY Apellido", RsTemp
While Not RsTemp.EOF
    lstDiputados.AddItem RsTemp.Fields("DFull") & Space(200) & ";" & RsTemp.Fields("id") & ";" & RsTemp.Fields("codigo_persona") & ";" & RsTemp.Fields("Provincia") & ";" & RsTemp.Fields("bloque_politico")
    RsTemp.MoveNext
Wend
RsTemp.Close
Set RsTemp = Nothing
End Sub
Private Sub CargaDiputadosFiltrados(filtro As String)
Dim RsTemp As ADODB.Recordset
Set RsTemp = New ADODB.Recordset
lstDiputados.Clear
SetearRs "SELECT Legisladores.apellido + ', ' + Legisladores.nombre AS DFull, Legisladores.id, Legisladores.codigo_persona, Legisladores.Provincia, Legisladores.bloque_politico FROM Legisladores WHERE Legisladores.tipo = 1 AND Legisladores.Apellido LIKE '" & filtro & "%' ORDER BY Apellido", RsTemp
While Not RsTemp.EOF
    lstDiputados.AddItem RsTemp.Fields("DFull") & Space(200) & ";" & RsTemp.Fields("id") & ";" & RsTemp.Fields("codigo_persona") & ";" & RsTemp.Fields("Provincia") & ";" & RsTemp.Fields("bloque_politico")
    RsTemp.MoveNext
Wend
RsTemp.Close
Set RsTemp = Nothing
End Sub
Private Sub CargaDiputadosActivos()
Dim RsTemp As ADODB.Recordset
Set RsTemp = New ADODB.Recordset
lstDiputados.Clear
SetearRs "SELECT Legisladores_activos.apellido + ', ' + Legisladores_activos.nombre AS DFull, Legisladores_activos.id, Legisladores.codigo_persona, Legisladores.Provincia, Legisladores.bloque_politico FROM Legisladores_activos ORDER BY Apellido", RsTemp
While Not RsTemp.EOF
    lstDiputados.AddItem RsTemp.Fields("DFull") & Space(200) & ";" & RsTemp.Fields("id") & ";" & RsTemp.Fields("codigo_persona") & ";" & RsTemp.Fields("Provincia") & ";" & RsTemp.Fields("bloque_politico")
    RsTemp.MoveNext
Wend
RsTemp.Close
Set RsTemp = Nothing
End Sub
Private Sub CargaDiputadosActivosFiltrados(filtro As String)
Dim RsTemp As ADODB.Recordset
Set RsTemp = New ADODB.Recordset
lstDiputados.Clear
SetearRs "SELECT Legisladores_activos.apellido + ', ' + Legisladores_activos.nombre AS DFull, Legisladores_activos.id, Legisladores.codigo_persona, Legisladores.Provincia, Legisladores.bloque_politico FROM Legisladores_activos WHERE Apellido LIKE '" & filtro & "%' ORDER BY Apellido", RsTemp
While Not RsTemp.EOF
    lstDiputados.AddItem RsTemp.Fields("DFull") & Space(200) & ";" & RsTemp.Fields("id") & ";" & RsTemp.Fields("codigo_persona") & ";" & RsTemp.Fields("Provincia") & ";" & RsTemp.Fields("bloque_politico")
    RsTemp.MoveNext
Wend
RsTemp.Close
Set RsTemp = Nothing
End Sub
Private Sub RevisaListas()
Dim i As Integer
Dim b As Integer
If lstSeleccionados.ListCount > 0 Then
    For i = 0 To lstSeleccionados.ListCount - 1
        For b = 0 To lstDiputados.ListCount - 1
            If lstDiputados.List(b) = lstSeleccionados.List(i) Then
                lstDiputados.RemoveItem b
            End If
        Next b
    Next i
End If
End Sub
