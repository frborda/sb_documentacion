VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmImprimirTotalActas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imprimir todas las actas de una sesión"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmSesion 
      Caption         =   "Sesión XXX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   7320
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Continuar"
         Height          =   495
         Left            =   5400
         TabIndex        =   4
         Top             =   960
         Width           =   1815
      End
      Begin ComctlLib.ProgressBar Barra 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1560
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
         Min             =   1e-4
      End
      Begin VB.Label lblActasPendientes 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1755
         TabIndex        =   10
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1755
         TabIndex        =   9
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Acta XXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6120
         TabIndex        =   8
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Actas Pendientes : "
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Top             =   1110
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Actas Procesadas : "
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   750
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Actas Totales : "
         Height          =   195
         Left            =   270
         TabIndex        =   3
         Top             =   390
         Width           =   1395
      End
      Begin VB.Label lblActa 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1755
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmImprimirTotalActas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strListaActas As String
Private xSesion       As Long
Private xTotalActas() As Long


' Private Sub ImprimirActa(pSesion, pActa)
Public Sub ImprimirActa(mPeriodo As String, mSesion As String)
    On Error GoTo TrapError


    If PermisosTotales.ConsultaActas = 0 Then
        MsgBox "No posee permisos para consulta de actas", vbInformation + vbOKOnly, "Usuario sin permisos!!"
        Exit Sub
    End If

    Dim m_Report As New rptActas
    Dim rstActa  As New ADODB.Recordset
    'Dim fViewer  As frmVisor
    Dim sql      As String
   
    'Set fViewer = New frmVisor
    

        sql = " SELECT DetalleActas.Departamento, DetalleActas.Bloque_político , DetalleActas.Resultado, actas.Período_Legislativo, actas.Sesión, actas.Número_de_Acta, actas.Versión_Acta, actas.Ultima_Versión_Acta, actas.Nombre_del_Acta, actas.Fecha, actas.Hora, actas.Miembros_del_cuerpo, actas.Desempate, actas.Votacion,  actas.Presentes_Identificables, actas.Presidente, actas.Presentes_No_Identificables, actas.Presentes_Total, actas.Ausentes_Total, actas.Votos_Afirm_Identificables, actas.Votos_Afirm_No_Identificables, actas.Votos_Afirm_Desempate, actas.Votos_Afirm_Total, actas.Votos_Neg_Identificables, actas.Votos_Neg_No_Identificables,  actas.Votos_Neg_Desempate, actas.Votos_Neg_Total, actas.Abstenciones_Identificables, actas.Abstenciones_No_Identificables, actas.Abstenciones_Total, actas.Fecha_Modificacion, actas.Hora_Modificacion, actas.Usuario_Modificacion, actas.IP_Modificacion, actas.Observaciones, actas.Base_de_Mayoria AS BaseM,  actas.Tipo_de_Mayoria, actas.Miembros_del_cuerpo AS Expr1, " & _
              " actas.Tipo_de_operación , tipoop.Tipo_de_operación AS descTipoOp, TipoMayoriaQuorum.descripcion AS descTipoMayQuo, " & _
              " basemay.Descripcion AS descBaseMay, tipmay.Descripcion AS descTipoMay, RTRIM(Legisladores.apellido) + ', ' + RTRIM(Legisladores.nombre) AS Legislador, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre) AS DetalleLegislador ,LEFT(actas.Período_Legislativo,3) + ' Período ' +  case SUBSTRING(actas.Período_Legislativo,4,1) WHEN 'O' THEN 'Legislativo' WHEN 'E' THEN 'Extraordinario' WHEN 'P' THEN 'Preparatoria' END + ' - ' + CAST(actas.Sesión AS Varchar(5))+ 'ª Sesión ' + case SUBSTRING(actas.Período_Legislativo,5,2) WHEN 'T' THEN 'de Tablas' WHEN 'E' THEN 'Especial' WHEN 'A' THEN 'Asamblea Legislativa'WHEN 'O' THEN 'Ordinaria' WHEN 'X' THEN 'Extraordinaria' WHEN 'P' THEN 'Preparatoria' END + CASE WHEN '1' IS NULL THEN '' ELSE ' - Próximo Nº de acta: ' + CAST(actas.Número_de_Acta AS Varchar(5)) END AS DescripcionPeriodoLegislativo, TipoMayoriaQuorum.descripcion AS DescripcionTipoQuorum " & _
              " FROM actas LEFT OUTER JOIN detalleactas ON actas.Período_Legislativo = detalleactas.Período_Legislativo " & _
              " AND actas.Sesión = " & _
              " detalleactas.Sesión AND actas.Número_de_Acta = detalleactas.Nro_de_Acta AND Actas.Versión_Acta = detalleactas.Versión_Acta  LEFT OUTER JOIN Legisladores ON actas.Presidente = Legisladores.id LEFT OUTER JOIN Legisladores AS DetalleLegis ON detalleactas.Legislador_asignado = DetalleLegis.id LEFT OUTER JOIN tipmay ON actas.Tipo_de_Mayoria = tipmay.Tipo_de_Mayoria LEFT OUTER JOIN basemay ON actas.Base_de_Mayoria = basemay.identificador_en_mensajes LEFT OUTER JOIN TipoMayoriaQuorum ON actas.Tipo_de_Quorum = TipoMayoriaQuorum.codigo LEFT OUTER JOIN" & _
              " tipoop ON actas.Tipo_de_operación = tipoop.identificador_en_mensajes LEFT OUTER JOIN Paramorden P ON P.Resultado = DetalleActas.Resultado " & _
              " WHERE (Actas.Período_Legislativo='" & mPeriodo & "') AND (Actas.Sesión=" & mSesion & ") AND (Actas.Versión_Acta=0) AND Actas.Tipo_de_Operación = 'votnom' " & _
              " ORDER BY Actas.Número_de_Acta, P.Orden, RTRIM(DetalleLegis.apellido) + ', ' + RTRIM(DetalleLegis.nombre)"
        
        'm_Report.Section5.Suppress = True
        m_Report.Detail.Visible = False
        'm_Report.Cuadro1.Suppress = True
        'm_Report.Cuadro1.Visible = False 'REVAP-091029A
        'm_Report.Texto25.Suppress = True
        m_Report.Texto25.Visible = False
        'm_Report.Línea5.Suppress = True
        m_Report.Línea5.Visible = False
        'm_Report.Texto23.Suppress = True
        m_Report.Texto23.Visible = False
        'm_Report.Texto1.Suppress = True
        m_Report.Texto1.Visible = False
        'm_Report.Línea6.Suppress = True
        m_Report.Línea6.Visible = False
        'm_Report.Línea7.Suppress = True
        m_Report.Línea7.Visible = False
        'm_Report.Línea8.Suppress = True
'        m_Report.Línea8.Visible = False
        'm_Report.Línea9.Suppress = True
'        m_Report.Línea9.Visible = False
        'm_Report.Texto22.Suppress = True
        m_Report.Texto22.Visible = False
    
    SetearRs sql, rstActa
    'm_Report.Database.SetDataSource rstActa
    Set m_Report.DataControl1.Recordset = rstActa
    If VistaPrevia = True Then
        m_Report.Show vbModal
    Else
        m_Report.Run False
        m_Report.PrintReport False
    End If
'    fViewer.CRViewer1.ReportSource = m_Report
'    If PermisosTotales.ImprimeActas = 1 Then
'        fViewer.CRViewer1.EnablePrintButton = True
'    Else
'        fViewer.CRViewer1.EnablePrintButton = False
'    End If
'
'    fViewer.CRViewer1.ViewReport
'    fViewer.CRViewer1.Zoom 100
'    fViewer.Show vbModal
    
    Set rstActa = Nothing
    'Set fViewer = Nothing
    Set m_Report = Nothing
Exit Sub
TrapError:
    Select Case Err.Number
        Case Else
            MsgBox "Error Nº " & Err.Number & Chr(10) & Err.Description & Chr(10) & "Originado en " & Err.Source
            Resume
    End Select
Return

End Sub


Public Property Let ListaActas(ByVal vNewValue As Variant)
    strListaActas = vNewValue
    lblActa.Caption = strListaActas
    lblActasPendientes.Caption = strListaActas
End Property
Public Property Let SesionActual(ByVal vNewValue As Variant)
    xSesion = vNewValue
    frmSesion.Caption = "Sesión " & Trim(Str(xSesion)) & " :  "
End Property

Private Sub Command1_Click()
    Call ImprimirActa("126ot", "1")
End Sub

