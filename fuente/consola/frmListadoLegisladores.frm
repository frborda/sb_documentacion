VERSION 5.00
Begin VB.Form frmListadoLegisladores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Legisladores"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Filtro por Sexo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   32
      Top             =   2280
      Width           =   4335
      Begin VB.ComboBox cmbSexo 
         Height          =   315
         ItemData        =   "frmListadoLegisladores.frx":0000
         Left            =   120
         List            =   "frmListadoLegisladores.frx":000D
         TabIndex        =   33
         Text            =   "Ningun tipo de filtro (Se muestran ambos sexos)"
         Top             =   330
         Width           =   4095
      End
   End
   Begin VB.Frame grpApellido 
      Caption         =   "Filtro por Apellido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   0
      TabIndex        =   27
      Top             =   2280
      Width           =   4320
      Begin VB.ComboBox cmbApellido2 
         Height          =   315
         ItemData        =   "frmListadoLegisladores.frx":004B
         Left            =   720
         List            =   "frmListadoLegisladores.frx":0052
         TabIndex        =   31
         Text            =   "Ningun tipo de filtro (Se muestran todos)"
         Top             =   795
         Width           =   3225
      End
      Begin VB.ComboBox cmbApellido1 
         Height          =   315
         ItemData        =   "frmListadoLegisladores.frx":0081
         Left            =   720
         List            =   "frmListadoLegisladores.frx":0088
         TabIndex        =   30
         Text            =   "Ningun tipo de filtro (Se muestran todos)"
         Top             =   315
         Width           =   3225
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta :"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Desde :"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame grpSeccion 
      Caption         =   "Filtro por Sección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   25
      Top             =   1320
      Width           =   4335
      Begin VB.ComboBox cmbSeccion 
         Height          =   315
         ItemData        =   "frmListadoLegisladores.frx":00B7
         Left            =   120
         List            =   "frmListadoLegisladores.frx":00BE
         TabIndex        =   26
         Text            =   "Ningun tipo de filtro (Sin Sección)"
         Top             =   330
         Width           =   4095
      End
   End
   Begin VB.Frame grmAgrupacion 
      Caption         =   "Filtro por Agrupación Política"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   23
      Top             =   1320
      Width           =   4335
      Begin VB.ComboBox cmbAgrupaciones 
         Height          =   315
         ItemData        =   "frmListadoLegisladores.frx":00E7
         Left            =   120
         List            =   "frmListadoLegisladores.frx":00EE
         TabIndex        =   24
         Text            =   "Ningun tipo de filtro (Se muestran todos)"
         Top             =   330
         Width           =   4095
      End
   End
   Begin VB.Frame grpBloque 
      Caption         =   "Filtro por Bloque Político"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   21
      Top             =   360
      Width           =   4335
      Begin VB.ComboBox cmbBloques 
         Height          =   315
         ItemData        =   "frmListadoLegisladores.frx":011D
         Left            =   120
         List            =   "frmListadoLegisladores.frx":0124
         TabIndex        =   22
         Text            =   "Ningun tipo de filtro (Se muestran todos)"
         Top             =   330
         Width           =   4095
      End
   End
   Begin VB.Frame grpFiltroActivos 
      Caption         =   "Filtro de Legisladores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   18
      Top             =   360
      Width           =   4335
      Begin VB.OptionButton optListarTodos 
         Caption         =   "Listar todos"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optSoloActivos 
         Caption         =   "Listar solo activos"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   8760
      Width           =   4335
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "frmListadoLegisladores.frx":0153
      Top             =   8040
      Width           =   6615
   End
   Begin VB.CommandButton cmdPorAlfabeticoControl 
      Caption         =   "Alfabético para Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   8040
      Width           =   4335
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "frmListadoLegisladores.frx":019D
      Top             =   7320
      Width           =   6615
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmListadoLegisladores.frx":01FA
      Top             =   6720
      Width           =   6615
   End
   Begin VB.TextBox lblPorDistrito 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "frmListadoLegisladores.frx":0244
      Top             =   6120
      Width           =   6615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frmListadoLegisladores.frx":028D
      Top             =   5520
      Width           =   6615
   End
   Begin VB.TextBox lblAgrupacionEtiqueta 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "frmListadoLegisladores.frx":02D0
      Top             =   4920
      Width           =   6615
   End
   Begin VB.TextBox txtPorBanca 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmListadoLegisladores.frx":0317
      Top             =   4320
      Width           =   6615
   End
   Begin VB.TextBox txtListadoAlfabetico 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frmListadoLegisladores.frx":0353
      Top             =   3720
      Width           =   6615
   End
   Begin VB.CommandButton cmdConFotografias 
      Caption         =   "Con Fotografías"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   7320
      Width           =   4335
   End
   Begin VB.CommandButton cmdPorSexo 
      Caption         =   "Por Sexo y Bloque Político"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   6720
      Width           =   4335
   End
   Begin VB.CommandButton cmdPorProvincia 
      Caption         =   "Por Secciones Electorales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   6120
      Width           =   4335
   End
   Begin VB.CommandButton cmdPorBloquePolitico 
      Caption         =   "Por Bloque Político"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   5520
      Width           =   4335
   End
   Begin VB.CommandButton cmdPorAgrupacionPolitica 
      Caption         =   "Por Agrupación Política"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   4920
      Width           =   4335
   End
   Begin VB.CommandButton cmdPorBanca 
      Caption         =   "Por Número de Orden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   4320
      Width           =   4335
   End
   Begin VB.CommandButton cmdPorAlfabetico 
      Caption         =   "Alfabético"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   4335
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   11040
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   11040
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Label Label1 
      Caption         =   "Listar/Imprimir/Exportar por orden :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmListadoLegisladores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Rs As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1
Dim Combo2Actualizado As Boolean
Private Sub cmdConFotografias_Click()
Call imprimirUnActa("fotografia")
End Sub

Private Sub cmdPorAgrupacionPolitica_Click()
Call imprimirUnActa("agrupacion")
End Sub
Private Sub cmdPorAlfabetico_Click()
'imprimirUnActa "alfabetico"
imprimirUnActa "alfabetico"
End Sub
Private Sub cmdPorAlfabetico_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub cmdPorAlfabeticoControl_Click()
imprimirUnActa "controlid"
End Sub

Private Sub cmdPorBanca_Click()
imprimirUnActa "banca"
End Sub
Private Sub cmdPorBloquePolitico_Click()
Call imprimirUnActa("bloque")
End Sub

Private Sub cmdPorProvincia_Click()
imprimirUnActa ("secciones")
End Sub

Private Sub cmdPorSexo_Click()
imprimirUnActa ("sexoybloque")
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Public Sub imprimirUnActa(strTipoOperacion As String)
    On Error GoTo TrapError


    If PermisosTotales.ConsultaActas = 0 Then 'por corregir
        MsgBox "No posee permisos para consulta de actas", vbInformation + vbOKOnly, "Usuario sin permisos!!"
        Exit Sub
    End If
    Dim m_Report
    Dim rstListado  As New ADODB.Recordset
    Dim rstImagen As New ADODB.Recordset
    Dim sql      As String
    Dim ComandoExiste As Boolean
    Dim Filtro_Bloque As String
    Dim Filtro_Seccion As String
    Dim Filtro_Agrupacion As String
    Dim Filtro_Nombre1 As String
    Dim Filtro_Nombre2 As String
    Dim Filtro_Sex As String
    Dim sql_activos As String
    Dim sql_todos As String
    If cmbSeccion.ListIndex > 0 Then
        Filtro_Seccion = cmbSeccion.List(cmbSeccion.ListIndex)
    End If
    If cmbBloques.ListIndex > 0 Then
        Filtro_Bloque = cmbBloques.List(cmbBloques.ListIndex)
    End If
    If cmbAgrupaciones.ListIndex > 0 Then
        Filtro_Agrupacion = cmbAgrupaciones.List(cmbAgrupaciones.ListIndex)
    End If
    If cmbSexo.ListIndex > 0 Then
        Filtro_Sex = cmbSexo.List(cmbSexo.ListIndex)
    End If
    If cmbApellido1.ListIndex > 0 And cmbApellido2.ListIndex > 0 Then
        Filtro_Nombre1 = cmbApellido1.List(cmbApellido1.ListIndex)
        Filtro_Nombre2 = cmbApellido2.List(cmbApellido2.ListIndex)
    End If
    If cmbApellido1.ListIndex > 0 And cmbApellido2.ListIndex = 0 Then
        MsgBox "El filtro de apellidos no se aplicará : " & vbCrLf _
        & "Deben seleccionarse dos apellidos."
    End If
    If cmbApellido2.ListIndex > 0 And cmbApellido1.ListIndex = 0 Then
        MsgBox "El filtro de apellidos no se aplicará : " & vbCrLf _
        & "Deben seleccionarse dos apellidos."
    End If
    ComandoExiste = False
    sql_activos = "SELECT     Legisladores.id,Legisladores.nombre, Legisladores.apellido, Legisladores.sexo, Legisladores.grupo_politico, CASE WHEN Legisladores.bloque_politico = '' THEN 'Sin Bloque Político' ELSE Legisladores.bloque_politico END As bloque_politico, Legisladores.PICTURE, " _
              & "Legisladores.fotografia," & SQL_Provincia & " AS Distrito,distritos.distrito AS Distrito_Solo,secciones.seccion AS Seccion_Sola " _
              & ",ISNULL(legisladores_activos.DESKID,Legisladores.IndiceBanca) AS Banca FROM Legisladores INNER JOIN " _
              & "distritos ON Legisladores.distrito = distritos.id_distrito INNER JOIN " _
              & "secciones ON distritos.seccion = secciones.id_seccion INNER JOIN" _
              & " legisladores_activos ON Legisladores.id = legisladores_activos.ID"
    sql_todos = "SELECT     Legisladores.id,Legisladores.nombre, Legisladores.apellido, Legisladores.sexo, Legisladores.grupo_politico, CASE WHEN Legisladores.bloque_politico = '' THEN 'Sin Bloque Político' ELSE Legisladores.bloque_politico END As bloque_politico, Legisladores.PICTURE, " _
              & "Legisladores.fotografia," & SQL_Provincia & " AS Distrito,distritos.distrito AS Distrito_Solo,secciones.seccion AS Seccion_Sola " _
              & ",ISNULL(legisladores_activos.DESKID,Legisladores.IndiceBanca) AS Banca FROM Legisladores INNER JOIN " _
              & "distritos ON Legisladores.distrito = distritos.id_distrito INNER JOIN " _
              & "secciones ON distritos.seccion = secciones.id_seccion LEFT OUTER JOIN" _
              & " legisladores_activos ON Legisladores.id = legisladores_activos.ID"
    Select Case strTipoOperacion
    Case "alfabetico"
        Set m_Report = New rptListado
        m_Report.Caption = "Por Orden Alfabético"
        m_Report.lblOrden.Caption = "Por Orden Alfabético"
        If optSoloActivos.Value = True Then
            sql = sql_activos
        Else
            sql = sql_todos
        End If
        sql = sql & FiltrarSQL(Filtro_Seccion, Filtro_Bloque, Filtro_Agrupacion, Filtro_Nombre1, Filtro_Nombre2, Filtro_Sex) & " ORDER BY Legisladores.apellido"
        ComandoExiste = True
        If Not DISTRITO_HABILITADO Then
            m_Report.lblSeccionDistrito.Visible = False
            m_Report.fldDistrito.Visible = False
        End If
    Case "agrupacion"
        Set m_Report = New rptAgrupacionPolitica
        m_Report.Caption = "Ordenado por Agrupación Política"
        m_Report.lblOrden.Caption = "Ordenado por Agrupación Política"
        If optSoloActivos.Value = True Then
            sql = sql_activos
        Else
            sql = sql_todos
        End If
        sql = sql & FiltrarSQL(Filtro_Seccion, Filtro_Bloque, Filtro_Agrupacion, Filtro_Nombre1, Filtro_Nombre2, Filtro_Sex) & " ORDER BY Legisladores.grupo_politico"
        ComandoExiste = True
    Case "banca"
        Set m_Report = New rptListado
        m_Report.Caption = "Ordenado por Nro. de Orden"
        m_Report.lblOrden.Caption = "Ordenado Por Nro. de Orden"
        If optSoloActivos.Value = True Then
            sql = sql_activos
        Else
            sql = sql_todos
        End If
        sql = sql & FiltrarSQL(Filtro_Seccion, Filtro_Bloque, Filtro_Agrupacion, Filtro_Nombre1, Filtro_Nombre2, Filtro_Sex) & " ORDER BY Banca"
        ComandoExiste = True
        If Not DISTRITO_HABILITADO Then
            m_Report.lblSeccionDistrito.Visible = False
            m_Report.fldDistrito.Visible = False
        End If
    Case "bloque"
        Set m_Report = New rptAgrupacionPolitica
        m_Report.Caption = m_Report.Caption & " Ordenado por Bloque Político"
        m_Report.lblOrden.Caption = "Ordenado por Bloque Político"
        m_Report.GroupHeader1.DataField = "bloque_politico"
        m_Report.fldGroupTitle.DataField = "bloque_politico"
        m_Report.lblBloque.Caption = "Agrupación Política"
        m_Report.fldBloque.DataField = "grupo_politico"
        If Not AGRUPACION_POLITICA_HABILITADA Then
            m_Report.lblBloque.Visible = False
            m_Report.fldBloque.Visible = False
        End If
        If Not DISTRITO_HABILITADO Then
            m_Report.lblSeccionDistrito.Visible = False
            m_Report.fldDistrito.Visible = False
        End If
        If optSoloActivos.Value = True Then
            sql = sql_activos
        Else
            sql = sql_todos
        End If
        sql = sql & FiltrarSQL(Filtro_Seccion, Filtro_Bloque, Filtro_Agrupacion, Filtro_Nombre1, Filtro_Nombre2, Filtro_Sex) & " ORDER BY Legisladores.bloque_politico"
        ComandoExiste = True
    Case "secciones"
        Set m_Report = New rptAgrupacionPolitica
        m_Report.Caption = m_Report.Caption & " Ordenado por Secciones Electorales"
        m_Report.lblOrden.Caption = "Ordenado por Secciones Electorales"
        m_Report.GroupHeader1.DataField = "Seccion_Sola"
        m_Report.fldGroupTitle.DataField = "Seccion_Sola"
        m_Report.lblBloque.Caption = "Agrupación Política"
        m_Report.fldBloque.DataField = "grupo_politico"
        If Not AGRUPACION_POLITICA_HABILITADA Then
            m_Report.lblBloque.Visible = False
            m_Report.fldBloque.Visible = False
        End If
        m_Report.lblSeccionDistrito.Caption = "Distrito"
        m_Report.fldDistrito.DataField = "Distrito_Solo"
        
        If optSoloActivos.Value = True Then
            sql = sql_activos
        Else
            sql = sql_todos
        End If
        sql = sql & FiltrarSQL(Filtro_Seccion, Filtro_Bloque, Filtro_Agrupacion, Filtro_Nombre1, Filtro_Nombre2, Filtro_Sex) & " ORDER BY Seccion_Sola"
        ComandoExiste = True
    Case "sexoybloque"
        Set m_Report = New rptSexo
        If Not AGRUPACION_POLITICA_HABILITADA Then
            m_Report.lblBloque.Visible = False
            m_Report.fldBloque.Visible = False
        End If
        If Not DISTRITO_HABILITADO Then
            m_Report.lblSeccionDistrito.Visible = False
            m_Report.fldDistrito.Visible = False
        End If
        If optSoloActivos.Value = True Then
            sql = sql_activos
        Else
            sql = sql_todos
        End If
        sql = sql & FiltrarSQL(Filtro_Seccion, Filtro_Bloque, Filtro_Agrupacion, Filtro_Nombre1, Filtro_Nombre2, Filtro_Sex) & " ORDER BY sexo,bloque_politico"
        ComandoExiste = True
        
    Case "fotografia"
        Set m_Report = New rptFotografia
        If Not AGRUPACION_POLITICA_HABILITADA Then
            m_Report.fldAgrupacionPolitica.Visible = False
        End If
        If Not DISTRITO_HABILITADO Then
            m_Report.fldDistrito.Visible = False
        End If
        If optSoloActivos.Value = True Then
            sql = sql_activos
        Else
            sql = sql_todos
        End If
        sql = sql & FiltrarSQL(Filtro_Seccion, Filtro_Bloque, Filtro_Agrupacion, Filtro_Nombre1, Filtro_Nombre2, Filtro_Sex) & " ORDER BY sexo,bloque_politico,apellido,nombre"
        ComandoExiste = True
    Case "controlid"
        'Set m_Report = New rptControlID
        'If Not AGRUPACION_POLITICA_HABILITADA Then
        '    m_Report.fldAgrupacionPolitica.Visible = False
        '    m_Report.lblAgrupacionPolitica.Visible = False
        '    m_Report.fldBloque.Width = 2500
        'End If
        'If Not DISTRITO_HABILITADO Then
        '    m_Report.lblSeccionDistrito.Visible = False
        '    m_Report.fldDistrito.Visible = False
        'End If
        'If optSoloActivos.Value = True Then
        '    sql = sql_activos
        'Else
        '    sql = sql_todos
        'End If
        'sql = sql & FiltrarSQL(Filtro_Seccion, Filtro_Bloque, Filtro_Agrupacion, Filtro_Nombre1, Filtro_Nombre2, Filtro_Sex) & " ORDER BY Legisladores.apellido"
        'ComandoExiste = True
    End Select
    If ComandoExiste = True Then
        m_Report.lblFecha.Caption = Format(Date, "DD/MM/YYYY")
        m_Report.lblFecha.Caption = m_Report.lblFecha.Caption & " " & Format(Time, "HH:MM")
        SetearRs sql, rstListado
        Set m_Report.DataControl1.Recordset = rstListado
        If VistaPrevia = True Then
            m_Report.Show vbModal
        Else
            m_Report.Run False
            m_Report.PrintReport False
        End If
    End If
        Set rstListado = Nothing
        Set m_Report = Nothing
Exit Sub
TrapError:
    Select Case err.Number
        Case Else
            MsgBox "Error Nº " & err.Number & Chr(10) & err.Description & Chr(10) & "Originado en " & err.Source
            End
    End Select
Return
End Sub
Private Function FiltrarSQL(Filtro_S As String, Filtro_B As String, Filtro_A As String, Filtro_Apellido1 As String, Filtro_Apellido2 As String, Filtro_Sexo As String) As String
        If Filtro_S <> "" Then
            FiltrarSQL = " WHERE (secciones.seccion = " & "'" & Filtro_S & "')"
        End If
        If Filtro_B <> "" Then
            If FiltrarSQL = "" Then
                FiltrarSQL = " WHERE (Legisladores.bloque_politico = " & "'" & Filtro_B & "')"
            Else
                FiltrarSQL = FiltrarSQL & " AND (Legisladores.bloque_politico = " & "'" & Filtro_B & "')"
            End If
        End If
        If Filtro_A <> "" Then
            If FiltrarSQL = "" Then
                FiltrarSQL = " WHERE (Legisladores.grupo_politico = " & "'" & Filtro_A & "')"
            Else
                FiltrarSQL = FiltrarSQL & " AND (Legisladores.grupo_politico = " & "'" & Filtro_A & "')"
            End If
        End If
        If Filtro_Apellido1 <> "" And Filtro_Apellido2 <> "" Then
            If FiltrarSQL = "" Then
                FiltrarSQL = " WHERE (Legisladores.apellido >= " & "'" & Filtro_Apellido1 & "')" & " AND (Legisladores.apellido <= " & "'" & Filtro_Apellido2 & "')"
            Else
                FiltrarSQL = FiltrarSQL & " AND (Legisladores.apellido <= " & "'" & Filtro_Apellido2 & "')"
            End If
        End If
        If Filtro_Sexo <> "" Then
            Dim filtro As Integer
            If Filtro_Sexo = "Masculino" Then
                filtro = 1
            Else
                filtro = 0
            End If
            If FiltrarSQL = "" Then
                FiltrarSQL = " WHERE (Legisladores.sexo = " & "'" & filtro & "')"
            Else
                FiltrarSQL = FiltrarSQL & " AND (Legisladores.sexo = " & "'" & filtro & "')"
            End If
        End If
        If FiltrarSQL = "" Then
            FiltrarSQL = " WHERE (Legisladores.es_legislador = 1)"
        Else
            FiltrarSQL = FiltrarSQL & " AND (Legisladores.es_legislador = 1)"
        End If
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEnter Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    Combo2Actualizado = False
    LlenarCombo "bloques_politicos", cmbBloques
    LlenarCombo "agrupaciones_politicas", cmbAgrupaciones
    LlenarCombo "seccion", cmbSeccion
    LlenarCombo "apellidos", cmbApellido1
    Call HabilitarControles
End Sub
Private Sub HabilitarControles()
    
    lblAgrupacionEtiqueta.Visible = AGRUPACION_POLITICA_HABILITADA
    cmdPorAgrupacionPolitica.Visible = AGRUPACION_POLITICA_HABILITADA
    cmdPorProvincia.Visible = DISTRITO_HABILITADO
    lblPorDistrito.Visible = DISTRITO_HABILITADO
    
    grmAgrupacion.Visible = AGRUPACION_POLITICA_HABILITADA
    cmbAgrupaciones.Visible = AGRUPACION_POLITICA_HABILITADA
    grpSeccion.Visible = DISTRITO_HABILITADO
    cmbSeccion.Visible = DISTRITO_HABILITADO
    If Not AGRUPACION_POLITICA_HABILITADA Then
        cmbSeccion.Left = 120
        grpSeccion.Left = 0
    End If
    
End Sub
Private Sub LlenarCombo(tipo As String, combo As ComboBox)
    Dim strSql As String
    Dim temp As String
    Set Rs = New ADODB.Recordset
    Select Case tipo
    Case "bloques_politicos"
        strSql = "SELECT Bloque_Político FROM Bloques ORDER BY Bloque_Político"
    Case "agrupaciones_politicas"
        strSql = "SELECT Agrupación_Política FROM Grupos ORDER BY Agrupación_Política"
    Case "seccion"
        strSql = "SELECT seccion FROM secciones ORDER BY seccion"
    Case "apellidos"
        If optSoloActivos.Value = True Then
        strSql = "SELECT Legisladores.apellido,Legisladores.nombre , Legisladores.sexo, Legisladores.grupo_politico, Legisladores.bloque_politico, Legisladores.PICTURE, " _
             & "Legisladores.fotografia," & SQL_Provincia & " AS Distrito " _
             & "FROM Legisladores INNER JOIN " _
             & "distritos ON Legisladores.distrito = distritos.id_distrito INNER JOIN " _
             & "secciones ON distritos.seccion = secciones.id_seccion INNER JOIN " _
             & "legisladores_activos ON Legisladores.id = legisladores_activos.ID"
        Else
        strSql = "SELECT Legisladores.apellido,Legisladores.nombre, Legisladores.IndiceBanca,Legisladores.sexo, Legisladores.grupo_politico, Legisladores.bloque_politico, Legisladores.PICTURE, " _
              & "Legisladores.fotografia," & SQL_Provincia & " AS Distrito " _
              & "FROM Legisladores INNER JOIN " _
              & "distritos ON Legisladores.distrito = distritos.id_distrito INNER JOIN " _
              & "secciones ON distritos.seccion = secciones.id_seccion"
        End If
    End Select
    Datos.SetearRsW strSql, Rs
    temp = combo.List(0)
    combo.Clear
    combo.AddItem temp
    combo.Text = temp
    If tipo = "apellidos" Then
        temp = cmbApellido2.List(0)
        cmbApellido2.Clear
        cmbApellido2.Text = temp
        cmbApellido2.AddItem temp
    End If
    If Rs.EOF = False Then
        Rs.MoveFirst
        combo.AddItem Trim(Rs.Fields(0))
        If tipo = "apellidos" Then
            cmbApellido2.AddItem Trim(Rs.Fields(0))
        End If
        For i = 1 To (Rs.RecordCount - 1)
            Rs.MoveNext
            combo.AddItem Trim(Rs.Fields(0))
            If tipo = "apellidos" Then
                cmbApellido2.AddItem Trim(Rs.Fields(0))
            End If
        Next i
    End If
    Rs.Close
End Sub
Private Sub optListarTodos_Click()
LlenarCombo "apellidos", cmbApellido1
End Sub
Private Sub optSoloActivos_Click()
LlenarCombo "apellidos", cmbApellido1
End Sub
