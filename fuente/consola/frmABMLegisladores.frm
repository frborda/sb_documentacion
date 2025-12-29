VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmABMLegisladores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mantenimiento de Datos de Legisladores"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameActualizando 
      Caption         =   "Sincronizar Enrolador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   120
      TabIndex        =   57
      Top             =   6120
      Visible         =   0   'False
      Width           =   10455
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Actualizando la base de datos. Por favor espere."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         TabIndex        =   58
         Top             =   840
         Width           =   3795
      End
   End
   Begin VB.Data base_enrolador 
      Caption         =   "Access"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8340
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   100
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Número de Orden e Identificación Dactilar: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   120
      TabIndex        =   34
      Top             =   6120
      Width           =   10455
      Begin VB.CommandButton btnSincronizarEnrolador 
         Caption         =   "Sincronizar Enrolador"
         Height          =   495
         Left            =   390
         TabIndex        =   53
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CommandButton cmdDesasignarBanca 
         Caption         =   "Desasignar Nro. de Orden"
         Height          =   465
         Left            =   5400
         TabIndex        =   37
         ToolTipText     =   "Quita todas las bancas asignadas al legislador actual"
         Top             =   330
         Width           =   2085
      End
      Begin VB.TextBox lblBanca 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Left            =   390
         TabIndex        =   19
         Text            =   "0"
         Top             =   285
         Width           =   2805
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   465
         Left            =   3240
         TabIndex        =   20
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   820
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdAbrirArchivoHuella 
         Caption         =   "Abrir Archivo con Huella Dactilar"
         Height          =   525
         Left            =   7560
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "Asignar Nro. de Orden"
         Enabled         =   0   'False
         Height          =   465
         Left            =   3510
         TabIndex        =   21
         Top             =   330
         Width           =   1845
      End
      Begin VB.Label lblFlagHuella 
         AutoSize        =   -1  'True
         Caption         =   "Tiene huella asignada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   38
         Top             =   1770
         Width           =   1890
      End
      Begin VB.Label Label10 
         Caption         =   "Permite importar y exportar los datos de legisladores y tablas relacionadas"
         Height          =   375
         Left            =   3240
         TabIndex        =   35
         Top             =   1200
         Width           =   5835
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Legislador: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   31
      Top             =   1080
      Width           =   10455
      Begin VB.TextBox txtFotografia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   6960
         TabIndex        =   51
         Top             =   3840
         Width           =   2565
      End
      Begin VB.CommandButton btnElegirImagen 
         Caption         =   "Seleccionar imágen"
         Height          =   495
         Left            =   6960
         TabIndex        =   17
         Top             =   4200
         Width           =   2535
      End
      Begin VB.ComboBox cmbDistrito 
         Height          =   315
         Left            =   1920
         TabIndex        =   14
         Top             =   4320
         Width           =   4605
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   1080
         Width           =   2200
      End
      Begin VB.ComboBox cmbBloque 
         Height          =   315
         Left            =   1920
         TabIndex        =   13
         Top             =   3765
         Width           =   2200
      End
      Begin VB.ComboBox cmbSexo 
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   1875
         Width           =   2200
      End
      Begin VB.TextBox txtApellido 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   1500
         Width           =   2200
      End
      Begin VB.TextBox txtInicioActividad 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   2280
         Width           =   1245
      End
      Begin VB.ComboBox cmbMandato 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   2760
         Width           =   4605
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   41
         Top             =   360
         Width           =   2205
      End
      Begin VB.CheckBox chkPersonalMantenimiento 
         Caption         =   "Personal de Mantenimiento"
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
         Left            =   6690
         TabIndex        =   15
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtFechaNacimiento 
         Height          =   315
         Left            =   5910
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.ComboBox cmbAgrupacionPolitica 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   3240
         Width           =   2205
      End
      Begin VB.CheckBox chkEsLegislador 
         Caption         =   "Es Legislador"
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
         Left            =   6690
         TabIndex        =   16
         Top             =   960
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker dtFechaNacimiento 
         Height          =   315
         Left            =   7170
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20578305
         CurrentDate     =   38027
      End
      Begin MSComCtl2.DTPicker dtInicioActividad 
         Height          =   315
         Left            =   3240
         TabIndex        =   52
         Top             =   2280
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20578305
         CurrentDate     =   38027
      End
      Begin VB.Label lblIdentifica 
         Alignment       =   1  'Right Justify
         Caption         =   "Id.:"
         Height          =   255
         Left            =   960
         TabIndex        =   59
         Top             =   445
         Width           =   735
      End
      Begin VB.Image picFoto 
         Height          =   1935
         Left            =   7440
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Fotografía "
         Height          =   285
         Left            =   6600
         TabIndex        =   50
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label lblDistrito 
         Alignment       =   1  'Right Justify
         Caption         =   "Sección / Distrito  Electoral : "
         Height          =   405
         Left            =   120
         TabIndex        =   48
         Top             =   4230
         Width           =   1725
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre : "
         Height          =   285
         Left            =   1080
         TabIndex        =   47
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Bloque Político : "
         Height          =   285
         Left            =   480
         TabIndex        =   46
         Top             =   3795
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Apellido : "
         Height          =   285
         Left            =   1080
         TabIndex        =   45
         Top             =   1515
         Width           =   765
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Sexo : "
         Height          =   285
         Left            =   840
         TabIndex        =   44
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio de actividad : "
         Height          =   285
         Left            =   420
         TabIndex        =   43
         Top             =   2295
         Width           =   1425
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Mandato : "
         Height          =   285
         Left            =   480
         TabIndex        =   42
         Top             =   2775
         Width           =   1365
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha de Nacimiento : "
         Height          =   285
         Left            =   4035
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblAgrupacionPolitica 
         Alignment       =   1  'Right Justify
         Caption         =   "Agrupación Política : "
         Height          =   285
         Left            =   240
         TabIndex        =   32
         Top             =   3285
         Width           =   1605
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   915
      Left            =   100
      ScaleHeight     =   855
      ScaleWidth      =   6225
      TabIndex        =   29
      Top             =   50
      Width           =   6280
      Begin VB.CommandButton Borrar 
         Caption         =   "Eliminar"
         Height          =   855
         Left            =   2490
         Picture         =   "frmABMLegisladores.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmBuscar 
         Caption         =   "&Buscar"
         Height          =   855
         Left            =   4980
         Picture         =   "frmABMLegisladores.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdPin 
         Caption         =   "PIN"
         Height          =   855
         Left            =   3735
         Picture         =   "frmABMLegisladores.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Nuevo 
         Caption         =   "&Nuevo"
         Height          =   855
         Left            =   0
         Picture         =   "frmABMLegisladores.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton Grabar 
         Caption         =   "&Grabar"
         Height          =   855
         Left            =   1245
         Picture         =   "frmABMLegisladores.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   915
      Left            =   6480
      ScaleHeight     =   855
      ScaleWidth      =   1245
      TabIndex        =   28
      Top             =   50
      Width           =   1300
      Begin VB.CommandButton Salir 
         Caption         =   "&Salir"
         Height          =   855
         Left            =   0
         Picture         =   "frmABMLegisladores.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   390
      Left            =   100
      ScaleHeight     =   330
      ScaleWidth      =   690
      TabIndex        =   27
      Top             =   8355
      Width           =   750
      Begin VB.CommandButton cmdPrimero 
         Height          =   325
         Left            =   0
         Picture         =   "frmABMLegisladores.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdAnterior 
         Height          =   325
         Left            =   345
         Picture         =   "frmABMLegisladores.frx":079E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   390
      Left            =   7080
      ScaleHeight     =   330
      ScaleWidth      =   690
      TabIndex        =   0
      Top             =   8355
      Width           =   750
      Begin VB.CommandButton cmdUltimo 
         Height          =   325
         Left            =   345
         Picture         =   "frmABMLegisladores.frx":0930
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdSiguiente 
         Height          =   325
         Left            =   0
         Picture         =   "frmABMLegisladores.frx":0AC2
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Actualizando la base de datos. Por favor espere."
      Height          =   435
      Left            =   7380
      TabIndex        =   56
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Actualizando la base de datos. Por favor espere."
      Height          =   435
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Actualizando la base de datos. Por favor espere."
      Height          =   435
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Bloque Político : "
      Height          =   285
      Left            =   840
      TabIndex        =   49
      Top             =   5430
      Width           =   1125
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha de Nacimiento : "
      Height          =   285
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   1785
   End
   Begin VB.Label lblId 
      Caption         =   "nothing"
      Height          =   255
      Left            =   8040
      TabIndex        =   39
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblRecordSet 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0/0 Legisladores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   990
      TabIndex        =   30
      Top             =   8355
      Width           =   5940
   End
End
Attribute VB_Name = "frmABMLegisladores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rs            As ADODB.Recordset
Dim RsB           As ADODB.Recordset ' RecordSet usado para buscar la banca asignada a un legislador
Attribute RsB.VB_VarHelpID = -1
Dim RsHuella      As ADODB.Recordset ' RecordSet usado para almacenar las huellas
Dim Directorio_Imagen As String
Dim Puede_Cargar_Foto As Boolean
Dim Extension_Imagen As String
Private strHuella As String
Private strHuellas(10) As String
Public arrSecciones As String
Private blEsNuevo As Boolean
Const VERSION_ENROLADOR = 1101 'hcdn 2011, 906 La plata
Public Sub VersionDatosSqv()
    Dim strSQLInsert As String
    strSQLInsert = "UPDATE config SET version_datos_sqv='" & Replace(Date, "/", "_") & "_" & Replace(Time(), ":", "") & "'"
    SenteciaSQl strSQLInsert
End Sub
Private Function buscarIndexDistrito(strDistrito)
    Dim arrSec, sec
    
    Dim nPointer As Integer
    buscarIndexDistrito = 1
    arrSec = Split(arrSecciones, "||")
    'MsgBox (uboun(arrSec))
    For nPointer = 0 To UBound(arrSec) - 1
        sec = Split(arrSec(nPointer), "|")
        If CInt(sec(0)) = CInt(strDistrito) Then
            buscarIndexDistrito = sec(1)
            Exit Function
        End If
    Next
End Function
Private Function buscarIndexSeccion(strDistrito)
    Dim arrSec, sec
    
    Dim nPointer As Integer
    buscarIndexSeccion = 1
    arrSec = Split(arrSecciones, "||")
    'MsgBox (uboun(arrSec))
    For nPointer = 0 To UBound(arrSec) - 1
        sec = Split(arrSec(nPointer), "|")
        If sec(0) = strDistrito Then
            buscarIndexSeccion = sec(1)
            Exit Function
        End If
    Next
End Function

Private Sub Borrar_Click()
    Dim xRespuesta As Integer
    If Rs.RecordCount > 0 Then
        xRespuesta = MsgBox("¿Seguro de eliminar al legislador actual?", vbQuestion + vbYesNo)
        If xRespuesta = vbYes Then
            'borrar las huellas en cascada
            SenteciaSQl "DELETE FROM huellas where idlegislador='" & Rs.Fields("id") & "'"
            'borra de legisladores activos
            SenteciaSQl "DELETE FROM legisladores_activos where id='" & Rs.Fields("id") & "'"
            'borro de legisladores el registro activo
            Rs.Delete
            Call VersionDatosSqv
            Call Limpiar
            If Rs.RecordCount > 0 Then
                Rs.MoveFirst
                Call MostrarRegistro
            End If
        End If
    End If
End Sub

Private Sub btnElegirImagen_Click()
    Dim strArchivo As String
    Dim ArrayArchivo() As String
    ReDim ArrayArchivo(0 To 1)
    ArrayArchivo = BuscarArchivoImagen
    Directorio_Imagen = ArrayArchivo(0)
    strArchivo = ArrayArchivo(0)   ' Elegir archivo
    If strArchivo <> "" Then
        txtFotografia.Text = ArrayArchivo(1)
        picFoto.Picture = LoadPicture(Directorio_Imagen)
    End If
End Sub

Private Sub btnSincronizarEnrolador_Click()
    'sincronizacion
    Dim sPathBase As String
    Dim nImportados, nExportados As Integer
    Dim cnn As ADODB.Connection
    Dim strSql, strSQLInsert As String
    Dim nRegistrosNuevos As Integer
    Dim strIdLegislador As String
    
    Dim Nuevo As New FileSystemObject, Nuevo1
    nRegistrosNuevos = 0
        
    If strArchivoEnrolamiento <> "" Then
        sPathBase = strDirectorioEnrolamiento & strArchivoEnrolamiento
        If Dir(sPathBase) <> "" Then
            nRegistrosNuevos = 1
        Else
            MsgBox "No existe el archivo .MDB" & Chr(13) & "Archivo: " & sPathBase, vbCritical + vbOKOnly, "Actualización de la base de datos de Legisladores"
        End If
    Else
        MsgBox "El archivo .MDB no está configurado.", vbCritical + vbOKOnly, "Actualización de la base de datos de Legisladores"
    End If
    
    If nRegistrosNuevos > 0 Then
        Dim confirmaactualizacion
        confirmaactualizacion = MsgBox("¿Actualizar la base de datos de Legisladores?", vbYesNo + vbInformation + vbSystemModal, "Actualización de la base de datos de Legisladores")
        'MsgBox confirmaactualizacion
        If confirmaactualizacion = 6 Then 'si
            frameActualizando.Visible = True
            nImportados = 0
            nExportados = 0
            'genera el backup de enrolamiento (importacion)
            If Dir(strDirectorioEnrolamiento & strArchivoEnrolamiento) <> "" Then
                ' Dim Nuevo As New FileSystemObject, Nuevo1
                 Dim strDestino As String
                 strDestino = strDirectorioEnrolamiento & "RESGUARDO\Enrolador_resguardo_" & Replace(Date, "/", "_") & "_" & Replace(Time, ":", "") & ".mdb"
                 'MsgBox (strDestino)
                 Set Nuevo1 = Nuevo.GetFile(strDirectorioEnrolamiento & strArchivoEnrolamiento)
                 Nuevo1.Copy (strDestino)
            End If
            '-------------------------------------------------------
            'copio del access al SQL los que tienen cambios
            '-------------------------------------------------------
            
            Dim RsTemp As ADODB.Recordset
            Dim rsLegisladorAccess As Integer
            Dim rst As ADODB.Recordset
            Dim n As Integer
            Dim fecha As String
            
            Dim strHuellaLimpia As String
                        
            Set cnn = New ADODB.Connection
            Set rst = New ADODB.Recordset
            cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51; Data Source=" & sPathBase & ";"
            cnn.Open
                        
            base_enrolador = sPathBase
            base_enrolador.DatabaseName = sPathBase
            'base_enrolador.Exclusive = True
            base_enrolador.RecordSource = "Legisladores"
            base_enrolador.Refresh
            'base_enrolador.Recordset.MoveFirst
            base_enrolador.Recordset.FindFirst ("modifica_sqv=1")
            If Not base_enrolador.Recordset.EOF Then '(False = base_enrolador.Recordset.NoMatch) Then
                
                rsLegisladorAccess = base_enrolador.Recordset.AbsolutePosition
                'recorro el access
                While (Not base_enrolador.Recordset.EOF) And (False = base_enrolador.Recordset.NoMatch)
                
                    
                    If Trim(base_enrolador.Recordset.Fields("id")) <> "" Then
                        strIdLegislador = Trim(base_enrolador.Recordset.Fields("id"))
                        If IsDate(base_enrolador.Recordset.Fields("inicio_actividad").Value) Then
                            fecha = base_enrolador.Recordset.Fields("inicio_actividad").Value
                        Else
                            fecha = "01/01/1900"
                        End If
                        'me fijo si existe
                        Set RsTemp = New ADODB.Recordset
                        SetearRsW "SELECT id FROM legisladores where id='" & strIdLegislador & "'", RsTemp
                        'recorro el SQLSERVER
                        If Not RsTemp.EOF Then
                        'si existe hago UPDATE
                            strSQLInsert = "UPDATE legisladores SET "
                            strSQLInsert = strSQLInsert & "nombre='" & Trim(base_enrolador.Recordset.Fields("nombre")) & "',"
                            strSQLInsert = strSQLInsert & "apellido='" & Trim(base_enrolador.Recordset.Fields("apellido")) & "',"
                            strSQLInsert = strSQLInsert & "sexo=" & (base_enrolador.Recordset.Fields("sexo")) & ","
                            strSQLInsert = strSQLInsert & "bloque_politico='" & Trim(base_enrolador.Recordset.Fields("bloque_politico")) & "',"
                            strSQLInsert = strSQLInsert & "grupo_politico='" & Trim(base_enrolador.Recordset.Fields("grupo_politico")) & "',"
                            strSQLInsert = strSQLInsert & "Tipo='" & Trim(base_enrolador.Recordset.Fields("Tipo")) & "',"
                            If Not IsNull(base_enrolador.Recordset.Fields("es_legislador").Value) Then
                                strSQLInsert = strSQLInsert & "es_legislador=" & base_enrolador.Recordset.Fields("es_legislador").Value & ","
                            Else
                                strSQLInsert = strSQLInsert & "es_legislador=0,"
                            End If
                            If Not IsNull(fecha) Then
                                strSQLInsert = strSQLInsert & "inicio_actividad='" & fecha & "',"
                            Else
                                strSQLInsert = strSQLInsert & "inicio_actividad='01/01/1900',"
                            End If
                            strSQLInsert = strSQLInsert & "mandato='" & Trim(base_enrolador.Recordset.Fields("mandato")) & "',"
                            If Not IsNull(base_enrolador.Recordset.Fields("distrito").Value) Then
                                strSQLInsert = strSQLInsert & "distrito=" & base_enrolador.Recordset.Fields("distrito").Value & ","
                            Else
                                strSQLInsert = strSQLInsert & "distrito=0,"
                            End If
                            strSQLInsert = strSQLInsert & "fotografia='" & Trim(base_enrolador.Recordset.Fields("fotografia")) & "',"
                            strSQLInsert = strSQLInsert & "Numero_Minucias='" & Trim(base_enrolador.Recordset.Fields("Numero_Minucias")) & "'"
                            strSQLInsert = strSQLInsert & " where id='" & strIdLegislador & "'"
                        Else 'no existe
                        
                        'si no existe hago INSERT
                            strIdLegislador = trae_ultimo_id_legislador()
                            strSQLInsert = "INSERT INTO legisladores (id,nombre,apellido,sexo,bloque_politico,grupo_politico,tipo,es_legislador,inicio_actividad,mandato,distrito,fotografia,numero_minucias) values ("
                            'strSQLInsert = strSQLInsert & "'" & Trim(base_enrolador.Recordset.Fields("id").Value) & "',"
                            strSQLInsert = strSQLInsert & "'" & strIdLegislador & "',"
                            
                            strSQLInsert = strSQLInsert & "'" & base_enrolador.Recordset.Fields("nombre").Value & "',"
                            strSQLInsert = strSQLInsert & "'" & base_enrolador.Recordset.Fields("apellido").Value & "',"
                            strSQLInsert = strSQLInsert & "" & base_enrolador.Recordset.Fields("sexo").Value & ","
                            strSQLInsert = strSQLInsert & "'" & base_enrolador.Recordset.Fields("bloque_politico").Value & "',"
                            strSQLInsert = strSQLInsert & "'" & base_enrolador.Recordset.Fields("grupo_politico").Value & "',"
                            strSQLInsert = strSQLInsert & "'" & base_enrolador.Recordset.Fields("tipo").Value & "',"
                            If Not IsNull(base_enrolador.Recordset.Fields("es_legislador").Value) Then
                                strSQLInsert = strSQLInsert & "" & base_enrolador.Recordset.Fields("es_legislador").Value & ","
                            Else
                                strSQLInsert = strSQLInsert & "0,"
                            End If
                            
                            If Not IsNull(fecha) Then
                                strSQLInsert = strSQLInsert & "'" & fecha & "',"
                            Else
                                strSQLInsert = strSQLInsert & "'01/01/1900',"
                            End If
                            
                            strSQLInsert = strSQLInsert & "'" & base_enrolador.Recordset.Fields("mandato").Value & "',"
                            If Not IsNull(base_enrolador.Recordset.Fields("distrito").Value) Then
                                strSQLInsert = strSQLInsert & "" & base_enrolador.Recordset.Fields("distrito").Value & ","
                            Else
                                strSQLInsert = strSQLInsert & "0,"
                            End If
                            strSQLInsert = strSQLInsert & "'" & base_enrolador.Recordset.Fields("fotografia").Value & "',"
                            strSQLInsert = strSQLInsert & "'" & base_enrolador.Recordset.Fields("numero_minucias").Value & "')"
                            
                        End If
                        'MsgBox (strSQLInsert)
                        RsTemp.Close
                        SenteciaSQl strSQLInsert
                        nImportados = nImportados + 1
                        strSQLInsert = "DELETE FROM huellas where idlegislador='" & strIdLegislador & "'"
                        SenteciaSQl strSQLInsert
                        
                        
                        For n = 1 To 10
                            If Not IsNull(base_enrolador.Recordset.Fields("huella_" & n).Value) Then
                                strHuellaLimpia = Trim(base_enrolador.Recordset.Fields("huella_" & n).Value)
                                'Replace(Trim(base_enrolador.Recordset.Fields("huella_" & n).Value), " ", "")
                            Else
                                strHuellaLimpia = ""
                            End If
                            
                            strSQLInsert = ""
                            strSQLInsert = "INSERT INTO huellas (huella,idlegislador,nrohuella,Indice_Huella) VALUES ("
                            strSQLInsert = strSQLInsert & "'" & strHuellaLimpia & "',"
                            strSQLInsert = strSQLInsert & "'" & strIdLegislador & "',"
                            strSQLInsert = strSQLInsert & "" & n & ","
                            strSQLInsert = strSQLInsert & "" & base_enrolador.Recordset.Fields("indice_huella" & n).Value & ")"
                            If strHuellaLimpia <> "" Then
                                SenteciaSQl strSQLInsert
                            End If
                        Next
                        
                        
                    End If
                    'RsTemp.MoveNext
                    base_enrolador.Recordset.FindNext ("modifica_sqv=1")
                Wend
                'r.Close
            End If
            cnn.Close
            
            
            base_enrolador.DatabaseName = ""
            base_enrolador = ""
            
            '-------------------------------------------------------
            'copio del SQL al access
            '-------------------------------------------------------
            'me fijo si existe el archivo template
            'para evitar hacer el shrink a mano
            '--------------------------------------
            'sPathBase = strDirectorioEnrolamiento & Replace(strArchivoEnrolamiento, ".mdb", ".plantilla")
            'If Dir(sPathBase, vbArchive) <> "" Then
            '    'existe lo copio y trabajo sobre el template
            '     Set cnn = Nothing
            '     strDestino = strDirectorioEnrolamiento & "\" & strArchivoEnrolamiento
            '     Set Nuevo1 = Nuevo.GetFile(strDirectorioEnrolamiento & strArchivoEnrolamiento)
            '     Nuevo1.Copy (strDestino)
            'End If
            '---------------------------------
            
            ' Bloque Político
            ' ------------------------------------------------------
            Set cnn = New ADODB.Connection
            Set rst = New ADODB.Recordset
            cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51; Data Source=" & sPathBase & ";"
            'cnn.Open
            
            Set RsTemp = New ADODB.Recordset
            SetearRs "SELECT * FROM Bloques ORDER BY Bloque_Político", RsTemp
            If RsTemp.RecordCount > 0 Then
                cnn.Open
                RsTemp.MoveFirst
                rst.Open "DELETE * FROM  Bloques", cnn, adOpenDynamic, adLockOptimistic
                While Not RsTemp.EOF
                    strSQLInsert = "INSERT INTO bloques (Bloque_Político,clave,bancaminima,bancamaxima) values ("
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(0).Value & "',"
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(1).Value & "',"
                    If Not IsNull(RsTemp.Fields(2).Value) Then
                        strSQLInsert = strSQLInsert & "" & RsTemp.Fields(2).Value & ","
                    Else
                        strSQLInsert = strSQLInsert & "0,"
                    End If
                    
                    If Not IsNull(RsTemp.Fields(3).Value) Then
                        strSQLInsert = strSQLInsert & "" & RsTemp.Fields(3).Value & ")"
                    Else
                        strSQLInsert = strSQLInsert & "0)"
                    End If
                    
                    rst.Open strSQLInsert, cnn, adOpenDynamic, adLockOptimistic
                    RsTemp.MoveNext
                Wend
                cnn.Close
            End If
            RsTemp.Close
            
            Set RsTemp = New ADODB.Recordset
            ' distrito electoral
            ' ------------------------------------------------------
            
            SetearRs "SELECT * FROM distritos ORDER BY id_distrito", RsTemp
            If RsTemp.RecordCount > 0 Then
                cnn.Open
                RsTemp.MoveFirst
                rst.Open "DELETE * FROM  distritos", cnn, adOpenDynamic, adLockOptimistic
                While Not RsTemp.EOF
                    strSQLInsert = "INSERT INTO distritos (id_distrito,seccion,distrito) values ("
                    strSQLInsert = strSQLInsert & "" & RsTemp.Fields(0).Value & ","
                    strSQLInsert = strSQLInsert & "" & RsTemp.Fields(1).Value & ","
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(2).Value & "')"
                    
                    rst.Open strSQLInsert, cnn, adOpenDynamic, adLockOptimistic
                    RsTemp.MoveNext
                Wend
                cnn.Close
            End If
            RsTemp.Close
            
            Set RsTemp = New ADODB.Recordset
            ' seccion electoral
            ' ------------------------------------------------------
            
            SetearRs "SELECT * FROM secciones ORDER BY id_seccion", RsTemp
            If RsTemp.RecordCount > 0 Then
                cnn.Open
                RsTemp.MoveFirst
                rst.Open "DELETE * FROM  secciones", cnn, adOpenDynamic, adLockOptimistic
                While Not RsTemp.EOF
                    strSQLInsert = "INSERT INTO secciones (id_seccion,seccion) values ("
                    strSQLInsert = strSQLInsert & "" & RsTemp.Fields(0).Value & ","
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(1).Value & "')"
                    rst.Open strSQLInsert, cnn, adOpenDynamic, adLockOptimistic
                    RsTemp.MoveNext
                Wend
                cnn.Close
            End If
            RsTemp.Close
            
            
             Set RsTemp = New ADODB.Recordset
            ' mandatos
            ' ------------------------------------------------------
            
            SetearRs "SELECT * FROM mandatos ORDER BY fecha_mandato", RsTemp
            If RsTemp.RecordCount > 0 Then
                cnn.Open
                RsTemp.MoveFirst
                rst.Open "DELETE * FROM  mandatos", cnn, adOpenDynamic, adLockOptimistic
                While Not RsTemp.EOF
                    strSQLInsert = "INSERT INTO mandatos (fecha_mandato) values ("
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(0).Value & "')"
                    rst.Open strSQLInsert, cnn, adOpenDynamic, adLockOptimistic
                    RsTemp.MoveNext
                Wend
                cnn.Close
            End If
            RsTemp.Close
            
             Set RsTemp = New ADODB.Recordset
            ' grupos politicos
            ' ------------------------------------------------------
            
            SetearRs "SELECT * FROM grupos ORDER BY Agrupación_Política", RsTemp
            If RsTemp.RecordCount > 0 Then
                cnn.Open
                RsTemp.MoveFirst
                rst.Open "DELETE * FROM  grupos", cnn, adOpenDynamic, adLockOptimistic
                While Not RsTemp.EOF
                    strSQLInsert = "INSERT INTO grupos (Agrupación_Política) values ("
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(0).Value & "')"
                    rst.Open strSQLInsert, cnn, adOpenDynamic, adLockOptimistic
                    RsTemp.MoveNext
                Wend
                cnn.Close
            End If
            RsTemp.Close
            
            'legisladores
            Dim RsTemp2 As ADODB.Recordset
            Set RsTemp2 = New ADODB.Recordset
            
            
            SetearRs "SELECT id,nombre,apellido,sexo,bloque_politico,grupo_politico,tipo,es_legislador,pin,inicio_actividad,mandato,distrito,fotografia,numero_minucias FROM legisladores ORDER BY id", RsTemp
            If RsTemp.RecordCount > 0 Then
                cnn.Open
                RsTemp.MoveFirst
                rst.Open "DELETE * FROM  legisladores", cnn, adOpenDynamic, adLockOptimistic
                While Not RsTemp.EOF
                    strSQLInsert = "INSERT INTO legisladores (id,nombre,apellido,sexo,bloque_politico,grupo_politico,tipo,es_legislador,inicio_actividad,mandato,distrito,fotografia,numero_minucias,modifica_sqv) values ("
                    strSQLInsert = strSQLInsert & "'" & Trim(RsTemp.Fields(0).Value) & "',"
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(1).Value & "',"
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(2).Value & "',"
                    strSQLInsert = strSQLInsert & "" & RsTemp.Fields(3).Value & ","
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(4).Value & "',"
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(5).Value & "',"
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(6).Value & "',"
                    strSQLInsert = strSQLInsert & "" & RsTemp.Fields(7).Value & ","
                    If Not IsNull(RsTemp.Fields(9).Value) Then
                        strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(9).Value & "',"
                    Else
                        strSQLInsert = strSQLInsert & "'01/01/1900',"
                    End If
                    
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(10).Value & "',"
                    If Not IsNull(RsTemp.Fields(11).Value) Then
                        strSQLInsert = strSQLInsert & "" & RsTemp.Fields(11).Value & ","
                    Else
                        strSQLInsert = strSQLInsert & "0,"
                    End If
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(12).Value & "',"
                    strSQLInsert = strSQLInsert & "'" & RsTemp.Fields(13).Value & "',"
                    strSQLInsert = strSQLInsert & "0)" 'modifica SQV
                    
                    'MsgBox (strSQLInsert)
                    rst.Open strSQLInsert, cnn, adOpenDynamic, adLockOptimistic
                    'grabo las huellas para este legislador
                    SetearRs "select * from huellas where idlegislador =" & (RsTemp.Fields(0).Value) & " order by nrohuella", RsTemp2
                    If RsTemp2.RecordCount > 0 Then
                        RsTemp2.MoveFirst
                        While Not RsTemp2.EOF
                            strSQLInsert = "UPDATE LEGISLADORES SET huella_" & RsTemp2("nrohuella").Value & "='" & RsTemp2("huella").Value & "',"
                            If Not IsNull(RsTemp2("indice_huella").Value) Then
                                strSQLInsert = strSQLInsert & "indice_huella" & RsTemp2("nrohuella").Value & "=" & RsTemp2("indice_huella").Value
                            Else
                                strSQLInsert = strSQLInsert & "indice_huella" & RsTemp2("nrohuella").Value & "=0 "
                            End If
                            strSQLInsert = strSQLInsert & " where id='" & Trim(RsTemp.Fields(0).Value) & "'"
                            rst.Open strSQLInsert, cnn, adOpenDynamic, adLockOptimistic
                            RsTemp2.MoveNext
                        Wend
                    End If
                    RsTemp2.Close
                    
                    nExportados = nExportados + 1
                    '-----------
                    RsTemp.MoveNext
                Wend
                cnn.Close
            End If
            RsTemp.Close
            
            'hay que trae la huella tambien para cada legislador las 10 huellas
            '--------------------------
            'genera backup de exportacion
            If Dir(strDirectorioEnrolamiento & strArchivoEnrolamiento) <> "" Then
                 
                 'Dim strDestino As String
                 strDestino = strDirectorioEnrolamiento & "RESGUARDO\SQV_resguardo_" & Replace(Date, "/", "_") & "_" & Replace(Time, ":", "") & ".mdb"
                 'MsgBox (strDestino)
                 Set Nuevo1 = Nuevo.GetFile(strDirectorioEnrolamiento & strArchivoEnrolamiento)
                 Nuevo1.Copy (strDestino)
            End If
            If nImportados > 0 Then
                Call VersionDatosSqv
            End If
            MsgBox "La sincronización de datos de Legisladores se realizó con éxito." & Chr(13) & "Se importaron " & nImportados & " legisladores." & Chr(13) & "Se exportaron " & nExportados & " legisladores.", vbExclamation + vbOKOnly, "Actualización de la base de datos de Legisladores"
            frameActualizando.Visible = False
            Unload Me
            frmABMLegisladores.Show vbModal
        Else
            MsgBox "No se ha realizado la sincronización de datos de Legisladores.", vbCritical + vbOKOnly, "Actualización de la base de datos de Legisladores"
        End If
    Else
        MsgBox "No hay registros para actualizar.", vbCritical + vbOKOnly, "Actualización de la base de datos de Legisladores"
    End If
End Sub

Private Sub cmBuscar_Click()
    Dim strId       As String
    Dim blCondicion As Boolean
    
    If PermisosTotales.ConsultaABMLegislador = 0 Then
        MsgBox "El usuario no tiene permisos para esta tarea", vbInformation + vbOKOnly, "Usuario sin permisos!!"
        Exit Sub
    End If
    
    frmrSelLegislador.Show vbModal, Me
    strId = Trim(lblId.Caption)
    blCondicion = True
    If LCase(strId) <> "nothing" And LCase(strId) <> "id" Then
        With Rs
            If .RecordCount > 0 Then
                .MoveFirst
                While blCondicion
                    If .Fields("id").Value = strId Then
                        Call MostrarRegistro
                        blCondicion = False
                    Else
                        .MoveNext
                    End If
                Wend
            End If
        End With
    End If
End Sub

Private Function ValorCampo(xEtiqueta As String, strCadena As String) As String

    Dim nDesde As Long
    Dim nLargo As Long
    
    nDesde = InStr(1, strCadena, xEtiqueta) + Len(Trim(xEtiqueta)) + 2
    If nDesde > 0 Then
        nLargo = InStr(nDesde, strCadena, vbCrLf) - nDesde
        ValorCampo = Trim(Mid(strCadena, nDesde, nLargo))
    Else
        ValorCampo = ""
    End If
End Function

Private Function ValorHuella(xEtiqueta As String, strCadena As String) As String

    Dim nDesde As Long
    Dim nLargo As Long
    
    nDesde = InStr(1, strCadena, "Huella " & Trim(xEtiqueta)) + Len(Trim("Huella " & xEtiqueta)) + 2
    If nDesde > 11 Then
        nLargo = InStr(nDesde, strCadena, "Indice del Template") - nDesde
        ValorHuella = Trim(Mid(strCadena, nDesde, nLargo))
    Else
        ValorHuella = ""
    End If
End Function
Private Sub cmdAbrirArchivoHuella_Click()
    
    
    Dim strArchivo     As String          ' Nombre del archivo a abrir
    Dim strCadena      As String          ' contenido del archivo abierto
  
    Dim strIndice      As String
    Dim strDni         As String
    Dim strApellido    As String
    Dim strNombre      As String
    
    Dim xDesp          As Long
    Dim X              As Long
    Dim xBloque        As String
    Dim xBloqueControl As String
    Dim nHasta As Long
    Dim nDesde As Long
    Dim i As Integer
    Dim nHuellasEncontradas As Integer
    
    

    
    Dim xRespuesta     As Long
    Dim strMensaje     As String
    Dim strHuellaResguardo As String
    
    strArchivo = BuscarArchivoHuella   ' Elegir archivo
    
    If Len(Trim(strArchivo)) > 0 Then  ' Obtener contenido del archivo ".huella"
        strCadena = LeerContenidoArchivo(strArchivo)
        If VERSION_ENROLADOR = 906 Then 'La plata
        
        Else 'CBA2003
            If Len(strCadena) Mod 144 <> 0 Then
                MsgBox "El archivo seleccionado no coincide con el formato del enrolador." & Chr(10) & "Reemplacelo y reintente la operación", vbCritical + vbOKOnly, "Error en formato de archivo huella."
                Exit Sub 'error
            End If
        End If
    Else
        Exit Sub
    End If
    ' ---------------------------------------------------------------
    ' Analizar contenido de archivo
    ' ---------------------------------------------------------------
    If VERSION_ENROLADOR = 906 Then 'La plata
        'strIndice = ""
        strDni = ValorCampo("D.N.I.", strCadena)
        strApellido = ValorCampo("Apellido", strCadena)
        strNombre = ValorCampo("Nombre", strCadena)
    Else 'CBA2003
        strIndice = Trim(Mid(strCadena, 9, 4))
        strDni = Trim(Mid(strCadena, 15, 8))
        strApellido = Trim(HexATexto(Mid(strCadena, 23, 60), 30))
        strNombre = Trim(HexATexto(Mid(strCadena, 83, 60), 30))
    End If
    ' ---------------------------------------------------------------
    ' Levantar la huella
    ' ---------------------------------------------------------------
    ' Inicializar y guardar la huella anterior
    If VERSION_ENROLADOR = 906 Then 'La plata
        nHuellasEncontradas = 0
        For i = 1 To 10
            strHuella = ""
            strHuella = ValorHuella(Trim(Str(i)), strCadena)
            strHuella = Replace(strHuella, vbCrLf, "")
            strHuella = Replace(strHuella, " ", "")
            strHuellas(i) = strHuella
            If strHuella > "" Then
                nHuellasEncontradas = nHuellasEncontradas + 1
            End If
        Next i
    Else
        strHuellaResguardo = strHuella
        strHuella = ""
        xDesp = 8
        xBloqueControl = "09"
        For X = 1 To Len(strCadena) Step 144 ' Leer archivo linea a linea
            ' Esta es una nueva linea - obtener indice de linea y bloque
            If Mid(strCadena, X, 8) <> "SRLEGI ^" Then
                'error
                MsgBox "Informacion no reconocida en posicion: " & Str(X) & " (linea: " & Int(Str(X / 144)) & ") " & vbCrLf & Mid(strCadena, X, 40) & "..."
                Exit Sub
            End If
            xBloque = Mid(strCadena, X + xDesp + 4, 2)
            If xBloque <> "00" Then
                xBloqueControl = xBloque
                strHuella = strHuella & Mid(strCadena, X + xDesp + 4 + 2, 128)
            End If
        Next X
    End If
    ' ---------------------------------------------------------------
    ' Verificar que los datos levantados del archivo sean de la misma
    ' persona que se esta editando en el formulario
    ' ---------------------------------------------------------------
    '& "                  ID: " & strIndice & Chr(10
    strMensaje = "Se han levantado los datos de: " & Chr(10) _
               & "            APELLIDO: " & strApellido & Chr(10) _
               & "              NOMBRE: " & strNombre & Chr(10) _
               & " Huellas Encontradas: " & Trim(Str(nHuellasEncontradas)) & Chr(10) & Chr(10) _
               & "      ¿Desea continuar con la operación?  "
               '& "                 DNI: " & strDni & Chr(10) _
    xRespuesta = MsgBox(strMensaje, vbInformation + vbYesNo)
    If xRespuesta = vbYes Then
        'txtId.Text = strIndice
        txtApellido = strApellido
        txtNombre = strNombre
        'txtDni = strDni
    Else
        ' si no confirma, vuelve a la huella anterior
        strHuella = strHuellaResguardo
    End If
          
End Sub


Private Sub MostrarRegistro()
    Dim xPos   As Long
    Dim xMax   As Long
    Dim strSql As String
    ' --------------------------------------------------------------------------------
    ' Mostrar Datos del legislador
    ' --------------------------------------------------------------------------------
    With Rs
        xPos = .AbsolutePosition
        xMax = .RecordCount
        If xMax > 0 Then
            lblRecordSet.Caption = Trim(Str(xPos)) & "/" & Trim(Str(xMax)) & " Legisladores"
        Else
            lblRecordSet.Caption = "0/0 Legisladores"
            Call Limpiar
        End If
        txtID.Text = .Fields("Id").Value
        txtApellido.Text = .Fields("Apellido").Value
        txtNombre.Text = .Fields("Nombre").Value
        cmbSexo.Text = IIf(.Fields("Sexo").Value = 1, "Masculino", "Femenino")
        cmbBloque.Text = IIf(IsNull(.Fields("Bloque_politico").Value), "", .Fields("Bloque_politico"))
        cmbAgrupacionPolitica.Text = IIf(IsNull(.Fields("grupo_politico").Value), "", .Fields("grupo_politico").Value)
        'txtDepartamento.Text = .Fields("Departamento").Value
        If Not IsNull(.Fields("fecha_nacimiento").Value) Then
            txtFechaNacimiento.Text = .Fields("fecha_nacimiento").Value
        Else
            txtFechaNacimiento.Text = ""
        End If
        If Not IsNull(.Fields("inicio_actividad").Value) Then
            txtInicioActividad.Text = .Fields("inicio_actividad").Value
        Else
            txtInicioActividad.Text = ""
        End If
        If Not IsNull(.Fields("distrito").Value) Then
           ' cmbDistrito.Text = .Fields("seccion").Value
            cmbDistrito.ListIndex = buscarIndexDistrito(Rs.Fields("distrito").Value) - 1
        Else
            cmbDistrito.Text = ""
            
        End If
       
        If Not IsNull(.Fields("mandato").Value) Then
            cmbMandato.Text = .Fields("mandato").Value
        Else
            cmbMandato.Text = ""
        End If
       
'        MsgBox (strDirectorioEnrolamiento & .Fields("fotografia").Value)
        If Not IsNull(.Fields("fotografia").Value) And Trim(.Fields("fotografia").Value) <> "" Then
                On Error GoTo SinFoto
                'If Dir(LCase(strDirectorioEnrolamiento) & "IMAGENES\" & Trim(.Fields("fotografia").Value), vbArchive) <> "" And UCase(Right(Trim(.Fields("fotografia").Value), 3)) <> "PNG" Then
                Dim dir_completo As String
                dir_completo = strDirectorioEnrolamiento & "IMAGENES\" & Trim(.Fields("fotografia").Value)
                If Dir(dir_completo, vbArchive) <> "" Then
                    picFoto.Picture = LoadPicture(strDirectorioEnrolamiento & "IMAGENES\" & Trim(.Fields("fotografia").Value))
                    txtFotografia.Text = .Fields("fotografia").Value
                Else
                '    picFoto.Picture = LoadPicture("C:\SinFoto.jpg")
                    'MsgBox "La imágen no se puede mostrar debido a que fue borrada manualmente de la carpeta IMAGENES." & vbCrLf _
                    & "No obstante, la misma sigue en la base de datos y se imprimirá en los listados."
                    txtFotografia.Text = "Vista previa no disponible (Borrado manualmente)"
                End If
        Else
            picFoto.Picture = LoadPicture("")
            txtFotografia.Text = "Sin Fotografía"
        End If
        'txtDni.Text = .Fields("Dni").Value
        chkPersonalMantenimiento.Value = IIf(.Fields("Tipo").Value = 1, 0, 1)
        chkEsLegislador.Value = IIf(.Fields("Es_Legislador").Value = 1, 1, 0)
        
        If IsNull(.Fields("template").Value) Then
            strHuella = ""
        Else
            strHuella = .Fields("template").Value
        End If
        ' si tiene huella digital
        
        If VERSION_ENROLADOR = 906 Then ' la plata
            If trae_huellas(txtID.Text, "C", 0) > 0 Then
                lblFlagHuella.Caption = "Tiene huella registrada"
            Else
                lblFlagHuella.Caption = "NO tiene huella registrada"
            End If
        Else
            If Len(strHuella) > 0 Then
                    lblFlagHuella.Caption = "Tiene huella registrada"
            Else
                lblFlagHuella.Caption = "NO tiene huella registrada"
            End If
        End If
        ' buscar banca asignada
        strSql = "SELECT Id, DeskId, OrdenPresidente FROM legisladores_activos WHERE (ID = " & Trim(.Fields("id").Value) & ")"
    End With
    ' --------------------------------------------------------------------------------
    ' Mostrar datos de la banca vinculada al legislador
    ' --------------------------------------------------------------------------------
    SetearRsW strSql, RsB
    If RsB.RecordCount = 1 Then
        lblBanca.Text = RsB.Fields("DeskId").Value
    ElseIf RsB.RecordCount > 1 Then
        MsgBox "El legislador " & UCase(txtApellido.Text & ", " & txtNombre.Text) _
             & " tiene mas de una banca asignada." & Chr(10) _
             & "Se eliminaran todas las bancas asociadas al legislador " & UCase(txtApellido.Text & ", " & txtNombre.Text) & "." & Chr(10) _
             & "Posteriormente, se deberá proceder a reasignar su banca", vbCritical + vbInformation, "Error en coherencia de datos"
        Screen.MousePointer = 11
        RsB.MoveFirst
        While Not RsB.EOF
            RsB.Delete
            RsB.MoveNext
        Wend
        Screen.MousePointer = 0
        lblBanca.Text = ""
    ElseIf RsB.RecordCount = 0 Then
        lblBanca.Text = ""
    End If
    RsB.Close
    ' --------------------------------------------------------------------------------
    ' Solo habilitar la posibilidad de asignar bancas a aquellas personas que son legisladores
    ' --------------------------------------------------------------------------------
    If chkEsLegislador.Value = 1 Then
        cmdAsignar.Enabled = True
    Else
        cmdAsignar.Enabled = False
    End If
SinFoto:
'MsgBox (Err.Description)
Select Case Err.Number
        Case Is <> 0
            picFoto.Picture = LoadPicture("")
            txtFotografia.Text = ""
            Exit Sub
    End Select
   ' picFoto.Picture = LoadPicture("")
   ' txtFotografia.Text = ""
   
End Sub

Private Sub cmdAsignar_Click()
    Dim xNuevaBanca  As Long
    Dim strSql       As String
    Dim strIdLeg     As String
    Dim RsAsignacion As ADODB.Recordset
    Dim xResp        As Long
    Dim strTemp1     As String
    Dim strTemp2     As String
    
    If Trim(lblBanca.Text) = "" Then
        Exit Sub
    End If
    xNuevaBanca = Int(lblBanca.Text)
    strIdLeg = Trim(txtID.Text)
    
    strSql = "SELECT Legisladores.id, Legisladores.apellido, Legisladores.nombre, legisladores_activos.DESKID " _
           & "FROM Legisladores INNER JOIN legisladores_activos ON Legisladores.id = legisladores_activos.ID " _
           & "WHERE (legisladores_activos.DESKID = " & Trim(Str(xNuevaBanca)) & ")"
    
    Set RsAsignacion = New ADODB.Recordset
    SetearRsW strSql, RsAsignacion
    With RsAsignacion
        Select Case .RecordCount
            Case 0
                strSql = "delete from legisladores_activos where deskid = " & Trim(Str(xNuevaBanca)) & " or id = '" & strIdLeg & "'"
                SenteciaSQl strSql
                strSql = "INSERT INTO legisladores_activos (Id, DeskId, OrdenPresidente) VALUES ('" & strIdLeg & "'," & Trim(Str(xNuevaBanca)) & ", 99)"
                SenteciaSQl strSql
            Case 1
                strTemp1 = UCase(.Fields("Apellido").Value & ", " & .Fields("Nombre").Value)
                strTemp2 = UCase(txtApellido.Text & ", " & txtNombre.Text)
                xResp = MsgBox("Actualmente la banca " & Trim(Str(xNuevaBanca)) & " esta asignada al legislador " & strTemp1 & Chr(10) _
                             & "¿Desea desasignar al legislador " & strTemp1 & " para asignar la banca " & Trim(Str(xNuevaBanca)) & " al legislador " & strTemp2 & "?", vbQuestion + vbYesNo)
                If xResp = vbYes Then
                    strTemp1 = Trim(.Fields("Id").Value)
                    strSql = "UPDATE legisladores_activos SET Id = '" & strIdLeg & "' WHERE id = '" & strTemp1 & "'"
                    SenteciaSQl strSql
                Else
                    lblBanca.Text = ""
                End If
        End Select
    End With
End Sub

Private Sub cmdDesasignarBanca_Click()

    Dim strSql   As String
    Dim strIdLeg As String
    Dim xResp    As Long
    Dim strLeg   As String
    strIdLeg = Trim(txtID.Text)
    strLeg = UCase(txtApellido.Text & ", " & txtNombre)
    
    xResp = MsgBox("Esta a punto de desasignar al legislador " & strLeg & " de la banca XX." & Chr(10) _
                  & "¿Desea Continuar?", vbQuestion + vbYesNo)
    If xResp = vbYes Then
        strSql = "DELETE FROM legisladores_activos WHERE (ID = '" & strIdLeg & "')"
        SenteciaSQl strSql
        lblBanca.Text = ""
    End If
    
End Sub


Private Sub Command1_Click()

End Sub

Private Sub dtInicioActividad_Change()
 With txtInicioActividad
        .Locked = False
        .Text = ""
        .Text = dtInicioActividad.Value
        .Locked = True
    End With
End Sub

Private Sub Form_Load()
    blEsNuevo = False
    strHuella = ""
    
    
    Dim strSql  As String
    Dim strNewDir As String
    
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    Set RsB = New ADODB.Recordset ' RecordSet dinámico responsable de la informacion de la banca asociada al legislador
    Dim NewDir As New FileSystemObject, NewDir2
    Call LlenarCombos
    Call SetearRsLegisladores
    Call VerificarPines
    Call HabilitarControles
    Call MostrarRegistro
    
End Sub

Private Sub HabilitarControles()
    lblAgrupacionPolitica.Visible = AGRUPACION_POLITICA_HABILITADA
    cmbAgrupacionPolitica.Visible = AGRUPACION_POLITICA_HABILITADA
    lblDistrito.Visible = DISTRITO_HABILITADO
    cmbDistrito.Visible = DISTRITO_HABILITADO
End Sub
Private Sub VerificarPines()
    With Rs
        If .RecordCount > 0 Then
            While Not .EOF
               
               ' If Trim(.Fields("dni").Value) <> "" Then
                '    If Trim(.Fields("pin").Value) = "" Then
                 '       .Fields("pin").Value = Encripta.EncryptString(Format(.Fields("dni").Value, "00000000"))
                  '      .Update
                   ' End If
                'Else
                  '  .Fields("dni").Value = Format(.Fields("id").Value, "00000000")
                    'MPU.Fields("pin").Value = Encripta.EncryptString(Format(.Fields("id").Value, "00000000"))
                    'MPU.Update
                'End If
                .MoveNext
            Wend
        End If
        .MoveFirst
        '.Resync
    End With
End Sub

Private Sub grabar_huella_legislador(idlegislador, nrohuella, Huella)
    Dim strSql As String
    Dim Cn As ADODB.Connection
    Set Cn = New ADODB.Connection
    With Cn
        .ConnectionString = strconexion
        .CommandTimeout = 15
        .CursorLocation = adUseClient
        .Open
    End With
    Cn.BeginTrans
        strSql = "INSERT INTO huellas (idlegislador,nrohuella,huella) VALUES ('" & idlegislador & "'," & nrohuella & ",'" & Huella & "')"
        Cn.Execute (strSql)
    Cn.CommitTrans
End Sub
Private Function trae_huellas(idlegislador, tipo, nrohuella)
    Dim strSql As String
    Dim rsHuellasLegislador As ADODB.Recordset
    Select Case tipo
        Case "C" 'cantidad de huellas
            strSql = "SELECT count(id) as cant_huellas FROM huellas WHERE idlegislador ='" & idlegislador & "'"
        Case "H" 'devuelve una huella determinada
            strSql = "SELECT huella FROM huellas WHERE idlegislador ='" & idlegislador & "' and nrohuella = " & nrohuella
    End Select
    Set rsHuellasLegislador = New ADODB.Recordset
    SetearRsW strSql, rsHuellasLegislador
    DoEvents
    If rsHuellasLegislador.RecordCount > 0 Then
        rsHuellasLegislador.MoveFirst
        Select Case tipo
        Case "C" 'cantidad de huellas
             trae_huellas = CInt(rsHuellasLegislador.Fields("cant_huellas").Value)
        Case "H" 'devuelve una huella determinada
             trae_huellas = rsHuellasLegislador.Fields("huella").Value
    End Select
   
    Else
        trae_huellas = -1 'error no encontro registros
    End If
    Set rsHuellasLegislador = Nothing
End Function

Private Function trae_ultimo_id_legislador()
    Dim strSql As String
    Dim RsIDLegislador As ADODB.Recordset
    strSql = "SELECT TOP 1 id FROM Legisladores ORDER BY cast(id as integer) DESC"
    Set RsIDLegislador = New ADODB.Recordset
    SetearRsW strSql, RsIDLegislador
    DoEvents
    If RsIDLegislador.RecordCount > 0 Then
        RsIDLegislador.MoveFirst
        trae_ultimo_id_legislador = Val(RsIDLegislador.Fields("id").Value) + 1
    Else
        trae_ultimo_id_legislador = 1001
    End If
    Set RsIDLegislador = Nothing
End Function
Private Sub CargarLegisladoresActivos()
    Dim X As Long
    Dim strSql As String
    Dim Cn As ADODB.Connection
    Set Cn = New ADODB.Connection
    With Cn
        .ConnectionString = strconexion
        .CommandTimeout = 15
        .CursorLocation = adUseClient
        .Open
    End With
    Cn.BeginTrans
        strSql = "delete from legisladores_activos "
        Cn.Execute (strSql)
        For X = 1 To 70
            strSql = "INSERT INTO legisladores_activos (Id, DeskId, OrdenPresidente) VALUES ('" & Trim(Str(X)) & "','" & Trim(Str(X)) & "'," & X + 1 & ")"
            Cn.Execute (strSql)
        Next X
    Cn.CommitTrans
End Sub

Private Sub Grabar_Click()
    On Error GoTo TrapError
    Dim strId                 As String
    Dim strApellido           As String
    Dim strNombre             As String
    Dim xSexo                 As Long
    Dim strBloque             As String
    Dim strAgrupacionPolitica As String
    Dim strDepartamento       As String
    Dim vFechaNacimiento      As Variant
    Dim strDni                As String
    Dim strTipo               As String
    Dim xEsLegislador         As Long
    Dim strSQLDelete          As String
    Dim nrohuella             As Integer
    Dim strMandato, strFinMandato, strDistrito, strFotografia, nDistrito As String
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    
    ' ------------------------------------------------------------------------------
    ' Validar datos ingresado por usuarios
    ' ------------------------------------------------------------------------------
    ' id
  '  If Trim(txtID.Text) = "" Then
  '      MsgBox "El código de identificación único de legislador es un dato obligatorio", vbInformation + vbOKOnly
  '      txtID.SetFocus
  '      Exit Sub
  '  End If
    ' nombre
    If Trim(txtNombre.Text) = "" Then
        MsgBox "El nombre del legislador es un dato obligatorio", vbCritical + vbOKOnly
        txtNombre.SetFocus
        Exit Sub
    End If
    ' Apellido
    If Trim(txtApellido.Text) = "" Then
        MsgBox "El apellido del legislador es un dato obligatorio", vbCritical + vbOKOnly
        txtApellido.SetFocus
        Exit Sub
    End If
    ' Grupo Politico
    If AGRUPACION_POLITICA_HABILITADA Then
        If Trim(cmbAgrupacionPolitica.Text) = "" Then
            MsgBox "El grupo político es un dato obligatorio", vbCritical + vbOKOnly
            cmbAgrupacionPolitica.SetFocus
            Exit Sub
        End If
    Else
        cmbAgrupacionPolitica.Text = " "
    End If
    If DISTRITO_HABILITADO Then
        If Trim(cmbDistrito.Text) = "" Then
            MsgBox "El Distrito Electoral es un dato obligatorio", vbInformation + vbOKOnly
            cmbDistrito.SetFocus
            Exit Sub
        Else
            strDistrito = cmbDistrito.Text
            nDistrito = cmbDistrito.ItemData(cmbDistrito.ListIndex)
            'MsgBox (cmbSeccion.ItemData(cmbSeccion.ListIndex))
        End If
    Else
        cmbDistrito.Text = " "
        strDistrito = " "
        nDistrito = "0"
    End If
    ' DNI
  '  If Trim(txtDni.Text) = "" Then
  '      MsgBox "El DNI del legislador es un dato obligatorio", vbCritical + vbOKOnly
  '      txtDni.SetFocus
  '      Exit Sub
  '  End If
    ' ------------------------------------------------------------------------------
    ' Leer valores ingresados por el usuario
    ' ------------------------------------------------------------------------------
    'strId = txtId.Text
    'strId = Trim(txtDni.Text)
    strApellido = txtApellido.Text
    strNombre = txtNombre.Text
    xSexo = IIf(Trim(LCase(cmbSexo.Text)) = "masculino", 1, 0)
    strBloque = cmbBloque.Text
    strMandato = cmbMandato.Text
   
    If Extension_Imagen = "" Then Extension_Imagen = ".jpg"
   
    strFotografia = txtID.Text & Extension_Imagen
    If Directorio_Imagen <> "" Then
        FileCopy Directorio_Imagen, strDirectorioEnrolamiento & "IMAGENES\" & txtID.Text & Extension_Imagen
        If Insertar_Imagen("Select PICTURE from Legisladores WHERE id = '" & txtID.Text & "'", "PICTURE", Directorio_Imagen) = False Then
            strFotografia = ""
        End If
    End If

    
    'nSeccion = cmbDistrito.ItemData(cmbDistrito.ListIndex)
'    strDistrito = cmbDistrito.Text
    strAgrupacionPolitica = cmbAgrupacionPolitica.Text
   ' strDepartamento = txtDepartamento.Text
    vFechaNacimiento = txtFechaNacimiento.Text
    txtInicioActividad = txtFechaNacimiento.Text
    'strDni = txtDni.Text
    strTipo = IIf(chkPersonalMantenimiento.Value = 0, 1, 0)
    xEsLegislador = IIf(chkEsLegislador.Value = 1, 1, 0)
    
    ' ------------------------------------------------------------------------------
    ' Grabar en base de datos
    ' ------------------------------------------------------------------------------
    With Rs
        If blEsNuevo Then
            .AddNew
            'busco el ultimo ID y le sumo 1
            SetearRs "select top 1 id as last_id from legisladores order by id desc", RsTemp

            If Not RsTemp.EOF Then
                strId = RsTemp("last_id") + 1
            Else
                strId = Replace(Replace(Date, "/", "") & Replace(Replace(Time, ":", ""), " ", ""), "/", "")
            End If
    
            .Fields("id").Value = Trim(strId)
            
                
            '.Fields("pin").Value = strDni
            '.Fields("IndiceBanca").Value = -99
        End If
        
        If txtInicioActividad <> "" And Not IsNull(txtInicioActividad) Then
            .Fields("inicio_actividad").Value = txtInicioActividad
        End If
        .Fields("mandato").Value = Trim(strMandato)
        .Fields("sexo").Value = xSexo
        .Fields("apellido").Value = Trim(strApellido)
        .Fields("nombre").Value = Trim(strNombre)
        .Fields("sexo").Value = xSexo
        .Fields("grupo_politico").Value = Trim(strAgrupacionPolitica)
        .Fields("bloque_politico").Value = Trim(strBloque)
        .Fields("departamento").Value = Trim(strDepartamento)
        .Fields("fotografia").Value = Trim(strFotografia)
        If vFechaNacimiento <> "" And Not IsNull(vFechaNacimiento) Then
            .Fields("fecha_nacimiento").Value = vFechaNacimiento
        End If
        '.Fields("dni").Value = Trim(strDni)
        .Fields("Tipo").Value = Trim(strTipo)
        .Fields("Distrito").Value = Trim(nDistrito)
        .Fields("es_legislador").Value = xEsLegislador
        If Not VERSION_ENROLADOR = 906 Then ' la plata
            .Fields("template1").Value = strHuella
        End If
        If Not blEsNuevo Then .Fields("pin").Value = Encripta.EncryptString(strDni)
        .Update
        DoEvents
        If blEsNuevo Then
            .Requery
            .MoveLast
            blEsNuevo = False
        End If
    End With
    'borro las huellas para este usuario
    'julian
    If strHuellas(1) <> "" Then
        strSQLDelete = "DELETE FROM huellas where idlegislador='" & Trim(strId) & "'"
        SenteciaSQl strSQLDelete
        
        'recorro el array de huellas y grabo cada huella
        For nrohuella = 1 To UBound(strHuellas) Step 1
            If strHuellas(nrohuella) <> "" Then
                Call grabar_huella_legislador(Trim(strId), nrohuella, strHuellas(nrohuella))
            End If
        Next
    End If
    Call HabilitarBotones(True)
    Call MostrarRegistro
Exit Sub
TrapError:
    
    Select Case Err.Number
        Case -2147217873
            MsgBox "Esta a punto de cometerse un error de integridad de datos." & Chr(10) & "Revise los valores de AGRUPACION POLITICA y BLOQUE POLITICO antes de continuar con la operación.", vbInformation + vbOKOnly
            Exit Sub
        Case Else
            MsgBox "Error Nº " & Err.Number & Chr(10) & Err.Description & Chr(10) & "Originado en " & Err.Source
            Exit Sub
    End Select
End Sub

Private Sub lblBanca_Change()
    If Trim(lblBanca.Text) = "" Then
        cmdDesasignarBanca.Enabled = False
    Else
        cmdDesasignarBanca.Enabled = True
    End If
End Sub



Private Sub Nuevo_Click()
    Call HabilitarBotones(False)
    Call Limpiar
    blEsNuevo = True
    txtID.Text = trae_ultimo_id_legislador()
    txtNombre.SetFocus
End Sub

Private Sub Salir_Click()
    If Salir.Caption = "&Salir" Then
        Unload Me
    Else
        blEsNuevo = False
        Call HabilitarBotones(True)
        Call MostrarRegistro
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub lblBanca_GotFocus()
    With lblBanca
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub lblBanca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Exit Sub
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then ' Solo permite ingresar caracteres entre 0 y 9
        KeyAscii = 0
    End If
End Sub

Private Sub txtId_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub UpDown1_DownClick()
    Dim xValor As Long
    xValor = Int(Val(lblBanca.Text))
    xValor = xValor - 1
    If xValor < 0 Then
        xValor = 0
    End If
    lblBanca.Text = Str(xValor)
End Sub
Private Sub UpDown1_UpClick()
    Dim xValor As Long
    xValor = 0 + Int(Val(lblBanca.Text))
    xValor = xValor + 1
    If xValor > 70 Then
        xValor = 70
    End If
    lblBanca.Text = Str(xValor)
End Sub
Private Sub HabilitarBotones(blEstado As Boolean)
    Nuevo.Enabled = blEstado
    Grabar.Enabled = True
    Borrar.Enabled = blEstado
    cmdPin.Enabled = blEstado
    cmBuscar.Enabled = blEstado
    If Not blEstado Then
        Salir.Caption = "&Cancelar"
    Else
        Salir.Caption = "&Salir"
    End If
    cmdPrimero.Enabled = blEstado
    cmdUltimo.Enabled = blEstado
    cmdSiguiente.Enabled = blEstado
    cmdAnterior.Enabled = blEstado
    lblRecordSet.Enabled = blEstado
End Sub
Private Sub LlenarCombos()
    Dim strSql As String
    Dim nIndex As Integer
    
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    ' -------------------------------------------------------
    ' Combo Sexo
    ' -------------------------------------------------------
    With cmbSexo
        .Clear
        .AddItem "Masculino"
        .AddItem "Femenino"
    End With
    ' -------------------------------------------------------
    ' Bloque Político
    ' -------------------------------------------------------
    strSql = "SELECT Bloque_Político FROM Bloques ORDER BY Bloque_Político"
    SetearRs strSql, RsTemp
    With cmbBloque
        .Clear
        If RsTemp.RecordCount > 0 Then
            RsTemp.MoveFirst
            While Not RsTemp.EOF
                .AddItem Trim(RsTemp.Fields(0).Value)
                RsTemp.MoveNext
            Wend
        End If
        RsTemp.Close
    End With
    ' -------------------------------------------------------
    ' Agrupación Política
    ' -------------------------------------------------------
    strSql = "SELECT Agrupación_Política From Grupos ORDER BY Agrupación_Política"
    SetearRs strSql, RsTemp
    With cmbAgrupacionPolitica
        .Clear
        If RsTemp.RecordCount > 0 Then
            RsTemp.MoveFirst
            While Not RsTemp.EOF
                .AddItem Trim(RsTemp.Fields(0).Value)
                RsTemp.MoveNext
            Wend
        End If
        RsTemp.Close
    End With
    ' -------------------------------------------------------
    ' Seccion / DIstrito Electoral
    ' -------------------------------------------------------
    arrSecciones = ""
    nIndex = 1
    strSql = "SELECT d.id_distrito,d.seccion,d.distrito,s.seccion as seccionstr from distritos d INNER JOIN secciones s ON s.id_seccion = d.seccion order by seccionstr,distrito"
    SetearRs strSql, RsTemp
    With cmbDistrito
        .Clear
        If RsTemp.RecordCount > 0 Then
            RsTemp.MoveFirst
            While Not RsTemp.EOF
                .AddItem RsTemp.Fields("seccionstr").Value & " --> " & RsTemp.Fields("distrito").Value
                .ItemData(cmbDistrito.NewIndex) = RsTemp.Fields("id_distrito").Value
                'arrSecciones = arrSecciones & RsTemp.Fields("seccionstr").Value & " --> " & RsTemp.Fields("distrito").Value & "|" & nIndex & "||"
                arrSecciones = arrSecciones & RsTemp.Fields("id_distrito").Value & "|" & nIndex & "||"
                RsTemp.MoveNext
                nIndex = nIndex + 1
            Wend
        End If
        RsTemp.Close
    End With
    
    ' -------------------------------------------------------
    ' inicio mandato
    ' -------------------------------------------------------
    strSql = "SELECT fecha_mandato From mandatos ORDER BY fecha_mandato"
    SetearRs strSql, RsTemp
    With cmbMandato
        .Clear
        If RsTemp.RecordCount > 0 Then
            RsTemp.MoveFirst
            While Not RsTemp.EOF
                .AddItem RsTemp.Fields(0).Value
                RsTemp.MoveNext
            Wend
        End If
        RsTemp.Close
    End With
    
    Set RsTemp = Nothing
End Sub
Private Sub SetearRsLegisladores()
    Dim strSql As String
    strSql = "SELECT l.* FROM Legisladores l ORDER BY Apellido"
    Set Rs = New ADODB.Recordset
    SetearRsW strSql, Rs
    DoEvents
    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
    End If
End Sub
Private Sub cmdPrimero_Click()
    If PermisosTotales.ConsultaABMLegislador = 0 Then
        MsgBox "El usuario no tiene permisos para esta tarea", vbInformation + vbOKOnly, "Usuario sin permisos!!"
        Exit Sub
    End If
    Rs.MoveFirst
    Call MostrarRegistro
End Sub
Private Sub cmdSiguiente_Click()
    Directorio_Imagen = ""
    Extension_Imagen = ""
    Puede_Cargar_Foto = False
    If PermisosTotales.ConsultaABMLegislador = 0 Then
        MsgBox "El usuario no tiene permisos para esta tarea", vbInformation + vbOKOnly, "Usuario sin permisos!!"
        Exit Sub
    End If
    If Rs.AbsolutePosition < Rs.RecordCount Then
        Rs.MoveNext
    Else
        Rs.MoveLast
    End If
    Call MostrarRegistro
End Sub
Private Sub cmdAnterior_Click()
    Directorio_Imagen = ""
    Extension_Imagen = ""
    Puede_Cargar_Foto = False
    If PermisosTotales.ConsultaABMLegislador = 0 Then
        MsgBox "El usuario no tiene permisos para esta tarea", vbInformation + vbOKOnly, "Usuario sin permisos!!"
        Exit Sub
    End If
    If Rs.AbsolutePosition > 1 Then
        Rs.MovePrevious
    Else
        Rs.MoveFirst
    End If
    Call MostrarRegistro
End Sub
Private Sub cmdUltimo_Click()
    If PermisosTotales.ConsultaABMLegislador = 0 Then
        MsgBox "El usuario no tiene permisos para esta tarea", vbInformation + vbOKOnly, "Usuario sin permisos!!"
        Exit Sub
    End If
    Rs.MoveLast
    Call MostrarRegistro
End Sub
Private Sub Limpiar()
    txtID.Text = ""
    txtApellido.Text = ""
    txtNombre.Text = ""
    cmbSexo.Text = ""
    cmbBloque.Text = ""
    cmbAgrupacionPolitica.Text = ""
    cmbDistrito.Text = ""
    cmbDistrito.Text = ""
    cmbMandato.Text = ""
    
    txtInicioActividad.Text = ""
   ' txtDepartamento.Text = ""
    dtFechaNacimiento.Value = Date
    txtFechaNacimiento.Text = ""
    'txtDni.Text = ""
    chkPersonalMantenimiento.Value = 0
    chkEsLegislador.Value = 0
    lblBanca.Text = ""
    picFoto.Picture = LoadPicture("")
    txtFotografia.Text = ""
End Sub
Private Sub cmdPin_Click()
    Dim xRespuesta As Long
    Dim strIdLeg   As String
    Dim strDni     As String
    xRespuesta = MsgBox("¿Esta seguro de modificar el PIN del legislador " & UCase(txtApellido.Text) & "?", vbQuestion + vbYesNo)
    If xRespuesta = vbYes Then
        strIdLeg = Trim(txtID.Text)
        'strDni = Trim(txtDni.Text)
        frmPIN.ID_Legislador = strIdLeg
        frmPIN.DNI = strDni
        frmPIN.Show vbModal
    End If
End Sub
Private Sub dtFechaNacimiento_Change()
    With txtFechaNacimiento
        .Locked = False
        .Text = ""
        .Text = dtFechaNacimiento.Value
        .Locked = True
    End With
End Sub

' ---------------------------------------------------------------------------------------------------------
' RUTINAS PARA MANEJO DE ARCHIVOS EXTENSION "HUELLA"
' ---------------------------------------------------------------------------------------------------------
Private Function HexATexto(strTexto As String, nLong As Long) As String
    Dim i As Long
    'convierte un texto que contenga pares hexadecimales codificados como string en un string ascii
    HexATexto = ""
    For i = 1 To Len(strTexto) Step 2
        HexATexto = HexATexto & HexAChr(Mid(strTexto, i, 2))
    Next
    HexATexto = CerosIzquierda(HexATexto, nLong)
End Function
Private Function HexAChr(strHex) As String
    'convierte dos digitos hexadecimales codificados como string en un caracter ascii
    Dim nDecimal As Long
    nDecimal = 0
    nDecimal = DigitoHexADec(Mid(strHex, 1, 1)) * 16
    nDecimal = nDecimal + DigitoHexADec(Mid(strHex, 2, 1))
    HexAChr = Chr(nDecimal)
End Function
Private Function DigitoHexADec(charHex) As Long
    If charHex >= "0" And charHex <= "9" Then
        DigitoHexADec = Asc(charHex) - 48
    Else
        DigitoHexADec = Asc(charHex) - 55
    End If
End Function
Private Function TextoAHex(strTexto As String, nLong As Long) As String
    Dim i As Long
    TextoAHex = ""
    For i = 1 To Len(strTexto)
        TextoAHex = TextoAHex & Hex(Asc(Mid(strTexto, i, 1)))
    Next
    TextoAHex = CerosIzquierda(TextoAHex, nLong)
End Function
Private Function BuscarArchivoImagen() As String()
    ' Buscar archivos de imagen
    Dim DatosArray() As String
    ReDim DatosArray(0 To 1)
    With CommonDialog1
        .CancelError = True
        On Error GoTo ErrHandler
        .Flags = cdlOFNHideReadOnly 'Or cdlOFNExplorer
        '.Filter = "Archivos JPG (*.jpg)|*.jpg|Archivos BMP (*.bmp)|*.bmp"
        .Filter = "Archivos JPG (*.jpg)|*.jpg"
        .FilterIndex = 1
        .InitDir = strDirectorioEnrolamiento & "IMAGENES\"
        .ShowOpen
        DatosArray(0) = .FileName
        DatosArray(1) = .FileTitle
        Extension_Imagen = Right(.FileTitle, 4)
        BuscarArchivoImagen = DatosArray()
    End With
Exit Function
ErrHandler:
    DatosArray(0) = ""
    DatosArray(1) = ""
    BuscarArchivoImagen = DatosArray() ' El usuario ha hecho clic en el botón Cancelar
    Exit Function
End Function

Private Function BuscarArchivoHuella() As String
    ' Buscar archivos extension "huella"
    With CommonDialog1
        .CancelError = True
        On Error GoTo ErrHandler
        .Flags = cdlOFNHideReadOnly
        .Filter = "Todos los archivos (*.*)|*.*|Archivos de Enrolador (*.huella)|*.huella"
        .FilterIndex = 2
        .ShowOpen
        BuscarArchivoHuella = .FileName
    End With
Exit Function
ErrHandler:
    BuscarArchivoHuella = "" ' El usuario ha hecho clic en el botón Cancelar
    Exit Function
End Function
Private Function LeerContenidoArchivo(strFile As String) As String
    ' Recibe el nombre de un archivo y devuelve el string contenido en dicho archivo
    Dim xFile As Long
    xFile = FreeFile               ' # de archivo disponible por el sistema operativo
    LeerContenidoArchivo = ""
    Open strFile For Binary As #xFile
        LeerContenidoArchivo = Space(LOF(xFile))
        Get #xFile, , LeerContenidoArchivo
    Close #xFile
End Function
Private Function CerosIzquierda(strText As String, nLong As Long) As String
    If nLong > Len(strText) Then
        CerosIzquierda = Left(String(nLong - Len(strText), "0") & strText, nLong)
    Else
        CerosIzquierda = Right(strText, nLong)
    End If
End Function


