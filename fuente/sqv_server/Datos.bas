Attribute VB_Name = "Datos"
Public renglonesExtra() As String
Public PresidenteEstuvoMantenimiento As Boolean
Public PrimeraVezCeros As Boolean
Public BancasDeshabilitadas(256) As Boolean
Public ModoMant As Boolean
Public CuentaSQL As Long
Public VectorDesconectadas(256) As Boolean
Public VectorControlDoble(256) As Integer
Public VectorControlDobleTick(256) As Long
Public PrimerControlLarga As Boolean
Public Imprimio As Boolean
Public sinIdentificarCongelado As Boolean

' ------------------------------------------------------------------------
' Valores totales para control de Cartel
' ------------------------------------------------------------------------
Type DatosCartel
    Presentes                         As Long
    Ausentes                          As Long
    Resultado                         As String
    Afirmativos                       As Long
    Negativos                         As Long
    Abstenciones                      As Long
    MinimoVotosParaAfirmativo         As Long
    LeyendaQuorum                     As String
    LeyendaTiempo                     As String
    LeyendaTipoOperacion              As String
    LeyendaMinimoVotosParaAfirmativo  As String
End Type
' ------------------------------------------------------------------------
' Valores totales para control de estados
' ------------------------------------------------------------------------
Type EstadoServer
    EnIdentificacion(0 To 256)        As Boolean
    VectorColor()                     As String
    VectorPresencia()                 As String
    VectorIdentificacion()            As String
    VectorResultados()                As String
    VectorResultadosCong()            As String
    VectorPresenciaCong()             As String
    VectorIdentificacionCong()        As String
    VectorIdentificacionHabilitados() As String
    VectorAbstencion()                As String
    TipoDeOperacion                   As String
    OcupadosNoIdentificados           As Long
    Sesion                            As Long
    NroActa                           As Long
    BaseMayoria                       As String
    TipoMayoria                       As String
    strError                          As String
    EstadoVotacion_y_PasList          As String
    'ModalidadVotacion                 As String
    PendientesEmitirVotos             As Long
    MensajeAlOperador                 As String
    TituloDelActa                     As String
    GrabarAutomaticamente             As Integer
    ListarAutomaticamente             As Integer
    TipoMayoriaQuorum                 As String
    TipoDeAbstencion                  As String
    PeriodoLegislativo                As String
    ModoMantenimientoBancas           As Integer
    ActaGrabada                       As Integer
    SolicitudGrabarManual             As Integer
    TiempoParaVotacion                As Long  ' tiempo asignado para votacion
    IP_Consola                        As String
    ModoNormalMantSistema             As Integer
    Modo_Ident_Nom                    As Integer '1: Habilita identificacion en modo no nominal, 0: solo se permite identificacion durante operaciones nominales (votnom, paslis)
    IdentificadorDeFormulario         As String
    CartelEncendido                   As Integer
    Presentes                         As Long
    Ausentes                          As Long
    EstadoSesion                      As String
    FechaVotacion                     As Date
    HoraVotacion                      As String
    LimpiarResultados                 As Integer
    PresentesCongelados               As Integer
    AusentesCongelados                As Integer
    OcupadosNoIdentificadosCongelados As Integer
    AbstencionistasAutorizados        As Long 'esta variable se deber reemplazar por una de cartel.
    VMantBanca()                      As Long
    VMantInfo()                       As String
    VMantIdentificacion()             As String
    VMantEstado()                     As String
    MantIdentificaciones              As String
    MantPresencias                    As String
    MantListaPendientes               As String
    MantListaFallas                   As String
    MantCantPendientes                As Long
    MantCantFallas                    As Long
    VTipoIdentificacion()             As String     'Blanco significa el valor por omision que es huella dactilar. "T" Teclado
    Modo_Presencia_Nom                As Integer    '1: se considera presentes solo a los presentes identificados; 0: Todos los presentes cuentan para el total de presentes para quorum
    Reunion                           As Long
    Orador                            As String
    OradorNombre                      As String
    OradorAgrupacionPolitica          As String
    OradorDistrito                    As String
    OradorSexo                        As String
    VectorError()                     As String
    ModoVotaPresidente                As Boolean    'Verdadero: el presidente puede emitir voto
    ResultadoVotoPresidente           As String     'Resultado del voto del presidente (fuera empate)
    EsperarVotoPresidente             As Boolean
    PresidenteHabilitadoParaVotar     As Boolean
    ExtensionDeTiempoPorPresidente    As Boolean
    Expresiones_Minoria               As Boolean
    BancaEnPrueba                     As Integer
End Type
Type MensajeSistema
    sTipo                             As String
    sComponente                       As String
    sObjeto                           As String
    sAtributo                         As String
    sValor                            As String
    sComentario                       As String
End Type

Type CartelMural
    strQuorum             As String
    strPresentes          As String
    strAusentes           As String
    strAfirmativos        As String
    strNegativos          As String
    strAbtenciones        As String
    strResultado          As String
    strSesion             As String
    strOrdenDia           As String
    strTitulo             As String
    strTipoVota           As String
    strTiempoVota         As String
    strMayoria            As String
    strLineaCartel10      As String
    strLineaCartel11      As String
    strAtributo03         As String
    strAtributo04         As String
    strAtributo05         As String
    strAtributo10         As String
    strAtributo11         As String
End Type

Type BaseYTipo
    strBase               As String
    strTipo               As String
End Type
' ----------------------------------------------------------------------
' Colores
' ----------------------------------------------------------------------
' Req.: Pasarlo a variable, manipulables desde BD
' y hacer un ABM para poder cambiar de colores a gusto del operador
' el ABM debe estar en la consola.
Global Const cGRIS = 0
Global Const cBLANCO = 1
Global Const cAMARILLO = 2
Global Const cROJO = 3
Global Const cCELESTE = 4
Global Const cNARANJA = 5
Global Const cVERDE = 6
Global Const cNEGRO = 7
Global Const cOLIVA = 8
Global Const cAZUL = 9
Global Const cMARRON = 10
Global Const cMarronClaro = 11
Global Const SEPARADOR_VECTOR As String = ";"
Global Const cUltimoPanelMant As Long = 6

Global Const ERROR_SIN_ERROR As String = " "
Global Const ERROR_IOC As String = "W"

Public EtiquetasCartel As BaseYTipo
' Objeto Encriptador
Public Encripta As New clsEncriptador
Public PrimeraVezControl As Boolean 'Para no repetir msjs de cancel de teclado en abs automatica

Global Const AGRUPACION_POLITICA_HABILITADA = False
Global Const DISTRITO_HABILITADO = True
' ----------------------------------------------------------------------
' Tipos de identificacion
' ----------------------------------------------------------------------
Global Const TIPO_IDENTIFICACION_HUELLA As String = " "
Global Const TIPO_IDENTIFICACION_TECLADO As String = "T"
' ----------------------------------------------------------------------
' Valores GLOBALES
' ----------------------------------------------------------------------
Global strConexion              As String        ' string de conexion
Global EstadoActual             As EstadoServer  ' seguimiento de recinto
Global CartelActual             As DatosCartel   ' seguimiento de recinto
Global xMiembrosDelCuerpo       As Long          ' Cantidad total de legisladores
Global xUltimaBanca             As Long          ' Ultima banca, considerando que la primera es la cero
Global strVersion               As String        ' compilacion actual de sqv server
' Valores de control de procesamiento
Global xtiempoInicioVotac       As Long
Global xTiempoEsperaPaseLista   As Long
Global xMinimoParaQuorum        As Long
Global xMinimoParaQuorumEntero  As Long
Global flSwitchExitoso          As Boolean
Global flBancaIdentifPosExitosa As Boolean
Global xSensibilidadReintentos  As Long
Global xSegundosFinOperacion    As Long 'traer de config
Global xPresidenteLegislador    As Boolean
Global Const PermitirVotarAlPresidente As Boolean = False
Global xPresidenteAnteriorLegislador    As Boolean
Global xBancaPruebaScan         As Long
Global xVectorIdentificacionHabilitados() As String
Global xNombreUltimoIdentificado As String
Global xLogSQVPrueba             As String
Global nLogSQVPrueba             As Long
Global blBanderaPruebas          As Boolean
' Variables para control de actualizacion del cartel
Global xControlCartelTipoOperacion      As String
Global xControlCartelEstadoOperacion    As String
Global sCartel                          As CartelMural
Global xHuboEmpate    As Boolean
Global xHuboDesempate As Boolean
Global xCierreEmpateOperador As Boolean
Global xVotoSenadorEmpate As String
Global ultimoResultadoEvaluado As String


'configuracion levantador de aplicaciones
Global strPuerto   As String
Global strIpServer As String
Global strExeSqv   As String
Global strExeSb    As String

' Variables para control de votacion
Global tFinVotacion As Date
Global xTipoVotacion As String 'puede valer votnom o votnum. Se utiliza para diferencia votaciones numericas de nominales cuando se trabaja en modo nominal.

Global mColores()               As String


' ----------------------------------------------------------------------
' Estados de las bancas
' ----------------------------------------------------------------------
Global Const PRESENTE              As String = "1"
Global Const AUSENTE               As String = "0"
Global Const NO_IDENTIFICADO       As String = "0"
Global Const BANCA_INHABILITADA    As String = "X"
Global Const ABSTENCION            As String = " "
Global Const AFIRMATIVO            As String = "s"
Global Const NEGATIVO              As String = "n"
Global Const ABSTENCION_AUTORIZADA As String = "a"

Global Const FORMATOFECHA As String = "dd/mm/yyyy"

Global xBancaDuplicada             As Long
Global flExitoPierdeID As Boolean
Global flExitoPierdeIdDup As Boolean
Global flExitoPierdeIdDupConPresdte As Boolean

'A02
Global mColoresFuente()               As String
'A02 END

' ----------------------------------------------------------------------
' Formularios de presentacion del recinto. 091011
' ----------------------------------------------------------------------
Global Const cFORMULARIO_VERSION              As String = "09"
Global Const cFORMULARIO_COLOR_FONDO          As String = &H80000008
Global Const cFORMULARIO_MOSTRAR_BANCAS       As String = False

'-------------------------
' Para facilitar el seguimiento del log
'--------------------------
Global strUltimoMensaje_SQV_SB As String
Global strUltimoMensaje_SB_SQV As String
Global Const cSEVERIDAD_MINIMA = "0" ' subir este nivel para reducir los logs generados

Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Function PAS(Periodo As String, Sesion As Long, Nuevo_Estado As String) As Long
    Dim rsTemp As ADODB.Recordset
    Dim Consulta As String
    Set rsTemp = New ADODB.Recordset
    frmMain.SetearRsAux "SELECT Próximo_Acta FROM sesion WHERE Sesión = '" & Sesion & "' AND Período_Legislativo = '" & Periodo & "'", rsTemp
    If Not rsTemp.EOF Then
        PAS = rsTemp.Fields(0)
    Else
        Consulta = "INSERT INTO sesion (Período_Legislativo,Sesión,Fecha_de_Inicio,Próximo_Acta,Estado_sesión,Prorroga) " & _
        " VALUES ('" & Periodo & "'," & Sesion & ",'" & Date & "',1,'" & Nuevo_Estado & "',0)"
        frmMain.EjecutarSQL Consulta
        PAS = 1
    End If
    rsTemp.Close
    Set rsTemp = Nothing
End Function
    Sub InicializarValores()
    Dim X As Long
    EstadoActual.FechaVotacion = DateAdd("s", -10, Now)
    With EstadoActual
        ' ------------------------------------------------------------------------
        ' Inicializar vectores
        ' ------------------------------------------------------------------------
        ReDim .VectorPresencia(0 To xUltimaBanca)
        ReDim .VectorColor(0 To xUltimaBanca)
        ReDim .VectorIdentificacion(0 To xUltimaBanca)
        ReDim .VectorResultados(0 To xUltimaBanca)
        ReDim .VectorIdentificacionCong(0 To xUltimaBanca)
        ReDim .VectorPresenciaCong(0 To xUltimaBanca)
        ReDim .VectorResultadosCong(0 To xUltimaBanca)
        ReDim .VectorIdentificacionHabilitados(0 To xUltimaBanca)
        ReDim .VMantEstado(0 To xUltimaBanca)
        ReDim .VectorAbstencion(0 To xUltimaBanca)
        ReDim .VTipoIdentificacion(0 To xUltimaBanca)
        
        ReDim .VMantBanca(0 To cUltimoPanelMant)
        ReDim .VMantInfo(0 To cUltimoPanelMant)
        ReDim .VMantIdentificacion(0 To cUltimoPanelMant)
        ReDim .VectorError(0 To xUltimaBanca)
        
        For X = 0 To (xUltimaBanca)
            .VectorPresencia(X) = BANCA_INHABILITADA
            .VectorIdentificacion(X) = NO_IDENTIFICADO
            .VectorError(X) = ERROR_SIN_ERROR
            .VectorColor(X) = AsignarColor(X)
            .VectorResultados(X) = ABSTENCION
            .VectorResultadosCong(X) = ABSTENCION
            .VectorIdentificacionCong(X) = NO_IDENTIFICADO
            .VectorPresenciaCong(X) = BANCA_INHABILITADA
            .VMantEstado(X) = ABSTENCION
            .VTipoIdentificacion(X) = TIPO_IDENTIFICACION_HUELLA
        Next X
    End With
    Call ResetearPresidente
    EstadoActual.MantIdentificaciones = ""
    EstadoActual.MantPresencias = ""
    EstadoActual.MantCantFallas = 0
    EstadoActual.MantCantPendientes = 0
    EstadoActual.MantListaFallas = " "
    EstadoActual.MantListaPendientes = " "
    strVersion = "Versión " & App.Major & "." & App.Revision & "." & App.Minor
    EstadoActual.Modo_Ident_Nom = 0 ' Inicializa con scanners no habilitados en quorum
    EstadoActual.Modo_Presencia_Nom = 0 'Funcion no soportada aun, falta recibir el mensaje de control de la consola para permitir que se cuenten como presentes a los identificados.
End Sub
Sub Main()
    ' Verificar que no haya una instancia previa de SQV Server en RAM
    If App.PrevInstance = True Then
        End
    End If
    
    Encripta.Password = "ClaveInvulnerable350"
    strConexion = DeterminarStringConexion    ' Armar string de conexion comun
    If strConexion = "" Then
        frmConfig.Show 1
    End If
    Call LeerConfig                           ' Determinar cantidad de legisladores
    Call InicializarValores                   ' Valores de estado inciales
    'Call SetearCacheBancas                   ' HCDN 2011 No se usa
    'frmMain.Show
    'frmMain.Visible = False
    frmCartel2011.Show
    'frmNoFrame.Show
End Sub

Private Sub SetearCacheBancas()
    
    Dim strSql       As String
    Dim rsCache      As ADODB.Recordset
    Dim cnCache      As ADODB.Connection
    Dim strSecuencia As String
    Dim xBancaActual As Long
    Dim xUltimaBanca As Long
    Dim strUpdate    As String
    
    Set rsCache = New ADODB.Recordset
    Set cnCache = New ADODB.Connection
    ' ------------------------------------------------------------------------------------
    ' Conectar a la base de datos
    ' ------------------------------------------------------------------------------------
    strSql = "SELECT BancasCercanas.Banca, Legisladores.IndiceBanca " _
           & "FROM BancasCercanas INNER JOIN legisladores_activos ON BancasCercanas.BancaCercana = legisladores_activos.DESKID INNER JOIN " _
           & "Legisladores ON legisladores_activos.ID = Legisladores.id"
    GoSub ConexionBase
    ' ------------------------------------------------------------------------------------
    ' Armar secuencia de envio de huellas
    ' ------------------------------------------------------------------------------------
    With rsCache
        If .RecordCount > 0 Then
            .MoveFirst
            xBancaActual = .Fields("Banca").Value
            xUltimaBanca = xBancaActual
            While Not .EOF
                xBancaActual = .Fields("Banca").Value
                If xBancaActual <> xUltimaBanca Then
                    strUpdate = "update BancasIP SET IdString = '" & strSecuencia & "' WHERE BancaNumero = " & Str(xUltimaBanca)
                    cnCache.Execute (strUpdate)
                    xUltimaBanca = xBancaActual
                    strSecuencia = ""
                End If
                strSecuencia = strSecuencia & strString(4, Hex(Fix(.Fields("IndiceBanca").Value)), "0", "I")
                xUltimaBanca = xBancaActual
                .MoveNext
            Wend
            strUpdate = "update BancasIP SET IdString = '" & strSecuencia & "' WHERE BancaNumero = " & Str(xUltimaBanca)
            cnCache.Execute (strUpdate)
            xUltimaBanca = xBancaActual
            strSecuencia = ""
        End If
    End With
    ' ------------------------------------------------------------------------------------
    ' Cerrar conexion
    ' ------------------------------------------------------------------------------------
    rsCache.Close
    cnCache.Close
    Set rsclose = Nothing
    Set cnCache = Nothing
Exit Sub
ConexionBase:
    With cnCache
        .ConnectionString = strConexion
        .ConnectionTimeout = 15
        .CursorLocation = adUseClient
        .Open
    End With
    With rsCache
        .ActiveConnection = cnCache
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .Source = strSql
        .LockType = adLockOptimistic
        .Open
    End With
Return
End Sub
Private Function strString(xTam As Long, strValor As String, strRelleno As String, Optional strTipo As String = "D") As String
    If Len(Trim(strValor)) < xTam Then
        If strTipo = "I" Then
            strString = String(xTam - Len(Trim(strValor)), strRelleno) & strValor
        Else
            strString = strValor & String(xTam - Len(Trim(strValor)), strRelleno)
        End If
    Else
        strString = Left(strValor, xTam)
    End If

End Function

Function DeterminarStringConexion() As String
On Error GoTo TrapError:
    Dim xFile      As Long
    Dim strCadena  As String
    Dim strCadenaP As String  ' string de conexion de base de datos de prueba
    Dim strArchivo As String
    Dim CnConf     As ADODB.Connection
    Dim RsConf     As ADODB.Recordset
    
    Set CnConf = New ADODB.Connection
    Set RsConf = New ADODB.Recordset
    
    strArchivo = App.Path & "\sqv.dat"
    xFile = FreeFile
    ' ----------------------------------------------------------------------
    ' Abrir archivo de configuracion y levantar cadena de conexion
    ' ----------------------------------------------------------------------
    Open strArchivo For Binary As #xFile
        strCadena = Space(LOF(xFile))
        Get #xFile, , strCadena
    Close #xFile
    ' ----------------------------------------------------------------------
    ' Desencriptar cadena de conexion
    ' ----------------------------------------------------------------------
    strCadena = Encripta.EncryptString(strCadena)
    If Trim(strCadena) = "" Then
        MsgBox "Error al abrir archivo de configuración inicial. Deberá generar nuevamente la configuración de acceso a la base de datos", vbInformation, "SQV Server Informa"
        DeterminarStringConexion = ""
    Else
        ' Conectarse a base de datos config: NOTA 091001 Revisar que el sqv.dat este arrancando en sqv_config
        With CnConf
            .ConnectionString = strCadena
            .ConnectionTimeout = 15
            .CursorLocation = adUseClient
            .Open
        End With
        With RsConf
            .ActiveConnection = CnConf
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
            .Source = "SELECT * FROM Configuracion"
            .LockType = adLockOptimistic
            .Open
            ' ---------------------------------------------------------------------------
            ' Leer cadena de conexion
            ' ---------------------------------------------------------------------------
            DoEvents
            .MoveFirst
            While Not .EOF
                If .Fields("Variable").Value = "base_prueba" Then
                    strCadenaP = .Fields("Valor").Value
                End If
                If .Fields("Variable").Value = "base_vigente" Then
                    strCadena = Trim(.Fields("Valor").Value)
                    DeterminarStringConexion = strCadena
                End If
                .MoveNext
            Wend
        End With
        ' ---------------------------------------------------------------------------
        ' Mostrar la base de datos que se esta utilizando (produccion o pruebas)
        ' ---------------------------------------------------------------------------
        If strCadenaP = strCadena Then
            blBanderaPruebas = True
        Else
            blBanderaPruebas = False
        End If
       
        ' ---------------------------------------------------------------------------
        ' Cerrar conexion con base SQV_Config
        ' ---------------------------------------------------------------------------
        RsConf.Close
        CnConf.Close
        Set RsConf = Nothing
        Set CnConf = Nothing
    End If
    Exit Function
TrapError:
    Select Case err.Number
        Case 6
            MsgBox "Error N° " & err.Number & Chr(10) & err.Description & " Originado en " & err.Source
            End
        Case 3709 Or -2147467259
            MsgBox "Error N° " & err.Number & Chr(10) & err.Description & " Originado en " & err.Source & vbCrLf & " Verifique que la tabla de configuracion de la base de datos config tenga correctamente configurado el 'string' de conexión."
            End
        Case Else
            If MsgBox("Error N° " & err.Number & Chr(10) & err.Description & " Originado en " & err.Source & vbCrLf & " Verifique los parámetros de conexión con la base de datos. " & vbCrLf & "Para reintentar indique SI", vbQuestion + vbYesNo, "Confirma la operación?") = vbYes Then
                'Resume
            Else
                End
            End If
    End Select
End Function



Sub LeerConfig()
    Dim strSql                    As String
    Dim xLegisladores             As Long
    Dim xLegisladores_Habilitados As Long
    Dim Cn                        As ADODB.Connection
    Dim rs                        As ADODB.Recordset
    
    Set Cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    xLegisladores = 0
    xLegisladores_Habilitados = 0
    ' ----------------------------------------------------------------------
    ' Abrir una conexion a la base de datos
    ' ----------------------------------------------------------------------
    With Cn
        .ConnectionString = strConexion
        .CursorLocation = adUseServer
        .ConnectionTimeout = 25
        .Open
    End With
    ' ----------------------------------------------------------------------
    ' Setear Rs para levantar cantidad de legisladores totales
    ' ----------------------------------------------------------------------
    strSql = "SELECT * FROM config"
    rs.Open strSql, Cn, adOpenDynamic, adLockOptimistic
    With rs
        xMiembrosDelCuerpo = .Fields("Cantidad_de_Legisladores").Value
        xtiempoInicioVotac = .Fields("Segundos_de_inicio_operacion").Value
        xSegundosFinOperacion = .Fields("Segundos_de_fin_operacion").Value
        xTiempoEsperaPaseLista = .Fields("Tiempo_espera_Pase_de_Lista").Value
        xUltimaBanca = .Fields("cantidad_de_bancas").Value - 1
        xSensibilidadReintentos = .Fields("Sensib_scan_neg").Value
        .Close
    End With
    ' ----------------------------------------------------------------------
    ' Ver si hay tantos legisladores habilitados como totales
    ' ----------------------------------------------------------------------
    strSql = "SELECT * FROM legisladores_activos"
    rs.CursorLocation = adUseClient
    rs.Open strSql, Cn, adOpenDynamic, adLockOptimistic
    xLegisladores_Habilitados = rs.RecordCount
    If xLegisladores_Habilitados < xMiembrosDelCuerpo Then
        ' hacer un log en archivos para esto
        'Call frmMain.AltaLogGeneral("SQV SERVER", "La configuración del sistema registra " & Str(xMiembrosDelCuerpo) & " Legisladores totales dentro del cuerpo, pero solo hay " & xLegisladores_Habilitados & " legisladores activos registrados.")
        'MsgBox "La configuración del sistema registra " & Str(xMiembrosDelCuerpo) & " Legisladores totales dentro del cuerpo, pero solo hay " & xLegisladores_Habilitados & " legisladores activos registrados.", vbCritical + vbInformation, "SQV AUDITORIA"
    End If
    
    rs.Close
    strSql = "SELECT * FROM legisladores"
    rs.CursorLocation = adUseClient
    rs.Open strSql, Cn, adOpenDynamic, adLockOptimistic
    xLegisladores = rs.RecordCount
    If xLegisladores_Habilitados > xLegisladores Then
        ' hacer un log en archivos para esto
        'Call frmMain.AltaLogGeneral("SQV SERVER", "Existen solo " & Str(xLegisladores) & " Legisladores registrados, pero hay " & Str(xLegisladores_Habilitados) & " legisladores habilitados. Error en la gestión de datos básicos.")
        'MsgBox "Existen solo " & Str(xLegisladores) & " Legisladores registrados, pero hay " & Str(xLegisladores_Habilitados) & " legisladores habilitados. Error en la gestión de datos básicos.", vbCritical + vbInformation, "SQV AUDITORIA"
    End If
    Cn.Close
    Set rs = Nothing
    Set Cn = Nothing
End Sub

Public Function AsignarColor(xBancaColor As Long) As String
        
    Dim xMostrarResultadoVotnum As Boolean

    xMostrarResultadoVotnum = False ' Verdadero si se desea mostrar resultados en votnum

    If EstadoActual.VectorError(xBancaColor) = ERROR_IOC Then
                AsignarColor = cAZUL 'ERROR DE SWITCH DE BANCA
    ElseIf (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum") _
        And (EstadoActual.EstadoVotacion_y_PasList = "finalizada" Or EstadoActual.EstadoVotacion_y_PasList = "empate") Then  'Presentacion de resultados
            If EstadoActual.VectorPresenciaCong(xBancaColor) = BANCA_INHABILITADA Then
                AsignarColor = cMARRON '=HCDN, ANTES=cNARANJA
            ElseIf EstadoActual.VectorPresenciaCong(xBancaColor) = AUSENTE Then
                AsignarColor = IIf(xBancaColor = 0, cCELESTE, cBLANCO) 'ANTES cGRIS) 'PRESIDENTE SIEMPRE IDENTIFICADO
                If xBancaColor = 0 Then
                    If EstadoActual.TipoDeOperacion = "votnom" Then
                        If EstadoActual.VectorResultados(xBancaColor) <> " " Then
                            If EstadoActual.VectorResultados(xBancaColor) = AFIRMATIVO Then
                                    AsignarColor = cVERDE
                                ElseIf EstadoActual.VectorResultados(xBancaColor) = NEGATIVO Then
                                    AsignarColor = cROJO
                            End If
                        Else
                            If EstadoActual.ResultadoVotoPresidente = AFIRMATIVO Then
                                AsignarColor = cVERDE
                            ElseIf EstadoActual.ResultadoVotoPresidente = NEGATIVO Then
                                AsignarColor = cROJO
                            End If
                        End If
                    Else
                        If (Trim(EstadoActual.VectorResultados(xBancaColor)) <> "") Then
                            EstadoActual.VectorColor(xBancaColor) = cGRIS
                        ElseIf (Trim(EstadoActual.ResultadoVotoPresidente <> "")) Then
                            EstadoActual.VectorColor(xBancaColor) = cGRIS
                        End If
                    End If
                End If
            ElseIf EstadoActual.VectorPresenciaCong(xBancaColor) = PRESENTE Then
                If Not (EstadoActual.VectorIdentificacion(xBancaColor) = NO_IDENTIFICADO) Then
                    AsignarColor = cCELESTE
                Else
                    AsignarColor = IIf(xBancaColor = 0, cCELESTE, cAMARILLO)
                End If
                If EstadoActual.VectorResultados(xBancaColor) = ABSTENCION_AUTORIZADA Then
                    AsignarColor = cNEGRO
                ElseIf EstadoActual.VectorResultados(xBancaColor) = AFIRMATIVO Then
                    If EstadoActual.TipoDeOperacion = "votnum" Then
                        If xMostrarResultadoVotnum Then
                            AsignarColor = cVERDE
                        Else
                            AsignarColor = cGRIS 'cOLIVA
                        End If
                    Else
                        AsignarColor = cVERDE
                    End If
                ElseIf EstadoActual.VectorResultados(xBancaColor) = NEGATIVO Then
                    If EstadoActual.TipoDeOperacion = "votnum" Then
                        If xMostrarResultadoVotnum Then
                            AsignarColor = cROJO
                        Else
                            AsignarColor = cGRIS 'cOLIVA
                        End If
                    Else
                        AsignarColor = cROJO
                    End If
                End If
            End If
    ElseIf EstadoActual.VectorPresencia(xBancaColor) = BANCA_INHABILITADA And xBancaColor <> 0 Then
        AsignarColor = cMARRON '=HCDN, ANTES=cNARANJA
    ElseIf EstadoActual.VectorPresencia(xBancaColor) = AUSENTE And xBancaColor <> 0 Then
        AsignarColor = IIf(xBancaColor = 0, cCELESTE, cBLANCO) ' ANTES cGRIS) PRESIDENTE SIEMPRE IDENTIFICADO
    ElseIf EstadoActual.VectorPresencia(xBancaColor) = PRESENTE Or xBancaColor = 0 Then
        If Not (EstadoActual.VectorIdentificacion(xBancaColor) = NO_IDENTIFICADO) Then
            AsignarColor = cCELESTE
        Else
            AsignarColor = IIf(xBancaColor = 0, cCELESTE, cAMARILLO)
        End If
        If EstadoActual.VectorResultados(xBancaColor) = ABSTENCION_AUTORIZADA Then
            AsignarColor = cNEGRO
        ElseIf (EstadoActual.TipoDeOperacion = "votnom" Or EstadoActual.TipoDeOperacion = "votnum") _
                And InStr("votando larga", EstadoActual.EstadoVotacion_y_PasList) > 0 Then ' Acuse de voto realizado
                If InStr(AFIRMATIVO & NEGATIVO, EstadoActual.VectorResultados(xBancaColor)) > 0 Then
                    AsignarColor = cGRIS 'cOLIVA 'FUERZA COLOR para no mostrar resultado durante la votacion.
                End If
        End If
    End If
End Function
Public Sub ResetearPresidente()
    If True Then '090403
        With EstadoActual
            .VectorPresencia(0) = AUSENTE
            xPresidenteAnteriorLegislador = False
            .VectorIdentificacion(0) = 0 'antes 1100
            .VectorColor(0) = AsignarColor(0)
            .VectorResultados(0) = ABSTENCION
            .VectorResultadosCong(0) = ABSTENCION
            .VectorIdentificacionCong(0) = NO_IDENTIFICADO
            .VectorPresenciaCong(0) = BANCA_INHABILITADA
            .VMantEstado(0) = ABSTENCION
            .VTipoIdentificacion(0) = TIPO_IDENTIFICACION_HUELLA
        End With
    End If
End Sub
Public Function SesionValida(Periodo As String, pSesion As Long) As Boolean
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
Dim Sesion As String
Sesion = Trim(Str(pSesion))
frmMain.SetearRsAux "SELECT * FROM sesion WHERE Período_Legislativo = '" & Periodo & "' AND Sesión = " & Sesion & " AND Estado_sesión = 'abierta'", rsTemp
If rsTemp.EOF Then
    SesionValida = False
Else
    SesionValida = True
End If
rsTemp.Close
Set rsTemp = Nothing
End Function
Public Function DevolverLeyendaTipo(codigo As String) As String
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
frmMain.SetearRsAux "SELECT Leyenda_para_cartel FROM tipmay WHERE Tipo_De_Mayoria = '" & codigo & "'", rsTemp
If rsTemp.EOF Then
    DevolverLeyendaTipo = "No encontrado"
Else
    DevolverLeyendaTipo = Trim(rsTemp.Fields(0))
End If
rsTemp.Close
Set rsTemp = Nothing
End Function
Public Function DevolverLeyendaBase(codigo As String) As String
Dim rsTemp As ADODB.Recordset
Set rsTemp = New ADODB.Recordset
frmMain.SetearRsAux "SELECT Leyenda_para_cartel FROM basemay WHERE identificador_en_mensajes= '" & codigo & "'", rsTemp
If rsTemp.EOF Then
    DevolverLeyendaBase = "No encontrado"
Else
    DevolverLeyendaBase = Trim(rsTemp.Fields(0))
End If
rsTemp.Close
Set rsTemp = Nothing
End Function
Public Sub MandarImprimir()
frmMain.EjecutarSQL ("UPDATE ComunicacionRapida SET ImprimirActa = 1")
End Sub
Public Sub BorrarImpresion()
frmMain.EjecutarSQL ("UPDATE ComunicacionRapida SET ImprimirActa = 0")
End Sub
Public Sub SetTiempoTranscurrido()
    frmMain.EjecutarSQL ("UPDATE ComunicacionRapida SET TiempoTranscurrido = 1")
End Sub
Public Sub UnsetTiempoTranscurrido()
    frmMain.EjecutarSQL ("UPDATE ComunicacionRapida SET TiempoTranscurrido = 0")
End Sub
