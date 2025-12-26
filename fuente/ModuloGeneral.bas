Attribute VB_Name = "ModuloGeneral"
Option Explicit

Public UltimoVoto As String
Public Transcurrido As Boolean
'Variables de Conexion SQL Server.
'***************************************************
Public Cn                        As ADODB.Connection
Public RsSQV                     As ADODB.Recordset
Public RsSB                      As ADODB.Recordset
Public RsBanca                   As ADODB.Recordset
Public RsCola                    As ADODB.Recordset
Public strConexionSQL            As String
Public CantidadInsertados       As Long 'Modificacion HCDN 15 Febrero

'***************************************************
' Defino Constantes                                *
'***************************************************
Global Const DEPURACION         As Boolean = False
Global Const cNIVEL_LOG         As Integer = 3 '3 todos, 2 algunos, 1 solo incidentes
Global Const MAX_BANCA          As Integer = 256
Global Const MODOLIGHT          As Boolean = True

'Variables Generales
'***************************************************
Public CountTimer                As Long
Public TickTLEVER                As Long
Public ContadorTLEVER            As Integer
Public msgSQV                    As MensajeSQV
Public Banca_ip()                As BancaIP
Public Skt2B()                   As SktBanca
Public B2Skt()                   As BancaSkt
Public tLegisladores()           As Legisladores
Public SecuenciaSalida()         As Long
Public UltimaSecuenciaEnviada()  As Long
Public BancasActivas             As Integer
Public ConexionesAbiertas        As Integer
Public StrProximo(0 To 256)      As String
Public strDato(0 To 256)         As String
Public nUltimoMensajeSQV         As Long ' me marca el ultimo numero que voy leyendo de la cola del SQV.
Public UltimoEstadodeBanca()     As String
Public UltimoEstadodePresencia() As String
Public UltimaSecuenciaBanca()    As String
'Public EstadoxEnviar()           As Boolean
Public sLegisla                  As TLegislador
Public strLegisla(0 To 10)       As String 'string para armar la cadena de legisladores
Public FlagTimerLeoCola          As Boolean
Public CountTimerAux             As Long
Public Version                   As String
Public BanderaReset()            As Boolean
Public FueConfigurada()          As Boolean
Public EnviandoHuellas()         As Boolean
Public EnvioCompletado(0 To MAX_BANCA)         As Boolean
'**********VECTORES PARA EL CORRECTO ENVIO DE LOS MENSAJES**********
Public LogEnvioCompletado(0 To MAX_BANCA)      As String
Public VectorEnvio(0 To MAX_BANCA) As String
Public VectorTicks(0 To MAX_BANCA) As Long
Public LogBancasMuertas(0 To MAX_BANCA) As String
'Para controlar la repeticion de mensajes (secuencia) en uso?
Public Control_Secuencia(0 To MAX_BANCA)       As String
Public VectorSAUTOD(0 To MAX_BANCA)     As String
Public TICK_LOG As Long
Public Prefijo_Tick As String

'------------------------------------------------------------------------
' Funcion que devuelve tick del procesador desde que se inicio el windows
'------------------------------------------------------------------------
Public Declare Function GetTickCount Lib "kernel32" () As Long


'***************************************************
' Conexion con el SQL Server -----------------------
'***************************************************

Public Sub AbrirConexionSQLServer()
    Dim strcad     As String
    Set Cn = New ADODB.Connection
    'Cadena de Conexion de la base sqv_config
    If True Then 'vmGen
        strConexionSQL = "PROVIDER=SQLOLEDB.1;PASSWORD=hcdn11;PERSIST SECURITY INFO=TRUE;USER ID=SQV;INITIAL CATALOG=SQV_Config;DATA SOURCE=10.1.1.5"
    Else 'SBA
        strConexionSQL = "Provider=SQLOLEDB.1;Password=unipaas;Persist Security Info=True;" _
                              & "User ID=sqv;Initial Catalog=sqv_config;Data Source=siprevo"
    End If
    With Cn
        .ConnectionString = strConexionSQL
        .CursorLocation = adUseServer
        .ConnectionTimeout = 30
        .Open
    End With
    'Cargo los Recorset
    Set RsSQV = New ADODB.Recordset
    Set RsSB = New ADODB.Recordset
    Set RsBanca = New ADODB.Recordset
    'Leo la base viegente de Uso y cambio la conexion
    strcad = "SELECT valor FROM configuracion WHERE variable = 'base_vigente'"
    SetearRsBanca (strcad)
    strConexionSQL = RsBanca.Fields(0).Value
    If InStr(strConexionSQL, "prueba") > 0 Then
        FormMain.lblBasePrueba.Visible = True
    End If
    With Cn
        .Close
        .ConnectionString = strConexionSQL
        .CursorLocation = adUseServer
        .ConnectionTimeout = 30
        .Open
        Call MostrarErr("Conexión OK.")
        'Call MostrarErr("Cadena de Conexion : " & strConexionSQL)
    End With
End Sub


'**************************************************
'* Cargo tabla de Legisladores en un Type
'**************************************************
Public Sub CargarLegisladores()
    Dim strcad   As String
    Dim cont     As Integer
    strcad = "SELECT * From Legisladores ORDER BY Tipo DESC, CAST(id AS int)"
    Call SetearRsBanca(strcad)
    ReDim tLegisladores(0 To RsBanca.RecordCount)
    cont = 0
    If Not RsBanca.RecordCount = 0 Then
        RsBanca.MoveFirst
        While Not RsBanca.EOF
            tLegisladores(cont).sBanca = RsBanca!indicebanca
            tLegisladores(cont).sId = Val(RsBanca!id)
            tLegisladores(cont).sMantenimiento = IIf(RsBanca!tipo = 0, True, False)
            cont = cont + 1
            RsBanca.MoveNext
        Wend
    End If
End Sub

'************************************************************************
' Paso Numero de id Legislador y me devuelve el indice donde esta Cargado
'************************************************************************

Public Function indiceLegislador(idlegislador As Long) As Integer
    Dim naUX As Long
    For naUX = 0 To UBound(tLegisladores)
        If tLegisladores(naUX).sId = idlegislador Then
            indiceLegislador = tLegisladores(naUX).sBanca
            naUX = UBound(tLegisladores)
        End If
    Next
End Function

Public Function MensajeTNACKNporTimeOut(fMensaje As String) As Boolean
    Dim sSecuencia As String
    Dim sComando  As String
    
    sSecuencia = Mid(fMensaje, 1, 1)
    sComando = Mid(fMensaje, 2, 6)
    MensajeTNACKNporTimeOut = False
    If sComando = "TNACKN" Then MensajeTNACKNporTimeOut = (sSecuencia >= "A" And sSecuencia <= "Z")
End Function

'***************************************************
' Interpreto los Datos que llegan desde las Bancas *
'***************************************************
Public Sub InterpretaDatosSkt(fSocket As Integer, fMensaje As String)
On Error GoTo GetE
    Dim msgSQV      As MensajeSQV
    Dim nIdent      As String
    Dim sMensaje    As String
    Dim sSecuencia  As String
    Dim nVersion    As String
    Dim Borrar      As Boolean
    Dim i As Integer
    Dim aTemp As String
    Borrar = True
    fMensaje = SacaNulos(fMensaje)
    If Len(fMensaje) = 0 Then Exit Sub
    'Guardamo log de datos recibidos :)
    '****************************************
    If Asc(Left(fMensaje, 1)) = 0 Then fMensaje = Mid(fMensaje, 2, 1000) '090225
    If Not Mid(fMensaje, 2, 6) = "TESTAD" Then
        'Call GuardarLog(Str(fSocket), "LOG1 - " & fMensaje)
    End If
    sSecuencia = Mid(fMensaje, 1, 1)
    sMensaje = Mid(fMensaje, 2, Len(fMensaje) - 1)
    Select Case UCase(Mid(fMensaje, 2, 6))
        ' Tipos de Casos que se Pueden Presentar
        '***************************************
        Case "TIDVAL" 'Identificacion Valida
            'check
            If Control_Secuencia(fSocket) = fMensaje Then
                'Log_DEBUG ("******** WARNING SOCKET " & fSocket & ": Se evitó mensaje TIDVAL repetido")
                Exit Sub
            Else
                Control_Secuencia(fSocket) = fMensaje
            End If
            '---------------------------------------------------
            ' Envia Autentificacion Positiva
            '---------------------------------------------------
            'nIdent = Str(Int(Mid(sMensaje, 9, Len(sMensaje) - 8)))
            ' nIdent = Mid(sMensaje, 9, Len(sMensaje) - 8) ' cambiado 090908
            nIdent = Mid(sMensaje, 8, 16) '090908
            
            With msgSQV
                .sTipo = "mevt"
                .sObjeto = Str(Skt2B(fSocket).Banca)
                .sComponente = "term.auth"
                .sAtributo = "result"
                .sValor = nIdent
            End With
            Call InsertarMsgSQV(msgSQV)
            'Call FormMain.EnviarxSkt(sSecuencia, fSocket, "SACKID")
            'Manejo de Secuencia de busqueda para las bancas. :)
            '*****************************************************
            If tLegisladores(indiceLegislador(Val(nIdent))).sMantenimiento = False Then
                If InStr(Mid(Banca_ip(fSocket).tBancaSecuencia, 1, 12), strString(4, Hex(Fix(indiceLegislador(Val(nIdent)))), "0", "I")) = 0 Then
                    Banca_ip(fSocket).tBancaSecuencia = strString(4, Hex(Fix(indiceLegislador(Val(nIdent)))), "0", "I") & Banca_ip(fSocket).tBancaSecuencia
                End If
            End If
        Case "TIDINV" 'Identificacion Invalida
            '---------------------------------------------------
            ' Envia Autentificacion Negativa
            '---------------------------------------------------
            'Call EliminarMensajeCola(Skt2B(fSocket).Banca, fMensaje, True)
            'Se eliminan los mensajes de la banca ya que si da TIDINV esta pidiendo identificacion.
            'En esta instancia no hay necesidad de envio de STATUS, no hay Votacion, etc.
            'Se sabe que esta sentado porque sino no estaría en EIDRXH
            With msgSQV
                .sTipo = "mevt"
                .sObjeto = Str(Skt2B(fSocket).Banca)
                .sComponente = "term.auth"
                .sAtributo = "result"
                .sValor = "negative"
            End With
            Call InsertarMsgSQV(msgSQV)
            Call FormMain.EnviarxSkt(sSecuencia, fSocket, "SACKNL")
        Case "TIDOUT" 'Tiempo de Espera de Huella Agotado
            '---------------------------------------------------
            ' Envia Autentificacion TimeOut
            '---------------------------------------------------
            With msgSQV
                .sTipo = "mevt"
                .sObjeto = Str(Skt2B(fSocket).Banca)
                .sComponente = "term.auth"
                .sAtributo = "result"
                .sValor = "timeout"
            End With
            Call InsertarMsgSQV(msgSQV)
            'continua reintentando permanentemente
            Call FormMain.EnviarxSkt(sSecuencia, fSocket, "SACKNL")
        Case "TVOTOX" 'Respuesta a un Voto
            '---------------------------------------------------
            ' Envia Cual pulsador de Votacion se Ejecuto
            '---------------------------------------------------
            With msgSQV
                .sTipo = "mevt"
                .sObjeto = Str(Skt2B(fSocket).Banca)
                Select Case Mid(sMensaje, 8, 1)
                    Case "S"
'                        If Transcurrido = False Then
'                            Call FormMain.EnviarxSkt("X", fSocket, "SACKVT S")
'                            UltimoVoto = "S" & Transcurrido
'                        End If
                        .sComponente = "term.keyb.si"
                    Case "N"
'                        If Transcurrido = False Then
'                            Call FormMain.EnviarxSkt("X", fSocket, "SACKVT N")
'                            UltimoVoto = "N" & Transcurrido
'                        End If
                        .sComponente = "term.keyb.no"
                    Case "A"
                        .sComponente = "term.keyb.ab"
                End Select
                .sAtributo = "state"
                .sValor = "on"
            End With
'            If Transcurrido = False Then
                Call InsertarMsgSQV(msgSQV)
'            End If
        Case "TESTAD" 'Respuesta a Un pedido de Estado
             '---------------------------------------------------
             ' Envia ESTADO DEL TERMINAL
             '---------------------------------------------------
             With msgSQV
                .sTipo = "mevt"
                .sObjeto = Str(Skt2B(fSocket).Banca)
                .sComponente = "term"
                .sAtributo = "state"
                .sComentario = UCase(Mid(sMensaje, 8, 6))
                Banca_ip(fSocket).tEstado = UCase(Mid(sMensaje, 8, 6))
                .sValor = "ok"
             End With
             If Not (Left(Trim(LCase(UltimoEstadodeBanca(fSocket))), 2) = "ok") Then
                Call InsertarMsgSQV(msgSQV)
                UltimoEstadodeBanca(fSocket) = "ok"
             End If
             '---------------------------------------------------
             ' Envia ESTADO DEL switch
             '---------------------------------------------------
             With msgSQV
                .sTipo = "mevt"
                .sObjeto = Str(Skt2B(fSocket).Banca)
                .sComponente = "term.seat"
                .sAtributo = "switch"
                If Mid(sMensaje, 15, 1) = "A" Then
                    .sValor = "open" 'A u otro valor
                ElseIf Mid(sMensaje, 15, 1) = "P" Then
                    .sValor = "closed" 'P
                Else
                    FormMain.ErrorSwitchBanca1 (fSocket)
                End If
             End With
             If Not UCase(msgSQV.sValor) = UCase(UltimoEstadodePresencia(fSocket)) Then
                msgSQV.sComentario = "TESTAD:" & sMensaje
                Call InsertarMsgSQV(msgSQV)
                UltimoEstadodePresencia(fSocket) = UCase(msgSQV.sValor)
             End If
             If Not (FueConfigurada(fSocket)) Then
                 FueConfigurada(fSocket) = True
                 Call FormMain.Configurar(fSocket)
             End If
             If Skt2B(fSocket).Banca = 123 Then
                fMensaje = fMensaje
            End If
             If InStr(fMensaje, "EINACT") > 0 Or InStr(fMensaje, "EIDACP") > 0 Or InStr(fMensaje, "EIDRXH") > 0 Then
                With msgSQV
                   .sTipo = "mevt"
                   .sObjeto = Str(Skt2B(fSocket).Banca)
                   .sComponente = "term"
                   If InStr(fMensaje, "EINACT") > 0 Then
                        .sAtributo = "einact"
                    ElseIf InStr(fMensaje, "EIDACP") > 0 Then
                        .sAtributo = "eidacp"
                    ElseIf InStr(fMensaje, "EIDRXH") > 0 Then
                        .sAtributo = "eidrxh"
                    End If
                   .sValor = LCase(Mid(sMensaje, 15, 1))
                   .sComentario = "Aviso de EINACT"
                End With
                Call InsertarMsgSQV(msgSQV)
             End If
        
        Case "TPRESE" 'Se cerro el switch de la Banca
            '---------------------------------------------------
            ' Envia Switch close a consola... Legislador Sentado
            '---------------------------------------------------
            With msgSQV
                .sTipo = "mevt"
                .sObjeto = Str(Skt2B(fSocket).Banca)
                .sComponente = "term.seat"
                .sAtributo = "switch"
                .sValor = "closed"
                .sComentario = "TPRESE"
                UltimoEstadodePresencia(fSocket) = UCase(msgSQV.sValor)
            End With
            Call InsertarMsgSQV(msgSQV)
            'Envio Ack
            Call FormMain.EnviarxSkt(sSecuencia, fSocket, "SACKNL")
            If Not (FueConfigurada(fSocket)) Then
                FueConfigurada(fSocket) = True
                Call FormMain.Configurar(fSocket)
            End If
        
        Case "TAUSEN" 'Se Abrio el switch de la Banca
            '---------------------------------------------------
            ' Envia Switch close a consola... Legislador Sentado
            '---------------------------------------------------
            With msgSQV
                .sTipo = "mevt"
                .sObjeto = Str(Skt2B(fSocket).Banca)
                .sComponente = "term.seat"
                .sAtributo = "switch"
                .sValor = "open"
                .sComentario = "TAUSEN"
                UltimoEstadodePresencia(fSocket) = UCase(.sValor)
            End With
            Call InsertarMsgSQV(msgSQV)
            'Envio Ack
            Call FormMain.EnviarxSkt(sSecuencia, fSocket, "SACKNL")
        
        Case "TACKNL" 'Respuesta ACK
            '------------------------------------------------------------
            ' Recibo Proceso de Acknoledge
            '------------------------------------------------------------
            ' SE PODRIA Procesar el switch si se recibe en la posicion 15 P o A
            If Mid(sMensaje, 14, 1) = "?" Then
                    'Call ErrorSwitchBanca1(fSocket)
                Call FormMain.ErrorSwitchBanca1(fSocket)
            End If
            If EnviandoSNUVER = True Then
                RespuestasSNUVER = RespuestasSNUVER + 1
            End If
            If VectorSAUTOD(Skt2B(fSocket).Banca) <> "0" Then
                If InStr(fMensaje, "EIDACPP") Then
                    With msgSQV
                        .sTipo = "mevt"
                        .sObjeto = Str(Skt2B(fSocket).Banca)
                        .sComponente = "term.auth"
                        .sAtributo = "result"
                        .sValor = VectorSAUTOD(Skt2B(fSocket).Banca)
                        .sComentario = "SAUTOD"
                    End With
                    Call InsertarMsgSQV(msgSQV)
                    VectorSAUTOD(Skt2B(fSocket).Banca) = "0"
                End If
            End If
        Case "TLEVER"
            If GetTickCount - TickTLEVER < 10000 Then
                ContadorTLEVER = ContadorTLEVER + 1
            Else
                ContadorTLEVER = 0
            End If
            If ContadorTLEVER > 220 Then
                FormMain.txtLog.Text = "YA PUEDE ENVIAR LAS HUELLAS CON SEGURIDAD."
            End If
            TickTLEVER = GetTickCount
            '------------------------------------------------------------
            ' Recibo NUMERO DE VERSION
            '------------------------------------------------------------
'            Dim strcad As String
'            Dim vSQV As String
'            Dim Version_SQV_Actualizada As Boolean
'            Dim rsT As ADODB.Recordset
'            Set rsT = New ADODB.Recordset
'            strcad = "SELECT TOP 1 version_datos_sqv FROM config"
'            SetearRsW strcad, rsT
'            rsT.MoveFirst
'            vSQV = rsT.Fields(0).Value
'            Set rsT = Nothing
'            Set rsT = New ADODB.Recordset
'            Version_SQV_Actualizada = True
'            strcad = "SELECT version_datos_sqv FROM BancasIP WHERE BancaNumero = " & Skt2B(fSocket).Banca
'            SetearRsW strcad, rsT
'            rsT.MoveFirst
'            If (rsT.Fields(0).Value <> vSQV) Then
'                Version_SQV_Actualizada = False
'            End If
'            rsT.Close
'            Set rsT = Nothing
            nVersion = Mid(sMensaje, 8, 12)
            If Skt2B(fSocket).Banca > 0 And Not Trim(nVersion) = Trim(Banca_ip(fSocket).tVersion) Then
'                Call MostrarErr("Banca : " & Str(Skt2B(fSocket).Banca) & " ---- Version : Banca:[" & nVersion & "] BD:[" & Banca_ip(fSocket).tVersion & "] No COINCIDEN !" & IIf(Version_SQV_Actualizada, "", "Actualizar SQV"), Str(Skt2B(fSocket).Banca))
                'Call MostrarErr("Banca : " & Str(Skt2B(fSocket).Banca) & " ---- Version : Banca:[" & nVersion & "] BD:[" & Banca_ip(fSocket).tVersion & "] No COINCIDEN !", Str(Skt2B(fSocket).Banca))
                'Call MostrarErr("Banca : " & Str(Skt2B(fSocket).Banca) & " ---- Version DESACTUALIZADA", Str(Skt2B(fSocket).Banca))
            End If
            Call GuardarVersionBanca(Str(Skt2B(fSocket).Banca), nVersion)
        Case "TNACKN" 'Respuesta que no se puedo procesar el Comando
            If InStr(fMensaje, "SNUVER") > 0 Then
                fMensaje = fMensaje
            End If
            If InStr(fMensaje, "SFINVT") Or InStr(fMensaje, "SLIMVT") Or InStr(fMensaje, "SVOTAR") Then
                'Se habria enviado un brc y la banca no estaba en condiciones de emitir o cancelar el voto.
                'Log_DEBUG ("------------------------SE ENVIO SCANCL------------------- A Banca" & Str(Skt2B(fSocket).Banca))
                Call EliminarMensajeCola(Skt2B(fSocket).Banca, fMensaje, False) ' True elimina todos los mensajes
            End If
            If InStr(fMensaje, "SFINNU") Or InStr(fMensaje, "ELISVTP") Or InStr(fMensaje, "SVOTNU") Then
                'Se habria enviado un brc y la banca no estaba en condiciones de emitir o cancelar el voto.
                Call EliminarMensajeCola(Skt2B(fSocket).Banca, fMensaje, False) ' True elimina todos los mensajes
            End If
            If InStr(fMensaje, "SACKID") Then
                'La banca mandó un TIDVAL (un diputado se identificó correctamente
                'Al mismo tiempo que se le intentó asignar una ID desde la consola.
                Call EliminarMensajeCola(Skt2B(fSocket).Banca, fMensaje, False) ' True elimina todos los mensajes
            End If
            If InStr(fMensaje, "SIDRXH") > 0 Then
                If Not (InStr(fMensaje, "EINACTP") > 0) Then 'Evita matarla si no está en EINACT
                'Si esta en EINACT despues de los reintentos, muere
                    Call EliminarMensajeCola(Skt2B(fSocket).Banca, fMensaje, False)
                    Log_DEBUG ("Se evitó muerte de banca TNACKN SIDRXH : " & fSocket & " | CASO: " & fMensaje)
                    If Not MODOLIGHT Then
                        Call Log_Banca(Trim(Str(fSocket)) & ".txt", Now() & "  - Se evitó muerte de banca TNACKN SIDRXH : " & fSocket & " | CASO: " & fMensaje)
                    End If
                End If
            End If
            If MensajeTNACKNporTimeOut(fMensaje) Then
                If InStr(fMensaje, "SRLEGI") > 0 Then
                    Borrar = False
                    Call FormMain.WSocketClose(fSocket, "TNACKN Timeout Se pide reintento (normal para SRLEGI) " & fMensaje)
                Else
                    Borrar = True
                    Log_DEBUG (fSocket & "TNACKN Timeout ELIMINADO de la cola (normal)" & fMensaje)
                End If
            Else
                Borrar = False 'banana dos
            End If
            If False Then
                RsCola.Filter = "(Tick<>'0') AND (Socket=" & fSocket & ") AND (Secuencia='" & sSecuencia & "')"
                If Not RsCola.EOF Or RsCola.RecordCount < 0 Then
                    RsCola.MoveFirst
                    'Antes de borrar tengo que preguntar si es alguna secuencia conocida
                    Select Case Left(RsCola!mensaje, 6)
                        Case "SRLEGI"
                        If MensajeTNACKNporTimeOut(fMensaje) Then
                            Borrar = True
                            EnviandoHuellas(Skt2B(fSocket).Banca) = False 'mandarina
                        Else
                            Borrar = False 'banana dos
                        End If
                    End Select
                    RsCola.UpdateBatch
                    RsCola.Filter = adFilterNone
                    If RsCola.RecordCount <> 0 Then
                        RsCola.MoveFirst
                    End If
                End If
            End If
        Case "EIDACP" 'Respuesta a un sackid
                
        Case "TVERRE" 'TVERREADY: La terminal termino de procesar todas las huellas enviadas
            'Call GuardarLog(Str(fSocket), fMensaje)
            EnviandoHuellas(Skt2B(fSocket).Banca) = False
            Call MostrarEnviandoHuellas
            Call EnviarSktxCola(Str(Skt2B(fSocket).Banca), "SLEVER")
            'If B2Skt(Val(.sObjeto)).Estado = True Then
            '    Call FormMain.WSocketClose(B2Skt(Val(.sObjeto)).Socket)
            'Else
            '    With msgSQV
            '        .sTipo = "mevt"
            '        .sObjeto = .sObjeto
            '        .sComponente = "term"
            '        .sAtributo = "state"
            '        .sValor = "off"
            '    End With
            '    Call InsertarMsgSQV(msgSQV)
            '    UltimoEstadodeBanca(B2Skt(Val(.sObjeto)).Socket) = "off"
            'End If
            
        Case Else
            Call Log_DEBUG("Err Msg no reconocido: " & Str(fSocket) & " " & fMensaje)
            If FormMain.WSocket(fSocket).State = 7 Then
                Call FormMain.EnviarxSkt("f", fSocket, "SRESET")
            End If
            'AQUI CERRAR
    End Select
    'actualizar ultimo comando procesado por socket
    'Elimino mensaje de la cola de mensaje
    '***********************************************
    If Borrar Then
        Call EliminarMensajeCola(fSocket, fMensaje, False)
    End If
    
Exit Sub
GetE:
If Err.Number = 28 Then
    End
Else
    Log_DEBUG (Err.Description)
    Resume Next
End If
End Sub

Private Sub GuardarVersionBanca(xBanca As String, nVersion As String)

Dim strcad As String
    Set RsBanca = New ADODB.Recordset
    strcad = "SELECT * FROM BANCASIP WHERE BancaNumero =(" & Trim(xBanca) & ")" & " ORDER BY BANCANUMERO ASC"
    Call SetearRsBanca(strcad)
    RsBanca.MoveFirst
    Do While Not RsBanca.EOF
        RsBanca!version_datos_banca = IIf(RsBanca!BancaNumero = 0, "No aplicable", nVersion)
        RsBanca.MoveNext
    Loop

End Sub


'***************************************************
' Interpreto los Datos que llegan desde el SQV     *
'***************************************************
Public Sub InterpretaDatosSQV(fMsgSQV As MensajeSQV)
    Dim msgSQV        As MensajeSQV
    Dim naUX          As Long
    Dim fDesconocido  As Boolean
    
    fDesconocido = True ' Si no entra en ninguna Funcion es un comando desconocido...
    With fMsgSQV
        'Guardo log de mensajes
        '**********************************************
        'Call GuardarLog(.sObjeto, , .sTipo & "/" & .sObjeto & "/" & .sComponente & "?" & .sAtributo & "=" & .sValor & "/" & .sComentario)
        
        'Interpretacion de los datos
        '***************************************************************************************
        '---------------------------------------------------------------------
        ' Simulaciones
        '---------------------------------------------------------------------
        If .sTipo = "simulacion_voto" Then
            Dim x As Integer
            For x = 1 To ConexionesAbiertas - 1
                Call FormMain.EnviarxSkt("*", x, "SCONFG 0000000")
            Next x
        End If
        '---------------------------------------------------------------------
        ' Actualización de IPs
        '---------------------------------------------------------------------
        If .sTipo = "ips" And .sValor = "update" Then
            FormMain.ActualizaIPs
        End If
        '---------------------------------------------------------------------
        ' Primero Procesamos los datos tipo = mset
        '---------------------------------------------------------------------
        If .sTipo = "mset" And LCase(.sComponente) = "term.auth" And LCase(.sAtributo) = "action" Then
               Select Case LCase(.sValor)
                    Case "auth_data_refresh"
                        'Enviar Tabla Legisladores
                        '******************************************************
                         fDesconocido = False
                         'Call EnviarHuellas   ' Envia huellas de los legisladores.
                     
                     Case "auth_start"
                        ' ----------------------------------------------------
                        ' Proceso de Identificacion solo huella
                        ' ----------------------------------------------------
                        fDesconocido = False
                        'Call EnviarSktxCola(.sObjeto, "SCANCL")
                        'Call EnviarSktxCola(.sObjeto, "STATUS") ' ^003C005F")
                        'If IsNumeric(.sObjeto) Then
                        '    Call FormMain.EnviarxSkt("f", Int(.sObjeto), "STST01")
                        'End If
                        'Call EnviarSktxCola(.sObjeto, "STST01")
                        If InStr(.sObjeto, ";") = 0 And InStr(.sObjeto, "brc") = 0 Then
                            Call FormMain.EnviarxSkt("f", Val(.sObjeto), "SCANCL")
                            Call FormMain.EnviarxSkt("*", Val(.sObjeto), "SIDRXH")
                        Else
                            Call EnviarSktxCola(.sObjeto, "SCANCL")
                            Call EnviarSktxCola(.sObjeto, "SIDRXH") ' ^003C005F")
                        End If
                     
                     Case "auth_key_start"
                        ' ----------------------------------------------------
                        ' Proceso de Identificacion solo Teclado
                        ' ----------------------------------------------------
                        fDesconocido = False
                        Call EnviarSktxCola(.sObjeto, "SCANCL")
                        Call EnviarSktxCola(.sObjeto, "SIDRNX") 'Pedido p/ Teclado
                     
                     Case "auth_test_key_start"
                        ' ----------------------------------------------------
                        ' Proceso de Identificacion solo Teclado
                        ' ----------------------------------------------------
                        fDesconocido = False
                        Call EnviarSktxCola(.sObjeto, "SCANCL")
                        Call EnviarSktxCola(.sObjeto, "SIDRNX ^" & Banca_ip(1).tbancaMinMaxMan) 'Pedido p/ Teclado Mant.
                     
                     
                     Case "auth_restart"
                        ' ----------------------------------------------------
                        ' Proceso de Identificacion solo huella
                        ' ----------------------------------------------------
                        fDesconocido = False
                        'If IsNumeric(.sObjeto) Then
                        '    Call FormMain.EnviarxSkt("f", Int(.sObjeto), "STST02")
                        'End If
                        Call EnviarSktxCola(.sObjeto, "SCANCL")
                        Call EnviarSktxCola(.sObjeto, "SIDRXH")
                     Case "auth_test"
                        'Proceso identificacion mantenimiento
                        fDesconocido = False
                        Call EnviarSktxCola(.sObjeto, "SCANCL")
                        Call EnviarSktxCola(.sObjeto, "SIDRXH ^" & Banca_ip(1).tbancaMinMaxMan)
                
                End Select
        End If
        
        '----------------------------------------------------------
        'Sincronizar datos de bancas
        '----------------------------------------------------------
        If .sTipo = "mset" And LCase(.sComponente) = "term.mon" And LCase(.sAtributo) = "action" And LCase(.sValor) = "sync" Then
            fDesconocido = False
            Call FormMain.SincronizarBancas(.sObjeto)
        End If

        '----------------------------------------------------------
        'Poner la banca en estado inicial
        '----------------------------------------------------------
        If LCase(.sComponente) = "term.mon" And LCase(.sAtributo) = "action" And LCase(.sValor) = "reset" Then
            fDesconocido = False
            If .sObjeto = "brc" Then
                For naUX = 0 To ConexionesAbiertas - 1
                    If FormMain.WSocket(naUX).State = 7 Then
                        Call FormMain.WSocketClose(CInt(naUX), "SQVSRV brc")
                        UltimoEstadodeBanca(naUX) = "off"
                    End If
                Next
                Call FormMain.ReconectarBancas
            Else
                If InStr(.sComentario, "modo_deshabilitar") > 0 Then
                    If BancasDeshabilitadas(Val(.sObjeto)) = False Then
                        BancasDeshabilitadas(Val(.sObjeto)) = True
                        Cn.Execute ("UPDATE BancasDeshabilitadas SET habilitada = 0 WHERE banca = " & .sObjeto)
                    Else
                        BancasDeshabilitadas(Val(.sObjeto)) = False
                        Cn.Execute ("UPDATE BancasDeshabilitadas SET habilitada = 1 WHERE banca = " & .sObjeto)
                    End If
                End If
                If B2Skt(Val(.sObjeto)).Estado = True Then
                    Call FormMain.WSocketClose(B2Skt(Val(.sObjeto)).Socket, "SQVSRV")
                    With msgSQV
                        .sTipo = "mevt"
                        .sObjeto = fMsgSQV.sObjeto
                        .sComponente = "term"
                        .sComentario = "a01"
                        .sAtributo = "state"
                        .sValor = "off"
                    End With
                    Call InsertarMsgSQV(msgSQV)
                    UltimoEstadodeBanca(B2Skt(Val(.sObjeto)).Socket) = "off"
                Else
                    With msgSQV
                        .sTipo = "mevt"
                        .sObjeto = .sObjeto
                        .sComponente = "term"
                        .sComentario = "a01"
                        .sAtributo = "state"
                        .sValor = "off"
                    End With
                    Call InsertarMsgSQV(msgSQV)
                    UltimoEstadodeBanca(B2Skt(Val(.sObjeto)).Socket) = "off"
                End If
            End If
        End If
        '----------------------------------------------------------
        'Poner la banca en estado inicial
        '----------------------------------------------------------
        If LCase(.sComponente) = "term.mon" And LCase(.sAtributo) = "action" And LCase(.sValor) = "resethard" Then
            fDesconocido = False
            'FALTA PONER MANEJO DE BRC
            If .sObjeto = "brc" Then
                For naUX = 0 To ConexionesAbiertas - 1
                    If FormMain.WSocket(naUX).State = 7 Then
                        Call FormMain.EnviarxSkt("f", B2Skt(naUX).Socket, "SRESET")
                    End If
                Next
                Call FormMain.ReconectarBancas
            Else
                Call FormMain.EnviarxSkt("f", B2Skt(Val(.sObjeto)).Socket, "SRESET")
            End If
            FueConfigurada(B2Skt(Val(.sObjeto)).Socket) = False
        End If
        'Proceso de Votacion
        '***************************************************************************************
        '----------------------------------------------------------
        'Acuse de recibo de Voto SI
        '----------------------------------------------------------
        If LCase(.sComponente) = "term.ledk1" And LCase(.sAtributo) = "state" And LCase(.sValor) = "on" Then
            fDesconocido = False
            Call FormMain.EnviarxSkt("X", B2Skt(Val(.sObjeto)).Socket, "SACKVT S")
        End If
        '----------------------------------------------------------
        'Acuse de recibo de Voto NO
        '----------------------------------------------------------
        If LCase(.sComponente) = "term.ledk2" And LCase(.sAtributo) = "state" And LCase(.sValor) = "on" Then
            fDesconocido = False
            Call FormMain.EnviarxSkt("X", B2Skt(Val(.sObjeto)).Socket, "SACKVT N")
        End If
        '----------------------------------------------------------
        'Acuse de recibo de Voto ABSTENCION (no usado)
        '----------------------------------------------------------
        If LCase(.sComponente) = "term.ledk3" And LCase(.sAtributo) = "state" And LCase(.sValor) = "on" Then
            fDesconocido = False
        End If
                
        '----------------------------------------------------------
        ' Inicio de Votacion ' Si se recibe solo On, se asume que es numerica.
        '----------------------------------------------------------
        If LCase(.sComponente) = "term.keyb" And LCase(.sAtributo) = "state" And LCase(.sValor) = "on" Then
            fDesconocido = False
            Call EnviarSktxCola(.sObjeto, "SVOTNU 03") '090403
            If cNIVEL_LOG > 1 Then Call MostrarErr("Inicio votacion numerica por omision. No se recibio tipo de votacion.", .sObjeto)
        End If
        '----------------------------------------------------------
        'Fin de Votacion compatibilidad: Provisorio hasta que hande SQV NUEVO (No diferencia entre votacion nominal y numerica en .sValor, mas abajo se procesan cada caso por separado)
        '----------------------------------------------------------
        If LCase(.sComponente) = "term.keyb" And LCase(.sAtributo) = "state" And LCase(.sValor) = "off" Then
            fDesconocido = False
            'Call EnviarSktxCola(.sObjeto, "SFINVT")
            Call EnviarSktxCola(.sObjeto, "SFINNU")
            'If cNIVEL_LOG > 1 Then Call MostrarErr("Fin votacion numerica por omision. No se recibio tipo de votacion.", .sObjeto)
        End If
    
        '----------------------------------------------------------
        ' Inicio de Votacion Numerica'
        '----------------------------------------------------------
        If LCase(.sComponente) = "term.keyb" And LCase(.sAtributo) = "state" And LCase(.sValor) = "onvotnum" Then
            Dim iX As Integer
            Dim AEnviar As String
            fDesconocido = False
            If ModoSimulacion = True Then
                If SimulacionTipo = 1 Then
                    'Call EnviarSktxCola(.sObjeto, "SCONFG 1111111")
                    AEnviar = "SCONFG 1111111"
                Else
                    'Call EnviarSktxCola(.sObjeto, "SCONFG 0000000")
                    AEnviar = "SCONFG 0000000"
                End If
                For iX = 1 To ConexionesAbiertas - 1
                    Call FormMain.EnviarxSkt("f", iX, "SVOTNU 03")
                    'Call FormMain.EnviarxSkt("*", iX, AEnviar)
                Next iX
            Else
                Call EnviarSktxCola(.sObjeto, "SVOTNU 03") '090403
            End If
            'If cNIVEL_LOG > 2 Then If LCase(Trim(.sObjeto)) = "brc" Then Call MostrarErr("Inicio de Votacion Numerica " & Now, .sObjeto)
        End If
        '----------------------------------------------------------
        'Fin de Votacion Numerica
        '----------------------------------------------------------
        If LCase(.sComponente) = "term.keyb" And LCase(.sAtributo) = "state" And LCase(.sValor) = "offvotnum" Then
            fDesconocido = False
            'Call EnviarSktxCola(.sObjeto, "SFINVT")
            Call EnviarSktxCola(.sObjeto, "SFINNU")
            'If cNIVEL_LOG > 2 Then If LCase(Trim(.sObjeto)) = "brc" Then Call MostrarErr("Fin de Votacion Numerica " & Now, .sObjeto)
        End If
                
        '----------------------------------------------------------
        ' Inicio de Votacion Nominal
        '----------------------------------------------------------
        If LCase(.sComponente) = "term.keyb" And LCase(.sAtributo) = "state" And LCase(.sValor) = "onvotnom" Then
            fDesconocido = False
            Call EnviarSktxCola(.sObjeto, "SVOTAR 03")  '090403
            'If cNIVEL_LOG > 2 Then If LCase(Trim(.sObjeto)) = "ETbrc" Then Call MostrarErr("Inicio de Votacion Nominal " & Now, .sObjeto)
        End If
        '----------------------------------------------------------
        'Fin de VotacionNominal
        '----------------------------------------------------------
        If LCase(.sComponente) = "term.keyb" And LCase(.sAtributo) = "state" And LCase(.sValor) = "offvotnom" Then
            fDesconocido = False
            Call EnviarSktxCola(.sObjeto, "SFINVT")
            'If cNIVEL_LOG > 2 Then Call MostrarErr("Fin de Votacion Nominal " & Now, .sObjeto)
        End If
        '----------------------------------------------------------
        'Limpiar Votos
        '----------------------------------------------------------
        If (LCase(.sComponente) = "term.ledk1" Or LCase(.sComponente) = "term.ledk2") And LCase(.sAtributo) = "state" Then
            If LCase(.sValor) = "offvotnom" Then
                fDesconocido = False
                Call EnviarSktxCola(.sObjeto, "SLIMVT")
                'Call FormMain.EnviarxSkt("}", B2Skt(Val(.sObjeto)).Socket, "SLIMVT")
            ElseIf LCase(.sValor) = "off" Or LCase(.sValor) = "offvotnum" Then
                fDesconocido = False
                'Call EnviarSktxCola(.sObjeto, "SLIMVT")
                Call EnviarSktxCola(.sObjeto, "SLIMVT")
'                If .sObjeto = "brc" Then
'                    For naUX = 0 To ConexionesAbiertas - 1
'                        If FormMain.WSocket(naUX).State = 7 Then
'                            Call EnviarSktxCola(Str(Skt2B(naUX).Banca), "SLIMVT")
'                        End If
'                    Next
'                Else
'                    Call EnviarSktxCola(.sObjeto, "SLIMVT")
'                End If
                
                'Log_DEBUG (.sObjeto & "SLIMVT" & .sValor)
                'Call FormMain.EnviarxSkt("}", 256, "SLIMVT")
                'Call EnviarSktxCola(.sObjeto, "SLIMNU") ' PENDIENTE de implementacion en banca nec 110117
                'Call EnviarSktxCola(.sObjeto, "SCANCL")
            End If
            'Call GuardarCola
        End If
        '****************************************************************************************
        '----------------------------------------------------------
        ' Enviar Informacion al Display
        '----------------------------------------------------------        If LCase(.sComponente) = "term.keyb" Then
        If LCase(.sComponente) = "term.display" And LCase(.sAtributo) = "text" Then
            fDesconocido = False
            '091012: Sin control de cola
            If "SINFOR" = "No implementado" Then
                'No implementado en la banca HCDN 2011
                Call FormMain.EnviarxSkt("X", B2Skt(Val(.sObjeto)).Socket, "SINFOR " & Left(.sValor, 40))
            End If
            ' PARA AGREGAR CONTROL DE COLA:
            'Call EnviarSktxCola(.sObjeto, "SINFOR " & Left(.sValor, 40))
        End If
        '-----------------------------------------------------------------------
        ' Enviar Informacion Acknowledge de Confirmacion de Identificacion
        '-----------------------------------------------------------------------
        If LCase(.sComponente) = "term.led1" And LCase(.sAtributo) = "state" Then
            fDesconocido = False
            Dim Separados() As String
            If LCase(.sValor) = "on" Then
                'If "SACKID" = "NO IMPLEMENTADO" Then
'                    Call EnviarSktxCola(.sObjeto, "SCANCL")
'                    Call EnviarSktxCola(.sObjeto, "SIDRXH")
                    Call EnviarSktxCola(.sObjeto, "SACKID")
'                    Call FormMain.EnviarxSkt("f", Val(.sObjeto), "SCANCL")
'                    Call FormMain.EnviarxSkt("f", Val(.sObjeto), "SIDRXH")
'                    Call FormMain.EnviarxSkt("f", Val(.sObjeto), "SACKID")
                'End If
            ElseIf Left(LCase(.sValor), 9) = "on_manual" Then
                'Call EnviarSktxCola(.sObjeto, "SCANCL")
                'Call EnviarSktxCola(.sObjeto, "SAUTOD " & strString(10, Trim(Mid(LCase(.sValor), InStr(1, LCase(.sValor), "|") + 1, 10)), "0", "I")) 'Autentificacion directa: Identificacion manual por el operador.
                If .sObjeto <> "0" Then
                    Call FormMain.EnviarxSkt("f", Val(.sObjeto), "SCANCL")
                End If
                Call FormMain.EnviarxSkt("f", Val(.sObjeto), "SAUTOD " & strString(10, Trim(Mid(LCase(.sValor), InStr(1, LCase(.sValor), "|") + 1, 10)), "0", "I")) 'Autentificacion directa: Identificacion manual por el operador.
                Separados = Split(.sValor, "|")
                VectorSAUTOD(Val(.sObjeto)) = strString(16, Hex(Fix(Val(Trim(Separados(1))))), "0", "I")
            End If
        End If
        '----------------------------------------------------------
        ' Cancela el modo de Identificacion
        '----------------------------------------------------------
        If LCase(.sComponente) = "term.auth" And LCase(.sAtributo) = "action" And LCase(.sValor) = "auth_cancel" Then
            fDesconocido = False
            'If IsNumeric(.sObjeto) Then
            '    Call FormMain.EnviarxSkt("f", Int(.sObjeto), "STST03")
            'End If
            Call EnviarSktxCola(.sObjeto, "SCANCL")
        End If
        '-----------------------------------------------------------
        ' Preguntamos el STATUS a las bancas
        '-----------------------------------------------------------
        If .sTipo = "mget" And (LCase(.sComponente) = "term" Or LCase(.sComponente) = "term.mon") And LCase(.sAtributo) = "state" Then
            fDesconocido = False
            If LCase(.sObjeto) = "brc" Then
                For naUX = 0 To ConexionesAbiertas - 1
                    UltimoEstadodePresencia(naUX) = "NULO"
                    UltimoEstadodeBanca(naUX) = "NULO"
                    If Banca_ip(naUX).Estado = False Then
                        msgSQV.sTipo = "mevt"
                        msgSQV.sObjeto = Str(Skt2B(naUX).Banca)
                        msgSQV.sComponente = "term"
                        msgSQV.sComentario = "a02"
                        msgSQV.sAtributo = "state"
                        msgSQV.sValor = "off"
                        Call InsertarMsgSQV(msgSQV)
                        UltimoEstadodeBanca(naUX) = "off"
                    End If
                Next
            Else
               If B2Skt(Val(.sObjeto)).Estado = True Then
                    If Banca_ip(B2Skt(Val(.sObjeto)).Socket).Estado = False Then
                        msgSQV.sTipo = "mevt"
                        msgSQV.sObjeto = .sObjeto
                        msgSQV.sComponente = "term"
                        msgSQV.sAtributo = "state"
                        msgSQV.sComentario = "a03"
                        msgSQV.sValor = "off"
                        Call InsertarMsgSQV(msgSQV)
                        UltimoEstadodeBanca(B2Skt(Val(.sObjeto)).Socket) = "off"
                    End If
                Else
                        msgSQV.sTipo = "mevt"
                        msgSQV.sObjeto = .sObjeto
                        msgSQV.sComponente = "term"
                        msgSQV.sComentario = "a04"
                        msgSQV.sAtributo = "state"
                        msgSQV.sValor = "off"
                        UltimoEstadodeBanca(B2Skt(Val(.sObjeto)).Socket) = "off"
                        Call InsertarMsgSQV(msgSQV)
                End If
            End If
            If LCase(.sObjeto) = "brc" Then
                Call EnviarSktxCola(.sObjeto, "STATUS")
            End If
        End If
        '-----------------------------------------------------------
        ' Comando Para hacer Shutdown
        '-----------------------------------------------------------
        If LCase(.sComponente) = "sb" And LCase(.sAtributo) = "shutdown" And LCase(.sValor) = "now" Then
           fDesconocido = False
           Unload FormMain
        End If
        If fDesconocido = True Then
            'Comando Desconocido
            Call Log_DEBUG("Comando desconocido" & msgSQV.sTipo & " a banca" & msgSQV.sObjeto)
        End If
    End With
End Sub


' Seteamos Recordset Varios
'*************************************************************************
Public Sub SetearRsSQV(strCadena As String)
    If RsSQV.State = adStateOpen Then
        RsSQV.Close
        RsSQV.Source = strCadena
        RsSQV.Open
    Else
        RsSQV.CursorLocation = adUseClient
        RsSQV.Open strCadena, Cn, adOpenDynamic, adLockOptimistic
    End If
    If RsSQV.RecordCount > 0 Then
        RsSQV.MoveFirst
    End If
End Sub


Public Sub SetearRsSB(strCadena As String)
    If RsSB.State = adStateOpen Then
        RsSB.Close
        RsSB.Source = strCadena
        RsSB.Open
    Else
        RsSB.CursorLocation = adUseClient
        RsSB.Open strCadena, Cn, adOpenDynamic, adLockOptimistic
    End If
    If RsSB.RecordCount > 0 Then
        RsSB.MoveFirst
    End If
End Sub

Public Sub SetearRsBanca(strCadena As String)
    If RsBanca.State = adStateOpen Then
        RsBanca.Close
        RsBanca.Source = strCadena
        RsBanca.Open
    Else
        RsBanca.CursorLocation = adUseClient
        RsBanca.Open strCadena, Cn, adOpenDynamic, adLockOptimistic
    End If
    If RsBanca.RecordCount > 0 Then
        RsBanca.MoveFirst
    End If
End Sub
'*************************************************************************************


Public Sub CargarGrilla()
    Dim naUX    As Long
    Dim fEstado As String
    Dim nFil    As Integer
    
    'FormMain.Grilla.Rows = 2
    'FormMain.Grilla.ColWidth(0) = 900
    'FormMain.Grilla.ColWidth(1) = 1450
    'FormMain.Grilla.Rows = 100
    'FormMain.Grilla.Row = 0
    'FormMain.Grilla.Col = 0
    'FormMain.Grilla.Text = "Banca"
    'FormMain.Grilla.Col = 1
    'FormMain.Grilla.Text = "Estado"
    'FormMain.Grilla.Row = 1
    'FormMain.Grilla.Col = 0
    'FormMain.Grilla.Text = ""
    'FormMain.Grilla.Col = 1
    'FormMain.Grilla.Text = ""
    'nFil = 1
    'For naUX = 0 To ConexionesAbiertas - 1
    '    If FormMain.WSocket(naUX).State = 0 Then
    '        fEstado = " Desconectada"
    '        FormMain.Grilla.Row = nFil
    '        FormMain.Grilla.Col = 0
    '        FormMain.Grilla.Text = Str(Skt2B(naUX).Banca)
    '        FormMain.Grilla.Col = 1
    '        FormMain.Grilla.Text = fEstado
    '        nFil = nFil + 1
    '    End If
    'Next
   '
End Sub


'***************************************************************************************************


'Elimina la cola del SQV, con -1 elimina completa y sino desde el id que le paso
'*******************************************************************************

Public Sub EliminarMensajesSQV(fId As Long)
    On Error GoTo Trap_Error
    Dim strSql As String
    strSql = ""
    'Cn.Execute "sqv_InsertarHistoricoMensajesSB"
    If fId = -1 Then
        strSql = "TRUNCATE TABLE sqv_sb_mensajes"
    Else
        strSql = "delete from sqv_sb_mensajes WHERE id < " & Str(fId)
    End If
    If strSql <> "" Then
        Cn.Execute (strSql)
    End If
Exit Sub
Trap_Error:
    Select Case Err.Number
        Case Else
            MsgBox "Error Nº " & Err.Number & Chr(10) & Err.Description & "Originado en " & Err.Source
            Resume
    End Select
End Sub

'Inserta mensaje para el SQV
'*********************************************

Public Sub InsertarMsgSQV(fSQV As MensajeSQV)
    'On Error GoTo Trap_Error
    Static nTotalTicksInserts As Long
    Dim strInsert As String
    Dim nTick As Long
    strInsert = "insert_sb_sqv_mensajes('" & fSQV.sTipo & "','" & fSQV.sComponente & "','" & fSQV.sObjeto & "', '" & fSQV.sAtributo & "', '" & fSQV.sValor & "',  '" & Trim(Left(Trim(fSQV.sComentario), 10)) & "','" & Format(Now(), "YYYYMMDD") & "','" & Format(Now(), "HH:mm:ss") & "')"
    nTick = GetTickCount
    Cn.Execute (strInsert)
    nTotalTicksInserts = nTotalTicksInserts + GetTickCount - nTick
    'If cNIVEL_LOG > 2 And nTotalTicksInserts Mod 1000 < 50 Then Call MostrarErr(" > Insert " & Str(nTotalTicksInserts / 1000) & " seg.", 0) 'banana
    CantidadInsertados = CantidadInsertados + 1
    If CantidadInsertados > 50000 Then 'Seguridad extra para desbordamiento
        CantidadInsertados = 0
    End If
    FormMain.lblMensajesInsertados.Caption = CantidadInsertados
Exit Sub
Trap_Error:
    Select Case Err.Number
        Case Else
            MsgBox "InsertarMsqSQV " & vbCrLf & strInsert & vbCrLf & "Error Nº " & Err.Number & Chr(10) & Err.Description & "Originado en " & Err.Source
            Exit Sub
    End Select
End Sub

' Leo Mensajes del SQV
'*********************************************

Public Sub LeerMensajesSQV()
    Dim strSql As String
    Dim msgSQV As MensajeSQV
    
    strSql = "SELECT * FROM sqv_sb_mensajes WHERE id > " & Str(nUltimoMensajeSQV)
    Call SetearRsSQV(strSql)
    
    While Not RsSQV.EOF
        ' ---------------------------------------------------------------------------
        ' Levantar mensajes no leidos
        ' ---------------------------------------------------------------------------
        With msgSQV
            .sTipo = RsSQV.Fields("Tipo").Value
            .sObjeto = RsSQV.Fields("Objeto").Value
            .sComponente = RsSQV.Fields("Componente").Value
            .sAtributo = RsSQV.Fields("Atributo").Value
            .sValor = RsSQV.Fields("Valor").Value
            .sComentario = RsSQV.Fields("comentario").Value
        End With
        ' ---------------------------------------------------------------------------
        ' Aca se interpretan los mensajes recibidos
        ' ---------------------------------------------------------------------------
        Call InterpretaDatosSQV(msgSQV)
        
        ' ---------------------------------------------------------------------------
        ' Proximo mensaje
        ' ---------------------------------------------------------------------------
        nUltimoMensajeSQV = RsSQV.Fields("id").Value
        RsSQV.MoveNext
        DoEvents
    Wend
End Sub

'Funcion que muestra Algunos errores en la pantalla
'***************************************************

Public Sub MostrarErr(fMensaje As String, Optional fObjeto As String)
    FormMain.txtLog.Text = Left(Left(fMensaje, 60) & vbCrLf & FormMain.txtLog.Text, 2000)
    If fObjeto = "" Then fObjeto = "sb"
    'Call GuardarLog(fObjeto, "ERR - " & fMensaje)
End Sub


'Enviar Tabla completa de Huellas a los Legisladores.
'****************************************************

Public Sub EnviarHuellas(fObjeto As String, Optional fTipo As Integer = -1)
    Dim strcad                      As String
    Dim naUX                        As Long
    Dim nIndice                     As Long
    Dim nEnvio                      As Integer
    Dim strLegisladorSecuencia      As String
    Dim strMantenimientoSecuencia   As String
    Dim fLSec                       As Boolean
    'Dim bNuevoTemplate As New ADODB.Stream
    'Funcion para Enviar Tabla de Legisladores.
    '**************************************************************
    
    'Cadena de SQL
    'strcad = "SELECT * FROM Legisladores WHERE (es_legislador = 1) ORDER BY Tipo DESC, CAST(id AS int)"
    strcad = "SELECT * FROM Legisladores_SIN_USO WHERE (es_legislador = 1) ORDER BY Tipo DESC, CAST(id AS int)"
    Call SetearRsBanca(strcad)
    fLSec = True
    If Not RsBanca.EOF Then
        RsBanca.MoveFirst
        nIndice = 0
        nEnvio = 0
        'Primera Secuencia de Legislador
        strLegisladorSecuencia = strString(4, Hex(Fix(nEnvio)), "0", "I")
        Version = Version = Format(Day(Now), "00") & Format(Month(Now), "00") & Format(Hour(Now), "00") & Format(Minute(Now), "00")
        Do While Not RsBanca.EOF
                If Not (IsNull(RsBanca!Picture)) Then
                    'bNuevoTemplate.Type = adTypeBinary
                    'bNuevoTemplate.Open
                    'bNuevoTemplate.Write RsBanca!Picture
                    
                    'bNuevoTemplate.Type = adTypeBinary
                    
                    'bNuevoTemplate = RsBanca!template11 'VER
                    With sLegisla
                         .sId = RsBanca!id
                         .sNombre = NullCadena(RsBanca!nombre)
                         .sApellido = NullCadena(RsBanca!apellido)
                         .sTemplate11 = RsBanca!Picture
                         MsgBox BinAHex(.sTemplate11)
                         If Len(Trim(.sTemplate11)) > 0 Then
                            If Val(RsBanca!tipo) = 0 Then
                               .sNombre = .sNombre & " " & Now
                            End If
                            strLegisla(nIndice) = "SRLEGI ^"
                            strLegisla(nIndice) = strLegisla(nIndice) & BinAHex(.sTemplate11)
                            'strLegisla(nIndice) = strLegisla(nIndice) & TextoAHex(LongFija(.sNombre, 30), 60)
                            'MsgBox BinAHex(.sTemplate11)
                            'Call EnviarSktxCola(fObjeto, strLegisla(nIndice))
                            nIndice = nIndice + 1
                            'For naUX = 1 To 9
                            '    strLegisla(nIndice) = "SRLEGI ^"
                            '    strLegisla(nIndice) = strLegisla(nIndice) & strString(4, Hex(Fix(nEnvio)), "0", "I")
                            '    strLegisla(nIndice) = strLegisla(nIndice) & Format(naUX, "00")
                            '    strLegisla(nIndice) = strLegisla(nIndice) & Mid(.sTemplate, (naUX - 1) * 128 + 1, 128)
                                'Call EnviarSktxCola(fObjeto, strLegisla(nIndice))
                                'Guardo en la tabla el numero de envio para cada legislador
                            '    nIndice = nIndice + 1
                            'Next
                            'Call EnviarSktxCola(fObjeto, "SINFOR " & Chr(13) & "RX Huella: " & Str(nEnvio))
                            'RsBanca!indicebanca = nEnvio
                            nEnvio = nEnvio + 1
                         End If
                    End With
                nIndice = 0
                End If
            RsBanca.MoveNext
        Loop
        'Guardamos ultima Secuencia de Mantenimiento
        'strMantenimientoSecuencia = strMantenimientoSecuencia & strString(4, Hex(Fix(nEnvio + 1)), "0", "I")
        RsBanca.Close
        'Marcamos fin de Registro
        'strLegisla(nIndice) = "SRLEGI ^" & CerosIzquierda(Hex(Fix(nEnvio)), 4) & "00" & String(128, "F") & vbCrLf
        'Call EnviarSktxCola(fObjeto, strLegisla(nIndice))
        Call EnviarSktxCola(fObjeto, "SNUVER ^" & Version)
        Call EnviarSktxCola(fObjeto, "SCANCL")
        'Grabamos Secuencia enviadas
        strcad = "SELECT * FROM BANCASIP ORDER BY BANCANUMERO ASC"
        Call SetearRsBanca(strcad)
        RsBanca.MoveFirst
        Do While Not RsBanca.EOF
            RsBanca!secuencialegislador = strLegisladorSecuencia
            RsBanca!secuenciamantenimiento = strMantenimientoSecuencia
            RsBanca!Version = Version
            RsBanca.MoveNext
        Loop
    End If
End Sub
Public Function SetearRsW(pCadena As String, ByRef pRst As ADODB.Recordset) As Boolean
    On Error GoTo TrapError
    
    SetearRsW = False
    
    With pRst
        If .State = adStateOpen Then
            .Close
        End If
        .Source = pCadena
        .ActiveConnection = Cn
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open
        DoEvents
        If Not .BOF And Not .EOF Then
            SetearRsW = True
        End If
    End With
Exit Function
TrapError:
    Select Case Err.Number
        Case Else
            MsgBox "Error N° " & Err.Number & Chr(10) & Err.Description & "Originado en " & Err.Source
            Resume
    End Select
Return

End Function

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
             trae_huellas = Replace(rsHuellasLegislador.Fields("huella").Value, " ", "")
    End Select
   
    Else
        trae_huellas = -1 'error no encontro registros
    End If
    Set rsHuellasLegislador = Nothing
End Function


'Enviar Tabla huellas completa de Huellas a los Legisladores.
'****************************************************
Public Sub EnviarMultiHuellas(fObjeto As String, Optional fTipo As Integer = -1)
    Dim strcad                      As String
    Dim naUX                        As Long
    Dim nIndice                     As Long
    Dim nEnvio                      As Integer
    Dim strLegisladorSecuencia      As String
    Dim strMantenimientoSecuencia   As String
    Dim fLSec                       As Boolean
    Dim strCadenaHuella             As String
    Dim nHuella                     As Integer
    Dim nCantHuellas                As Integer
    Dim xHuella                     As String
    Dim xUltimaHuella               As String
    Dim i As Integer
    Dim lFin As Boolean
    Dim strSQLWhere                 As String
    Dim nCantLegisladoresConHuella  As Integer
    Dim strVersionDatosSQV          As String
    Dim nEspera As Long
    Dim strPrefijoVersionBanca  As String
    'Funcion para Enviar Tabla de Legisladores.
    '**************************************************************
    'Busca en config el valor de version de datos sqv
    Set RsBanca = New ADODB.Recordset
    'Leo la base viegente de Uso y cambio la conexion
    strcad = "SELECT TOP 1 version_datos_sqv FROM config"
    SetearRsBanca (strcad)
    strVersionDatosSQV = RsBanca.Fields(0).Value
    strPrefijoVersionBanca = Replace(RsBanca.Fields(0).Value, "_", "")
    strPrefijoVersionBanca = Left(strPrefijoVersionBanca, 4) & Mid(strPrefijoVersionBanca, 9, 4)
    'Marco "en proceso"
    Set RsBanca = New ADODB.Recordset
    strcad = "SELECT * FROM BANCASIP " & IIf(LCase(fObjeto) = "brc", "", "WHERE BancaNumero " & IIf(InStr(fObjeto, ";"), "in (" & Trim(fObjeto) & ")", " =(" & Trim(fObjeto) & ")")) & " ORDER BY BANCANUMERO ASC"
    Call SetearRsBanca(strcad)
    RsBanca.MoveFirst
    Do While Not RsBanca.EOF
        RsBanca!version_datos_banca = "ERROR_PROCESANDO"
        RsBanca.MoveNext
    Loop
    
    
    strSQLWhere = "Legisladores.es_legislador >= 0" 'Envia todas las huellas, legisladores y mantenimiento.
    'strSQLWhere = "es_legislador = 0" 'Envia solo las huellas de mantenimiento.
    'strSQLWhere = "es_legislador = 1" 'Envia solo las huellas de legisladores .
    strcad = "SELECT     COUNT(*) AS cantidad FROM         (SELECT     idlegislador FROM          huellas INNER JOIN Legisladores ON huellas.idlegislador = Legisladores.id Where (" & strSQLWhere & " ) GROUP BY huellas.idlegislador) DERIVEDTBL"
    Call SetearRsBanca(strcad)
    If Not RsBanca.EOF Then ' si se encuentra al menos un legislador
        nCantLegisladoresConHuella = RsBanca!cantidad
    End If
    RsBanca.Close
    If nCantLegisladoresConHuella > 0 Then
        'SELECT     huellas.idlegislador FROM         huellas INNER JOIN                       Legisladores ON huellas.idlegislador = Legisladores.id Where (Legisladores.es_legislador = 1) GROUP BY huellas.idlegislador
        'Cadena de SQL
        strcad = "SELECT * FROM Legisladores WHERE (" & strSQLWhere & " ) ORDER BY Tipo DESC, CAST(id AS int)"
        Call SetearRsBanca(strcad)
        fLSec = True
        If Not RsBanca.EOF Then ' si se encuentra al menos un legislador
            RsBanca.MoveFirst
            nIndice = 0
            nEnvio = 0
            'Primera Secuencia de Legislador
            strLegisladorSecuencia = strString(4, Hex(Fix(nEnvio)), "0", "I")
            'Version = Format(Day(Now), "00") & Format(Month(Now), "00") & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(nCantLegisladoresConHuella, "0000") 'cantidad de personas a enviar en decimal de 4 bytes
            Version = strPrefijoVersionBanca & Format(nCantLegisladoresConHuella, "0000") 'cantidad de personas a enviar en decimal de 4 bytes
            Call EnviarSktxCola(fObjeto, "SNUVER " & Version)
            Do While Not RsBanca.EOF
                    With sLegisla
                         .sId = RsBanca!id
                         .sNombre = NullCadena(RsBanca!nombre)
                         .sApellido = NullCadena(RsBanca!apellido)
                         .sDNI = NullCadena(RsBanca!dni)
                         .sClase = IIf(RsBanca!es_legislador = 1, "S", "V") ' Considerar el caso de mantenimiento
                         .sIcono = "0"
                         nCantHuellas = trae_huellas(RsBanca!id, "C", 0)
                         If nCantHuellas > 0 Then
                            If Val(RsBanca!tipo) = 0 Then
                               If fLSec = True Then
                                    'Ultima Secuencia de Legislador
                                    If nEnvio = 0 Then
                                          'Si no hay Legisladores para enviar.
                                          strLegisladorSecuencia = strLegisladorSecuencia & "0001"
                                    Else
                                          strLegisladorSecuencia = strLegisladorSecuencia & strString(4, Hex(Fix(nEnvio)), "0", "I")
                                    End If
                                    'Primera Secuencia de Mantenimiento
                                    strMantenimientoSecuencia = strString(4, Hex(Fix(nEnvio)), "0", "I")
                                    fLSec = False
                               End If
                               '.sNombre = .sNombre & " " & Now
                            End If
    
                            nHuella = 0
                            i = 0
                            lFin = False
                            Do While Not lFin
                                i = i + 1 ' arranca de 1
                                If i <= 10 Then
                                    xHuella = trae_huellas(RsBanca!id, "H", i)
                                    
                                    If Len(Trim(xHuella)) = 2048 Then 'si obtiene un dato valido en longitud
                                        If nHuella + 1 <= nCantHuellas - 1 Then 'la ultima huella no la manda por aqui, sino con los datos del legislador
                                            strCadenaHuella = "SRLEGI 0" & strString(4, Hex(Fix((nIndice * 10) + nHuella)), "0", "I") & LongFija(xHuella, 2048)
                                            Call EnviarSktxCola(fObjeto, strCadenaHuella)
                                            'strCadenaHuella = ""
                                            nHuella = nHuella + 1
                                        Else
                                            lFin = True
                                        End If
                                        xUltimaHuella = xHuella
                                        'If nHuella >= nCantHuellas Then lFin = True
                                    Else
                                        If xHuella <> "-1" Then 'la huella es invalida, entonces tengo menos huellas para enviar
                                            Call GuardarLog(fObjeto, "huella invalida: " & strString(4, Hex(Fix((nIndice * 10) + nHuella)), "0", "I") & xHuella)
                                            nCantHuellas = nCantHuellas - 1
                                            If nHuella > nCantHuellas - 1 Then 'fin
                                                lFin = True
                                            End If
                                        End If
                                    End If
                                Else
                                    lFin = True
                                End If
                            Loop 'fin envio huellas 1 a (ncanthuellas - 1)
                                                   
                            'Envio de datos del legislador con la ultima huella
                            strCadenaHuella = "SRLEGI 1"
                            strCadenaHuella = strCadenaHuella & strString(4, Hex(Fix((nIndice * 10) + nHuella)), "0", "I") 'id
                            strCadenaHuella = strCadenaHuella & LongFija(.sApellido, 44)
                            strCadenaHuella = strCadenaHuella & LongFija(.sNombre, 44)
                            strCadenaHuella = strCadenaHuella & LongFija(Format(.sId, "0000000000"), 10)
                            strCadenaHuella = strCadenaHuella & LongFija(.sClase, 1)
                            strCadenaHuella = strCadenaHuella & LongFija(.sIcono, 1)
                            strCadenaHuella = strCadenaHuella & LongFija(xUltimaHuella, 2048)
                            Call EnviarSktxCola(fObjeto, strCadenaHuella)
                            'incrementa numero de legislador enviado
                            nIndice = nIndice + 1
                            
                            'Call EnviarSktxCola(fObjeto, "SINFOR " & Chr(13) & "RX Huella: " & Str(nEnvio))
                            RsBanca!indicebanca = nEnvio
                            nEnvio = nEnvio + 1
                         End If
                    End With
                'nIndice = 0
                strCadenaHuella = ""
                RsBanca.MoveNext
                'For nEspera = 1 To 100000
                '    DoEvents
                'Next
            Loop
            'Guardamos ultima Secuencia de Mantenimiento
            strMantenimientoSecuencia = strMantenimientoSecuencia & strString(4, Hex(Fix(nEnvio + 1)), "0", "I")
            RsBanca.Close
            'Marcamos fin de Registro
            ' strLegisla(nIndice) = "SRLEGI ^" & CerosIzquierda(Hex(Fix(nEnvio)), 4) & "00" & String(128, "F") & vbCrLf
            'Call EnviarSktxCola(fObjeto, strLegisla(nIndice))
            'Call EnviarSktxCola(fObjeto, "SNUVER ^" & Version)
            'Call EnviarSktxCola(fObjeto, "SCANCL")
            'Grabamos Secuencia enviadas
            Set RsBanca = New ADODB.Recordset
            strcad = "SELECT * FROM BANCASIP " & IIf(LCase(fObjeto) = "brc", "", "WHERE BancaNumero " & IIf(InStr(fObjeto, ";"), "in (" & Trim(fObjeto) & ")", " =(" & Trim(fObjeto) & ")")) & " ORDER BY BANCANUMERO ASC"
            Call SetearRsBanca(strcad)
            RsBanca.MoveFirst
            Do While Not RsBanca.EOF
                RsBanca!secuencialegislador = strLegisladorSecuencia
                RsBanca!secuenciamantenimiento = strMantenimientoSecuencia
                RsBanca!Version = IIf(RsBanca!BancaNumero = 0, "No aplicable", Version)
                RsBanca!version_datos_sqv = strVersionDatosSQV  'aqui
                'RsBanca!version_datos_banca = Version
                RsBanca.MoveNext
            Loop
            
            If fObjeto = "brc" Then
                For naUX = 0 To ConexionesAbiertas - 1
                    Banca_ip(naUX).tVersion = Version
                Next naUX
            Else
                Banca_ip(B2Skt(Val(fObjeto)).Socket).tVersion = Version
            End If
            'sleveraqui TVERREADY
            'Call EnviarSktxCola(fObjeto, "SLEVER")
            'For naUX = 0 To ConexionesAbiertas - 1
            '    Call FormMain.EnviarxSkt("X", B2Skt(naUX).Socket, "SLEVER")
            'Call MostrarErr("SLEVER", fObjeto)
            'Next
        End If
    Else
        Call MostrarErr("No hay huellas para enviar en la base de datos")
    End If
End Sub

Public Sub EnviarHuellasHCDN(fObjeto As String, Optional fTipo As Integer = -1)
    Dim strcad                      As String
    Dim naUX                        As Long
    Dim nIndice                     As Long
    Dim nEnvio                      As Integer
    Dim strLegisladorSecuencia      As String
    Dim strMantenimientoSecuencia   As String
    Dim fLSec                       As Boolean
    Dim strCadenaHuella             As String
    Dim nHuella                     As Integer
    Dim nCantHuellas                As Integer
    Dim xHuella                     As String
    Dim xUltimaHuella               As String
    Dim i As Integer
    Dim lFin As Boolean
    Dim strSQLWhere                 As String
    Dim nCantLegisladoresConHuella  As Integer
    Dim strVersionDatosSQV          As String
    Dim nEspera As Long
    Dim strPrefijoVersionBanca  As String
    Dim nProcesados As Integer
    Dim nEnviarMaximoHuellas As Integer
    Dim nTicks As Long
    EnviandoSNUVER = False
    RespuestasSNUVER = 0
    Do While RsCola.RecordCount > 3000
        DoEvents
    Loop
    nEnviarMaximoHuellas = 0 ' 0 = SIN LIMITE
    FormMain.txtEnviando.Text = ""
    'Funcion para Enviar Tabla de Legisladores.
    '**************************************************************
    'Busca en config el valor de version de datos sqv
    Set RsBanca = New ADODB.Recordset
    'Leo la base viegente de Uso y cambio la conexion
    strcad = "SELECT TOP 1 version_datos_sqv FROM config"
    SetearRsBanca (strcad)
    strVersionDatosSQV = RsBanca.Fields(0).Value
    strPrefijoVersionBanca = Replace(RsBanca.Fields(0).Value, "_", "")
    strPrefijoVersionBanca = Left(strPrefijoVersionBanca, 4) & Mid(strPrefijoVersionBanca, 9, 4)
    'Marco "en proceso"
    Set RsBanca = New ADODB.Recordset
    strcad = "SELECT * FROM BANCASIP " & IIf(LCase(fObjeto) = "brc", "", IIf(InStr(fObjeto, ";"), "WHERE substring('" & Trim(fObjeto) & "' ,bancanumero*2+1,1)='1'", "WHERE BancaNumero =(" & Trim(fObjeto) & ")")) & " ORDER BY BANCANUMERO ASC"
    Call SetearRsBanca(strcad)
    RsBanca.MoveFirst
    Do While Not RsBanca.EOF
        RsBanca!version_datos_banca = "ERROR_PROCESANDO"
        RsBanca.MoveNext
    Loop
    MostrarErr ("Se enviaran huellas a " & " cantidad " & RsBanca.RecordCount & " bancas: Patron: " & fObjeto)
    'strcad = "SELECT     COUNT(*) AS cantidad FROM         (SELECT     idlegislador FROM          huellas INNER JOIN Legisladores ON huellas.idlegislador = Legisladores.id Where (" & strSQLWhere & " ) GROUP BY huellas.idlegislador) DERIVEDTBL"
    'strcad = "SELECT COUNT(*) AS cantidad FROM legisladores_sb JOIN legisladores_activos ON legisladores_activos.id = legisladores_sb.id WHERE template IS NOT NULL"
    strcad = "SELECT     COUNT(*) " & _
" FROM         (SELECT     a.id" & _
"                       FROM          legisladores_sb a INNER JOIN" & _
"                                              legisladores_activos ON legisladores_activos.id = a.id" & _
"                       Where a.tipo = 1  AND (legisladores_activos.descripcion <> 'Activo sin incorporar')" & _
"                       Union" & _
"                       SELECT     b.id" & _
"                       FROM         legisladores_sb b" & _
"                       WHERE     b.tipo = 0) U INNER JOIN" & _
"                      legisladores_sb sb ON U.id = sb.id"
    Call SetearRsBanca(strcad)
    If Not RsBanca.EOF Then ' si se encuentra al menos un legislador
        nCantLegisladoresConHuella = RsBanca.Fields(0)
        If nEnviarMaximoHuellas > 0 And nCantLegisladoresConHuella > nEnviarMaximoHuellas Then nCantLegisladoresConHuella = nEnviarMaximoHuellas
    End If
    RsBanca.Close
    If True Then
        'SELECT     huellas.idlegislador FROM         huellas INNER JOIN                       Legisladores ON huellas.idlegislador = Legisladores.id Where (Legisladores.es_legislador = 1) GROUP BY huellas.idlegislador
        'Cadena de SQL
        strSQLWhere = "es_legislador >= 0" 'Envia todas las huellas, legisladores y mantenimiento.
        'strSQLWhere = "es_legislador = 0" 'Envia solo las huellas de mantenimiento.
        'strSQLWhere = "es_legislador = 1" 'Envia solo las huellas de legisladores .
        'strcad = "SELECT * FROM Legisladores_SIN_USO WHERE (" & strSQLWhere & " ) ORDER BY Tipo DESC, CAST(id AS int)"
        strcad = "SELECT     sb.*" & _
" FROM         (SELECT     a.id" & _
"                       FROM          legisladores_sb a INNER JOIN" & _
"                                              legisladores_activos ON legisladores_activos.id = a.id" & _
"                       Where a.tipo = 1 AND (legisladores_activos.descripcion <> 'Activo sin incorporar')" & _
"                       Union" & _
"                       SELECT     b.id" & _
"                       FROM         legisladores_sb b" & _
"                       WHERE     b.tipo = 0) U INNER JOIN" & _
"                      legisladores_sb sb ON U.id = sb.id" & _
" ORDER BY sb.tipo DESC,U.id"
        Call SetearRsBanca(strcad)
        fLSec = True
        If Not RsBanca.EOF Then ' si se encuentra al menos un legislador
            RsBanca.MoveFirst
            nIndice = 0
            nEnvio = 0
            strLegisladorSecuencia = strString(4, Hex(Fix(nEnvio)), "0", "I")
            'Primera Secuencia de Legislador
            'Version = Format(Day(Now), "00") & Format(Month(Now), "00") & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(nCantLegisladoresConHuella, "0000") 'cantidad de personas a enviar en decimal de 4 bytes
            Version = strPrefijoVersionBanca & Format(nCantLegisladoresConHuella, "0000") 'cantidad de personas a enviar en decimal de 4 bytes
            Dim lTick As Long
            lTick = GetTickCount
            While RsCola.RecordCount > 100
                If GetTickCount - lTick > 3000 Then
                    For i = 0 To 256
                        Call FormMain.EnviarxSkt("f", i, "STATUS")
                        lTick = GetTickCount
                    Next i
                End If
                Call EnviarDatosSkt
                DoEvents
            Wend
            Call EnviarSktxCola(fObjeto, "SNUVER " & Version)
            If InStr(fObjeto, ";") > 0 Then
                EnviandoSNUVER = True
                While RespuestasSNUVER < 220
                    FormMain.txtLog.Text = "SNUVERs Aceptados: " & RespuestasSNUVER
                    DoEvents
                Wend
                EnviandoSNUVER = False
                RespuestasSNUVER = 0
                EnviandoSNUVER = True
                Call EnviarSktxCola(fObjeto, "SNUVER " & Version)
                While RespuestasSNUVER < 220
                    FormMain.txtLog.Text = "SNUVERs(2) Aceptados: " & RespuestasSNUVER
                    DoEvents
                Wend
                EnviandoSNUVER = False
            Else
                Dim FDelay As Long
                FDelay = GetTickCount
                While GetTickCount - FDelay < 2000
                    FormMain.txtLog.Text = "Borrando huellas(1)..."
                    DoEvents
                Wend
                FDelay = GetTickCount
                While GetTickCount - FDelay < 2000
                    FormMain.txtLog.Text = "Borrando huellas(2)..."
                    DoEvents
                Wend
            End If
            lTick = GetTickCount
            Call EstablecerEnviandoHuellas(fObjeto)
            nProcesados = 0
            FormMain.prgBAR.Value = 0
            FormMain.prgBAR.Max = RsBanca.RecordCount
            Do While Not RsBanca.EOF
                    With sLegisla
                        If Not IsNull(RsBanca!template) Then
                         .sTemplate11 = RsBanca!template
                         .sId = RsBanca!id
                         .sNombre = NullCadena(RsBanca!nombre)
                         .sApellido = NullCadena(RsBanca!apellido)
                         .sDNI = NullCadena(RsBanca!dni)
                         .sClase = IIf(RsBanca!es_legislador = 1, "S", "V") ' Considerar el caso de mantenimiento
                         .sIcono = "0"
                         nCantHuellas = 1
                         If nCantHuellas > 0 Then
                            If Val(RsBanca!tipo) = 0 Then
                               If fLSec = True Then
                                    'Ultima Secuencia de Legislador
                                    If nEnvio = 0 Then
                                          'Si no hay Legisladores para enviar.
                                          strLegisladorSecuencia = strLegisladorSecuencia & "0001"
                                    Else
                                          strLegisladorSecuencia = strLegisladorSecuencia & strString(4, Hex(Fix(nEnvio)), "0", "I")
                                    End If
                                    'Primera Secuencia de Mantenimiento
                                    strMantenimientoSecuencia = strString(4, Hex(Fix(nEnvio)), "0", "I")
                                    fLSec = False
                               End If
                               '.sNombre = .sNombre & " " & Now
                            End If
    
                            nHuella = 0
                            i = 0
                            lFin = False
                            'Envio de datos del legislador con la ultima huella
                            xUltimaHuella = BinAHex(.sTemplate11)
                        
                            strCadenaHuella = "SRLEGI "
                            strCadenaHuella = strCadenaHuella & xUltimaHuella
                            FormMain.txtTemp.Text = strCadenaHuella
                            
                            'MsgBox strCadenaHuella
                            Call EnviarSktxCola(fObjeto, strCadenaHuella) 'banana2
                            DoEvents
                            'incrementa numero de legislador enviado
                            nIndice = nIndice + 1
                            
                            'Call EnviarSktxCola(fObjeto, "SINFOR " & Chr(13) & "RX Huella: " & Str(nEnvio))
                            'RsBanca!indicebanca = nEnvio
                            nEnvio = nEnvio + 1
                            If Len(fObjeto) > 1 Then
                                While RsCola.RecordCount > 100
                                    Call EnviarDatosSkt
                                    DoEvents
                                Wend
                            Else
                                While RsCola.RecordCount > 0
                                    Call EnviarDatosSkt
                                    DoEvents
                                Wend
                            End If
                            FormMain.prgBAR.Value = FormMain.prgBAR.Value + 1
                         End If
                      End If
                    End With
                'nIndice = 0
                strCadenaHuella = ""
                If Not RsBanca.EOF Then
                    RsBanca.MoveNext
                Else
                    Exit Do
                End If
'                If Not RsBanca.EOF Then
'                    RsBanca.MoveNext
'                    'nTicks = GetTickCount
'
'                    Do While RsCola.RecordCount > 1000
'                        DoEvents
'                    Loop
'
'                    'Do While GetTickCount - nTicks < 100
'                    '    For nEspera = 1 To 50 '100
'                    '        DoEvents
'                    '    Next
'                    'Loop
'                    nProcesados = nProcesados + 1
'                    Call MostrarErr("Envio de huellas a cola, procesadas: " & nProcesados & " Mensajes en cola ultima consulta:" & RsCola.RecordCount, 0)
'                    FormMain.lblRegistrosEnCola.Caption = RsCola.RecordCount
'                    If nEnviarMaximoHuellas > 0 And nProcesados > nEnviarMaximoHuellas Then
'                        Exit Do
'                    End If
'                Else
'                    Call MostrarErr("Fin de legisladores no esperado: Procesados " & nProcesados)
'                    Exit Do
'                End If
            Loop
            Call MostrarErr("Fin envio de huellas a cola, procesadas: " & nProcesados & " Mensajes en cola ultima consulta:" & RsCola.RecordCount, 0)
            FormMain.prgBAR.Value = 0
            'Guardamos ultima Secuencia de Mantenimiento
            strMantenimientoSecuencia = strMantenimientoSecuencia & strString(4, Hex(Fix(nEnvio + 1)), "0", "I")
            RsBanca.Close
            'Marcamos fin de Registro
            ' strLegisla(nIndice) = "SRLEGI ^" & CerosIzquierda(Hex(Fix(nEnvio)), 4) & "00" & String(128, "F") & vbCrLf
            'Call EnviarSktxCola(fObjeto, strLegisla(nIndice))
            'Call EnviarSktxCola(fObjeto, "SNUVER ^" & Version)
            'Call EnviarSktxCola(fObjeto, "SCANCL")
            'Grabamos Secuencia enviadas
            Set RsBanca = New ADODB.Recordset
            'strcad = "SELECT * FROM BANCASIP " & IIf(LCase(fObjeto) = "brc", "", "WHERE BancaNumero " & IIf(InStr(fObjeto, ";"), "in (" & Trim(fObjeto) & ")", " =(" & Trim(fObjeto) & ")")) & " ORDER BY BANCANUMERO ASC"
            strcad = "SELECT * FROM BANCASIP " & IIf(LCase(fObjeto) = "brc", "", IIf(InStr(fObjeto, ";"), "WHERE substring('" & Trim(fObjeto) & "' ,bancanumero*2+1,1)='1'", "WHERE BancaNumero =(" & Trim(fObjeto) & ")")) & " ORDER BY BANCANUMERO ASC"
            Call SetearRsBanca(strcad)
            RsBanca.MoveFirst
            Do While Not RsBanca.EOF
                RsBanca!secuencialegislador = strLegisladorSecuencia
                RsBanca!secuenciamantenimiento = strMantenimientoSecuencia
                RsBanca!Version = IIf(RsBanca!BancaNumero = 0, "No aplicable", Version)
                RsBanca!version_datos_sqv = strVersionDatosSQV  'aqui
                'RsBanca!version_datos_banca = Version
                RsBanca.MoveNext
            Loop
            
            If fObjeto = "brc" Then
                For naUX = 0 To ConexionesAbiertas - 1
                    Banca_ip(naUX).tVersion = Version
                Next naUX
            Else '0;1;1;0  bancas 0 1 2 3
                If InStr(fObjeto, ";") Then
                    For i = 1 To Len(Trim(fObjeto)) Step 2
                        If Mid(fObjeto, i, 1) = "1" Then
                            'MsgBox ("antes" & Banca_ip(B2Skt((i - 1) / 2).Socket).tVersion)
                            Banca_ip(B2Skt((i - 1) / 2).Socket).tVersion = Version
                            'MsgBox ("despues" & Banca_ip(B2Skt((i - 1) / 2).Socket).tVersion)
                        End If
                    Next i
                Else
                    Banca_ip(B2Skt(Val(fObjeto)).Socket).tVersion = Version
                End If
            End If
            'sleveraqui TVERREADY
            'Call EnviarSktxCola(fObjeto, "SLEVER")
            'For naUX = 0 To ConexionesAbiertas - 1
            '    Call FormMain.EnviarxSkt("X", B2Skt(naUX).Socket, "SLEVER")
            'Call MostrarErr("SLEVER", fObjeto)
            'Next
        End If
    Else
        Call MostrarErr("No hay huellas para enviar en la base de datos")
    End If
End Sub


Public Sub EnviarPresencia(fSocket As Integer, Optional fEstado As Boolean = True)
'NO SE USA?
        Call EnviarSktxCola(Str(fSocket), "SCANCL")
    If fEstado = True Then
        Call EnviarSktxCola(Str(fSocket), "SCONFG ^1E1801003C03")
    Else
        Call EnviarSktxCola(Str(fSocket), "SCONFG ^1E1800003C03")
    End If
        Call EnviarSktxCola(Str(fSocket), "SCANCL")
        Call EnviarSktxCola(Str(fSocket), "STATUS")
End Sub

Public Sub GuardarLog(fSocket1 As String, Optional fmensajeSkt1 As String, Optional fmensajeSqv1 As String)
    Dim nBanca         As String
    Dim sOrigen        As String
    Dim sDescripcion   As String
    Dim sCmdLog   As String
    
    nBanca = " "
    sOrigen = " "
    sDescripcion = " "
    sCmdLog = " "
    
    'Exit Sub
    If Len(fmensajeSkt1) = 0 Then
        sOrigen = "Servidor SQV"
        sDescripcion = fmensajeSqv1
    Else
        sOrigen = "Terminal de Banca"
        sDescripcion = fmensajeSkt1
    End If
    If fSocket1 = "brc" Then
        nBanca = "brc"
    Else
        If fSocket1 = "sb" Then
            nBanca = "sb"
        Else
            If InStr(fSocket1, ";") > 0 Then
                nBanca = fSocket1
            Else
                'Convierto socket en Banca.
                nBanca = Str(Skt2B(fSocket1).Banca)
            End If
        End If
    End If
End Sub
Private Sub Log_Texto(sOrigen As String, nBanca As String, Descripcion As String, Comando As String)
'Open App.Path & "\log\log.txt" For Append As #1
'Print #1, Now & "      -    " & sOrigen & "|" & nBanca & "|" & Descripcion & "|" & Comando
'Close #1
End Sub
Public Sub Log_DEBUG(texto As String)
'Dim pFormateado As String
'If (GetTickCount - TICK_LOG) > 3600000 Then 'Si paso una hora
'    Prefijo_Tick = Now() 'Actualizo el prefijo del log
'    TICK_LOG = GetTickCount
'End If
'pFormateado = Replace(Prefijo_Tick, ":", ".")
'pFormateado = Replace(pFormateado, "/", "_")
'Open App.Path & "\log\logDEBUG_" & pFormateado & ".txt" For Append As #1
'Print #1, Now & "      -    " & texto
'Close #1
End Sub
Public Sub Log_Banca(archivo As String, texto As String)
'Open App.Path & "\log\" & archivo For Append As #1
'Print #1, texto
'Close #1
End Sub
Public Sub Log_Particular(archivo As String, texto As String)
'Open App.Path & "\log\" & archivo For Output As #1
'Print #1, texto
'Close #1
End Sub
Public Function SacaNulos(sOrigen As String)
Dim i As Integer
Dim sSinNulos As String
Dim sCaracter As String
Dim n As Integer
Dim sImprimible As String
n = 0
For i = 1 To Len(Trim(sOrigen))
    sCaracter = Mid(Trim(sOrigen), i, 1)
    If Asc(sCaracter) <> 0 Then ' valido
        sSinNulos = Mid(sSinNulos, 1, n) & sCaracter
        n = n + 1
    ElseIf True Then
        sImprimible = ""
        sSinNulos = Mid(sSinNulos, 1, n) & Trim(sImprimible)
        n = n + Len(Trim(sImprimible))
    End If
Next i
SacaNulos = sSinNulos
End Function

Public Function SacaInvalidos(sOrigen As String)
Dim i As Integer
Dim sSinNulos As String
Dim sCaracter As String
Dim n As Integer
Dim sImprimible As String
n = 0
For i = 1 To Len(Trim(sOrigen))
    sCaracter = Mid(Trim(sOrigen), i, 1)
    If Asc(sCaracter) >= 32 Then
        sSinNulos = Mid(sSinNulos, 1, n) & sCaracter
        n = n + 1
    ElseIf True Then
        sImprimible = "\" & Trim(Str(Asc(sCaracter)))
        sSinNulos = Mid(sSinNulos, 1, n) & Trim(sImprimible)
        n = n + Len(Trim(sImprimible))
    End If
Next i
SacaInvalidos = sSinNulos
End Function





' Fin Funciones principales.
'*****************************************************************************************


'Funciones Adicionales para tratado de conversiones
'*****************************************************************************************

'Funcion que elimina problema con string nulos
Public Function NullCadena(Optional strCadena As Variant) As String
    NullCadena = strCadena & ""
End Function

'*****************************************************************************************


' Funciones de Alejandro
'*****************************************************************************************
Public Function HexATexto(strTexto As String, nLong As Long) As String
Dim i As Long
    'convierte un texto que contenga pares hexadecimales codificados como string en un string ascii
    HexATexto = ""
    For i = 1 To Len(strTexto) Step 2
        HexATexto = HexATexto & HexAChr(Mid(strTexto, i, 2))
    Next
    HexATexto = CerosIzquierda(HexATexto, nLong)
End Function

Public Function HexAChr(strHex) As String
    'convierte dos digitos hexadecimales codificados como string en un caracter ascii
    Dim nDecimal As Long
    nDecimal = 0
    nDecimal = DigitoHexADec(Mid(strHex, 1, 1)) * 16
    nDecimal = nDecimal + DigitoHexADec(Mid(strHex, 2, 1))
    HexAChr = Chr(nDecimal)
End Function

Public Function DigitoHexADec(charHex) As Long
    If charHex >= "0" And charHex <= "9" Then
        DigitoHexADec = Asc(charHex) - 48
    Else
        DigitoHexADec = Asc(charHex) - 55
    End If
End Function

Public Function LongFija(strText As String, nLong As Long) As String
    If nLong > Len(strText) Then
        LongFija = Left(strText & Space(nLong - Len(strText)), nLong)
    Else
        LongFija = Left(strText, nLong)
    End If
End Function

Public Function TextoAHex(strTexto As String, nLong As Long) As String
Dim i As Long
    TextoAHex = ""
    For i = 1 To Len(strTexto)
        TextoAHex = TextoAHex & Hex(Asc(Mid(strTexto, i, 1)))
    Next
    TextoAHex = CerosIzquierda(TextoAHex, nLong)
End Function

Public Function BinAHex(dataBin() As Byte) As String
Dim i As Long
    BinAHex = ""
    For i = LBound(dataBin) To (UBound(dataBin))
        BinAHex = BinAHex & CerosIzquierda(Hex(dataBin(i)), 2)
    Next
    'TextoAHex = CerosIzquierda(TextoAHex, nLong)
End Function

Public Function CerosIzquierda(strText As String, nLong As Long) As String
    If nLong > Len(strText) Then
        CerosIzquierda = Left(String(nLong - Len(strText), "0") & strText, nLong)
    Else
        CerosIzquierda = Right(strText, nLong)
    End If
End Function

Public Function strString(xTam As Long, strValor As String, strRelleno As String, Optional strTipo As String = "D") As String
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
Private Function EstaIdentificado(Banca As Integer) As Boolean
Dim rsTemp As New ADODB.Recordset
Dim Bancas() As String
Set rsTemp = New ADODB.Recordset
SetearRsW "SELECT vector_identificacion FROM vector", rsTemp
If Not rsTemp.EOF Then
    Bancas = Split(rsTemp.Fields(0), ";")
    If Bancas(Banca) = "0" Then
        EstaIdentificado = False
    Else
        EstaIdentificado = True
    End If
End If
rsTemp.Close
Set rsTemp = Nothing
End Function


Public Function setVersion()
    setVersion = "5.02a 110202"
End Function
'**********************************************************************************
