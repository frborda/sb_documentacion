Attribute VB_Name = "ManejodeCola"
'***********************************************************************
' Modulo Solo para el manejo de la cola de Mensajes
'
'
'***********************************************************************

Private Const TimeOutMensaje          As Long = 6000   ' En milisegundos
Private Const TimeOutReintentos       As Long = 3000   ' En milisegundos
Global Const ModoSimulacion As Boolean = False
Global Const SimulacionTipo As Integer = 0 '0 para voto negativo. 1 para voto positivo
Public SecuenciaStatus As Integer
Private Const ReintentosMensajes      As Integer = 3   ' Reintentos de Mensajes
Private ReintentosSkt()               As Integer
Private ElimMsg                       As Boolean
Private tickRsCola                    As Long
Private DebugSecuencia As String
Public BancasDeshabilitadas(256) As Boolean
Public CantidadInserciones As Long
Public RespuestasSNUVER As Long
Public EnviandoSNUVER As Boolean

'Cargamos datos para la cola de mensajes

Public Sub setearCola()
    'Seteo Variables
    Dim naUX    As Integer
    
    ReDim ReintentosSkt(0 To ConexionesAbiertas)
    For naUX = 0 To ConexionesAbiertas
        ReintentosSkt(naUX) = 0
    Next
    ElimMsg = False
    
    tickRsCola = GetTickCount
    If (RsCola Is Nothing) = False Then
        RsCola.Close
        Set RsCola = Nothing
    End If
    
    Set RsCola = New ADODB.Recordset
    RsCola.Fields.Append "Secuencia", adChar, 1
    RsCola.Fields.Append "Socket", adInteger
    RsCola.Fields.Append "Mensaje", adVarChar, 5120 '200
    RsCola.Fields.Append "Tick", adDouble, 20
    RsCola.Fields.Append "TickAux", adDouble 'No se usa
    RsCola.Fields.Append "Reintentos", adInteger
    
    'RsCola.CursorLocation = adUseClient
    RsCola.CursorType = adOpenDynamic
    RsCola.CursorType = adOpenStatic
    RsCola.Open
End Sub


'Determino que tipo de datos son y los coloco en la cola de Mensajes.
'********************************************************************

Public Sub EnviarSktxCola(fObjeto As String, fMensaje As String)
    Dim naUX         As String
    Dim BancaAux    As Integer
    Dim SocketAux As Integer
    Dim nVector()    As String
    Dim nContVect    As Long

    'Funcion
    If LCase(fObjeto) = "brc" Then ' Aqui entramos si en Broadcast.!!
        For SocketAux = 0 To ConexionesAbiertas - 1
            If Banca_ip(SocketAux).Estado = True Then
                If Skt2B(SocketAux).Banca > 0 Or Not (InStr("SRLEGI;SLEVER;SNUVER", Mid(fMensaje, 1, 6)) > 0) Then
                    If (InStr("SRLEGI", Mid(fMensaje, 1, 6)) > 0) Then
                        If EnviandoHuellas(SocketAux) Then 'Estaba Naux
                            Call CargarCola(SocketAux, fMensaje) ' solo srlegi
                        End If
                    ElseIf InStr(fMensaje, "SNUVER") > 0 Then
                        Call FormMain.EnviarxSkt("f", SocketAux, fMensaje)
                    Else
                        Dim n As Long
                        'Evita enviar STATUS mientras esta enviando huellas
                        If Not EnviandoHuellas(SocketAux) Or Not (InStr("STATUS", Mid(fMensaje, 1, 6)) > 0) Then
                            n = GetTickCount
                            If VectorEnvio(SocketAux) = "0" Or GetTickCount - VectorTicks(SocketAux) > 3000 Then
                                'Si es la primera vez que se utiliza el vector (primer envio de mensaje)
                                Call CargarCola(SocketAux, fMensaje)
                                VectorEnvio(SocketAux) = fMensaje
                                VectorTicks(SocketAux) = n
                                'Si el ultimo mensaje se envio hace más de tres segundos
                                'Se envia para que la banca no muera (puede ser el caso de un STATUS en Quorum)
                                Log_DEBUG ("BANCA " & Skt2B(SocketAux).Banca & " : Se evitó duplicidad de mensaje " & fMensaje)
                                If Not MODOLIGHT Then
                                    Call Log_Banca(Trim(Str(SocketAux)) & ".txt", Now() & " - Se evitó duplicidad de mensaje " & fMensaje)
                                End If
                            ElseIf VectorEnvio(SocketAux) <> fMensaje Then
                                'Si el mensaje es distinto del ultimo enviado
                                Call CargarCola(SocketAux, fMensaje)
                            End If
                        End If
                    End If
                End If
            Else
'                If InStr(fMensaje, "SLIMVT") > 0 Then
'                    Call Log_DEBUG("Banca IP en estado FALSE, no se cargo cola : " & naUX & " Msj : " & fMensaje & " Conex : " & ConexionesAbiertas)
'                    Call Log_DEBUG("Estado de Skt2B : " & Skt2B(naUX).Banca & " ESTADO " & Skt2B(naUX).Estado)
'                End If
            End If
        Next
    Else
        If (InStr(1, fObjeto, ";")) > 0 Then ' Si es un Split por aqui
            nVector = Split(fObjeto, ";")
            'If UBound(nVector) - 1 < (ConexionesAbiertas - 1) Then
            '    nContVect = UBound(nVector) - 1
            'Else
            '    nContVect = ConexionesAbiertas - 1
            'End If
            For BancaAux = 0 To UBound(nVector) - 1
                If BancaAux = 21 Then
                   BancaAux = 21
                End If
                If nVector(BancaAux) = 1 And B2Skt(BancaAux).Estado = True Then
                    If BancaAux > 0 Or Not (InStr("SRLEGI;SLEVER;SNUVER", Mid(fMensaje, 1, 6)) > 0) Then
                        If (InStr("SRLEGI", Mid(fMensaje, 1, 6)) > 0) Then
                            If EnviandoHuellas(B2Skt(BancaAux).Socket) Then Call CargarCola(B2Skt(BancaAux).Socket, fMensaje) ' solo srlegi
                        Else
                            If Not EnviandoHuellas(B2Skt(BancaAux).Socket) Or Not (InStr("STATUS", Mid(fMensaje, 1, 6)) > 0) Then Call CargarCola(B2Skt(BancaAux).Socket, fMensaje)
                        End If
                    End If
                End If
            Next
        Else
            If B2Skt(Trim(Str(fObjeto))).Estado = True Then ' Ultimo Caso mensaje solo para una banca en particualar
                    If Not EnviandoHuellas(B2Skt(Trim(Str(fObjeto))).Socket) Or Not (InStr("STATUS", Mid(fMensaje, 1, 6)) > 0) Then
                        Call CargarCola(B2Skt(Trim(Str(fObjeto))).Socket, fMensaje)
                    End If
            End If
        End If
    End If

End Sub

Public Sub EstablecerEnviandoHuellas(fObjeto As String)
    Dim naUX         As String
    Dim BancaAux    As Integer
    Dim SocketAux As Integer
    
    Dim nVector()    As String
    Dim nContVect    As Long
    Dim aTemp As String
    Dim i As Integer
    
    For i = LBound(EnviandoHuellas) To UBound(EnviandoHuellas)
        EnviandoHuellas(i) = False
    Next i
    'Funcion
    If LCase(fObjeto) = "brc" Then ' Aqui entramos si en Broadcast.!!
        For SocketAux = 0 To ConexionesAbiertas - 1
            If Banca_ip(SocketAux).Estado = True Then
                If Skt2B(SocketAux).Banca > 0 Then
                    EnviandoHuellas(SocketAux) = True
                End If
            End If
        Next
    Else
        If (InStr(1, fObjeto, ";")) > 0 Then ' Si es un Split por aqui
            nVector = Split(fObjeto, ";")
            'If UBound(nVector) - 1 < (ConexionesAbiertas - 1) Then
            '    nContVect = UBound(nVector) - 1
            'Else
            '    nContVect = ConexionesAbiertas - 1
            'End If
            For BancaAux = 0 To UBound(nVector) - 1
                If nVector(BancaAux) = 1 And B2Skt(BancaAux).Estado = True Then
                    EnviandoHuellas(B2Skt(BancaAux).Socket) = True
                End If
            Next
        Else
            If B2Skt(Trim(Str(fObjeto))).Estado = True Then ' Ultimo Caso mensaje solo para una banca en particualar
                EnviandoHuellas(B2Skt(Trim(Str(fObjeto))).Socket) = True
            End If
        End If
    End If
    Call MostrarEnviandoHuellas
End Sub
Public Sub MostrarEnviandoHuellas()
Dim aTemp As String
Dim i As Integer
Dim nCant As Integer

    For i = LBound(EnviandoHuellas) To UBound(EnviandoHuellas)
        aTemp = Trim(aTemp) & IIf(EnviandoHuellas(i), "1", "0")
        nCant = nCant + IIf(EnviandoHuellas(i), 1, 0)
    Next i
    FormMain.txtEnviando.Text = aTemp
    If cNIVEL_LOG > 1 Then
        Call MostrarErr("Enviando faltan " & Str(nCant) & ":" & aTemp, Str(fSocket))
        FormMain.lblFaltanEnviar.Caption = nCant
    End If
End Sub
Public Function EstadoEnviandoHuellas() As Boolean
Dim i As Integer
Dim nCant As Integer
EstadoEnviandoHuellas = False
nCant = 0
For i = LBound(EnviandoHuellas) To UBound(EnviandoHuellas)
    nCant = nCant + IIf(EnviandoHuellas(i), 1, 0)
    If nCant > 1 Then
        i = UBound(EnviandoHuellas)
        EstadoEnviandoHuellas = True
    End If
Next i
End Function

'Funcion para cargar datos en la cola que es llamada por la funcion Anterior.
'****************************************************************************

Public Sub CargarCola(fSocket As Integer, fMensaje As String)
    'Si no hay mensajes y paso el timeout de la ultima vez que se entro aca entra
    If RsCola.RecordCount = 0 And (GetTickCount - tickRsCola) > 900000 Then
        'Reinicio la Cola para liberar memoria del recordset que con los delete no libera memoria
        If RsCola.State = 1 Then
            RsCola.Close
            Set RsCola = Nothing
            Set RsCola = New ADODB.Recordset
            RsCola.Fields.Append "Secuencia", adChar, 1
            RsCola.Fields.Append "Socket", adInteger
            RsCola.Fields.Append "Mensaje", adVarChar, 5120 '200            RsCola.Fields.Append "Mensaje", adVarChar, 200
            RsCola.Fields.Append "Tick", adDouble, 20
            RsCola.Fields.Append "TickAux", adDouble
            RsCola.Fields.Append "Reintentos", adInteger
            RsCola.Open
            tickRsCola = GetTickCount
            Call MostrarErr(Now() & " Se reinició la cola de mensajes.", 0)
        End If
    End If
    
    If Banca_ip(fSocket).Estado = True Then
        With RsCola
            .AddNew
            !secuencia = DarNuevaSecuencia(fSocket)
            !Socket = fSocket
            !mensaje = fMensaje
            !tick = "0"
            !reintentos = 0
            .UpdateBatch
        End With
    Else
        Log_DEBUG ("! CargarCola socket no habilitado " & fSocket & " Msj " & fMensaje)
    End If
    DoEvents
End Sub

Public Sub EliminarMensajeCola(fSocket As Integer, Optional fMensaje As String, Optional fTodo As Boolean = False)
    Dim fsecuencia         As String
    Dim sMensajeCompleto   As String
    Dim sPresencia         As String
    Dim sMensajeAck As String
'    If fSocket = 128 Then
'        FormMain.txtTemp = "ASDASD"
'    End If
    If Len(fMensaje) > 0 Then
        If Asc(Left(fMensaje, 1)) = 0 Then fMensaje = Mid(fMensaje, 2, 4000)
    End If
    'If fSocket = 50 Then Stop ' Solo Depuracion
    If fTodo = False Then
        fsecuencia = Mid(fMensaje, 1, 1) ' Secuencia
        sMensaje = Mid(fMensaje, 2, Len(fMensaje) - 1) ' Mensaje secuencia
        If Mid(sMensaje, 1, 6) = "TACKNL" Or Mid(sMensaje, 1, 6) = "TNACKN" Or Mid(sMensaje, 1, 6) = "TESTAD" Then
            sPresencia = Mid(fMensaje, 15, 1) 'Saca el estado de presencia de la banca
        End If
        sMensajeAck = Left(Trim(sMensaje), 14)
    End If
    'La secuencia de mensaje va de ASCII 125 al 250
    'Si es f va por afuera de la cola
    'Si es entre A y Z viene de la banca por lo cual no han sido ninguno de los dos encolados
    If fsecuencia = "f" Or (fsecuencia > "A" And fsecuencia < "Z") Then
        Exit Sub ' Significa que el mensaje fue enviado directamente
    End If
    ElimMsg = True
    If fTodo = False Then
        'RsCola.Filter = adFilterNone
        'If RsCola.RecordCount <> 0 Then
        '    RsCola.MoveFirst
        'End If

        If sMensajeAck = "TACKNL EIDRXHP" Then
            'If cNIVEL_LOG > 1 Then Call MostrarErr("RECIBIDO ACKNOWLEDGE SEC " & Str(Asc(fsecuencia)) & " BANCA " & Str(fSocket) & " Mensaje " & (sMensaje), Str(fSocket))
            If 1 = 1 Then ' RECIBIDO ACKNOWLEDGE
            End If
        Else
            If 1 = 1 Then ' RECIBIDO ACKNOWLEDGE
            End If
            
        End If

        RsCola.Filter = "(Tick<>'0') AND (Socket=" & fSocket & ") AND (Secuencia='" & fsecuencia & "')"
        
        'RsCola.Find "Socket = " & fSocket, , , 1
        
        'RsCola.Filter = ""
        'RsCola.Filter = "(Secuencia='" & fsecuencia & "')"
        'RsCola.MoveFirst
        If Not RsCola.EOF Or RsCola.RecordCount < 0 Then
            ' Si hay un elemento borro el primero
            RsCola.MoveFirst
            'Antes de borrar tengo que preguntar si es alguna secuencia conocida
            Select Case RsCola!mensaje
                Case "SNUVER"
                    'Entra en el estado y se borra o si no esta tambien
                    RsCola.Delete
                Case "SRLEGI" 'banana
                        RsCola.Delete
                Case "SVOTNU"
                    'Entra en el estado y se borra o si no esta tambien
                    If sMensajeAck = "TACKNL ELISVTP" Or sMensajeAck = "TNACKN ELISVTP" _
                    Or sPresencia = "A" Then
                        RsCola.Delete
                    End If
                Case "SVOTAR"
                    'Entra en el estado y se borra o si no esta tambien
                    If sMensajeAck = "TACKNL ELISVTP" Or sMensajeAck = "TNACKN ELISVTP" _
                        Or sPresencia = "A" Then
                        RsCola.Delete
                    End If
                Case "SIDRXH"
                    If sMensajeAck = "TACKNL EIDRXHP" Or Left(sMensajeAck, 6) = "TACKNL" Then
                    'Or sMensajeAck = "TNACKN EIDRXHP" Or sMensajeAck = "TNACKN EIDACPP" Or sPresencia = "A" Then
                        RsCola.Delete
                    End If
                    If Left(sMensajeAck, 6) = "TNACKN" Then
                        'Call EnviarxSkt("X", index, "SCANCL")  test091013  aquii
                        RsCola.Delete
                    End If

                    
                Case "SIDRNX"
                    If sMensajeAck = "TACKNL EIDRNXP" Or sMensajeAck = "TNACKN EIDRNXP" _
                    Or sPresencia = "A" Then
                        RsCola.Delete
                    End If
                Case "SIDRNH"
                    If sMensajeAck = "TACKNL EIDRNHP" Or sMensajeAck = "TNACKN EIDRNHP" _
                    Or sPresencia = "A" Then
                        RsCola.Delete
                    End If
                Case "SIDRDH"
                    If sMensajeAck = "TACKNL EIDRDHP" Or sMensajeAck = "TNACKN EIDRDHP" _
                    Or sPresencia = "A" Then
                        RsCola.Delete
                    End If
                Case "SCANCL"
                    If fSocket = 0 Then
                        fSocket = fSocket
                    End If
                    If Trim(sMensajeAck) = "TACKNL" _
                    Or sMensajeAck = "TACKNL EINACTP" Or sMensajeAck = "TNACKN EINACTP" _
                    Or sMensajeAck = "TACKNL EINACTA" Or sMensajeAck = "TNACKN EINACTA" Or sMensajeAck = "TACKNL EIDACPP" Then
                        RsCola.Delete
                    End If
                Case "SLIMVT"
                    'MsgBox "SLIMVT"
                    RsCola.Delete
                Case Else
                    RsCola.Delete
            End Select
            RsCola.UpdateBatch
        Else
            ' En este caso puede ser un reintento que llego el dato despues de que se envio de vuelta
            ' Ej. SCANCL no llego respuesta SCANCL,. => TACKNL EINACTP, borro no hay mas registro
            ' Nota creo que no hace falta porque si llega la segunda respuesta yo no puse la bandera en off para enviar la segunda peticion
            ' EstadoxEnviar(fSocket) = True
        End If
    Else
        RsCola.Filter = "(Socket=" & fSocket & ")"
        If Not RsCola.EOF Then
            RsCola.MoveFirst
            Do While Not RsCola.EOF
                RsCola.Delete
                RsCola.MoveNext
                DoEvents
            Loop
            RsCola.UpdateBatch
        End If
    End If
    RsCola.Filter = adFilterNone
    If RsCola.RecordCount <> 0 Then
        RsCola.MoveFirst
    End If
End Sub

'Public Sub EnviarDatosSocket()
'On Error GoTo ErrEOFBOF
'Dim nRecordCount   As Long
'    nRecordCount = RsCola.RecordCount
'    FormMain.TRsCola.Text = nRecordCount
'    If Not nRecordCount = 0 Then
'        RsCola.MoveFirst
'        Do While Not (RsCola.EOF Or RsCola.AbsolutePosition > 80) ' Ver si cambio por solo unos cuantos. ?
'            If ElimMsg = True Then
'                ElimMsg = False
'                'Exit Sub 'Sale porque se borro un mensaje de la cola y tiene que comensar de vuelta
'            End If
'            If RsCola!TICK <> 0 Then
'                'If (GetTickCount - RsCola!TICK) > TimeOutMensaje Or RsCola!REINTENTOS > 3 Then
'                If RsCola!reintentos > 3 Then
'                    Call MostrarErr("Se cerro banca " & Str(Skt2B(RsCola!Socket).Banca) & " Mensaje " & (RsCola!mensaje) & " Reint: " & Str(RsCola!reintentos - 1) & " Tick : " & Str(GetTickCount - RsCola!TICK))
'                    Call FormMain.WSocketClose(RsCola!Socket)
'                    Exit Sub
'                End If
'            End If
'            If ((GetTickCount - RsCola!TICK) > TimeOutReintentos) Or (RsCola!TICK = 0) Then
'                If EstadoxEnviar(RsCola!Socket) = True Then
'                    EstadoxEnviar(RsCola!Socket) = False
'                    RsCola!TICK = GetTickCount
'                    RsCola!reintentos = RsCola!reintentos + 1
'                    Call FormMain.EnviarxSkt(RsCola!SECUENCIA, RsCola!Socket, RsCola!mensaje)
'                Else
'                    '****************************************************
'                    ' La idea es que si no se puede enviar verifico
'                    ' que ya se haya enviado entonces reintento
'                    '****************************************************
'                    If RsCola!reintentos > 0 Then
'                        RsCola!TICK = GetTickCount
'                        RsCola!reintentos = RsCola!reintentos + 1
'                        Call FormMain.EnviarxSkt(RsCola!SECUENCIA, RsCola!Socket, RsCola!mensaje)
'                    End If
'                    'EstadoxEnviar(RsCola!Socket) = True
'                    'RsCola.MoveFirst
'                    'If RsCola!Socket = 70 Then Stop ' Depuracion
'                    'RsCola!TICK = GetTickCount
'                    'RsCola!REINTENTOS = RsCola!REINTENTOS + 1
'                End If
'            End If
'            RsCola.MoveNext
'            DoEvents
'            If ElimMsg = True Then
'                ElimMsg = False
'                'Exit Sub 'Sale porque se borro un mensaje de la cola y tiene que comensar de vuelta
'            End If
'        Loop
'    End If
'    Exit Sub
'ErrEOFBOF:
'    Select Case Err.Number
'        Case 3021
'            MostrarErr ("*------------------------------------------*")
'            Exit Sub
'        Case Else
'            Resume Next
'    End Select
'End Sub

'******************************************************
'Funcion Nueva que envia datos socket por socket...
'******************************************************
Public Sub EnviarDatosSkt()
'On Error GoTo ErrEOFBOF
    Dim xAux   As Integer
    Dim Enviado As Boolean
    Enviado = False
    If RsCola.RecordCount > 0 Then
        For xAux = 0 To ConexionesAbiertas
            'Indico que busque el primer soket que encuentre en la lista :)
            '***************************************************************
            If xAux = 256 Then
               ' MsgBox ("asd")
            End If
            RsCola.UpdateBatch
            RsCola.Find "Socket = " & Str(xAux), , , 1
            If RsCola.AbsolutePosition > 0 Then
'                If RsCola!Socket = 33 And InStr(RsCola!mensaje, "SLIMVT") Then
'                    Log_DEBUG ("W33 " & RsCola!mensaje)
'                End If
                ' Proceso los Datos
                If RsCola!tick <> 0 Then
                    If RsCola!reintentos > 10 Then 'esteeselquelomata
                        If (Left(RsCola!mensaje, 6) = "SRLEGI" And (RsCola!reintentos > 10) Or Not EnviandoHuellas(RsCola!Socket)) Or Not Left(RsCola!mensaje, 6) = "SRLEGI" Then
                            'If cNIVEL_LOG > 0 Then Call MostrarErr("01 Se cerro banca " & Str(Skt2B(RsCola!Socket).Banca) & " Reint: " & Str(RsCola!reintentos - 1) & " Tick : " & Str(GetTickCount - RsCola!tick) & " Mensaje " & Left(RsCola!mensaje, 30), Str(Skt2B(RsCola!Socket).Banca))
                            Log_DEBUG ("Se excedió el máximo de reintentos BANCA " & Str(Skt2B(RsCola!Socket).Banca) & " Mensaje " & Left(RsCola!mensaje, 30) & " Reintentos:" & RsCola!reintentos)
                            If Not MODOLIGHT Then
                                Call Log_Banca(Trim(Str(RsCola!Socket)) & ".txt", Now() & " - ERROR MAXIMO REINTENTOS " & sEnviar)
                            End If
                            Call FormMain.WSocketClose(RsCola!Socket, "SRLEGI MAX REINTENTOS")
                            Exit Sub
                        End If
                    Else
                        'MsgBox RsCola!mensaje
                    End If
                End If
                Dim TC As Long
                Dim TCTICK As Long
                TC = GetTickCount
                TCTICK = RsCola!tick
                If ((GetTickCount - RsCola!tick) > TimeOutReintentos) Or (RsCola!tick = 0) Then
                        RsCola!tick = GetTickCount
                        RsCola!reintentos = RsCola!reintentos + 1
                        Log_DEBUG ("SE ENVIO : " & RsCola!Socket & " ------ " & RsCola!mensaje)
                        If Not MODOLIGHT Then
                            Call Log_Banca(Trim(Str(RsCola!Socket)) & ".txt", Now() & " - Reintento de envio " & RsCola!reintentos & " - " & sEnviar)
                        End If
                        Call FormMain.EnviarxSkt(RsCola!secuencia, RsCola!Socket, RsCola!mensaje)
                ElseIf InStr(RsCola!mensaje, "SLIMVT") > 0 Then
                        Log_DEBUG ("NO SE ENVIO : " & RsCola!Socket & " ------ " & RsCola!mensaje)
                        Log_DEBUG ("TICK: " & RsCola!tick & " REINTENTOS: " & RsCola!reintentos & "| Supero timeout : " & IIf((GetTickCount - RsCola!tick) > TimeOutReintentos, "SI", "NO"))
                End If
            End If
        Next
    End If
    Exit Sub
ErrEOFBOF:
    Select Case Err.Number
        Case 3021
            MostrarErr ("*------------ - - - ---------------------------*")
            Exit Sub
    End Select
End Sub

Public Sub GuardarCola()
'On Error GoTo ErrEOFBOF
    Dim xAux   As Integer
    Dim Enviado As Boolean
    Enviado = False
    Log_DEBUG (">>>>>>>>>>>>>>>>>>>>>> Volcado de cola")
    If RsCola.RecordCount > 0 Then
        For xAux = 0 To ConexionesAbiertas
            Log_DEBUG ("------------- " & xAux)
            'Indico que busque el primer soket que encuentre en la lista :)
            '***************************************************************
            'RsCola.UpdateBatch
            RsCola.Filter = "Socket = " & Str(xAux)
            'RsCola.Find "Socket = " & Str(xAux), , , 1
            If Not RsCola.EOF Or RsCola.RecordCount < 0 Then
                RsCola.MoveFirst
                Do While Not (RsCola.EOF)
                    Log_DEBUG (RsCola!secuencia & " | " & RsCola!Socket & " | " & RsCola!mensaje & " | " & RsCola!tick & " | " & RsCola!reintentos & " | " & RsCola.AbsolutePosition)
                    RsCola.MoveNext
                Loop
            End If
            RsCola.Filter = adFilterNone
            If RsCola.RecordCount <> 0 Then
                RsCola.MoveFirst
            End If
        Next
    End If
    Log_DEBUG ("<<<<<<<<<<<<<<<<< Fin Volcado de cola")
    Exit Sub
ErrEOFBOF:
    Select Case Err.Number
        Case 3021
            MostrarErr ("*------------ - - - ---------------------------*")
            Exit Sub
    End Select
End Sub

Public Function DarNuevaSecuencia(fSocket As Integer) As String
    ' Puedo dar desde el 125 al 250
    Dim naUX
    If UltimaSecuenciaBanca(fSocket) = 250 Then
       UltimaSecuenciaBanca(fSocket) = 125
    Else
       UltimaSecuenciaBanca(fSocket) = UltimaSecuenciaBanca(fSocket) + 1
    End If
    DarNuevaSecuencia = Chr(UltimaSecuenciaBanca(fSocket))
End Function

