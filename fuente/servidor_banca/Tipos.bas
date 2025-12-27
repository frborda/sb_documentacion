Attribute VB_Name = "Tipos"
'*******************************************************
' Type Para informacion de las Bancas                  *
'*******************************************************

Type BancaIP
    Banca              As Integer
    IP                 As String
    Puerto             As String
    Estado             As Boolean
    LegisladorActivo   As Long
    tEstado            As String
    tSecLegislador     As String
    tSecMantenimiento  As String
    tSecBusqueda       As String
    tBancaMinMax       As String
    tbancaMinMaxMan    As String
    tBancaBusca        As String
    tBancaSecuencia    As String
    tVersion           As String

End Type

'*******************************************************
' Ingreso la Banca y me devuelve el Socket             *
'*******************************************************

Type BancaSkt
     Socket  As Integer
     Estado As Boolean
End Type

'******************************************************
' Ingreso el Socket y me devuelve el Numero de Banca  *
'******************************************************
Type SktBanca
     Banca As Integer
     Estado As Boolean
End Type

'*******************************************************
' Defino Estructura de la Tabla de los mensajes del SQV*
'*******************************************************

Type MensajeSQV
    sTipo                     As String
    sComponente               As String
    sObjeto                   As String
    sAtributo                 As String
    sValor                    As String
    sComentario               As String
End Type

Type TLegislador
     sId                      As String
     sNombre                  As String
     sApellido                As String
     sDNI                     As String
     sClase                   As String
     sIcono                   As String
     sTemplate                As String
     sTemplate11()            As Byte
End Type

'*********************************************************
'  Estructura de Legisladores
'*********************************************************


Type Legisladores
    sId            As Long
    sBanca         As Integer
    sMantenimiento As Boolean
End Type
