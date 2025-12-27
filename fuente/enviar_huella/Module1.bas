Attribute VB_Name = "Module1"
Public Declare Function GetTickCount Lib "kernel32" () As Long
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
