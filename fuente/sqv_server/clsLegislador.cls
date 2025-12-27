VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLegislador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales para almacenar los valores de las propiedades
Private mvarApellido As String 'copia local
Private mvarNombre As String 'copia local
Private mvarApellidoNombre As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarID As String 'copia local
Public Property Let ID(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ID = 5
    mvarID = vData
End Property


Public Property Get ID() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property



Public Property Let ApellidoNombre(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ApellidoNombre = 5
    mvarApellidoNombre = vData
End Property


Public Property Get ApellidoNombre() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ApellidoNombre
    ApellidoNombre = mvarApellidoNombre
End Property



Public Property Let Nombre(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Nombre = 5
    mvarNombre = vData
End Property


Public Property Get Nombre() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Nombre
    Nombre = mvarNombre
End Property



Public Property Let Apellido(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Apellido = 5
    mvarApellido = vData
End Property


Public Property Get Apellido() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Apellido
    Apellido = mvarApellido
End Property



