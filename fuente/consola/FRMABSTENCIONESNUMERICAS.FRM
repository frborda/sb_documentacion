VERSION 5.00
Begin VB.Form frmAbstencionesNumericas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Abstenciones para votación Numérica"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstBloques 
      Height          =   960
      Left            =   8040
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   120
      Width           =   3735
   End
   Begin VB.PictureBox Picture2 
      Height          =   915
      Left            =   1560
      ScaleHeight     =   855
      ScaleWidth      =   1245
      TabIndex        =   10
      Top             =   120
      Width           =   1300
      Begin VB.CommandButton Salir 
         Caption         =   "&Salir"
         Height          =   855
         Left            =   0
         Picture         =   "frmAbstencionesNumericas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   915
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1260
      TabIndex        =   9
      Top             =   120
      Width           =   1320
      Begin VB.CommandButton cmdPin 
         Caption         =   "&Abstener"
         Height          =   855
         Left            =   0
         Picture         =   "frmAbstencionesNumericas.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5625
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   11655
      Begin VB.CommandButton QuitarTodos 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5580
         TabIndex        =   8
         Top             =   1725
         Width           =   495
      End
      Begin VB.CommandButton QuitarUno 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5580
         TabIndex        =   7
         Top             =   1350
         Width           =   495
      End
      Begin VB.CommandButton AgregarTodos 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5580
         TabIndex        =   6
         Top             =   825
         Width           =   495
      End
      Begin VB.CommandButton AgregarUno 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5580
         TabIndex        =   5
         Top             =   450
         Width           =   495
      End
      Begin VB.ListBox lstBancasSeleccionadas 
         Height          =   4935
         Left            =   6150
         TabIndex        =   4
         Top             =   450
         Width           =   5355
      End
      Begin VB.ListBox lstBancasNoSeleccionadas 
         Height          =   4935
         Left            =   150
         MultiSelect     =   1  'Simple
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   450
         Width           =   5355
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bancas Seleccionadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6150
         TabIndex        =   2
         Top             =   180
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bancas NO seleccionadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   1
         Top             =   180
         Width           =   2775
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Abstener por Bloques : "
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
      Left            =   6000
      TabIndex        =   14
      Top             =   120
      Width           =   2010
   End
End
Attribute VB_Name = "frmAbstencionesNumericas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CargarBancas()
    Dim strSql As String
    Dim Rs     As ADODB.Recordset
    
    Screen.MousePointer = 11
    Set Rs = New ADODB.Recordset
       
    strSql = "SELECT Legisladores.id, legisladores_activos.DESKID AS banca, Legisladores.apellido, Legisladores.nombre, " _
           & "Legisladores.bloque_politico , dbo.Legisladores.grupo_politico " _
           & "FROM Legisladores INNER JOIN legisladores_activos ON Legisladores.id = legisladores_activos.ID " _
           & "ORDER BY legisladores_activos.DESKID"
    SetearRs strSql, Rs
    
    With Rs
        If .RecordCount > 0 Then
            lstBancasNoSeleccionadas.Clear
            While Not .EOF
                If .Fields("Banca").Value > 0 Then
                    lstBancasNoSeleccionadas.AddItem Format(.Fields("Banca").Value, "00") & ") " & .Fields("Apellido").Value & ",  " & .Fields("Nombre").Value & " - " & UCase(.Fields("bloque_politico").Value)
                End If
                .MoveNext
            Wend
        End If
        .Close
    End With
    Set Rs = Nothing
    Screen.MousePointer = 0
End Sub
Private Sub AgregarTodos_Click()
    Call QuitarTodos_Click
    Dim x As Long
    lstBancasSeleccionadas.Clear
    With lstBancasNoSeleccionadas
        .Visible = False
        If .ListCount > 0 Then
            For x = 0 To .ListCount - 1
                .ListIndex = x
                lstBancasSeleccionadas.AddItem .Text
            Next x
            .Visible = True
            .Clear
        End If
    End With
    Call DesBloques
End Sub

Private Sub AgregarUno_Click()
    Dim xIndice As Long
    Dim strVect As String
    Dim xVect() As String
    
    With lstBancasNoSeleccionadas
        If .ListIndex = -1 Then                      ' Si no se selecciono ninguna banca, se sale de la rutina
            Exit Sub
        End If
        .Visible = False
        For xIndice = 0 To .ListCount - 1            ' Se recorre la lista
            .ListIndex = xIndice                     ' Elemento por elemento
            If .Selected(xIndice) Then               ' y si un elemento de la lista esta seleccionado
                lstBancasSeleccionadas.AddItem .Text ' Se lo pasa a la otra lista
            Else                                     ' caso contrario
                .ListIndex = xIndice                 ' se guarda el item en un string,  separado con punto y coma
                strVect = strVect & Trim(.Text) & ";"
            End If
        Next xIndice
        .Clear                                       ' Se borra la lista
        If Trim(strVect) <> "" Then                  ' si hay elementos en el estring
            xVect = Split(strVect, ";")              ' Se crea uun arreglo
            For xIndice = 0 To UBound(xVect) - 1     ' y se la carga con los elementos deel arreglo (los NO seleccionados)
                .AddItem xVect(xIndice)
            Next xIndice
        End If
        .Visible = True
    End With
    Call DesBloques
End Sub

Private Sub cmdPin_Click()

    Dim strLista As String
    Dim x        As Long
    Dim strItem  As String
    Dim xPos     As Long
    strLista = ""
    With lstBancasSeleccionadas
        If .ListCount > 0 Then
            For x = 0 To .ListCount - 1
                .ListIndex = x
                strItem = .Text
                xPos = InStr(1, strItem, ")")
                If xPos > 0 Then
                    strItem = Mid(strItem, 1, xPos - 1)
                    strItem = Str(Val(strItem))
                End If
                strLista = strLista & strItem & ";"
                strItem = ""
            Next x
        End If
    End With
    ' strLista = "00000118;00000119;00000101;00000100;00000103;"
    MensajesSQV.AbstenerLeg (strLista)
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub CargarBloques()
    Dim strSql As String
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    strSql = "SELECT Bloque_Político From Bloques ORDER BY bancaminima"
    SetearRs strSql, RsTemp
    With RsTemp
        lstBloques.Clear
        If .RecordCount > 0 Then
            While Not .EOF
                lstBloques.AddItem .Fields(0).Value
                .MoveNext
            Wend
        End If
        .Close
    End With
    Set RsTemp = Nothing
End Sub
Private Sub Form_Load()
    'Call CargarBancas
    Call BancasAbstenidas
    Call CargarBloques
End Sub
Private Sub BancasAbstenidas()
    Dim strSql As String
    Dim Rs As ADODB.Recordset
    Dim strLista As String

    Screen.MousePointer = 11
    Set Rs = New ADODB.Recordset

    'Levanta el Vector de Abstenciones
    strSql = "SELECT vector_abstencion FROM Vector"
    SetearRs strSql, Rs

    Rs.MoveFirst
    strLista = Rs.Fields("vector_abstencion").Value
    strLista = Replace(strLista, ";", ",")
    strLista = Replace(strLista, ",,", ",0,")
    strLista = Replace(strLista, ",,", ",0,")
    If Right(strLista, 1) = "," Then
        strLista = Left(strLista, Len(strLista) - 1)
    End If
    If Left(strLista, 1) = "," Then
        strLista = Right(strLista, Len(strLista) - 1)
    End If
    If strLista = "" Then strLista = "99999"
    
    Rs.Close
    
    
    'Carga la Lista de No seleccionados
    strSql = "SELECT Legisladores.id, legisladores_activos.DESKID AS banca, Legisladores.apellido, Legisladores.nombre, " _
        & "Legisladores.bloque_politico , dbo.Legisladores.grupo_politico " _
        & "FROM Legisladores INNER JOIN legisladores_activos ON Legisladores.id = legisladores_activos.ID " _
        & "Where (legisladores_activos.DESKID not in (" & strLista & ")) " _
        & "ORDER BY legisladores_activos.DESKID"
    SetearRs strSql, Rs
    
    With Rs
        If .RecordCount > 0 Then
            lstBancasNoSeleccionadas.Clear
            While Not .EOF
                If .Fields("Banca").Value > 0 Then
                    lstBancasNoSeleccionadas.AddItem Format(.Fields("Banca").Value, "00") & ") " & .Fields("Apellido").Value & ", " & .Fields("Nombre").Value & " - " & UCase(.Fields("bloque_politico").Value)
                End If
                .MoveNext
            Wend
        End If
        .Close
    End With
    
    'Carga la Lista de seleccionados
    
    strSql = "SELECT Legisladores.id, legisladores_activos.DESKID AS banca, Legisladores.apellido, Legisladores.nombre, " _
        & "Legisladores.bloque_politico , dbo.Legisladores.grupo_politico " _
        & "FROM Legisladores INNER JOIN legisladores_activos ON Legisladores.id = legisladores_activos.ID " _
        & "Where (legisladores_activos.DESKID  in (" & strLista & ")) " _
        & "ORDER BY legisladores_activos.DESKID"
    SetearRs strSql, Rs
    
    With Rs
        If .RecordCount > 0 Then
            lstBancasSeleccionadas.Clear
            While Not .EOF
                If .Fields("Banca").Value > 0 Then
                    lstBancasSeleccionadas.AddItem Format(.Fields("Banca").Value, "00") & ") " & .Fields("Apellido").Value & ", " & .Fields("Nombre").Value & " - " & UCase(.Fields("bloque_politico").Value)
                End If
                .MoveNext
            Wend
        End If
        .Close
    End With

    Set Rs = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub lstBancasNoSeleccionadas_DblClick()
    Call AgregarUno_Click
End Sub
Private Sub lstBancasSeleccionadas_DblClick()
    Call QuitarUno_Click
End Sub
Private Sub DesBloques()
    Dim x     As Long
    With lstBloques
        .Visible = False
        For x = 0 To .ListCount - 1
            .ListIndex = x
            If .Selected(x) = True Then
                .Selected(x) = False
            End If
        Next x
        .ListIndex = 0
        .Visible = True
    End With
End Sub
Private Sub lstBloques_Click()
    Dim x     As Long
    Dim xFind As Long
    With lstBloques
        xFind = 0
        lstBancasNoSeleccionadas.Visible = False
        If .Selected(.ListIndex) = True Then
            For x = 0 To lstBancasNoSeleccionadas.ListCount - 1
                lstBancasNoSeleccionadas.ListIndex = x
                xFind = InStr(1, LCase(lstBancasNoSeleccionadas.Text), LCase(.Text))
                If xFind > 0 Then
                    lstBancasNoSeleccionadas.Selected(x) = True
                    xFind = 0
                End If
            Next x
        Else
            For x = 0 To lstBancasNoSeleccionadas.ListCount - 1
                lstBancasNoSeleccionadas.ListIndex = x
                xFind = InStr(1, LCase(lstBancasNoSeleccionadas.Text), LCase(.Text))
                If xFind > 0 Then
                    lstBancasNoSeleccionadas.Selected(x) = False
                    xFind = 0
                End If
            Next x
        End If
        lstBancasNoSeleccionadas.ListIndex = 0
        lstBancasNoSeleccionadas.Visible = True
    End With
End Sub

Private Sub QuitarTodos_Click()
    Dim x As Long
    lstBancasNoSeleccionadas.Clear
    lstBancasSeleccionadas.Clear
    Call CargarBancas
    Call DesBloques
End Sub

Private Sub QuitarUno_Click()
    If lstBancasSeleccionadas.ListIndex = -1 Then
        Exit Sub
    End If
    lstBancasNoSeleccionadas.AddItem lstBancasSeleccionadas.Text
    lstBancasSeleccionadas.RemoveItem (lstBancasSeleccionadas.ListIndex)
    Call DesBloques
End Sub

Private Sub Salir_Click()
    Unload Me
End Sub
