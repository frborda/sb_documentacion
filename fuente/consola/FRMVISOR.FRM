VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "CRVIEWER.DLL"
Begin VB.Form frmVisor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visor de Informes"
   ClientHeight    =   8145
   ClientLeft      =   2190
   ClientTop       =   2430
   ClientWidth     =   10830
   Icon            =   "frmVisor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   10830
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8115
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10515
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   1095
   End
   Begin ComctlLib.ImageList ImgLst 
      Left            =   7980
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "frmVisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo Form_Load_err

    Screen.MousePointer = vbHourglass
    Call Form_Resize
    Screen.MousePointer = vbDefault
    Exit Sub

' Let the user know about any errors that might have occurred.
Form_Load_err:
            MsgBox "Error: " + CStr(Err) + Chr(10) + Chr(13) + _
            Error(Err), , "Form Load"
End Sub

Private Sub Form_Resize()
    CRViewer1.Left = 0
    CRViewer1.Height = Me.Height - 330
    CRViewer1.Width = Me.Width - 100

End Sub

