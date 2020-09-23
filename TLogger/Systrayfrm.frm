VERSION 5.00
Object = "{33155A3D-0CE0-11D1-A6B4-444553540000}#1.0#0"; "SysTray.ocx"
Begin VB.Form Systemtrayfrm 
   ClientHeight    =   525
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   1560
   ClipControls    =   0   'False
   Icon            =   "Systrayfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   525
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin SysTray.SystemTray SystemTray1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      SysTrayText     =   ""
      IconFile        =   0
   End
End
Attribute VB_Name = "Systemtrayfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const sys_Add = 0
Private Const sys_Delete = 2
Private Sub Form_Load()
    If SystemTray1.IsIconLoaded = False Then
        SystemTray1.Icon = Val(Mainfrm.Icon)
        SystemTray1.SysTrayText = "Transaction Logger v1.0"
        SystemTray1.Action = sys_Add
    End If
End Sub
Private Sub SystemTray1_MouseDblClk(ByVal Button As Integer)
    Load Mainfrm
End Sub


