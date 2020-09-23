VERSION 5.00
Begin VB.Form Datapadfrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Datapad"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Richtextbox1 
      Height          =   3015
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Datapadfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    center Me
    DoEvents
    SetFormTopmost Datapadfrm
    Datapadfrm.Height = 0
    For X = 1 To 2940
        Datapadfrm.Height = Datapadfrm.Height + 1
        center Me
        DoEvents
        If Datapadfrm.Visible <> True Then
            Datapadfrm.Visible = True
        End If
    Next X
    Open (App.Path & "\datapad.dat") For Input As #1
    Richtextbox1 = Input(LOF(1), 1)
    Close #1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    For X = 1 To 2940
        Datapadfrm.Height = Datapadfrm.Height - 1
        center Me
        DoEvents
    Next X
    Datapadfrm.Height = 0
    Open (App.Path & "\datapad.dat") For Output As #1
    Print #1, Richtextbox1
    Close #1
End Sub


