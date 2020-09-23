VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Calendarfrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calendar"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2010
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   3545
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483630
      Appearance      =   0
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   -2147483634
      StartOfWeek     =   662831105
      TitleBackColor  =   -2147483638
      TrailingForeColor=   -2147483639
      CurrentDate     =   36750
   End
End
Attribute VB_Name = "Calendarfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MonthView1.Value = Date
center Me
    DoEvents
    SetFormTopmost Calendarfrm
    Calendarfrm.Height = 0
    For X = 1 To 1910
        Calendarfrm.Height = Calendarfrm.Height + 1
        center Me
        DoEvents
        If Calendarfrm.Visible <> True Then
            Calendarfrm.Visible = True
        End If
    Next X
End Sub

Private Sub Form_Unload(Cancel As Integer)
 For X = 1 To 1910
    Calendarfrm.Height = Calendarfrm.Height - 1
    center Me
    DoEvents
Next X
Calendarfrm.Height = 0
End Sub
