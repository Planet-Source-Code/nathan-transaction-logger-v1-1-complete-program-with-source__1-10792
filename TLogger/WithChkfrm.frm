VERSION 5.00
Begin VB.Form WithChkfrm 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Withdrawal From Checking"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4230
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Date.........."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Ammount......."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Description..."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "WithChkfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    check1 = IsNumeric(Text2.Text)
    check2 = IsNumeric(Left(Text1.Text, 2) & Mid(Text1.Text, 4, 2) & Right(Text1.Text, 2))
    If check2 = False Then
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        Text1.SetFocus
    ElseIf check1 = False Then
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2.Text)
        Text2.SetFocus
    Else
        If Mainfrm.can_we_edit = "N" Then
            Call Mainfrm.loadlist
        End If
        len1x = Len(Text1)
        len2x = Len(Text2)
        len3x = Len(Text3)
        If len1x < 10 Then
            Do
                Text1.Text = Text1.Text & " "
            Loop Until Len(Text1) = 10
        End If
        If len2x < 7 Then
            Do
                Text2.Text = " " + Text2.Text
            Loop Until Len(Text2) = 7
        End If
        If len3x < 30 Then
            Do
                Text3.Text = Text3.Text & " "
            Loop Until Len(Text3) = 30
        End If
        Mainfrm.Mainlst.AddItem Format(Date, "mm/dd/yyyy") & "     " & "Withdrawal From Checkings" & "     " & Text2 & "-      " & Text1 & "      " & Text3, 0
        Call Mainfrm.savelist
        Unload WithChkfrm
    End If
End Sub
Private Sub Form_Load()
    center Me
    DoEvents
    SetFormTopmost WithChkfrm
    WithChkfrm.Height = 0
    For X = 1 To 1125
        WithChkfrm.Height = WithChkfrm.Height + 1
        center Me
        DoEvents
        If WithChkfrm.Visible <> True Then
            WithChkfrm.Visible = True
        End If
    Next X
    Text1.SetFocus
End Sub
Private Sub Form_Unload(Cancel As Integer)
    For X = 1 To 1125
        WithChkfrm.Height = WithChkfrm.Height - 1
        center Me
        DoEvents
    Next X
    WithChkfrm.Height = 0
End Sub
Private Sub Text1_Change()
    If Len(Text1.Text) = 10 Then
        Text2.SetFocus
    End If
End Sub
Private Sub Text1_LostFocus()
    Text1.Text = Format(Text1.Text, "mm/dd/yyyy")
End Sub
Private Sub Text2_Change()
    If Left(Right(Text2.Text, 3), 1) = "." Then
        Text3.SetFocus
    End If
End Sub
Private Sub Text2_LostFocus()
    Text2.Text = Format(Text2.Text, "####.#0")
End Sub


