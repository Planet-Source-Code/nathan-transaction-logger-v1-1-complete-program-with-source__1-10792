VERSION 5.00
Begin VB.Form Editfrm 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Transaction"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
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
      Left            =   3840
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000009&
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
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000009&
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
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000009&
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
      MaxLength       =   8
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000009&
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
      MaxLength       =   25
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000009&
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
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
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
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
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
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
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
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction..."
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
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Recorded......"
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
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Editfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    holdindex = Mainfrm.Mainlst.ListIndex
    len1x = Len(Text1)
    len2x = Len(Text2)
    len3x = Len(Text3)
    len4x = Len(text4)
    len5x = Len(text5)
    If len1x < 10 Then
        Do
            Text1.Text = Text1.Text & " "
        Loop Until Len(Text1) = 10
    End If
    If len2x < 25 Then
        Do
            Text2.Text = Text2.Text + " "
        Loop Until Len(Text2) = 25
    End If
    If len3x < 7 Then
        Do
            Text3.Text = " " + Text3.Text
        Loop Until Len(Text3) = 7
    End If
    If len4x < 10 Then
        Do
            text4.Text = text4.Text & " "
        Loop Until Len(text4) = 10
    End If
    If len5x < 30 Then
        Do
            text5.Text = text5.Text & " "
        Loop Until Len(text5) = 30
    End If
    Mainfrm.Mainlst.RemoveItem (Mainfrm.Mainlst.ListIndex)
    Mainfrm.Mainlst.AddItem Text1 & "     " & Text2 & "      " & Text3 & "      " & text4 & "      " & text5, holdindex
    Call Mainfrm.savelist
    Unload Editfrm
End Sub
Private Sub Form_Load()
    center Me
    DoEvents
    SetFormTopmost Editfrm
    Editfrm.Height = 0
    For X = 1 To 1820
        Editfrm.Height = Editfrm.Height + 1
        center Me
        DoEvents
        If Editfrm.Visible <> True Then
            Editfrm.Visible = True
        End If
    Next X
    X = Mainfrm.Mainlst.List(Mainfrm.Mainlst.ListIndex)
    Text1.Text = Trim(Left(X, 10))
    Text2.Text = Trim(Mid(X, 16, 25))
    Text3.Text = Trim(Mid(X, 46, 8))
    text4.Text = Trim(Mid(X, 60, 10))
    text5.Text = Trim(Mid(X, 76, 30))
End Sub
Private Sub Form_Unload(Cancel As Integer)
    For X = 1 To 1820
        Editfrm.Height = Editfrm.Height - 1
        center Me
        DoEvents
    Next X
    Editfrm.Height = 0
    Mainfrm.can_we_edit = "Y"
End Sub


