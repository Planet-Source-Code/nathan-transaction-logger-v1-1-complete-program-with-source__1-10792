VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Mainfrm 
   BackColor       =   &H80000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1905
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mainfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox is_form_loading 
      Height          =   315
      Left            =   7920
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox can_we_edit 
      Height          =   315
      Left            =   7560
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   1320
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   8535
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   175
      Left            =   7105
      TabIndex        =   9
      Top             =   1710
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   318
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.ListBox List3 
      Height          =   270
      Left            =   6600
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      ItemData        =   "Mainfrm.frx":030A
      Left            =   7200
      List            =   "Mainfrm.frx":030C
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox Mainlst 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      ItemData        =   "Mainfrm.frx":030E
      Left            =   0
      List            =   "Mainfrm.frx":0315
      TabIndex        =   1
      Top             =   360
      Width           =   8535
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1650
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   "8/15/2000"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   8520
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label descriptionlbl 
      Caption         =   "Description"
      Height          =   255
      Left            =   5660
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label datelbl 
      Alignment       =   2  'Center
      Caption         =   "Date"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   795
   End
   Begin VB.Label ammountlbl 
      Caption         =   "Ammount"
      Height          =   255
      Left            =   3335
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label transactionlbl 
      Caption         =   "Tranaction"
      Height          =   255
      Left            =   1175
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label recordlbl 
      Caption         =   "Recorded"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu filenew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu fileopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu bara 
         Caption         =   "-"
      End
      Begin VB.Menu filesave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu saveas 
         Caption         =   "Save As"
      End
      Begin VB.Menu barw 
         Caption         =   "-"
      End
      Begin VB.Menu rstr 
         Caption         =   "Restore Data"
      End
      Begin VB.Menu bckup 
         Caption         =   "Backup Data"
      End
      Begin VB.Menu barnun 
         Caption         =   "-"
      End
      Begin VB.Menu filedel 
         Caption         =   "Delete"
      End
      Begin VB.Menu default 
         Caption         =   "Default File"
      End
      Begin VB.Menu baro 
         Caption         =   "-"
      End
      Begin VB.Menu prnt 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu bary 
         Caption         =   "-"
      End
      Begin VB.Menu exme 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu main 
      Caption         =   "&Transactions"
      Begin VB.Menu menudeposit 
         Caption         =   "Deposit"
         Begin VB.Menu menudepchk 
            Caption         =   "Checking"
         End
         Begin VB.Menu menudepsave 
            Caption         =   "Savings"
         End
      End
      Begin VB.Menu menuwithdrawel 
         Caption         =   "Withdrawl"
         Begin VB.Menu menuwithchk 
            Caption         =   "Checking"
         End
         Begin VB.Menu menuwithsav 
            Caption         =   "Savings"
         End
      End
      Begin VB.Menu transfer 
         Caption         =   "Transfer"
      End
      Begin VB.Menu adjust 
         Caption         =   "Adjustment"
         Begin VB.Menu chkadjust 
            Caption         =   "Checking"
            Begin VB.Menu posadjust 
               Caption         =   "Positive"
            End
            Begin VB.Menu negadjuts 
               Caption         =   "Negative"
            End
         End
         Begin VB.Menu svadjust 
            Caption         =   "Savings"
            Begin VB.Menu posadjust2 
               Caption         =   "Positive"
            End
            Begin VB.Menu savneg 
               Caption         =   "Negative"
            End
         End
      End
   End
   Begin VB.Menu menusort 
      Caption         =   "Transaction &List"
      Begin VB.Menu menusortdep 
         Caption         =   "Deposits"
         Begin VB.Menu tychk 
            Caption         =   "Checking"
         End
         Begin VB.Menu tysv 
            Caption         =   "Savings"
         End
      End
      Begin VB.Menu menusortwith 
         Caption         =   "Withdrawls"
         Begin VB.Menu wtsv 
            Caption         =   "Checking "
         End
         Begin VB.Menu wsv 
            Caption         =   "Savings"
         End
      End
      Begin VB.Menu menusorttran 
         Caption         =   "Transfers"
         Begin VB.Menu chk2sv 
            Caption         =   "Checking To Savings"
         End
         Begin VB.Menu sv2chk 
            Caption         =   "Savings To Checking"
         End
      End
      Begin VB.Menu adjustm 
         Caption         =   "Adjustments"
         Begin VB.Menu adjchk 
            Caption         =   "Checking"
            Begin VB.Menu adjpos 
               Caption         =   "Positive"
            End
            Begin VB.Menu adjneg 
               Caption         =   "Negative"
            End
         End
         Begin VB.Menu adjsav 
            Caption         =   "Savings"
            Begin VB.Menu posadjsav 
               Caption         =   "Positive"
            End
            Begin VB.Menu negadjsav 
               Caption         =   "Negative"
            End
         End
      End
      Begin VB.Menu barww 
         Caption         =   "-"
      End
      Begin VB.Menu details 
         Caption         =   "Details"
      End
      Begin VB.Menu barsome 
         Caption         =   "-"
      End
      Begin VB.Menu srch 
         Caption         =   "Search"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu menutransall 
         Caption         =   "Show All"
      End
   End
   Begin VB.Menu opt 
      Caption         =   "&Other"
      Begin VB.Menu datapad 
         Caption         =   "Datapad"
      End
      Begin VB.Menu calendar 
         Caption         =   "Calendar"
      End
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const LB_FINDSTRING = &H18F
Dim KeySection As String
Dim KeyKey As String
Dim KeyValue As String
Private Const sys_Add = 0
Private Const sys_Delete = 2
Sub get_totals()
    For X = 0 To Mainlst.ListCount
        dwp$ = RTrim(Mid(Mainlst.List(X), 16, 25))
        If dwp$ = "Deposit To Checkings" Then
            d2c = d2c + Val(Mid(Mainlst.List(X), 45, 8))
        ElseIf dwp$ = "Deposit To Savings" Then
            d2s = d2s + Val(Mid(Mainlst.List(X), 45, 8))
        ElseIf dwp$ = "Withdrawal From Checkings" Then
            w2c = w2c + Val(Mid(Mainlst.List(X), 45, 8))
        ElseIf dwp$ = "Withdrawal From Savings" Then
            w2s = w2s + Val(Mid(Mainlst.List(X), 45, 8))
        ElseIf dwp$ = "Savings To Checkings" Then
            s2c = s2c + Val(Mid(Mainlst.List(X), 45, 8))
        ElseIf dwp$ = "Checkings To Savings" Then
            c2s = c2s + Val(Mid(Mainlst.List(X), 45, 8))
        ElseIf dwp$ = "+ Checking Adjustment" Then
            pca = pca + Val(Mid(Mainlst.List(X), 45, 8))
        ElseIf dwp$ = "- Checking Adjustment" Then
            nca = nca + Val(Mid(Mainlst.List(X), 45, 8))
        ElseIf dwp$ = "+ Savings Adjustment" Then
            psa = psa + Val(Mid(Mainlst.List(X), 45, 8))
        ElseIf dwp$ = "- Savings Adjustment" Then
            nsa = nsa + Val(Mid(Mainlst.List(X), 45, 8))
        Else
        End If
    Next X
    total_checking = (d2c + s2c + pca) - (w2c + c2s + nca)
    total_savings = (d2s + c2s + psa) - (w2s + s2c + nsa)
    total_checking = Format(total_checking, "#.#0")
    total_savings = Format(total_savings, "#.#0")
    StatusBar1.Panels.Item(4).Text = "C  (" & total_checking & ")"
    StatusBar1.Panels.Item(5).Text = "S  (" & total_savings & ")"
End Sub
Private Sub loadini()
    Dim lngResult As Long
    Dim strFileName
    Dim strResult As String * 50
    strFileName = App.Path & "\default.ini"
    lngResult = GetPrivateProfileString(KeySection, KeyKey, strFileName, strResult, Len(strResult), strFileName)
    KeyValue = Trim(strResult)
End Sub
Private Sub saveini()
    Dim lngResult As Long
    Dim strFileName
    strFileName = App.Path & "\default.ini"
    lngResult = WritePrivateProfileString(KeySection, KeyKey, KeyValue, strFileName)
End Sub
Sub loadlist()
    List1.Visible = True
    Mainfrm.can_we_edit = "Y"
    menutransall.Checked = True
    Mainlst.Clear
    List2.Clear
    List3.Clear
    If is_form_loading = "Y" Then
        KeySection = "Default"
        KeyKey = "File"
        loadini
        If Trim(KeyValue) = "" Then
            KeySection = "Default"
            KeyKey = "File"
            KeyValue = App.Path & "\data.nfs"
            CommonDialog1.filename = KeyValue
            saveini
        End If
        CommonDialog1.filename = KeyValue
        If FileExists(CommonDialog1.filename) Then
            CommonDialog1.filename = CommonDialog1.filename
        Else
            Mainlst.Clear
            Call savelist
            Call loadlist
            KeySection = "Default"
            KeyKey = "File"
            KeyValue = App.Path & "\data.nfs"
            CommonDialog1.filename = KeyValue
            saveini
        End If
    End If
    Mainfrm.Caption = LCase(CommonDialog1.filename)
    Call check_record
    Dim Z As String
    Open CommonDialog1.filename For Input As #1
    Do While Not EOF(1): DoEvents
        Line Input #1, Z
        If Z <> "" Then
            Mainlst.AddItem Z
            f = f + 1
            DoEvents
            ProgressBar1.Value = (f * 100) / (FileLen(CommonDialog1.filename) / 108)
        End If
    Loop
    Close #1
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    StatusBar1.Panels.Item(3).Text = FileLen(CommonDialog1.filename) & " Bytes"
    Call get_totals
    Mainfrm.Show
    is_form_loading = "N"
    List1.Visible = False
    ProgressBar1.Value = 0
End Sub
Sub savelist()
    Mainfrm.Mainlst.Enabled = False
    Dim X
    Mainfrm.can_we_edit = "Y"
    menutransall.Checked = True
    CommonDialog1.filename = LCase(CommonDialog1.filename)
    Open CommonDialog1.filename For Output As #1
    Mainfrm.Caption = CommonDialog1.filename
    For X = 0 To Mainlst.ListCount - 1
        Print #1, Mainlst.List(X) + Chr(13)
        f = f + 1
        ProgressBar1.Value = (f * 100) / Mainfrm.Mainlst.ListCount
        DoEvents
    Next X
    Close #1
    DoEvents
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    StatusBar1.Panels.Item(3).Text = FileLen(CommonDialog1.filename) & " Bytes"
    Call get_totals
    Mainfrm.Mainlst.Enabled = True
    ProgressBar1.Value = 0
End Sub
Private Sub adjneg_Click()
    List1.Visible = True
    adjneg.Checked = True
    tychk.Checked = False
    menutransall.Checked = False
    tysv.Checked = False
    wtsv.Checked = False
    wsv.Checked = False
    chk2sv.Checked = False
    sv2chk.Checked = False
    adjpos.Checked = False
    posadjsav.Checked = False
    negadjsav.Checked = False
    getlc = FileLen(CommonDialog1.filename) / 108
    Mainlst.Clear
    Dim X As String
    Open (CommonDialog1.filename) For Input As #1
    Do While Not EOF(1): DoEvents
        Line Input #1, X
        If X <> "" Then
            dwp$ = Trim(Mid(X, 16, 25))
            f = f + 1
            ProgressBar1.Value = (f * 100) / getlc
            If dwp$ = "- Checking Adjustment" Then
                Mainlst.AddItem X
            End If
        End If
    Loop
    Close #1
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    List1.Visible = False
    ProgressBar1.Value = 0
End Sub
Private Sub adjpos_Click()
    List1.Visible = True
    adjpos.Checked = True
    tychk.Checked = False
    menutransall.Checked = False
    tysv.Checked = False
    wtsv.Checked = False
    wsv.Checked = False
    chk2sv.Checked = False
    sv2chk.Checked = False
    adjneg.Checked = False
    posadjsav.Checked = False
    negadjsav.Checked = False
    getlc = FileLen(CommonDialog1.filename) / 108
    Mainlst.Clear
    Dim X As String
    Open (CommonDialog1.filename) For Input As #1
    Do While Not EOF(1): DoEvents
        Line Input #1, X
        If X <> "" Then
            dwp$ = Trim(Mid(X, 16, 25))
            f = f + 1
            ProgressBar1.Value = (f * 100) / getlc
            If dwp$ = "+ Checking Adjustment" Then
                Mainlst.AddItem X
            End If
        End If
    Loop
    Close #1
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    List1.Visible = False
    ProgressBar1.Value = 0
End Sub
Private Sub bckup_Click()
    Dim X
    Mainfrm.can_we_edit = "Y"
    If menutransall.Checked <> True Then
        Call loadlist
    End If
    Open (App.Path & "\databu.nfs") For Output As #1
    For X = 0 To Mainlst.ListCount - 1
        Print #1, Mainlst.List(X) + Chr(13)
        f = f + 1
        ProgressBar1.Value = (f * 100) / Mainfrm.Mainlst.ListCount
        DoEvents
    Next X
    Close #1
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    StatusBar1.Panels.Item(3).Text = FileLen(CommonDialog1.filename) & " Bytes"
    Call get_totals
    ProgressBar1.Value = 0
End Sub
Private Sub calendar_Click()
    Calendarfrm.Show
End Sub
Private Sub chk2sv_Click()
    List1.Visible = True
    chk2sv.Checked = True
    tychk.Checked = False
    menutransall.Checked = False
    tysv.Checked = False
    wtsv.Checked = False
    wsv.Checked = False
    sv2chk.Checked = False
    adjpos.Checked = False
    adjneg.Checked = False
    posadjsav.Checked = False
    negadjsav.Checked = False
    Mainfrm.can_we_edit = "N"
    getlc = FileLen(CommonDialog1.filename) / 108
    Mainlst.Clear
    Dim X As String
    Open (CommonDialog1.filename) For Input As #1
    Do While Not EOF(1): DoEvents
        Line Input #1, X
        If X <> "" Then
            dwp$ = Trim(Mid(X, 16, 25))
            f = f + 1
            ProgressBar1.Value = (f * 100) / getlc
            If dwp$ = "Checkings To Savings" Then
                Mainlst.AddItem X
            End If
        End If
    Loop
    Close #1
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    List1.Visible = False
    ProgressBar1.Value = 0
End Sub
Private Sub datapad_Click()
    Datapadfrm.Show
End Sub
Private Sub default_Click()
    On Error GoTo Defaulterror:
    CommonDialog1.Filter = "Data Files (*.nfs)|*.nfs"
    CommonDialog1.DialogTitle = "Default File To Load"
    CommonDialog1.ShowOpen
    KeySection = "Default"
    KeyKey = "File"
    KeyValue = CommonDialog1.filename
    saveini
    CommonDialog1.filename = Mainfrm.Caption
Defaulterror:
End Sub
Private Sub details_Click()
    If can_we_edit = "N" Then
        Call loadlist
    End If
    If Mainlst.ListCount <> 0 Then
        Call clear_data
        For X = 0 To Mainlst.ListCount
            DoEvents
            ProgressBar1.Value = (X * 100) / Mainlst.ListCount
            dwp$ = RTrim(Mid(Mainlst.List(X), 16, 25))
            If dwp$ = "Deposit To Checkings" Then
                d2c = d2c + Val(Mid(Mainlst.List(X), 45, 8))
                Printfrm.text1.Caption = Val(Printfrm.text1) + 1
                Printfrm.text11.Caption = d2c
            ElseIf dwp$ = "Deposit To Savings" Then
                d2s = d2s + Val(Mid(Mainlst.List(X), 45, 8))
                Printfrm.text2.Caption = Val(Printfrm.text2) + 1
                Printfrm.text12.Caption = d2s
            ElseIf dwp$ = "Withdrawal From Checkings" Then
                w2c = w2c + Val(Mid(Mainlst.List(X), 45, 8))
                Printfrm.text3.Caption = Val(Printfrm.text3) + 1
                Printfrm.text13.Caption = w2c
            ElseIf dwp$ = "Withdrawal From Savings" Then
                w2s = w2s + Val(Mid(Mainlst.List(X), 45, 8))
                Printfrm.text4.Caption = Val(Printfrm.text4) + 1
                Printfrm.text14.Caption = w2s
            ElseIf dwp$ = "Savings To Checkings" Then
                s2c = s2c + Val(Mid(Mainlst.List(X), 45, 8))
                Printfrm.text5.Caption = Val(Printfrm.text5) + 1
                Printfrm.text15.Caption = s2c
            ElseIf dwp$ = "Checkings To Savings" Then
                c2s = c2s + Val(Mid(Mainlst.List(X), 45, 8))
                Printfrm.text6.Caption = Val(Printfrm.text6) + 1
                Printfrm.text16.Caption = c2s
            ElseIf dwp$ = "+ Checking Adjustment" Then
                pca = pca + Val(Mid(Mainlst.List(X), 45, 8))
                Printfrm.text7.Caption = Val(Printfrm.text7) + 1
                Printfrm.text17.Caption = pca
            ElseIf dwp$ = "- Checking Adjustment" Then
                nca = nca + Val(Mid(Mainlst.List(X), 45, 8))
                Printfrm.text8.Caption = Val(Printfrm.text8) + 1
                Printfrm.text18.Caption = nca
            ElseIf dwp$ = "+ Savings Adjustment" Then
                psa = psa + Val(Mid(Mainlst.List(X), 45, 8))
                Printfrm.text9.Caption = Val(Printfrm.text9) + 1
                Printfrm.text19.Caption = psa
            ElseIf dwp$ = "- Savings Adjustment" Then
                nsa = nsa + Val(Mid(Mainlst.List(X), 45, 8))
                Printfrm.text10.Caption = Val(Printfrm.text10) + 1
                Printfrm.text20.Caption = nsa
            Else
            End If
            Printfrm.text11.Caption = Format(Printfrm.text11.Caption, "##,###.#0")
            Printfrm.text12.Caption = Format(Printfrm.text12.Caption, "##,###.#0")
            Printfrm.text13.Caption = Format(Printfrm.text13.Caption, "##,###.#0")
            Printfrm.text14.Caption = Format(Printfrm.text14.Caption, "##,###.#0")
            Printfrm.text15.Caption = Format(Printfrm.text15.Caption, "##,###.#0")
            Printfrm.text16.Caption = Format(Printfrm.text16.Caption, "##,###.#0")
            Printfrm.text17.Caption = Format(Printfrm.text17.Caption, "##,###.#0")
            Printfrm.text18.Caption = Format(Printfrm.text18.Caption, "##,###.#0")
            Printfrm.text19.Caption = Format(Printfrm.text19.Caption, "##,###.#0")
            Printfrm.text20.Caption = Format(Printfrm.text20.Caption, "##,###.#0")
        Next X
        ProgressBar1.Value = 0
        center Printfrm
        DoEvents
        SetFormTopmost Printfrm
        Printfrm.Height = 0
        For Y = 1 To 2950
            Printfrm.Height = Printfrm.Height + 1
            center Printfrm
            DoEvents
            If Printfrm.Visible <> True Then
                Printfrm.Visible = True
            End If
        Next Y
    End If
End Sub
Private Sub exme_Click()
    If Systemtrayfrm.SystemTray1.IsIconLoaded = True Then
        Systemtrayfrm.SystemTray1.Action = sys_Delete
    End If
    End
End Sub
Private Sub filedel_Click()
    Response = MsgBox("Delete " & CommonDialog1.filename & "?", vbYesNo + vbCritical, "Warning")
    If Response = vbYes Then
        Kill CommonDialog1.filename
        Mainlst.Clear
        Mainfrm.Caption = ""
    Else
    End If
End Sub
Private Sub filenew_Click()
    On Error GoTo NewError:
    CommonDialog1.filename = ""
    CommonDialog1.Filter = "Data Files (*.nfs)|*.nfs"
    CommonDialog1.DialogTitle = "Create New Data File"
    CommonDialog1.ShowSave
    If CommonDialog1.filename <> "" Then
        If FileExists(CommonDialog1.filename) Then
            RETVAL = MsgBox("File: (" & CommonDialog1.filename & ") already exists." & Chr(13) & "Would you like to overwrite this file?", vbYesNo + vbInformation, "Warning")
            If RETVAL = vbYes Then
                Mainlst.Clear
                Call savelist
            Else
            End If
        Else
            Mainlst.Clear
            Call savelist
            Call loadlist
        End If
    ElseIf CommonDialog1.filename = "" Then
        CommonDialog1.filename = Mainfrm.Caption
    End If
NewError:
End Sub
Private Sub fileopen_Click()
    On Error GoTo OpenError:
    CommonDialog1.Filter = "Data Files (*.nfs)|*.nfs"
    CommonDialog1.DialogTitle = "Open Data File"
    CommonDialog1.ShowOpen
    Call loadlist
OpenError:
End Sub
Private Sub filesave_Click()
    If can_we_edit = "Y" Then
        Call savelist
    End If
End Sub
Private Sub Form_Load()
    If App.PrevInstance = True Then
        End
    End If
    is_form_loading = "Y"
    center Me
    Me.Show
    Call loadlist
End Sub
Sub clear_data()
    Printfrm.text1.Caption = "0"
    Printfrm.text2.Caption = "0"
    Printfrm.text3.Caption = "0"
    Printfrm.text4.Caption = "0"
    Printfrm.text5.Caption = "0"
    Printfrm.text6.Caption = "0"
    Printfrm.text7.Caption = "0"
    Printfrm.text8.Caption = "0"
    Printfrm.text9.Caption = "0"
    Printfrm.text10.Caption = "0"
    Printfrm.text11.Caption = "0.00"
    Printfrm.text12.Caption = "0.00"
    Printfrm.text13.Caption = "0.00"
    Printfrm.text14.Caption = "0.00"
    Printfrm.text15.Caption = "0.00"
    Printfrm.text16.Caption = "0.00"
    Printfrm.text17.Caption = "0.00"
    Printfrm.text18.Caption = "0.00"
    Printfrm.text19.Caption = "0.00"
    Printfrm.text20.Caption = "0.00"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Load Systemtrayfrm
    Mainfrm.Hide
    Unload Mainfrm
End Sub
Private Sub Mainlst_DblClick()
    Mainfrm.can_we_edit = "N"
    Load Editfrm
End Sub
Private Sub Mainlst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mainfrm.can_we_edit = "Y" Then
        If Mainlst.ListIndex >= 0 Then
            If Button And 2 Then
                Mainlst.RemoveItem (Mainlst.ListIndex)
                Call savelist
            Else
            End If
        Else
        End If
    Else
    End If
End Sub
Private Sub menudepchk_Click()
    Dep2Chkfrm.Show
End Sub
Private Sub menudepsave_Click()
    Dep2Savfrm.Show
End Sub
Private Sub menutransall_Click()
    If menutransall.Checked <> True Then
        Call loadlist
    Else
    End If
End Sub
Private Sub menuwithchk_Click()
    WithChkfrm.Show
End Sub
Private Sub menuwithsav_Click()
    WithSavfrm.Show
End Sub
Private Sub negadjsav_Click()
    List1.Visible = True
    negadjsav.Checked = True
    tychk.Checked = False
    menutransall.Checked = False
    tysv.Checked = False
    wtsv.Checked = False
    wsv.Checked = False
    chk2sv.Checked = False
    sv2chk.Checked = False
    adjpos.Checked = False
    adjneg.Checked = False
    posadjsav.Checked = False
    getlc = FileLen(CommonDialog1.filename) / 108
    Mainlst.Clear
    Dim X As String
    Open (CommonDialog1.filename) For Input As #1
    Do While Not EOF(1): DoEvents
        Line Input #1, X
        If X <> "" Then
            dwp$ = Trim(Mid(X, 16, 25))
            f = f + 1
            ProgressBar1.Value = (f * 100) / getlc
            If dwp$ = "- Savings Adjustment" Then
                Mainlst.AddItem X
            End If
        End If
    Loop
    Close #1
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    List1.Visible = False
    ProgressBar1.Value = 0
End Sub
Private Sub negadjuts_Click()
    neg = InputBox("Enter In Adjustment Ammount", "Adjustment")
    neg2 = IsNumeric(neg)
    If neg <> "" And neg2 = True Then
        neg = Format(neg, "####.#0")
        negx = Len(neg)
        If negx < 7 Then
            Do
                neg = " " + neg
            Loop Until Len(neg) = 7
        End If
        Mainfrm.Mainlst.AddItem Format(Date, "mm/dd/yyyy") & "     " & "- Checking Adjustment    " & "     " & neg & "-      " & Format(Date, "mm/dd/yyyy") & "      " & "Adjustment                    ", 0
        Call savelist
    Else
    End If
End Sub
Private Sub posadjsav_Click()
    List1.Visible = True
    posadjsav.Checked = True
    tychk.Checked = False
    menutransall.Checked = False
    tysv.Checked = False
    wtsv.Checked = False
    wsv.Checked = False
    chk2sv.Checked = False
    sv2chk.Checked = False
    adjpos.Checked = False
    adjneg.Checked = False
    negadjsav.Checked = False
    getlc = FileLen(CommonDialog1.filename) / 108
    Mainlst.Clear
    Dim X As String
    Open (CommonDialog1.filename) For Input As #1
    Do While Not EOF(1): DoEvents
        Line Input #1, X
        If X <> "" Then
            dwp$ = Trim(Mid(X, 16, 25))
            f = f + 1
            ProgressBar1.Value = (f * 100) / getlc
            If dwp$ = "+ Savings Adjustment" Then
                Mainlst.AddItem X
            End If
        End If
    Loop
    Close #1
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    List1.Visible = False
    ProgressBar1.Value = 0
End Sub
Private Sub posadjust_Click()
    pos = InputBox("Enter In Adjustment Ammount", "Adjustment")
    pos2 = IsNumeric(pos)
    If pos <> "" And pos2 = True Then
        pos = Format(pos, "####.#0")
        posx = Len(pos)
        If posx < 7 Then
            Do
                pos = " " + pos
            Loop Until Len(pos) = 7
        End If
        Mainfrm.Mainlst.AddItem Format(Date, "mm/dd/yyyy") & "     " & "+ Checking Adjustment    " & "     " & pos & "+      " & Format(Date, "mm/dd/yyyy") & "      " & "Adjustment                    ", 0
        DoEvents
        Call Mainfrm.savelist
    Else
    End If
End Sub
Private Sub posadjust2_Click()
    pos = InputBox("Enter In Adjustment Ammount", "Adjustment")
    pos2 = IsNumeric(pos)
    If pos <> "" And pos2 = True Then
        pos = Format(pos, "####.#0")
        posx = Len(pos)
        If posx < 7 Then
            Do
                pos = " " + pos
            Loop Until Len(pos) = 7
        End If
        Mainfrm.Mainlst.AddItem Format(Date, "mm/dd/yyyy") & "     " & "+ Savings Adjustment     " & "     " & pos & "+      " & Format(Date, "mm/dd/yyyy") & "      " & "Adjustment                    ", 0
        Call savelist
    Else
    End If
End Sub
Private Sub prnt_Click()
    On Error GoTo PrintError
    Call loadlist
    Printer.Font = "Courier New"
    Printer.FontSize = 8
    Printer.Print "Printed on " & Date & " at " & Time
    Printer.Print
    Printer.Print "Recorded       Transaction                    Ammount      Date            Description"
    Printer.Print "----------     -------------------------      -------      ----------      ------------------------------"
    For Y = 0 To Mainlst.ListCount - 1
        Printer.Print Mainlst.List(Y)
    Next Y
    Printer.EndDoc
PrintError:
End Sub
Private Sub rstr_Click()
    List1.Visible = True
    Mainfrm.can_we_edit = "Y"
    menutransall.Checked = True
    fcnt = FileLen(App.Path & "\databu.nfs") / 108
    Mainlst.Clear
    Dim Z As String
    Open (App.Path & "\databu.nfs") For Input As #1
    Do While Not EOF(1): DoEvents
        Line Input #1, Z
        If Z <> "" Then
            Mainlst.AddItem Z
            f = f + 1
            DoEvents
            ProgressBar1.Value = (f * 100) / fcnt
        End If
    Loop
    Close #1
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    StatusBar1.Panels.Item(3).Text = FileLen(CommonDialog1.filename) & " Bytes"
    Call get_totals
    Mainfrm.Show
    List1.Visible = False
    Call savelist
End Sub
Private Sub saveas_Click()
    On Error GoTo SaveErr
    CommonDialog1.Filter = "Data Files (*.nfs)|*.nfs"
    CommonDialog1.DialogTitle = "Save Data File"
    CommonDialog1.ShowSave
    Call savelist
    Mainfrm.Caption = CommonDialog1.filename
SaveErr:
End Sub
Private Sub savneg_Click()
    neg = InputBox("Enter In Adjustment Ammount", "Adjustment")
    neg2 = IsNumeric(neg)
    If neg <> "" And neg2 = True Then
        neg = Format(neg, "####.#0")
        negx = Len(neg)
        If negx < 7 Then
            Do
                neg = " " + neg
            Loop Until Len(neg) = 7
        End If
        Mainfrm.Mainlst.AddItem Format(Date, "mm/dd/yyyy") & "     " & "- Savings Adjustment     " & "     " & neg & "-      " & Format(Date, "mm/dd/yyyy") & "      " & "Adjustment                    ", 0
        Call Mainfrm.savelist
    Else
    End If
End Sub
Private Sub srch_Click()
    If menutransall.Checked <> True Then
        Call loadlist
    End If
    src$ = InputBox("Enter String To Search For", "Search")
    If src$ <> "" Then
        menutransall.Checked = False
        Mainfrm.can_we_edit = "N"
        List1.Visible = True
        For X = 0 To Mainlst.ListCount
            srchcnt = InStr(1, Mainlst.List(X), src$, 1)
            DoEvents
            If Mainlst.ListCount <> 0 Then
                ProgressBar1.Value = (X * 100) / Mainlst.ListCount
            End If
            If srchcnt <> 0 Then
                List2.AddItem (Mainlst.List(X))
            End If
        Next X
        Mainlst.Clear
        For Y = 0 To (List2.ListCount - 1)
            Mainlst.AddItem List2.List(Y)
        Next Y
        List1.Visible = False
        List2.Clear
        StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    End If
    tychk.Checked = False
    menutransall.Checked = False
    tysv.Checked = False
    wtsv.Checked = False
    wsv.Checked = False
    chk2sv.Checked = False
    sv2chk.Checked = False
    adjpos.Checked = False
    adjneg.Checked = False
    posadjsav.Checked = False
    negadjsav.Checked = False
    ProgressBar1.Value = 0
End Sub
Private Sub sv2chk_Click()
    List1.Visible = True
    sv2chk.Checked = True
    tychk.Checked = False
    menutransall.Checked = False
    tysv.Checked = False
    wtsv.Checked = False
    wsv.Checked = False
    chk2sv.Checked = False
    adjpos.Checked = False
    adjneg.Checked = False
    posadjsav.Checked = False
    negadjsav.Checked = False
    Mainfrm.can_we_edit = "N"
    getlc = FileLen(CommonDialog1.filename) / 108
    Mainlst.Clear
    Dim X As String
    Open (CommonDialog1.filename) For Input As #1
    Do While Not EOF(1): DoEvents
        Line Input #1, X
        If X <> "" Then
            dwp$ = Trim(Mid(X, 16, 25))
            f = f + 1
            ProgressBar1.Value = (f * 100) / getlc
            If dwp$ = "Savings To Checkings" Then
                Mainlst.AddItem X
            End If
        End If
    Loop
    Close #1
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    List1.Visible = False
    ProgressBar1.Value = 0
End Sub
Private Sub transfer_Click()
    Transferfrm.Show
End Sub
Private Sub tychk_Click()
    List1.Visible = True
    tychk.Checked = True
    menutransall.Checked = False
    tysv.Checked = False
    wtsv.Checked = False
    wsv.Checked = False
    chk2sv.Checked = False
    sv2chk.Checked = False
    adjpos.Checked = False
    adjneg.Checked = False
    posadjsav.Checked = False
    negadjsav.Checked = False
    Mainfrm.can_we_edit = "N"
    getlc = FileLen(CommonDialog1.filename) / 108
    Mainlst.Clear
    Dim X As String
    Open (CommonDialog1.filename) For Input As #1
    Do While Not EOF(1): DoEvents
        Line Input #1, X
        If X <> "" Then
            dwp$ = Trim(Mid(X, 16, 25))
            f = f + 1
            ProgressBar1.Value = (f * 100) / getlc
            If dwp$ = "Deposit To Checkings" Then
                Mainlst.AddItem X
            End If
        End If
    Loop
    Close #1
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    List1.Visible = False
    ProgressBar1.Value = 0
End Sub
Private Sub tysv_Click()
    List1.Visible = True
    tysv.Checked = True
    tychk.Checked = False
    menutransall.Checked = False
    wtsv.Checked = False
    wsv.Checked = False
    chk2sv.Checked = False
    sv2chk.Checked = False
    adjpos.Checked = False
    adjneg.Checked = False
    posadjsav.Checked = False
    negadjsav.Checked = False
    Mainfrm.can_we_edit = "N"
    getlc = FileLen(CommonDialog1.filename) / 108
    Mainlst.Clear
    Dim X As String
    Open (CommonDialog1.filename) For Input As #1
    Do While Not EOF(1): DoEvents
        Line Input #1, X
        If X <> "" Then
            dwp$ = Trim(Mid(X, 16, 25))
            f = f + 1
            ProgressBar1.Value = (f * 100) / getlc
            If dwp$ = "Deposit To Savings" Then
                Mainlst.AddItem X
            End If
        End If
    Loop
    Close #1
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    List1.Visible = False
    ProgressBar1.Value = 0
End Sub
Private Sub wsv_Click()
    List1.Visible = True
    wsv.Checked = True
    tychk.Checked = False
    menutransall.Checked = False
    tysv.Checked = False
    wtsv.Checked = False
    chk2sv.Checked = False
    sv2chk.Checked = False
    adjpos.Checked = False
    adjneg.Checked = False
    posadjsav.Checked = False
    negadjsav.Checked = False
    Mainfrm.can_we_edit = "N"
    getlc = FileLen(CommonDialog1.filename) / 108
    Mainlst.Clear
    Dim X As String
    Open (CommonDialog1.filename) For Input As #1
    Do While Not EOF(1): DoEvents
        Line Input #1, X
        If X <> "" Then
            dwp$ = Trim(Mid(X, 16, 25))
            f = f + 1
            ProgressBar1.Value = (f * 100) / getlc
            If dwp$ = "Withdrawal From Savings" Then
                Mainlst.AddItem X
            End If
        End If
    Loop
    Close #1
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    List1.Visible = False
    ProgressBar1.Value = 0
End Sub
Private Sub wtsv_Click()
    List1.Visible = True
    wtsv.Checked = True
    tychk.Checked = False
    menutransall.Checked = False
    tysv.Checked = False
    tychk.Checked = False
    wsv.Checked = False
    chk2sv.Checked = False
    sv2chk.Checked = False
    adjpos.Checked = False
    adjneg.Checked = False
    posadjsav.Checked = False
    negadjsav.Checked = False
    Mainfrm.can_we_edit = "N"
    getlc = FileLen(CommonDialog1.filename) / 108
    Mainlst.Clear
    Dim X As String
    Open (CommonDialog1.filename) For Input As #1
    Do While Not EOF(1): DoEvents
        Line Input #1, X
        If X <> "" Then
            dwp$ = Trim(Mid(X, 16, 25))
            f = f + 1
            ProgressBar1.Value = (f * 100) / getlc
            If dwp$ = "Withdrawal From Checkings" Then
                Mainlst.AddItem X
            End If
        End If
    Loop
    Close #1
    StatusBar1.Panels.Item(2).Text = Mainlst.ListCount & " Entries"
    List1.Visible = False
    ProgressBar1.Value = 0
End Sub
Sub check_record()
    lenchk = FileLen(CommonDialog1.filename)
    fcnt = lenchk / 108
    chklen = lenchk Mod 108
    If chklen <> 0 Then
        Open CommonDialog1.filename For Input As #3
        Do While Not EOF(3): DoEvents
            Line Input #3, K
            If K <> "" Then
                List3.AddItem K
            End If
        Loop
        Close #3
        For cc = 0 To List3.ListCount - 1
            If Len(List3.List(cc)) <> 105 Then
                a1 = Left((List3.List(cc)), 10)
                If Len(a1) < 10 Then
                    Do
                        a1 = a1 + " "
                    Loop Until Len(a1) = 10
                End If
                a2 = Mid((List3.List(cc)), 16, 25)
                If Len(a2) < 25 Then
                    Do
                        a2 = a2 + " "
                    Loop Until Len(a2) = 25
                End If
                a3 = Mid((List3.List(cc)), 46, 8)
                If Len(a3) < 8 Then
                    Do
                        a3 = " " + a3
                    Loop Until Len(a3) = 8
                End If
                a4 = Mid((List3.List(cc)), 60, 10)
                If Len(a4) < 10 Then
                    Do
                        a4 = a4 + " "
                    Loop Until Len(a4) = 10
                End If
                a5 = Mid((List3.List(cc)), 76, 30)
                If Len(a5) < 30 Then
                    Do
                        a5 = a5 + " "
                    Loop Until Len(a5) = 30
                End If
                List2.AddItem a1 & "     " & a2 & "     " & a3 & "      " & a4 & "      " & a5
            Else
                List2.AddItem (List3.List(cc))
            End If
        Next cc
        Open CommonDialog1.filename For Output As #4
        For ff = 0 To List2.ListCount - 1
            Print #4, List2.List(ff) + Chr(13)
        Next ff
        Close #4
    End If
End Sub


