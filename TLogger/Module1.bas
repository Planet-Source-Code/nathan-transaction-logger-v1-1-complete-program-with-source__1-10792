Attribute VB_Name = "Module1"
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H8
Public Const Flame_Height = 30
Type pix
    r As Integer
    g As Integer
    b As Integer
    c As Boolean
End Type
Public maxx As Integer
Public maxy As Integer
Public new_flame() As pix
Public old_flame() As pix
Public Sub SetFormTopmost(TheForm As Form)
    SetWindowPos TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
            SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE
End Sub
Sub timeout(duration)
    DoEvents
    starttime = Timer
    Do While Timer - starttime < duration
        DoEvents
    Loop
End Sub
Sub center(frmform As Form)
    frmform.Left = (Screen.Width - frmform.Width) / 2
    frmform.Top = (Screen.Height - frmform.Height) / 2
End Sub
Function FileExists(filename As String) As Integer
    On Error Resume Next
    X% = Len(Dir$(filename))
    If Err Or X% = 0 Then FileExists = False Else FileExists = True
End Function



