Attribute VB_Name = "Mod_Resources"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE


Public CTN As String
Public uCTN As Variant

Public uLeft As Variant
Public uTop As Variant
Public uTopMost As Variant
Public uStartup As Variant
Public uClockMASTrans As Variant

Public CPLLoaded As Boolean

Public Sub FormDrag(TheForm As Form)
   ReleaseCapture
   SendMessage TheForm.hwnd, &HA1, 2, 0&
End Sub
Public Sub MakeFormTop(hwnd As Long, Action As Boolean)
If Action = True Then
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
Else
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End If
End Sub

Public Sub OpenSetting()
  On Error GoTo There_Is_No_File
  'Open App.Path & "\Setting.dat" For Input Access Read As #1
    Line Input #1, uCTN
    Line Input #1, uLeft
    Line Input #1, uTop
    Line Input #1, uTopMost
    Line Input #1, uStartup
    Line Input #1, uClockMASTrans
  Close #1
  GoTo The_File_Exist
  
There_Is_No_File:
uCTN = "System"
uLeft = Screen.Width / 2 - 1500
uTop = Screen.Height / 2 - 1500
uTopMost = False
uStartup = False
uClockMASTrans = 255

The_File_Exist:
CTN = uCTN
End Sub

'Public Sub SaveSetting()
'On Error Resume Next
  'Open App.Path & "\Setting.dat" For Output As #1
   ' Print #1, uCTN
 '   Print #1, uLeft
  '  Print #1, uTop
   ' Print #1, uTopMost
    'Print #1, uStartup
    'Print #1, uClockMASTrans
  'Close #1
'End Sub

Sub Main()
'OpenSetting
CreateAboutFiles
Splash.Show
End Sub

Public Sub SetStartUp(Action As Boolean)
  If Action = True Then
   ' SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "7-Clock", "<NonRun>"
    'SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "7-Clock", App.Path & "\" & App.EXEName & ".exe"
  Else
    'SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "7-Clock", "<NonRun>"
  End If
End Sub

Private Sub CreateAboutFiles()
Dim FileRes As Integer
Dim Buffer() As Byte
On Error Resume Next
    Buffer = LoadResData("About", "7-Clock")
    FileRes = FreeFile
    Open App.Path & "\SPLASH\About.png" For Binary Access Write As #FileRes
    Put #FileRes, , Buffer
    Close #FileRes

    Buffer = LoadResData("SPLASH", "EXPLODING MARBLES!")
    FileRes = FreeFile
    Open App.Path & "\SPLASH\Splash.png" For Binary Access Write As #FileRes
    Put #FileRes, , Buffer
    Close #FileRes

End Sub

