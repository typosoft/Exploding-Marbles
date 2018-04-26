Attribute VB_Name = "modStyles"
Option Explicit

' Enable Visual Styles - requires manifest file

' Optionally, you can open the manifest file in notepad and update the
' version details for your application:
'                             VB .  VB .  0  .  VB
' Four-part version format: major.minor.build.revision
' Each part separated by periods can be 0-65535 inclusive.

' Add the manifest to your application's resource file as follows:
'  1  24  "YourApp.exe.manifest"
' When you add this entry to the resource you must add it as one line.

' Alternatively, you can place the XML manifest file in the same directory
' as your application's executable file. The operating system will load the
' manifest from the file system first, and then check the resource section
' of the executable. The file system version takes precedence.

Private Declare Function InitCommonControlsEx Lib "comctl32" (pInitCtrls As tInitCommonControlsEx) As Long
Private Type tInitCommonControlsEx
   dwSize As Long
   dwICC As Long
End Type

' Loads Common Controls v6 to enable Visual Styles
Private Const ICC_USEREX_CLASSES = &H200&

' These APIs prevent shutdown crashes - thanks to Amer Tahir
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private m_hMod As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExW" (pVerInfo As OSVERSIONINFO) As Long

Private Const OFS_MAXPATHNAME = 128&
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumberLo As Integer
    dwBuildNumberHi As Integer
    dwPlatformId As Long
    szCSDVersion(1 To OFS_MAXPATHNAME) As Integer
End Type
Private osvi As OSVERSIONINFO

Private mHasStyles As Boolean

' Is running on Windows XP or higher
Public Function HasStyles() As Boolean
   If osvi.dwOSVersionInfoSize = 0& Then
      osvi.dwOSVersionInfoSize = LenB(osvi)
      If GetVersionEx(osvi) Then
         mHasStyles = (osvi.dwMajorVersion + osvi.dwMinorVersion) > 5&
   End If: End If
   HasStyles = mHasStyles
End Function

' Initialize Styles
Public Sub InitStyles()
   On Error GoTo Abort
   Dim iccex As tInitCommonControlsEx
   If HasStyles Then
      m_hMod = LoadLibrary(StrPtr("shell32.dll"))
      iccex.dwSize = LenB(iccex)
      iccex.dwICC = ICC_USEREX_CLASSES
      InitCommonControlsEx iccex
   End If
Abort:
End Sub

Public Sub TermStyles()
   If m_hMod Then FreeLibrary m_hMod
   m_hMod = 0&
End Sub
