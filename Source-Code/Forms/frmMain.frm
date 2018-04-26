VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "       Michael Hardy's Exploding Marbles! - A Great 3D ™ Marbles Game!"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9615
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   511
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ****************************************************************
'
'                  Exploding Marbles
'                  Version 2.0 - 4.0
'                 Commercial Edition...
'               Created By Michael Hardy
'
'    Special Thanks To Stacie Hardy, Zoe Hardy, Dylan Plymale,
'    Sher Hardy, Paul Eldridge, Birdie Eldridge, Robert and Norma
'    Plymale, Microsoft, IMac, Linux, Mozilla and Everyone Else
'    Who Supported This Development...
'           -  I - L O V E - Y O U - A L L ! -
'
'    This Computer Game is Dedicated To My Dad (James Hardy)
'    Who Passed Away on January 22nd, 2008 and To My Uncle
'    (James (Bo) Eldridge) Who Passed Away On August 13th, 2008
'             I Love and Miss You Both Very Much...
'
'    Exploding Marbles is Released Under The EULA
'    License Agreement (EULA) and is distributed by
'    Michael Hardy and ® Hardy Creations Inc.
'
'    YOU MAY NOT TAKE CREDIT FOR THE MAKING OF THIS GAME NOR
'    MAY YOU UPLOAD THIS GAME TO A BBS ARCHIEVE...
'
'    YOU CANNOT SALE THIS GAME OR IT'S SOURCE CODE AT ANY TIME...

'    THIS GAME IS COMMERCIAL ALL DATA FILES, DOCUMENTATION
'    i.e GRAPHICS, SOUND EFFECTS, MUSIC AND ETC ARE COPYRIGHTED
'    BY MICHAEL JAMES HARDY AND MAY NOT BE USED WITHOUT HIS WRITTEN
'    PERMISSION... ANY VIOLATION OF THE LICENSE AND TERMS
'    WILL RESULT IN TERMINATION OF THE LICENSE AGREEMENT AND
'    CRIMINAL AND / OR CIVIL SETTINGS MAY APPLY...
'*****************************************************************

Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Const ICC_USEREX_CLASSES = &H200
Private Declare Function GetSystemMenu Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal bRevert As Long) As Long

Private Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, _
     ByVal nPosition As Long, _
     ByVal wFlags As Long) As Long
     
Private Const MF_BYPOSITION = &H400&
'Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFilename) As Long
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

'Private Declare Function SHBrowseForFolder Lib "shell32" _
 '                            (lpbi As BrowseInfo) As Long

'Private Declare Function SHGetPathFromIDList Lib "shell32" _
 '                            (ByVal pidList As Long, _
  '                           ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                             (ByVal lpString1 As String, ByVal _
                             lpString2 As String) As Long
'API DECLARATIONS
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
                 "GetPrivateProfileStringA" (ByVal lpApplicationName _
                 As String, ByVal lpKeyName As Any, ByVal lpDefault _
                 As String, ByVal lpReturnedString As String, ByVal _
                 nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
                 "WritePrivateProfileStringA" (ByVal lpApplicationName _
                 As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
                 ByVal lpFileName As String) As Long
 Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFilename) As Long


Private Declare Function SHBrowseForFolder Lib "shell32" _
                             (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                             (ByVal pidList As Long, _
                             ByVal lpBuffer As String) As Long

                
Dim bEnd As Byte 'This variable controls the flow within the game
                 'an 1 means end program as fast as possible.

Dim sLogo As New pgeSprite
Dim sBar As New pgeSprite
Dim sField As New pgeSprite
Dim sCursor As New pgeSprite
Dim sGameOver As New pgeSprite
Dim sCredits As New pgeSprite
Dim sCredits_Detail As New pgeSprite
Dim sCreditsinfotitle As New pgeSprite
Dim sCredits_Info As New pgeSprite
Dim sQualify As New pgeSprite
Dim sText_Menu(4) As New pgeSprite
Dim sText_GameMenu(3) As New pgeSprite
Dim sText_Settings(4) As New pgeSprite
Dim sText_HighScore(1) As New pgeSprite
Dim sMeter(2) As New pgeSprite
Dim sMeterDrag(2) As New pgeSprite
Dim sMarble(7) As New pgeSprite
Dim sSpecials(6) As New pgeSprite
Dim sSelect As New pgeSprite
Dim sHand As New pgeSprite
Dim sSpark() As New pgeSprite
Dim sSmoke() As New pgeSprite
Dim sSnow() As New pgeSprite
Dim sSplat() As New pgeSprite
Dim sStar() As New pgeSprite
Dim sGhoul() As New pgeSprite
Dim sMud() As New pgeSprite
Dim sExplosion() As New pgeSprite
Dim sScores() As New pgeSprite
Dim sRing() As New pgeSprite
Dim NumMud As Long
Dim NumRings As Long
Dim NumGhouls As Long
Dim NumSplats As Long
Dim NumSmoke As Long
Dim NumScores As Long
Dim NumSparks As Long
Dim NumSnow As Long
Dim NumStars As Long
Dim NumExplosions As Long
Dim sText As String
Dim bText As Boolean
'Dim fso As New FileSystemObject
Dim SkinNames(256) As String
'TODO:
Dim NoMoreMovesRes As Boolean
'END TODO:

Private SkinOb As New SKINCRAFTERLib.SCSkin3


Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type
Private Type OpenFilename
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    iFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Function DisableCloseButton(frm As Form) As Boolean

'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu


    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
    
    
    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)

    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)

   'Remove seperator bar
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    
    'Return True if both calls were successful
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)

End Function
Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function
Private Function ShowFileDialog() As String
    Dim ofn As OpenFilename
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = hwnd
    ofn.lpstrFilter = "Exploding Marbles Skin Files (*.emSkin)" & Chr$(0) & "*.emSkin" & Chr$(0) & Chr(0) & Chr(0)
    ofn.lpstrFile = String(256, 0)
    ofn.nMaxFile = 255
    ofn.lpstrTitle = "Open A Cool Skin For Michael's Exploding Marbles..."
    ofn.Flags = &H800000 + &H1000 + &H8 + &H4
    ofn.lpstrDefExt = "emSkin" + Chr(0)
    GetOpenFileName ofn
    If Mid(ofn.lpstrFile, 1, 1) <> Chr(0) Then ShowFileDialog = ofn.lpstrFile
End Function

Private Sub Form_Initialize()
 Dim fileName As String
 Dim sINIFile As String
 Dim x As Long
On Error Resume Next
x = InitCommonControls
InitCommonControlsVB
SkinOb.InitLicenKeys "SKINCRAFTER", "SKINCRAFTER.COM", "support@skincrafter.com", "DEMOSKINCRAFTERLICENCE"
SkinOb.DefineLanguage 2
SkinOb.InitDecoration 1
fileName = GetSetting(App.EXEName, "Skin", "SkinOb", "Skins\\Michael_Hardy's_Creative_Red.emSkin") '"Skins\\Scratch.skf") '"Skins\\CoolDragon.skf")

'If fileName <> "" Then
 '       SkinOb.LoadSkinFromFile fileName
  '      SkinOb.ApplySkin
        
        'SkinOb.UpdateControl 0
        'Else
        
        
SkinOb.LoadSkinFromFile fileName
    SkinOb.ApplySkin
    SkinOb.DoNotDecorate Splash.hwnd
'fileName = GetSetting("MIKESVISTAMARBLES\" + App.EXEName, "Settings", "SkinOb", fileName)
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
'Dim sINIFile As String
  If bText And KeyAscii > 32 Then
    sText = sText & Chr(KeyAscii)
  End If
  Dim fileName As String
  Select Case KeyAscii
  Case vbKeyTab

    'saveSetting "MIKESVISTAMARBLES\" + App.EXEName, "Settings", "SkinOb", fileName
    '
  fileName = ShowFileDialog
'fileName = GetSetting(App.EXEName, "skin", "skinob", "")
    If fileName <> "" Then
        SkinOb.LoadSkinFromFile fileName
        SkinOb.ApplySkin
        If MsgBox("Would You Like To Set This Skin As The Default Skin? The Default Skin Is The Skin That Is Loaded On Game Startup..., ", vbQuestion Or vbYesNo, "Set As The Default Skin?") = vbYes Then
    
        'SkinOb.UpdateControl 0
         Me.Top = (Screen.Height - Me.Height) / 2
 Me.Left = (Screen.Width - Me.Width) / 2
 Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen

'SaveSetting "Exploding Marbles", "Skin", "Style", fileName
SaveSetting App.EXEName, "skin", "SkinOb", fileName




' Save text1's value
'SaveSetting App.EXEName, "textboxes", "text1", text1.Text


    End If
     Me.Top = (Screen.Height - Me.Height) / 2
 Me.Left = (Screen.Width - Me.Width) / 2
 Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen

   ' Else
'MsgBox "The Theme You Have Selected Is Now Your Default Interface Theme...", vbInformation, "The Theme Has Been Set!"
     'Else
     End If
'Case vbKeyX
'SkinOb.RemoveSkin
'Case vbKeyS
'SkinOb.InitLicenKeys "SKINCRAFTER", "SKINCRAFTER.COM", "support@skincrafter.com", "DEMOSKINCRAFTERLICENCE"
'SkinOb.DefineLanguage 2
'SkinOb.InitDecoration 1
'fileName = GetSetting("Exploding Marbles", "Skin", "SkinOb", "Skins\\Scratch.skf")

'If fileName <> "" Then
 '       SkinOb.LoadSkinFromFile fileName
  '      SkinOb.ApplySkin
        
        'SkinOb.UpdateControl 0
        'Else
        
        
'SkinOb.LoadSkinFromFile fileName
 '   SkinOb.ApplySkin
    End Select
End Sub
Private Sub Form_Load()
On Error Resume Next
  frmMain.Visible = False
  
  Dim KeyCode As Integer
  Dim fileName As String
   Me.Top = (Screen.Height - Me.Height) / 2
 Me.Left = (Screen.Width - Me.Width) / 2
 
'fileName = GetSetting(App.EXEName, "skinob", "filename", "")




Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen
Splash.Show
Splash.SetFocus
Call DelayTime(2) 'Delay 2 Sec
frmMain.Visible = True
Splash.SetFocus
Call DelayTime(2) 'Delay 2 Sec
Unload Splash
frmMain.SetFocus
  bFps = False
 ' Me.Caption = "Exploding Marbles! " & "- A Cool 3D Marble Arcade Game..." & " - Loading Game..."
  LatestHigh = -1
  
  If InitEngine And InitGfx And InitSfx Then
  If Not ("Config") Then MkDir ("Config")
  'End If
    LoadScores
    LoadSettings
  '  Me.Caption = "Exploding Marbles! " & "- A Cool 3D Marble Arcade Game..."
    DoTitle
  End If
  
  'bEnd = 1
  'Unload Me
End Sub

Public Function InitEngine() As Boolean
  'this sub inits the graphics & sound engine, as well as setting up some stuff
  On Error GoTo ErrH
  InitEngine = True
  
  'Load form
  Me.Show
  DoEvents
  
  ReDim sSpark(0)
  ReDim sExplosion(0)
  ReDim sScores(0)
  ReDim sSnow(0)
  ReDim sSmoke(0)
  ReDim sStar(0)
  ReDim sSplat(0)
  ReDim sGhoul(0)
  ReDim sRing(0)
  ReDim sMud(0)
  
  AddFontResource App.Path & "\Data\Altea.TTF"
  AddFontResource App.Path & "\Data\lcdmb.ttf"
  
  'Init
  pEngine.Init Me.hwnd, True, , , False
  Set pTextures = TPool
  pKeyboard.Create Me.hwnd
  pSound.Init 44100, 64, 75
  pMouse.Create Me
  
  FontArial.Create ReturnFont("Arial", 10)
  MainFont.Create ReturnFont("Jokerman", 13)
  LedFont.Create ReturnFont("LCDMono", 18)
  
  'Setup
  pKeyboard.SetTimer 0.1
  Randomize Timer
    
  Exit Function
ErrH:
  MsgBox "Could not Initialize Game Engine. Try Installing DirectX 8.1 or Reinstalling Exploding Marbles.", vbCritical, "Oh, Oh!, A Dang Error!"
  InitEngine = False
End Function
'TODO:
Public Function NoMoreMoves() As Boolean
    'If Player.lBombs = 0 Then
        Dim bool As Boolean, CanMove As Boolean
        bool = True
        CanMove = False
        Dim x As Long, y As Long
        x = 0
        y = 0
        Dim l1 As Long, l2 As Long, l3 As Long
            For x = 0 To 7
                For y = 1 To 8
                    l1 = lGrid(x, y).lType
                    If x = 0 And y = 1 Then 'Top Left 1,1
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 0 And y = 2 Then 'Top Left 1,2
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 1, y + 1).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 0 And y = 3 Then 'Top Left 1,3
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                             CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 0 And y > 3 And y < 6 Then 'Left side
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x, y - 1).lType = l1 Or lGrid(x + 2, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 0 And y = 6 Then 'Bottom Left 1,6
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 0 And y = 7 Then 'Bottom Left 1,7
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 0 And y = 8 Then 'Bottom Left 1,8
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 1 And y = 8 Then 'Bottom Left 2,8
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x - 1, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 2 And y = 8 Then 'Bottom Left 3,8
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x + 1, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x - 2, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x - 1, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    'TODO: ElseIf y = 8 And x > 2 And x < 5 Then 'Bottom
                        'If lGrid(x - 2, y).lType = l1 Then
                            'If lGrid(x - 3, y).lType = l1 Or lGrid(x + 1, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                'CanMove = True
                            'End If
                        'End If
                        'If lGrid(x - 1, y).lType = l1 Then
                            'If lGrid(x - 3, y).lType = l1 Or lGrid(x + 2, y).lType = l1 Or lGrid(x - 2, -1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                'CanMove = True
                            'End If
                        'End If
                        'If lGrid(x + 1, y).lType = l1 Then
                            'If lGrid(x + 3, y).lType = l1 Or lGrid(x - 2, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Then
                                'CanMove = True
                            'End If
                        'End If
                        'If lGrid(x, y - 2).lType = l1 Then
                            'If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                'CanMove = True
                            'End If
                        'End If
                        'If lGrid(x, y - 1).lType = l1 Then
                            'If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                'CanMove = True
                            'End If
                        'End If
                        'If lGrid(x - 1, y - 1).lType = l1 Then
                            'If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                'CanMove = True
                            'End If
                        'End If
                        'If lGrid(x + 1, y - 1).lType = l1 Then
                            'If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Then
                                'CanMove = True
                            'End If
                        'End If
                    ElseIf x = 5 And y = 8 Then 'Bottom Right 6,8
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x + 1, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x + 2, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x - 1, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x - 1, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 6 And y = 8 Then 'Bottom Right 7,8
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 1, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 7 And y = 8 Then 'Bottom Right 8,8
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 7 And y = 7 Then 'Bottom Right 8,7
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 7 And y = 6 Then 'Bottom Right 8,6
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 7 And y = 5 Then 'Right 8,5
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 7 And y = 4 Then 'Right 8,4
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 7 And y = 3 Then 'Top Right 8,3
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x - 1, y - 2).lType = l1 Or lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 3, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 7 And y = 2 Then 'Top Right 8,2
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 7 And y = 1 Then 'Top Right 8,1
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 6 And y = 1 Then 'Top Right 7,1
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 5 And y = 1 Then 'Top Right 6,1
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x - 2, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x - 1, y).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 4 And y = 1 Then 'Top 5,1
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x - 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 3 And y = 1 Then 'Top 4,1
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x - 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 2 And y = 1 Then 'Top Left 3,1
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x - 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 1 And y = 1 Then 'Top Left 2,1
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x - 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x + 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 1 And y = 2 Then 'Top Left 2,2
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x + 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 1, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 1 And y = 3 Then 'Top Left 2,3
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 1 And y = 4 Then 'Top Left 2,4
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 1 And y = 5 Then 'Top Left 2,5
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 1 And y = 6 Then 'Top Left 2,6
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 1 And y = 7 Then 'Top Left 2,7
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 2 And y = 7 Then 'Top Left 3,7
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x + 1, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 3 And y = 7 Then 'Top Left 4,7
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 4 And y = 7 Then 'Top Left 5,7
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 5 And y = 7 Then 'Top Left 6,7
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x - 1, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 6 And y = 7 Then 'Top Left 7,7
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 1, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 6 And y = 6 Then 'Top Left 7,6
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 6 And y = 5 Then 'Top Left 7,5
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 6 And y = 4 Then 'Top Left 7,4
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 1, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 6 And y = 3 Then 'Top Left 7,3
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 6 And y = 2 Then 'Top Left 7,2
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 5 And y = 2 Then 'Top Left 6,2
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x - 1, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 4 And y = 2 Then 'Top Left 5,2
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 3 And y = 2 Then 'Top Left 4,2
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 2 And y = 2 Then 'Top Left 3,2
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x + 1, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 2 And y = 3 Then 'Top Left 3,3
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x + 1, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 2 And y = 4 Then 'Top Left 3,4
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x + 1, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 2 And y = 5 Then 'Top Left 3,5
                        'TODO:
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x + 1, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 2 And y = 6 Then 'Top Left 3,6
                        'TODO:
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x + 1, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x + 2, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 3 And y = 6 Then 'Top Left 4,6
                        'TODO:
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 4 And y = 6 Then 'Top Left 5,6
                        'TODO:
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 5 And y = 6 Then 'Top Left 6,6
                        'TODO:
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x - 1, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 5 And y = 5 Then 'Top Left 6,5
                        'TODO:
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x - 1, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 5 And y = 4 Then 'Top Left 6,4
                        'TODO:
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x - 1, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 5 And y = 3 Then 'Top Left 6,3
                        'TODO:
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x - 2, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x - 1, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y - 2).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 4 And y = 3 Then 'Top Left 5,3
                        'TODO:
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x = 3 And y = 3 Then 'Top Left 4,3
                        'TODO:
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y + 2).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 2, y - 1).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 2, y + 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    ElseIf x > 2 And y > 3 And x < 5 And y < 5 Then 'Misc
                        'TODO:
                        If lGrid(x - 2, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y).lType = l1 Then
                            If lGrid(x - 3, y).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x - 2, y + 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x - 2, y).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 2, y).lType = l1 Then
                            If lGrid(x + 3, y).lType = l1 Or lGrid(x - 1, y).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 2).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y - 1).lType = l1 Then
                            If lGrid(x, y - 3).lType = l1 Or lGrid(x - 1, y - 2).lType = l1 Or lGrid(x + 1, y - 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 1).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x, y + 2).lType = l1 Then
                            If lGrid(x, y + 3).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y - 1).lType = l1 Then
                            If lGrid(x - 1, y - 2).lType = l1 Or lGrid(x - 2, y - 1).lType = l1 Or lGrid(x + 1, y - 1).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x - 1, y + 1).lType = l1 Then
                            If lGrid(x - 2, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x - 1, y + 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y - 1).lType = l1 Then
                            If lGrid(x + 1, y - 2).lType = l1 Or lGrid(x + 1, y + 1).lType = l1 Or lGrid(x - 1, y - 1).lType = l1 Or lGrid(x + 2, y - 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                        If lGrid(x + 1, y + 1).lType = l1 Then
                            If lGrid(x + 1, y - 1).lType = l1 Or lGrid(x + 1, y + 2).lType = l1 Or lGrid(x - 1, y + 1).lType = l1 Or lGrid(x + 2, y + 1).lType = l1 Then
                                CanMove = True
                            End If
                        End If
                    End If
                    If CanMove = True Then
                        Exit For
                    End If
                Next
                If CanMove = True Then
                    Exit For
                End If
            Next
        If CanMove = False Then
        NoMoreMoves = True
        'DoEndGame = True
        Else
        NoMoreMoves = False
        End If
    'Else
        'NoMoreMoves = False
    'End If
End Function
'END TODO:
Public Sub DoTitle()
  'This is the main title screen
  Dim T As New pgeTimer
  Dim tRect As RECT 'We setup an RECT at the mouse pointer so we can
                    'use IntersectR to easy check pointer against sprites.
  Dim x As Long, y As Long
  Dim Hold As Boolean

tAgain:
  pSound.MusicPlayID "0", Settings.MusicVolume
  bEnd = 0
  
  'setup graphics needed in this screen.
  If Not Hold Then
    sLogo.SetPosition 900, 100
    sLogo.SetAutoPath 384, 100, 1000
    sLogo.SetColor 255, 255, 255, 0
    sLogo.SetAutoFade 255, 255, 255, 255, 1000
  End If
  Hold = False
  sText_Menu(0).SetPosition 900, 200
  sText_Menu(1).SetPosition 900, 275
  sText_Menu(2).SetPosition 900, 350
  'sText_Menu(3).SetPosition 900, 375
  sText_Menu(3).SetPosition 900, 425
  sText_Menu(0).SetAutoPath 384, 200, 1000
  sText_Menu(1).SetAutoPath 384, 275, 1000
  sText_Menu(2).SetAutoPath 384, 350, 1000
  'sText_Menu(3).SetAutoPath 384, 375, 1000
  sText_Menu(3).SetAutoPath 384, 425, 1000
  For x = 0 To 4
    sText_Menu(x).SetColor 255, 255, 255, 0
    sText_Menu(x).SetAutoFade 255, 255, 255, 255, 1000
  Next
  
  T.StartTime
  Do
    DoEvents
    
    'Poll & setup mouse pointer
    pMouse.Poll Me
    sCursor.SetPosition pMouse.g_cursorx, pMouse.g_cursory
    tRect.Left = pMouse.g_cursorx
    tRect.Top = pMouse.g_cursory
    tRect.Right = tRect.Left + 1
    tRect.bottom = tRect.Top + 1

    'Handle input, but only after menu items have appeared
    If T.GetTime > 1 Then
      'Check mouse vs menu options
      For x = 0 To 3
        If IntersectR(tRect, sText_Menu(x).GetDestRect) Then
          sText_Menu(x).SetColor 255, 255, 255, 100 + Abs(155 * Sine(T.GetTimeMs / 5))
          If pMouse.button1 Then
            Select Case x
              Case 0
                bEnd = 4
              Case 1
                bEnd = 3
              Case 2
                bEnd = 2
              Case 3
                bEnd = 5
            End Select
          End If
        Else
          sText_Menu(x).SetColor 255, 255, 255, 255
         'sText_Menu(x).SetColor 190, 150, 255, 160
        End If
      Next
      'Check keyboard
      If pKeyboard.KeyDown(DIK_ESCAPE) Then bEnd = 1
      If pKeyboard.KeyDown(DIK_F) Then bFps = Not bFps
     ' If pKeyboard.KeyDown(DIK_A) Then Call Credits
     '  If pKeyboard.KeyDown(DIK_F1) Then CreditsAbout
    End If
    
    
    '###Drawing sequence
    pEngine.Clear
    
    sBar.Render
    sField.Render
    sGameOver.Render
    sLogo.Render
    For x = 0 To 3
      sText_Menu(x).Render
    Next
    For x = 0 To 4
      sText_Settings(x).Render
    Next
    For x = 0 To 2
      sMeter(x).Render
      sMeterDrag(x).Render
    Next
    For x = 0 To 1
      sText_HighScore(x).Render
    Next
    For x = 0 To 2
      sText_GameMenu(x).Render
    Next
    
    If Player.lTime > 0 And T.GetTimeMs < 1000 Then
      sHand.SetColor 255, 255, 255, 255 - tob((T.GetTimeMs / 1000) * 255)
      sHand.Render
      LedFont.DrawText Player.lScore, ReturnRECT(20, 34, 115, 64), RGBA(0, 200, 0, 255 - tob((T.GetTimeMs / 1000) * 255)), DT_RIGHT
      y = -1
      For x = 9 To 0 Step -1
        If High(x).lScore > Player.lScore Then
          y = x
          Exit For
        End If
      Next
      If y = -1 Then
        LedFont.DrawText Player.lScore, ReturnRECT(20, 98, 115, 128), RGBA(0, 200, 0, 255) - tob((T.GetTimeMs / 1000) * 255), DT_RIGHT
      Else
        LedFont.DrawText High(y).lScore, ReturnRECT(20, 98, 115, 128), RGBA(0, 200, 0, 255 - tob((T.GetTimeMs / 1000) * 255)), DT_RIGHT
      End If
      LedFont.DrawText Player.lBombs, ReturnRECT(20, 168, 115, 198), RGBA(0, 200, 0, 255 - tob((T.GetTimeMs / 1000) * 255)), DT_RIGHT
    Else
      Player.lTime = 0
    End If
    MainFont.DrawText "Copyright © 2003-2014 Created By ® Michael Hardy", ReturnRECT(128, 480, 640, 512), RGBA(255, 255, 255, tob((T.GetTimeMs / 2000) * 200)), DT_CENTER Or DT_VCENTER
    MainFont.DrawText "To Load and Set A New Skin Press The ""TAB"" Key!", ReturnRECT(128, 450, 640, 480), RGBA(255, 255, 255, tob((T.GetTimeMs / 2000) * 200)), DT_CENTER Or DT_VCENTER
    MainFont.DrawText "Press ""ESC"" During the Game For The Menu!", ReturnRECT(128, 120, 640, 143), RGBA(255, 255, 255, tob((T.GetTimeMs / 2000) * 200)), DT_CENTER Or DT_VCENTER
    
    sCursor.Render
    
    If bFps Then
      FontArial.DrawText pEngine.lFPS, ReturnRECT(0, 0, 100, 15), RGBA(0, 255, 0, 255), DT_LEFT
    End If
    
    pEngine.Render
    '###End of drawing sequence
  Loop Until bEnd
  
  Select Case bEnd
    Case 1
      Unload Me
    Case 2
      Options
      If bEnd = 1 Then
        Unload Me
      Else
        GoTo tAgain
      End If
    Case 3
      Highscore
      If bEnd = 1 Then
        Unload Me
      Else
        GoTo tAgain
      End If
    Case 4
      NewGame
      DoGame
      If bEnd = 1 Then
        Unload Me
      ElseIf bEnd = 4 Then
        Hold = True
        GoTo tAgain
      Else
        GoTo tAgain
      End If
    Case 5
      Credits
  End Select
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If bEnd = 0 Then
    bEnd = 1
    Cancel = 1
    'SaveSettings
  End If
  SaveSettings
  SkinOb.DeInitDecoration
  ' SaveScores
  'SaveSettings
End Sub

Public Function InitGfx() As Boolean
  'This function loads all graphics
  On Error GoTo ErrH
  InitGfx = True
  Dim x As Long
  
  'Load textures
  pTextures.AddFromFile App.Path & "\data\graphics\title.png", "title"
  pTextures.AddFromFile App.Path & "\data\graphics\field.png", "field"
  'TODO:
  pTextures.AddFromFile App.Path & "\data\graphics\level1bg.png", "level1bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level2bg.png", "level2bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level3bg.png", "level3bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level4bg.png", "level4bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level5bg.png", "level5bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level6bg.png", "level6bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level7bg.png", "level7bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level8bg.png", "level8bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level9bg.png", "level9bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level10bg.png", "level10bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level11bg.png", "level11bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level12bg.png", "level12bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level13bg.png", "level13bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level14bg.png", "level14bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level15bg.png", "level15bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level16bg.png", "level16bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level17bg.png", "level17bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level18bg.png", "level18bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level19bg.png", "level19bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level20bg.png", "level20bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level21bg.png", "level21bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level22bg.png", "level22bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level23bg.png", "level23bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level24bg.png", "level24bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level25bg.png", "level25bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level26bg.png", "level26bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level27bg.png", "level27bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level28bg.png", "level28bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level29bg.png", "level29bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level30bg.png", "level30bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level31bg.png", "level31bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level32bg.png", "level32bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level33bg.png", "level33bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level34bg.png", "level34bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level35bg.png", "level35bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level36bg.png", "level36bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level37bg.png", "level37bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level38bg.png", "level38bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level39bg.png", "level39bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level40bg.png", "level40bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level41bg.png", "level41bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level42bg.png", "level42bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level43bg.png", "level43bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level44bg.png", "level44bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level45bg.png", "level45bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level46bg.png", "level46bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level47bg.png", "level47bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level48bg.png", "level48bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level49bg.png", "level49bg"
  pTextures.AddFromFile App.Path & "\data\graphics\level50bg.png", "level50bg"
  'END TODO:
  pTextures.AddFromFile App.Path & "\data\graphics\bar.png", "bar"
  pTextures.AddFromFile App.Path & "\data\graphics\cursor.png", "cursor"
  pTextures.AddFromFile App.Path & "\data\graphics\new_game.png", "new_game"
  pTextures.AddFromFile App.Path & "\data\graphics\highscores.png", "highscore"
  pTextures.AddFromFile App.Path & "\data\graphics\options.png", "options"
  ' pTextures.AddFromFile App.Path & "\data\graphics\creditsinfo.png", "creditsinfo"
  pTextures.AddFromFile App.Path & "\data\graphics\exit.png", "exit"
  pTextures.AddFromFile App.Path & "\data\graphics\settings_big.png", "settings_big"
  pTextures.AddFromFile App.Path & "\data\graphics\sfx_volume.png", "sfx_volume"
  pTextures.AddFromFile App.Path & "\data\graphics\music_volume.png", "music_volume"
  pTextures.AddFromFile App.Path & "\data\graphics\mouse_speed.png", "mouse_speed"
  pTextures.AddFromFile App.Path & "\data\graphics\back.png", "back"
  pTextures.AddFromFile App.Path & "\data\graphics\meter.png", "meter"
  pTextures.AddFromFile App.Path & "\data\graphics\drag.png", "drag"
  pTextures.AddFromFile App.Path & "\data\graphics\highscore_big.png", "highscore_big"
  pTextures.AddFromFile App.Path & "\data\graphics\mikesbutterfly.png", "splat"
  pTextures.AddFromFile App.Path & "\data\graphics\star.png", "star"
  pTextures.AddFromFile App.Path & "\data\graphics\spark.png", "spark"
  pTextures.AddFromFile App.Path & "\data\graphics\spark2.png", "spark2"
  pTextures.AddFromFile App.Path & "\data\graphics\skull.png", "snow"
  pTextures.AddFromFile App.Path & "\data\graphics\smoke.png", "smoke"
  pTextures.AddFromFile App.Path & "\data\graphics\ball01.png", "ball01"
  pTextures.AddFromFile App.Path & "\data\graphics\ball02.png", "ball02"
  pTextures.AddFromFile App.Path & "\data\graphics\ball03.png", "ball03"
  pTextures.AddFromFile App.Path & "\data\graphics\ball04.png", "ball04"
  pTextures.AddFromFile App.Path & "\data\graphics\ball05.png", "ball05"
  pTextures.AddFromFile App.Path & "\data\graphics\ball06.png", "ball06"
  pTextures.AddFromFile App.Path & "\data\graphics\ball07.png", "ball07"
  pTextures.AddFromFile App.Path & "\data\graphics\ball08.png", "ball08"
  pTextures.AddFromFile App.Path & "\data\graphics\select.png", "select"
  pTextures.AddFromFile App.Path & "\data\graphics\explosion.png", "explosion", RGBA(0, 0, 0, 255)
  pTextures.AddFromFile App.Path & "\data\graphics\explosion2.png", "explosion2", RGBA(0, 0, 0, 255)
  pTextures.AddFromFile App.Path & "\data\graphics\30.png", "30"
  pTextures.AddFromFile App.Path & "\data\graphics\40.png", "40"
  pTextures.AddFromFile App.Path & "\data\graphics\50.png", "50"
  pTextures.AddFromFile App.Path & "\data\graphics\x2.png", "x2"
  pTextures.AddFromFile App.Path & "\data\graphics\x3.png", "x3"
  pTextures.AddFromFile App.Path & "\data\graphics\x4.png", "x4"
  pTextures.AddFromFile App.Path & "\data\graphics\x5.png", "x5"
  pTextures.AddFromFile App.Path & "\data\graphics\x6.png", "x6"
  pTextures.AddFromFile App.Path & "\data\graphics\x7.png", "x7"
  pTextures.AddFromFile App.Path & "\data\graphics\x8.png", "x8"
  pTextures.AddFromFile App.Path & "\data\graphics\x9.png", "x9"
  pTextures.AddFromFile App.Path & "\data\graphics\x10.png", "x10"
  pTextures.AddFromFile App.Path & "\data\graphics\hand.png", "hand"
  pTextures.AddFromFile App.Path & "\data\graphics\resume.png", "resume"
   pTextures.AddFromFile App.Path & "\data\graphics\creditsinfo.png", "creditsinfo"
  pTextures.AddFromFile App.Path & "\data\graphics\end_game.png", "end_game"
  pTextures.AddFromFile App.Path & "\data\graphics\game_over.png", "game_over"
  pTextures.AddFromFile App.Path & "\data\graphics\qualify.png", "qualify"
  pTextures.AddFromFile App.Path & "\data\graphics\ghostmydad.png", "ghoul"
  pTextures.AddFromFile App.Path & "\data\graphics\ring.png", "ring"
  pTextures.AddFromFile App.Path & "\data\graphics\staciesfairy.png", "mud"
  pTextures.AddFromFile App.Path & "\data\graphics\spec_left_right.png", "xlr"
  pTextures.AddFromFile App.Path & "\data\graphics\spec_up_down.png", "xud"
  pTextures.AddFromFile App.Path & "\data\graphics\spec_all.png", "xall"
  pTextures.AddFromFile App.Path & "\data\graphics\spec_switch.png", "xswitch"
  pTextures.AddFromFile App.Path & "\data\graphics\spec_x2.png", "xx2"
  pTextures.AddFromFile App.Path & "\data\graphics\spec_shield.png", "xshield"
  pTextures.AddFromFile App.Path & "\data\graphics\spec_time.png", "xtime"
  pTextures.AddFromFile App.Path & "\data\graphics\credits.png", "credits"
  pTextures.AddFromFile App.Path & "\data\graphics\credits_detail.png", "credits_detail"
  pTextures.AddFromFile App.Path & "\data\graphics\creditsinfotitle.png", "Creditsinfotitle"
  pTextures.AddFromFile App.Path & "\data\graphics\credits_info.png", "credits_info"
  
  pTextures.AddFromFile App.Path & "\data\graphics\bomb.png", "bomb"
  
  pTextures("explosion").Columns = 2
  pTextures("explosion").Rows = 8
  pTextures("explosion").SpriteHeight = 64
  pTextures("explosion").SpriteWidth = 64
  pTextures("explosion2").Columns = 2
  pTextures("explosion2").Rows = 8
  pTextures("explosion2").SpriteHeight = 64
  pTextures("explosion2").SpriteWidth = 64
  pTextures("ball08").Columns = 2
  pTextures("ball08").Rows = 4
  pTextures("ball08").SpriteHeight = 64
  pTextures("ball08").SpriteWidth = 64
  
  'Setup global sprites
  sLogo.CreateFromTexture "title"
  sBar.CreateFromTexture "bar"
  sBar.bCenterScale = False
  sBar.SetPosition 0, 0
  sField.CreateFromTexture "field"
  sField.bCenterScale = False
  sField.SetPosition 128, 0
  sCursor.CreateFromTexture "cursor"
  sCursor.bCenterScale = False
  sText_Menu(0).CreateFromTexture "new_game"
  sText_Menu(1).CreateFromTexture "highscore"
  sText_Menu(2).CreateFromTexture "options"
 ' sText_Menu(3).CreateFromTexture "creditsinfo"
  sText_Menu(3).CreateFromTexture "exit"
  sText_GameMenu(0).CreateFromTexture "resume"
  sText_GameMenu(1).CreateFromTexture "options"
  sText_GameMenu(2).CreateFromTexture "creditsinfo"
  sText_GameMenu(3).CreateFromTexture "end_game"
  sText_Settings(0).CreateFromTexture "settings_big"
  sText_Settings(1).CreateFromTexture "sfx_volume"
  sText_Settings(2).CreateFromTexture "music_volume"
  sText_Settings(3).CreateFromTexture "mouse_speed"
  sText_Settings(4).CreateFromTexture "back"
  For x = 0 To 3
    sText_GameMenu(x).SetColor 255, 255, 255, 0
  Next
  For x = 0 To 4
    sText_Settings(x).SetColor 255, 255, 255, 0
  Next
  For x = 0 To 2
    sMeter(x).CreateFromTexture "meter"
    sMeter(x).SetColor 255, 255, 255, 0
    sMeterDrag(x).CreateFromTexture "drag"
    sMeterDrag(x).SetColor 255, 255, 255, 0
  Next
  sText_HighScore(0).CreateFromTexture "highscore_big"
  sText_HighScore(1).CreateFromTexture "back"
  For x = 0 To 1
    sText_HighScore(x).SetColor 255, 255, 255, 0
  Next
  sMarble(0).CreateFromTexture "ball01"
  sMarble(1).CreateFromTexture "ball02"
  sMarble(2).CreateFromTexture "ball03"
  sMarble(3).CreateFromTexture "ball04"
  sMarble(4).CreateFromTexture "ball05"
  sMarble(5).CreateFromTexture "ball06"
  sMarble(6).CreateFromTexture "ball07"
  sMarble(7).CreateFromTexture "ball08", 1, 8, 100
  For x = 0 To 7
    sMarble(x).bCenterScale = False
  Next
  sSelect.CreateFromTexture "select"
  sHand.CreateFromTexture "hand"
  sGameOver.CreateFromTexture "game_over"
  sGameOver.SetColor 255, 255, 255, 0
  sQualify.CreateFromTexture "qualify"
  sQualify.SetColor 255, 255, 255, 0
  sCredits.CreateFromTexture "credits"
  sCredits_Detail.CreateFromTexture "credits_detail"
  sCreditsinfotitle.CreateFromTexture "creditsinfotitle"
  sCredits_Info.CreateFromTexture "credits_info"
  sSpecials(0).CreateFromTexture "xlr"
  sSpecials(1).CreateFromTexture "xud"
  sSpecials(2).CreateFromTexture "xall"
  sSpecials(3).CreateFromTexture "xswitch"
  sSpecials(4).CreateFromTexture "xx2"
  sSpecials(5).CreateFromTexture "xshield"
  sSpecials(6).CreateFromTexture "xtime"

  Exit Function
ErrH:
  MsgBox "Could not init game graphic data. Please Reinstall.", vbCritical, "Error"
  InitGfx = False
End Function

Public Sub NewGame()
  'This sub resets to an new game
  'reset variables
  CurrentMusic = CurrentMusic + 1
  If CurrentMusic > MaxMusic Then CurrentMusic = 1
  pSound.MusicPlayID CurrentMusic, Settings.MusicVolume
  Player.lLevel = 1
  Player.lScore = 0
  Player.lBombs = 13
  Player.lTime = 90000
  Player.lDisplayTime = 90000
  'TODO:
  NoMoreMovesRes = False
  sField.CreateFromTexture "level1bg"
  sField.bCenterScale = False
  sField.SetPosition 128, 0
  'END TODO:
  'Create random playfield
Remake:
  Dim x As Long, y As Long
  For x = 0 To 7
    For y = 0 To 8
      lGrid(x, y).lType = Int(Rnd * 6) + 1
      lGrid(x, y).lY = (y * 64) - 64
      lGrid(x, y).lX = 128 + x * 64
      lGrid(x, y).lDead = 0
      lGrid(x, y).bSpecial = 0
    Next
  Next
  'make sure there are no three in a row of the same marble color
  'vertical
  Dim l As Long, l2 As Long
  For x = 0 To 7
    l = 0
    l2 = 0
    For y = 1 To 8
      If lGrid(x, y).lType = l Then
        l2 = l2 + 1
      Else
        l = lGrid(x, y).lType
        l2 = 1
      End If
      If l2 = 3 Then
        Do
          lGrid(x, y).lType = Int(Rnd * 6) + 1
        Loop Until lGrid(x, y).lType <> l
      End If
    Next
    If l2 = 3 Then
      Do
        lGrid(x, 8).lType = Int(Rnd * 6) + 1
      Loop Until lGrid(x, 8).lType <> l
    End If
  Next
  'horizontal
  For y = 1 To 8
    l = 0
    l2 = 0
    For x = 0 To 7
      If lGrid(x, y).lType = l Then
        l2 = l2 + 1
      Else
        l = lGrid(x, y).lType
        l2 = 1
      End If
      If l2 = 3 Then
        Do
          lGrid(x, y).lType = Int(Rnd * 6) + 1
        Loop Until lGrid(x, y).lType <> l
      End If
    Next
    If l2 = 3 Then
      Do
        lGrid(7, y).lType = Int(Rnd * 6) + 1
      Loop Until lGrid(7, y).lType <> l
    End If
  Next
  If CheckMatches Then GoTo Remake
End Sub

Public Sub Highscore()
  'This is the highscore screen
  bEnd = 0
  
  Dim T As New pgeTimer
  Dim tRect As RECT
  Dim x As Long
  Dim spark As New pgeSprite
  Dim fade As Long
  Dim f As New pgeTimer
  spark.CreateFromTexture "spark"
  
  'setup graphics needed in this screen.
  sText_HighScore(0).SetPosition 384, -450
  sText_HighScore(1).SetPosition 192, -100
  sText_HighScore(0).SetAutoPath 384, 100, 1000
  sText_HighScore(1).SetAutoPath 192, 450, 1000
  For x = 0 To 1
    sText_HighScore(x).SetColor 255, 255, 255, 0
    sText_HighScore(x).SetAutoFade 255, 255, 255, 255, 1000
  Next
  
  sLogo.SetAutoPath 900, 100, 1000
  sText_Menu(0).SetAutoPath 900, 200, 1000
  sText_Menu(1).SetAutoPath 900, 275, 1000
  sText_Menu(2).SetAutoPath 900, 350, 1000
 'sText_Menu(3).SetAutoPath 900, 375, 1000
   sText_Menu(3).SetAutoPath 900, 425, 1000
  sLogo.SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(0).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(1).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(2).SetAutoFade 255, 255, 255, 0, 1000
 ' sText_Menu(3).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(3).SetAutoFade 255, 255, 255, 0, 1000
  
  spark.Active = False
  
  T.StartTime
  f.StartTime
  Do
    DoEvents
    
    'Poll & setup mouse pointer
    pMouse.Poll Me
    sCursor.SetPosition pMouse.g_cursorx, pMouse.g_cursory
    tRect.Left = pMouse.g_cursorx
    tRect.Top = pMouse.g_cursory
    tRect.Right = tRect.Left + 1
    tRect.bottom = tRect.Top + 1

    'Handle input, but only after menu items have appeared
    If T.GetTime > 1 Then
      'Check mouse vs menu options
      If IntersectR(tRect, sText_HighScore(1).GetDestRect) Then 'back
        sText_HighScore(1).SetColor 255, 255, 255, 100 + Abs(155 * Sine(T.GetTimeMs / 5))
        If pMouse.button1 Then bEnd = 2
      Else
        sText_HighScore(1).SetColor 255, 255, 255, 255
      End If
      'Check keyboard
      If pKeyboard.KeyDown(DIK_ESCAPE) Then bEnd = 2
    End If

    '###Drawing sequence
    pEngine.Clear
    
    sBar.Render
    sField.Render
    sLogo.Render
    For x = 0 To 3
      sText_Menu(x).Render
    Next
    For x = 0 To 1
      sText_HighScore(x).Render
    Next
    sQualify.Render
    sGameOver.Render
    spark.Render
    
    'Sparks at the highscore sign
    If Not spark.Active And T.GetTime > 2 Then
      spark.Active = True
      spark.SetPosition sText_HighScore(0).GetUpperLeftCorner.x + 30 + Rnd * (sText_HighScore(0).GetWidth - 60), sText_HighScore(0).GetUpperLeftCorner.y + 80 + Rnd * (sText_HighScore(0).GetHeight - 160)
      spark.SetAutoRotation 1 + Rnd * 3, 10
      spark.SetScale 0.1, 0.1
      spark.SetAutoScale 1, 1, 1500
      spark.SetColor 255, 255, 255, 255
      spark.SetAutoFade 255, 255, 255, 0, 2500, True
    End If
    
    If T.GetTime > 1 And f.GetTimeMs > 50 Then
      fade = fade + 5
      f.StartTime
    End If
    
    If T.GetTime > 1 Then
      For x = 0 To 9
        If x = LatestHigh Then
          MainFont.DrawText x + 1 & ". " & High(x).sName, ReturnRECT(180, 160 + 25 * x, 450, 160 + 25 * (x + 1)), RGBA(255 - x * 15, 255 - x * 15, 0, tob(fade - x * 7)), DT_LEFT
          MainFont.DrawText CStr(High(x).lScore), ReturnRECT(500, 160 + 25 * x, 575, 160 + 25 * (x + 1)), RGBA(255 - x * 15, 255 - x * 15, 0, tob(fade - x * 7)), DT_RIGHT
        Else
          MainFont.DrawText x + 1 & ". " & High(x).sName, ReturnRECT(180, 160 + 25 * x, 450, 160 + 25 * (x + 1)), RGBA(255 - x * 15, 255 - x * 15, 255 - x * 15, tob(fade - x * 7)), DT_LEFT
          MainFont.DrawText CStr(High(x).lScore), ReturnRECT(500, 160 + 25 * x, 575, 160 + 25 * (x + 1)), RGBA(255 - x * 15, 255 - x * 15, 255 - x * 15, tob(fade - x * 7)), DT_RIGHT
        End If
      Next
    End If
    
    sCursor.Render
    
    If bFps Then
      FontArial.DrawText pEngine.lFPS, ReturnRECT(0, 0, 100, 15), RGBA(0, 255, 0, 255), DT_LEFT
    End If
    
    pEngine.Render
    '###End of drawing sequence
  Loop Until bEnd
  
  sText_HighScore(0).SetAutoPath 384, -450, 1000
  sText_HighScore(1).SetAutoPath 192, -100, 1000
  For x = 0 To 1
    sText_HighScore(x).SetAutoFade 255, 255, 255, 0, 1000
  Next
  LatestHigh = -1
End Sub

Public Sub Options()
  'This is the options screen
  bEnd = 0
  
  Dim T As New pgeTimer
  Dim tRect As RECT
  Dim x As Long, y As Long
  Dim bDrag As Integer
  
  'setup graphics needed in this screen.
  sText_Settings(0).SetPosition 384, 600
  sText_Settings(1).SetPosition 224, 700
  sText_Settings(2).SetPosition 224, 775
  sText_Settings(3).SetPosition 224, 850
  sText_Settings(4).SetPosition 288, 925
  For x = 0 To 4
    sText_Settings(x).SetColor 255, 255, 255, 0
    sText_Settings(x).SetAutoFade 255, 255, 255, 255, 1000
  Next
  sText_Settings(0).SetAutoPath 384, 100, 1000
  sText_Settings(1).SetAutoPath 224, 200, 1000
  sText_Settings(2).SetAutoPath 224, 275, 1000
  sText_Settings(3).SetAutoPath 224, 350, 1000
 sText_Settings(4).SetAutoPath 224, 425, 1000
 
  sMeter(0).SetPosition 496, 700
  sMeter(1).SetPosition 496, 775
  sMeter(2).SetPosition 496, 850
  sMeter(0).SetAutoPath 496, 200, 1000
  sMeter(1).SetAutoPath 496, 275, 1000
  sMeter(2).SetAutoPath 496, 350, 1000
  'meter length = 0-175
  sMeterDrag(0).SetPosition 408 + 175 * (Settings.SfxVolume / 255), 700
  sMeterDrag(1).SetPosition 408 + 175 * (Settings.MusicVolume / 255), 775
  sMeterDrag(2).SetPosition 408 + 175 * ((Settings.MouseSpeed - 0.5) / 4.5), 850
  sMeterDrag(0).SetAutoPath sMeterDrag(0).GetPosition.x, 200, 1000
  sMeterDrag(1).SetAutoPath sMeterDrag(1).GetPosition.x, 275, 1000
  sMeterDrag(2).SetAutoPath sMeterDrag(2).GetPosition.x, 350, 1000
  For x = 0 To 2
    sMeter(x).SetAutoFade 255, 255, 255, 255, 1000
    sMeter(x).SetColor 255, 255, 255, 0
    sMeterDrag(x).SetColor 255, 255, 255, 0
    sMeterDrag(x).SetAutoFade 255, 255, 255, 255, 1000
  Next
  
  sLogo.SetAutoPath 900, 100, 1000
  sText_Menu(0).SetAutoPath 900, 200, 1000
  sText_Menu(1).SetAutoPath 900, 275, 1000
  sText_Menu(2).SetAutoPath 900, 350, 1000
  'sText_Menu(3).SetAutoPath 900, 375, 1000
  sText_Menu(3).SetAutoPath 900, 425, 1000
  sLogo.SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(0).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(1).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(2).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(3).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(4).SetAutoFade 255, 255, 255, 0, 1000
  
  bDrag = -1
  T.StartTime
  Do
    DoEvents
    
    'Poll & setup mouse pointer
    pMouse.Poll Me
    sCursor.SetPosition pMouse.g_cursorx, pMouse.g_cursory
    tRect.Left = pMouse.g_cursorx
    tRect.Top = pMouse.g_cursory
    tRect.Right = tRect.Left + 1
    tRect.bottom = tRect.Top + 1

    'Handle input, but only after menu items have appeared
    If T.GetTime > 1 Then
      'Check mouse vs menu options
      If IntersectR(tRect, sText_Settings(4).GetDestRect) Then 'back
        sText_Settings(4).SetColor 255, 255, 255, 100 + Abs(155 * Sine(T.GetTimeMs / 5))
        If pMouse.button1 Then bEnd = 2
      Else
        sText_Settings(4).SetColor 255, 255, 255, 255
      End If
      For x = 0 To 2 'meter bars
        If IntersectR(tRect, sMeterDrag(x).GetDestRect) Then
          sMeterDrag(x).SetColor 255, 255, 255, 100 + Abs(155 * Sine(T.GetTimeMs / 5))
          If pMouse.button1 Then
            If bDrag = -1 Then bDrag = x
            If pMouse.g_cursorx >= 408 And pMouse.g_cursorx <= 583 Then
              sMeterDrag(bDrag).SetPosition pMouse.g_cursorx, sMeterDrag(bDrag).GetPosition.y
            End If
          Else
            Select Case bDrag
              Case 0
                Settings.SfxVolume = Int(255 * (sMeterDrag(0).GetPosition.x - 408) / 175)
              Case 1
                Settings.MusicVolume = Int(255 * (sMeterDrag(1).GetPosition.x - 408) / 175)
                pSound.MusicVolume Settings.MusicVolume
              Case 2
                Settings.MouseSpeed = 0.5 + 4.5 * (sMeterDrag(2).GetPosition.x - 408) / 175
                pMouse.g_Sensitivity = Settings.MouseSpeed
            End Select
            bDrag = -1
          End If
        Else
          sMeterDrag(x).SetColor 255, 255, 255, 255
        End If
      Next
      'Check keyboard
      If pKeyboard.KeyDown(DIK_ESCAPE) Then bEnd = 2
    End If

    '###Drawing sequence
    pEngine.Clear
    
    sBar.Render
    sField.Render
    sLogo.Render
    For x = 0 To 3
      sText_Menu(x).Render
    Next
    For x = 0 To 2
      sText_GameMenu(x).Render
    Next
    For x = 0 To 4
      sText_Settings(x).Render
    Next
    For x = 0 To 2
      sMeter(x).Render
      sMeterDrag(x).Render
    Next
    
    If Player.lTime > 0 Then
      sHand.Render
      LedFont.DrawText Player.lScore, ReturnRECT(20, 34, 115, 64), RGBA(0, 200, 0, 255), DT_RIGHT
      y = -1
      For x = 9 To 0 Step -1
        If High(x).lScore > Player.lScore Then
          y = x
          Exit For
        End If
      Next
      If y = -1 Then
        LedFont.DrawText Player.lScore, ReturnRECT(20, 98, 115, 128), RGBA(0, 200, 0, 255), DT_RIGHT
      Else
        LedFont.DrawText High(y).lScore, ReturnRECT(20, 98, 115, 128), RGBA(0, 200, 0, 255), DT_RIGHT
      End If
      LedFont.DrawText Player.lBombs, ReturnRECT(20, 168, 115, 198), RGBA(0, 200, 0, 255), DT_RIGHT
    End If
    
    
    sCursor.Render
    
    If bFps Then
      FontArial.DrawText pEngine.lFPS, ReturnRECT(0, 0, 100, 15), RGBA(0, 255, 0, 255), DT_LEFT
    End If
    
    pEngine.Render
    '###End of drawing sequence
  Loop Until bEnd
  
  sText_Settings(0).SetAutoPath 384, 600, 1000
  sText_Settings(1).SetAutoPath 224, 700, 1000
  sText_Settings(2).SetAutoPath 224, 775, 1000
  sText_Settings(3).SetAutoPath 224, 850, 1000
  sText_Settings(4).SetAutoPath 288, 925, 1000
  sMeter(0).SetAutoPath 496, 700, 1000
  sMeter(1).SetAutoPath 496, 775, 1000
  sMeter(2).SetAutoPath 496, 850, 1000
  sMeterDrag(0).SetAutoPath sMeterDrag(0).GetPosition.x, 700, 1000
  sMeterDrag(1).SetAutoPath sMeterDrag(1).GetPosition.x, 775, 1000
  sMeterDrag(2).SetAutoPath sMeterDrag(2).GetPosition.x, 850, 1000
  For x = 0 To 4
    sText_Settings(x).SetAutoFade 255, 255, 255, 0, 1000
  Next
  For x = 0 To 2
    sMeter(x).SetAutoFade 255, 255, 255, 0, 1000
    sMeterDrag(x).SetAutoFade 255, 255, 255, 0, 1000
  Next
  'SaveSettings
End Sub

Public Sub LoadScores()
  Dim file As New clsDatafile, x As Long
  If FileExist(App.Path & "\Config\HighScores.ini") Then
    file.fileName = App.Path & "\Config\HighScores.ini"
    For x = 0 To 9
      High(x).sName = file.ReadStr
      High(x).lScore = file.ReadNumber
    Next
  Else
    For x = 0 To 9
      High(x).sName = "Your Name Here..."
      High(x).lScore = 1000 - x * 100
    Next
  End If
End Sub

Private Sub Form_Terminate()
SkinOb.DeInitDecoration
 'SaveScores
'SaveSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  SaveScores
 ' SaveSettings
  RemoveFontResource App.Path & "\Data\Altea.TTF"
  RemoveFontResource App.Path & "\Data\lcdmb.ttf"
  SaveSettings
 ' SkinOb.DeInitDecoration
'FrmCredits.Show
 'SaveScores
  'SaveSettings
End Sub

Public Sub DoGame()
  'This is the game screen
  bEnd = 0
  
  Dim T As New pgeTimer
  Dim T2 As New pgeTimer
  Dim game As New pgeTimer
  Dim LevelandMoveTimer As New pgeTimer
  Dim tRect As RECT
  Dim x As Long, y As Long, z As Long
  Dim selx As Integer, sely As Integer
  Dim swapx As Integer, swapy As Integer
  Dim swap2x As Integer, swap2y As Integer
  Dim down As Boolean
  Dim down2 As Boolean
  Dim moving As Boolean 'moving is true when there are animations on screen, and prevents for example you to select other marbles
  Dim dead As Boolean 'dead is true when marbles are being destroyed, and prevents for example other marbles to fall down
  Dim tGrid(7, 8) As Integer
  Dim bonus As Integer
  Dim lSnd(9) As Long
  Dim sSkip As Single
  Dim l As Long, l2 As Long
  Dim DoEndGame As Boolean 'is true if game should be ended by force
  
  
  'setup graphics needed in this screen.
  sLogo.SetAutoPath 900, 100, 1000
  sText_Menu(0).SetAutoPath 900, 200, 1000
  sText_Menu(1).SetAutoPath 900, 275, 1000
  sText_Menu(2).SetAutoPath 900, 350, 1000
 ' sText_Menu(3).SetAutoPath 900, 375, 1000
  sText_Menu(3).SetAutoPath 900, 425, 1000
  sLogo.SetAutoFade 255, 255, 255, 0, 1000
  For x = 0 To 2
    sText_Menu(x).SetAutoFade 255, 255, 255, 0, 1000
  Next
  
  pKeyboard.SetTimerEx DIK_F5, 1
  
  sHand.SetPosition 66, 280
  sSelect.SetScale 0.6, 0.6
  
  T.StartTime
gAgain:
  T2.StartTime
  game.StartTime
  LevelandMoveTimer.StartTime
  
  selx = -1
  sely = -1
  swapx = -1
  
  Do
    DoEvents
    
    'Poll & setup mouse pointer
    pMouse.Poll Me
    sCursor.SetPosition pMouse.g_cursorx, pMouse.g_cursory
    tRect.Left = pMouse.g_cursorx
    tRect.Top = pMouse.g_cursory
    tRect.Right = tRect.Left + 1
    tRect.bottom = tRect.Top + 1
        
    dead = False
    moving = False
    'Process game
    If T.GetTimeMs > 5 Then
      'make sure 5ms has passed, so animations are played at an nice speed
      sSkip = T.GetTimeMs / 10
      
      
      For x = 0 To NumMud
        With sMud(x)
          If .Active Then
            .SetAutoMovement .GetAutoMovement.x, .GetAutoMovement.y + 0.05, 20
          End If
        End With
      Next
      
      
      For x = 0 To 7
        For y = 0 To 8
          With lGrid(x, y)
            If .lDead > 0 Then 'this marble is to be removed, show its death animation
              .lDead = .lDead + 1 * sSkip
              .lFlag(3) = .lFlag(3) + 1
              dead = True
              Select Case .lType
                Case 7 'ghouls
                  If .lDead >= 50 Then
                    .lType = 0
                    .lDead = 0
                    lSnd(3) = 0
                  Else
                    If .lFlag(0) = False And .lDead >= 2 Then
                      If lSnd(3) = 0 Or Not pSound.SfxIsPlaying(lSnd(3)) Then
                        lSnd(3) = pSound.SfxPlayExID("ghouls0" & Int(Rnd * 2) + 1, LOOP_OFF, Settings.SfxVolume)
                      End If
                      .lFlag(0) = True
                    End If
                    If .lFlag(3) < 3 Then AddGhoul CSng(.lX) + 34, CSng(.lY) + 34, -30 + Rnd * 60
                  End If
                Case 1 'ice explosion
                  If .lDead >= 50 Then
                    .lType = 0
                    .lDead = 0
                    lSnd(4) = 0
                  Else
                    If .lFlag(0) = False Then
                      If lSnd(4) = 0 Or Not pSound.SfxIsPlaying(lSnd(4)) Then
                        lSnd(4) = pSound.SfxPlayExID("smash0" & Int(Rnd * 3) + 1, LOOP_OFF, Settings.SfxVolume)
                      End If
                      .lFlag(0) = True
                    End If
                    If .lFlag(1) = False And .lDead >= 3 Then
                      AddIceExplosion CSng(.lX) + 34, CSng(.lY) + 34
                      .lFlag(1) = True
                    End If
                    If .lFlag(3) < 15 Then AddSnow CSng(.lX) + 34, CSng(.lY) + 34, Rnd * 360
                  End If
                Case 6 'splat
                  If .lDead >= 50 Then
                    .lType = 0
                    .lDead = 0
                    lSnd(6) = 0
                  Else
                    If .lFlag(0) = False And .lDead >= 2 Then
                      If lSnd(6) = 0 Or Not pSound.SfxIsPlaying(lSnd(6)) Then
                        lSnd(6) = pSound.SfxPlayExID("splat0" & Int(Rnd * 2) + 1, LOOP_OFF, Settings.SfxVolume)
                      End If
                      .lFlag(0) = True
                    End If
                    If .lFlag(3) < 20 Then AddSplat CSng(.lX) + 34, CSng(.lY) + 34, Rnd * 360
                  End If
                Case 2 'scale down & rotate, sparkle
                  If .lDead >= 50 Then
                    .lType = 0
                    .lDead = 0
                    lSnd(1) = 0
                  Else
                    If .lFlag(0) = False And .lDead >= 2 Then
                      If lSnd(1) = 0 Or Not pSound.SfxIsPlaying(lSnd(1)) Then
                        lSnd(1) = pSound.SfxPlayExID("spark0" & Int(Rnd * 2) + 1, LOOP_OFF, Settings.SfxVolume)
                      End If
                      .lFlag(0) = True
                    End If
                    If .lFlag(3) < 20 Then AddSpark CSng(.lX) + 34, CSng(.lY) + 34, Rnd * 360
                  End If
                Case 5 'fairy 'mud splat
                If .lDead >= 50 Then
                    .lType = 0
                    .lDead = 0
                    lSnd(3) = 0
                  Else
                    If .lFlag(0) = False And .lDead >= 2 Then
                      If lSnd(3) = 0 Or Not pSound.SfxIsPlaying(lSnd(3)) Then
                        lSnd(3) = pSound.SfxPlayExID("ghouls0" & Int(Rnd * 2) + 1, LOOP_OFF, Settings.SfxVolume)
                      End If
                      .lFlag(0) = True
                    End If
                    If .lFlag(3) < 3 Then AddMud CSng(.lX) + 34, CSng(.lY) + 34, -30 + Rnd * 60
                  End If
                  'If .lDead >= 50 Then
                   ' .lType = 0
                    '.lDead = 0
                    'lSnd(9) = 0
                  'Else
                   ' If .lFlag(0) = False And .lDead >= 2 Then
                    '  If lSnd(9) = 0 Or Not pSound.SfxIsPlaying(lSnd(9)) Then
                     '   lSnd(9) = pSound.SfxPlayExID("mud", LOOP_OFF, Settings.SfxVolume)
                     ' End If
                      '.lFlag(0) = True
                    'End If
                    'If .lFlag(3) < 15 Then AddMud CSng(.lX) + 34, CSng(.lY) + 34, Rnd * 360
                  'End If
                Case 3 'teleport out
                  If .lDead >= 75 Then
                    .lType = 0
                    .lDead = 0
                    lSnd(5) = 0
                  Else
                    If .lFlag(0) = False And .lDead >= 2 Then
                      If lSnd(5) = 0 Or Not pSound.SfxIsPlaying(lSnd(5)) Then
                        lSnd(5) = pSound.SfxPlayExID("teleport", LOOP_OFF, Settings.SfxVolume)
                      End If
                      .lFlag(0) = True
                    End If
                    If .lFlag(3) < 6 Then AddStar CSng(.lX) + 16 + Rnd * 32, CSng(.lY) + 16 + Rnd * 32
                  End If
                Case 4 'beep, boop
                  If .lDead >= 35 Then
                    .lType = 0
                    .lDead = 0
                    lSnd(8) = 0
                  Else
                    If .lFlag(0) = False And .lDead >= 2 Then
                      If lSnd(8) = 0 Or Not pSound.SfxIsPlaying(lSnd(8)) Then
                        lSnd(8) = pSound.SfxPlayExID("boop", LOOP_OFF, Settings.SfxVolume)
                      End If
                      .lFlag(0) = True
                    End If
                    If .lFlag(3) = 4 Then AddRing CSng(.lX) + 34, CSng(.lY) + 34
                  End If
                Case 8 'explode
                  If .lDead >= 50 Then
                    .lType = 0
                    .lDead = 0
                    lSnd(2) = 0
                  Else
                    If .lFlag(0) = False And .lDead >= 2 Then
                      If lSnd(2) = 0 Or Not pSound.SfxIsPlaying(lSnd(2)) Then
                        lSnd(2) = pSound.SfxPlayExID("explosion0" & Int(Rnd * 2) + 1, LOOP_OFF, Settings.SfxVolume)
                      End If
                      .lFlag(0) = True
                    End If
                    If .lFlag(3) < 6 Then AddSmoke CSng(.lX) + 34, CSng(.lY) + 34
                    If .lFlag(3) = 2 Then AddExplosion CSng(.lX) + 34, CSng(.lY) + 34
                  End If
              End Select
            End If
          End With
        Next
      Next
      
      'Check if there are empty spaces on the top row (not visible), add marbles
      For x = 0 To 7
        If lGrid(x, 0).lType = 0 Then
          If lGrid(x, 1).lY >= 0 Then
            If Player.lScore < 5000 Then
              lGrid(x, 0).lType = Int(Rnd * 6) + 1 'Add random marble
            ElseIf Player.lScore < 7500 Then
              lGrid(x, 0).lType = Int(Rnd * 7) + 1 'Add random marble
            Else
              lGrid(x, 0).lType = Int(Rnd * 8) + 1 'Add random marble
            End If
            'set the special
            z = Int(Rnd * 10000) + 1
            lGrid(x, 0).lFlag(3) = 0
            Select Case z
              Case 200 To 250 'clear left right
                lGrid(x, 0).bSpecial = 1
              Case 400 To 450 'clear up down
                lGrid(x, 0).bSpecial = 2
              Case 8000 To 8075 'clear all of same sort after 10000 pts
                If Player.lScore > 10000 Then lGrid(x, 0).bSpecial = 3
              Case 1600 To 1900 'switcher
                lGrid(x, 0).bSpecial = 4
                lGrid(x, 0).lFlag(3) = timeGetTime
              Case 600 To 850 'x2
                lGrid(x, 0).bSpecial = 5
              Case 2000 To 2200 'shield after 5000 pts
                If Player.lScore > 5000 Then
                  lGrid(x, 0).bSpecial = 6
                End If
              Case 2500 To 3000 'even higher chance of shield after 10000 pts
                If Player.lScore > 10000 Then
                  lGrid(x, 0).bSpecial = 6
                End If
              Case 6500 To 7500 'even higher chance of shield after 20000 pts
                If Player.lScore > 20000 Then
                  lGrid(x, 0).bSpecial = 6
                End If
              Case 5600 To 5900 'time
                lGrid(x, 0).bSpecial = 7
              Case Else
                lGrid(x, 0).bSpecial = 0
            End Select
            lGrid(x, 0).lY = -64
            lGrid(x, 0).lX = 128 + x * 64
          End If
        End If
      Next
      'If there are spaces under marbles, drop them down...
      If Not dead Then
        For x = 0 To 7
          For y = 7 To 0 Step -1
            If lGrid(x, y + 1).lType = 0 Then
              lGrid(x, y + 1).lType = lGrid(x, y).lType
              lGrid(x, y + 1).bSpecial = lGrid(x, y).bSpecial
              lGrid(x, y + 1).lY = lGrid(x, y).lY
              lGrid(x, y + 1).lFlag(3) = lGrid(x, y).lFlag(3)
              lGrid(x, y).lType = 0
              lGrid(x, y).lY = 0
              moving = True
            End If
          Next
        Next
      End If
      'Check if there are marbles to be swapped, animate them
      If swapx > -1 Then
        moving = True
        z = 0
        If swapx <> swap2x Then 'Swap is on x axis
          x = 128 + swap2x * 64
          y = 128 + swapx * 64
          If lGrid(swapx, swapy).lX < x - 4 * sSkip Then
            lGrid(swapx, swapy).lX = lGrid(swapx, swapy).lX + 4 * sSkip
          ElseIf lGrid(swapx, swapy).lX > x + 4 * sSkip Then
            lGrid(swapx, swapy).lX = lGrid(swapx, swapy).lX - 4 * sSkip
          Else
            lGrid(swapx, swapy).lX = x
            z = z + 1
          End If

          If lGrid(swap2x, swap2y).lX < y - 4 * sSkip Then
            lGrid(swap2x, swap2y).lX = lGrid(swap2x, swap2y).lX + 4 * sSkip
          ElseIf lGrid(swap2x, swap2y).lX > y + 4 * sSkip Then
            lGrid(swap2x, swap2y).lX = lGrid(swap2x, swap2y).lX - 4 * sSkip
          Else
            lGrid(swap2x, swap2y).lX = y
            z = z + 1
          End If
          
        Else 'swap is on y axis
          x = (swap2y - 1) * 64
          y = (swapy - 1) * 64
          
          If lGrid(swapx, swapy).lY > x + 4 * sSkip Then
            lGrid(swapx, swapy).lY = lGrid(swapx, swapy).lY - 4 * sSkip
          ElseIf lGrid(swapx, swapy).lY < x - 4 * sSkip Then
            lGrid(swapx, swapy).lY = lGrid(swapx, swapy).lY + 4 * sSkip
          Else
            lGrid(swapx, swapy).lY = x
            z = z + 1
          End If
          
          If lGrid(swap2x, swap2y).lY > y + 4 * sSkip Then
            lGrid(swap2x, swap2y).lY = lGrid(swap2x, swap2y).lY - 4 * sSkip
          ElseIf lGrid(swap2x, swap2y).lY < y - 4 * sSkip Then
            lGrid(swap2x, swap2y).lY = lGrid(swap2x, swap2y).lY + 4 * sSkip
          Else
            lGrid(swap2x, swap2y).lY = y
            z = z + 1
          End If

        End If
        
        If z = 2 Then 'if both has swapped nicely, swap them for real
          lGrid(swapx, swapy).lX = 128 + swapx * 64
          lGrid(swapx, swapy).lY = (swapy - 1) * 64
          lGrid(swap2x, swap2y).lX = 128 + swap2x * 64
          lGrid(swap2x, swap2y).lY = (swap2y - 1) * 64
          z = lGrid(swapx, swapy).lType
          lGrid(swapx, swapy).lType = lGrid(swap2x, swap2y).lType
          lGrid(swap2x, swap2y).lType = z
          z = lGrid(swapx, swapy).bSpecial
          lGrid(swapx, swapy).bSpecial = lGrid(swap2x, swap2y).bSpecial
          lGrid(swap2x, swap2y).bSpecial = z
          z = lGrid(swapx, swapy).lFlag(3)
          lGrid(swapx, swapy).lFlag(3) = lGrid(swap2x, swap2y).lFlag(3)
          lGrid(swap2x, swap2y).lFlag(3) = z
          swapx = -1
        End If
        
      Else
        'Check if marbles are falling, animate them
        For x = 0 To 7
          For y = 8 To 0 Step -1
            If lGrid(x, y).lType > 0 Then
              If lGrid(x, y).lY < y * 64 - 64 Then
                If y < 8 Then
                  If lGrid(x, y + 1).lY > lGrid(x, y).lY + 64 Then
                    If lGrid(x, y).lY + 8 * sSkip < y * 64 - 64 Then
                      lGrid(x, y).lY = lGrid(x, y).lY + 8 * sSkip
                    Else
                      lGrid(x, y).lY = y * 64 - 64
                    End If
                    moving = True
                  End If
                Else
                  If lGrid(x, y).lY + 8 * sSkip < y * 64 - 64 Then
                    lGrid(x, y).lY = lGrid(x, y).lY + 8 * sSkip
                  Else
                    lGrid(x, y).lY = y * 64 - 64
                  End If
                  moving = True
                End If
                If lGrid(x, y).lY = y * 64 - 64 Then
                  If lSnd(0) = 0 Or Not pSound.SfxIsPlaying(lSnd(0)) Then
                    lSnd(0) = pSound.SfxPlayExID("land0" & Int(Rnd * 2) + 1, LOOP_OFF, Settings.SfxVolume)
                  End If
                End If
              End If
            End If
          Next
        Next
      End If
      
      'Clear all matches
      If Not moving And Not dead Then
        If CheckMatches Then bonus = bonus + 1
        Erase tGrid
        'vertical
        For x = 0 To 7
          l = 0
          l2 = 0
          For y = 1 To 8
            If lGrid(x, y).lType = l Then
              l2 = l2 + 1
            Else
              If l2 >= 3 Then
                For z = y - l2 To y - 1
                  tGrid(x, z) = -1
                  If lGrid(x, z).bSpecial = 5 Then bonus = bonus + 1
                Next
                HandleScore x, y - 1, l2, bonus
              End If
              l = lGrid(x, y).lType
              l2 = 1
            End If
          Next
          If l2 >= 3 Then
            For z = y - l2 To y - 1
              tGrid(x, z) = -1
              If lGrid(x, z).bSpecial = 5 Then bonus = bonus + 1
            Next
            HandleScore x, y - 1, l2, bonus
          End If
        Next
        'horizontal
        For y = 1 To 8
          l = 0
          l2 = 0
          For x = 0 To 7
            If lGrid(x, y).lType = l Then
              l2 = l2 + 1
            Else
              If l2 >= 3 Then
                For z = x - l2 To x - 1
                  tGrid(z, y) = -1
                  If lGrid(z, y).bSpecial = 5 Then bonus = bonus + 1
                Next
                HandleScore x - 1, y, l2, bonus
              End If
              l = lGrid(x, y).lType
              l2 = 1
            End If
          Next
          If l2 >= 3 Then
            For z = x - l2 To x - 1
              tGrid(z, y) = -1
              If lGrid(z, y).bSpecial = 5 Then bonus = bonus + 1
            Next
            HandleScore x - 1, y, l2, bonus
          End If
        Next
        
        'Remove marbles
        For x = 0 To 7
          For y = 1 To 8
            If tGrid(x, y) = -1 Then
              lGrid(x, y).lDead = 1
              lGrid(x, y).lFlag(3) = 0
              lGrid(x, y).lFlag(0) = False
              lGrid(x, y).lFlag(1) = False
              lGrid(x, y).lFlag(2) = False
              Player.lScore = Player.lScore + 10 * bonus
              'Handle marble specials
              Select Case lGrid(x, y).bSpecial
                Case 1 'remove row
                  For z = 0 To 7
                    lGrid(z, y).lDead = 1
                  Next
                  pSound.SfxPlayExID "explosion0" & Int(Rnd * 2) + 1, LOOP_OFF, Settings.SfxVolume
                Case 2 'remove column
                  For z = 1 To 8
                    lGrid(x, z).lDead = 1
                  Next
                  pSound.SfxPlayExID "explosion0" & Int(Rnd * 2) + 1, LOOP_OFF, Settings.SfxVolume
                Case 3 'remove all of same type
                  For l = 0 To 7
                    For l2 = 1 To 8
                      If lGrid(l, l2).lType = lGrid(x, y).lType Then
                        lGrid(l, l2).lDead = 1
                      End If
                    Next
                  Next
                  pSound.SfxPlayExID "explosion0" & Int(Rnd * 2) + 1, LOOP_OFF, Settings.SfxVolume
                Case 6
                  lGrid(x, y).lDead = 0
                  lGrid(x, y).bSpecial = 0
                  pSound.SfxPlayExID "spec6", LOOP_OFF, Settings.SfxVolume
                Case 7
                  Player.lTime = Player.lTime + 5000
                  If Player.lTime > 120000 Then
                    Player.lScore = Player.lScore + 250
                  End If
              End Select
            End If
          Next
        Next
      End If
      
      
      '##Handle input
      'Check mouse
      If T2.GetTimeMs > 250 Then
        If pMouse.button1 Then
          If Not down And Not moving And Not dead Then 'make sure nothing is moving
            bonus = 0
            down = True 'to make sure you release the button first
            If sely = -1 And selx = -1 Then 'If nothing is selected, select
              selx = Int(((pMouse.g_cursorx - 128) / 64))
              sely = Int((pMouse.g_cursory / 64))
              If selx < 0 Then selx = 0
              Call pSound.SfxPlayExID("select", LOOP_OFF, Settings.SfxVolume)
            Else 'if something is selected, check to see if new selected is beside old selected
                 'if it is, check if the selected ones can be swapped to create an three in a row
                 'else deselect and play sound
              x = Int(((pMouse.g_cursorx - 128) / 64))
              y = Int((pMouse.g_cursory / 64))
              If x < 0 Then x = 0
              'if selected is other than new selected
              If x <> selx Or y <> sely Then
                'if new selected is adjacent to old selected setup them for an swap
                If (Abs(x - selx) = 1 And y = sely) Or (Abs(y - sely) = 1 And x = selx) Then
                  'ok, so the new selected one is next to the old selected one
                  'now check to see if swapping these will create three in a row
                  z = lGrid(selx, sely + 1).lType
                  lGrid(selx, sely + 1).lType = lGrid(x, y + 1).lType
                  lGrid(x, y + 1).lType = z
                  If CheckMatches Then 'Yes, three in a row will be created
                    swapx = x
                    swapy = y + 1
                    swap2x = selx
                    swap2y = sely + 1
                    Call pSound.SfxPlayExID("switch", LOOP_OFF, Settings.SfxVolume)
                  Else
                    swapx = -1
                    Call pSound.SfxPlayExID("deselect", LOOP_OFF, Settings.SfxVolume)
                  End If
                  z = lGrid(selx, sely + 1).lType
                  lGrid(selx, sely + 1).lType = lGrid(x, y + 1).lType
                  lGrid(x, y + 1).lType = z
                Else
                  'new selected is not adjacent to old selected play error sound.
                  Call pSound.SfxPlayExID("deselect", LOOP_OFF, Settings.SfxVolume)
                End If
                'reset selection
                selx = -1
                sely = -1
              End If
            End If
          End If
        Else
          down = False
        End If
        If pMouse.button2 Then
          If Not down2 And Not moving And Not dead And Player.lBombs > 0 Then 'make sure nothing is moving and there are bombs
            bonus = 1
            down2 = True
            x = Int(((pMouse.g_cursorx - 128) / 64))
            y = Int((pMouse.g_cursory / 64))
            AddExplosion 160 + x * 64, 32 + y * 64, 3 'display large explosion
            selx = x
            sely = y
            If selx + 1 <= 7 Then selx = selx + 1
            If sely + 1 <= 7 Then sely = sely + 1
            If x - 1 >= 0 Then x = x - 1
            If y - 1 >= 0 Then y = y - 1
            For swapx = x To selx
              For swapy = y + 1 To sely + 1
                lGrid(swapx, swapy).lDead = 1
              Next
            Next
            'Player.lTime = Player.lTime + 15000
            Player.lBombs = Player.lBombs - 1
            Call pSound.SfxPlayExID("explosion03", LOOP_OFF, Settings.SfxVolume)
            'dropping an bomb resets selection, as marbles rearrange
            swapx = -1
            selx = -1
            sely = -1
          End If
        Else
          down2 = False
        End If
      End If

      'Check keyboard
      If pKeyboard.KeyDown(DIK_ESCAPE) And Not DoEndGame And Not moving And Not dead And T2.GetTimeMs > 500 Then bEnd = 2
      If pKeyboard.KeyDown(DIK_F5) Then
        CurrentMusic = CurrentMusic + 1
        If CurrentMusic > MaxMusic Then CurrentMusic = 1
        pSound.MusicPlayID CurrentMusic, Settings.MusicVolume
      End If
      If pKeyboard.KeyDown(DIK_F12) Then DoEndGame = True
          
      T.StartTime
    End If
If pKeyboard.KeyDown(DIK_F1) Then
      CreditsAbout
      End If
    '###Drawing sequence
    pEngine.Clear
    
    sBar.Render
    'TODO:
    If Player.lScore > 1000 And Player.lScore < 26000 Then
        If Player.lScore > 1000 And Player.lScore < 2000 Then
            sField.CreateFromTexture "level2bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 2
        ElseIf Player.lScore > 2000 And Player.lScore < 3000 Then
            sField.CreateFromTexture "level3bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 3
        ElseIf Player.lScore > 3000 And Player.lScore < 4000 Then
            sField.CreateFromTexture "level4bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 4
        ElseIf Player.lScore > 4000 And Player.lScore < 5000 Then
            sField.CreateFromTexture "level5bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 5
        ElseIf Player.lScore > 5000 And Player.lScore < 6000 Then
            sField.CreateFromTexture "level6bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 6
        ElseIf Player.lScore > 6000 And Player.lScore < 7000 Then
            sField.CreateFromTexture "level7bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 7
        ElseIf Player.lScore > 8000 And Player.lScore < 9000 Then
            sField.CreateFromTexture "level8bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 8
        ElseIf Player.lScore > 9000 And Player.lScore < 10000 Then
            sField.CreateFromTexture "level9bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 9
        ElseIf Player.lScore > 10000 And Player.lScore < 11000 Then
            sField.CreateFromTexture "level10bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 10
        ElseIf Player.lScore > 11000 And Player.lScore < 12000 Then
            sField.CreateFromTexture "level11bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 11
        ElseIf Player.lScore > 12000 And Player.lScore < 13000 Then
            sField.CreateFromTexture "level12bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 12
        ElseIf Player.lScore > 13000 And Player.lScore < 14000 Then
            sField.CreateFromTexture "level13bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 13
        ElseIf Player.lScore > 14000 And Player.lScore < 15000 Then
            sField.CreateFromTexture "level14bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 14
        ElseIf Player.lScore > 15000 And Player.lScore < 16000 Then
            sField.CreateFromTexture "level15bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 15
        ElseIf Player.lScore > 16000 And Player.lScore < 17000 Then
            sField.CreateFromTexture "level16bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 16
        ElseIf Player.lScore > 17000 And Player.lScore < 18000 Then
            sField.CreateFromTexture "level17bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 17
        ElseIf Player.lScore > 18000 And Player.lScore < 19000 Then
            sField.CreateFromTexture "level18bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 18
        ElseIf Player.lScore > 19000 And Player.lScore < 20000 Then
            sField.CreateFromTexture "level19bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 19
        ElseIf Player.lScore > 20000 And Player.lScore < 21000 Then
            sField.CreateFromTexture "level20bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 20
        ElseIf Player.lScore > 21000 And Player.lScore < 22000 Then
            sField.CreateFromTexture "level21bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 21
        ElseIf Player.lScore > 22000 And Player.lScore < 23000 Then
            sField.CreateFromTexture "level22bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 22
        ElseIf Player.lScore > 23000 And Player.lScore < 24000 Then
            sField.CreateFromTexture "level23bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 23
        ElseIf Player.lScore > 24000 And Player.lScore < 25000 Then
            sField.CreateFromTexture "level24bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 24
        ElseIf Player.lScore > 25000 And Player.lScore < 26000 Then
            sField.CreateFromTexture "level25bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 25
        End If
    ElseIf Player.lScore > 26000 Then
        If Player.lScore > 26000 And Player.lScore < 27000 Then
            sField.CreateFromTexture "level26bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 26
        ElseIf Player.lScore > 27000 And Player.lScore < 28000 Then
            sField.CreateFromTexture "level27bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 27
        ElseIf Player.lScore > 28000 And Player.lScore < 29000 Then
            sField.CreateFromTexture "level28bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 28
        ElseIf Player.lScore > 29000 And Player.lScore < 30000 Then
            sField.CreateFromTexture "level29bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 29
        ElseIf Player.lScore > 30000 And Player.lScore < 31000 Then
            sField.CreateFromTexture "level30bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 30
        ElseIf Player.lScore > 31000 And Player.lScore < 32000 Then
            sField.CreateFromTexture "level31bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 31
        ElseIf Player.lScore > 32000 And Player.lScore < 33000 Then
            sField.CreateFromTexture "level32bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 32
        ElseIf Player.lScore > 33000 And Player.lScore < 34000 Then
            sField.CreateFromTexture "level33bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 33
        ElseIf Player.lScore > 34000 And Player.lScore < 35000 Then
            sField.CreateFromTexture "level34bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 34
        ElseIf Player.lScore > 35000 And Player.lScore < 36000 Then
            sField.CreateFromTexture "level35bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 35
        ElseIf Player.lScore > 36000 And Player.lScore < 37000 Then
            sField.CreateFromTexture "level36bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 36
        ElseIf Player.lScore > 37000 And Player.lScore < 38000 Then
            sField.CreateFromTexture "level37bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 37
        ElseIf Player.lScore > 38000 And Player.lScore < 39000 Then
            sField.CreateFromTexture "level38bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 38
        ElseIf Player.lScore > 39000 And Player.lScore < 40000 Then
            sField.CreateFromTexture "level39bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 39
        ElseIf Player.lScore > 40000 And Player.lScore < 41000 Then
            sField.CreateFromTexture "level40bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 40
        ElseIf Player.lScore > 41000 And Player.lScore < 42000 Then
            sField.CreateFromTexture "level41bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 41
        ElseIf Player.lScore > 42000 And Player.lScore < 43000 Then
            sField.CreateFromTexture "level42bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 42
        ElseIf Player.lScore > 43000 And Player.lScore < 44000 Then
            sField.CreateFromTexture "level43bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 43
        ElseIf Player.lScore > 44000 And Player.lScore < 45000 Then
            sField.CreateFromTexture "level44bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 44
        ElseIf Player.lScore > 45000 And Player.lScore < 46000 Then
            sField.CreateFromTexture "level45bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 45
        ElseIf Player.lScore > 46000 And Player.lScore < 47000 Then
            sField.CreateFromTexture "level46bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 46
        ElseIf Player.lScore > 47000 And Player.lScore < 48000 Then
            sField.CreateFromTexture "level47bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 47
        ElseIf Player.lScore > 48000 And Player.lScore < 49000 Then
            sField.CreateFromTexture "level48bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 48
        ElseIf Player.lScore > 49000 And Player.lScore < 50000 Then
            sField.CreateFromTexture "level49bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 49
        ElseIf Player.lScore > 50000 And Player.lScore < 51000 Then
            sField.CreateFromTexture "level50bg"
            sField.bCenterScale = False
            sField.SetPosition 128, 0
            Player.lLevel = 50
        ElseIf Player.lScore > 51000 Then
            DoEndGame = True
        End If
    End If
    'END TODO:
    sField.Render
    For x = 0 To 7
      For y = 0 To 8
        With lGrid(x, y)
          If .lDead > 0 Then 'this marble is to be removed, show its death animation
            Select Case .lType
              Case 7 'ghouls
                  sMarble(.lType - 1).SetPosition CSng(.lX) + 2, CSng(.lY) + 2 + (32 * (.lDead / 50))
                  sMarble(.lType - 1).SetScale 1, 1 * (1 - (.lDead / 50))
                  sMarble(.lType - 1).SetColor 255, 255, 255, tob(255 * (1 - (.lDead / 50)))
                  sMarble(.lType - 1).Render
              Case 1 'ice explosion
                  sMarble(.lType - 1).SetPosition CSng(.lX) + 2, CSng(.lY) + 2
                  sMarble(.lType - 1).SetScale 1, 1
                  sMarble(.lType - 1).SetColor 255, 255, 255, tob(255 * (1 - (.lDead / 50)))
                  sMarble(.lType - 1).Render
              Case 6 'splat
                  sMarble(.lType - 1).SetPosition CSng(.lX) + 2, CSng(.lY) + 2
                  sMarble(.lType - 1).SetScale 1, 1
                  sMarble(.lType - 1).SetColor 255, 255, 255, tob(255 * (1 - (.lDead / 50)))
                  sMarble(.lType - 1).Render
              Case 2 'scale down & rotate, sparkle
                  sMarble(.lType - 1).SetPosition CSng(.lX) + 2 + (32 * (.lDead / 50)), CSng(.lY) + 2 + (32 * (.lDead / 50))
                  sMarble(.lType - 1).SetRotation 360 * (.lDead / 50)
                  sMarble(.lType - 1).SetScale 1 * (1 - (.lDead / 50)), 1 * (1 - (.lDead / 50))
                  sMarble(.lType - 1).Render
              Case 5 'fairy 'mud splat
              sMarble(.lType - 1).SetPosition CSng(.lX) + 2, CSng(.lY) + 2 + (32 * (.lDead / 50))
                  sMarble(.lType - 1).SetScale 1, 1 * (1 - (.lDead / 50))
                  sMarble(.lType - 1).SetColor 255, 255, 255, tob(255 * (1 - (.lDead / 50)))
                  sMarble(.lType - 1).Render
                  'sMarble(.lType - 1).SetPosition CSng(.lX) + 2, CSng(.lY) + 2
                  'sMarble(.lType - 1).SetScale 1, 1
                  'sMarble(.lType - 1).SetColor 255, 255, 255, tob(255 * (1 - (.lDead / 50)))
                  'sMarble(.lType - 1).Render
              Case 3 'teleport out
                  sMarble(.lType - 1).SetPosition CSng(.lX) + 2 + (32 * (.lDead / 75)), CSng(.lY) + 2
                  sMarble(.lType - 1).SetScale 1 * (1 - (.lDead / 75)), 1
                  sMarble(.lType - 1).SetColor 255, 255, 255, 255
                  sMarble(.lType - 1).Render
              Case 4 'beep, boop
                  sMarble(.lType - 1).SetPosition CSng(.lX) + 2, CSng(.lY) + 2
                  sMarble(.lType - 1).SetScale 1, 1
                  sMarble(.lType - 1).SetColor 255, 255, 255, tob(255 * (1 - (.lDead / 35)))
                  sMarble(.lType - 1).Render
              Case 8 'explode
                  sMarble(.lType - 1).SetPosition CSng(.lX) + 2, CSng(.lY) + 2
                  sMarble(.lType - 1).SetScale 1, 1
                  sMarble(.lType - 1).SetColor 255, 255, 255, tob(255 * (1 - (.lDead / 50)))
                  sMarble(.lType - 1).Render
            End Select
          Else 'normal marble
            If lGrid(x, y).lType > 0 Then
              sMarble(.lType - 1).SetPosition CSng(.lX) + 2, CSng(.lY) + 2
              sMarble(.lType - 1).SetScale 1, 1
              sMarble(.lType - 1).SetRotation 0
              sMarble(.lType - 1).SetColor 255, 255, 255, tob((T2.GetTimeMs / 1000) * 255)
              sMarble(.lType - 1).Render
              If .bSpecial > 0 Then
                sSpecials(.bSpecial - 1).SetPosition CSng(.lX) + 34, CSng(.lY) + 34
                sSpecials(.bSpecial - 1).SetColor 255, 255, 255, tob((T2.GetTimeMs / 2000) * 200)
                sSpecials(.bSpecial - 1).SetScale 1, 1
                sSpecials(.bSpecial - 1).SetRotation 0
                sSpecials(.bSpecial - 1).Render
                If .bSpecial = 4 And .lFlag(3) > 0 Then
                  If timeGetTime - .lFlag(3) >= 10000 Then
                    .lFlag(3) = timeGetTime
                    If y > 0 Then
                      .lType = .lType + 1
                      If .lType > 8 Then .lType = 1
                      AddRing .lX + 34, .lY + 34
                      pSound.SfxPlayExID "spec4", LOOP_OFF, Settings.SfxVolume
                    End If
                  End If
                End If
              End If
            End If
          End If
        End With
      Next
    Next
    
    'Selection box
    If selx <> -1 And sely <> -1 Then
      sSelect.SetPosition 160 + selx * 64, 32 + sely * 64
      sSelect.SetColor 255, 255, 255, 55 + Abs(200 * Sine(T2.GetTimeMs / 5))
      sSelect.Render
    End If
    
    'Game Timer
    If game.GetTimeMs > 100 Then
      Player.lTime = Player.lTime - game.GetTimeMs
      game.StartTime
    End If
    If Player.lTime > 120000 Then Player.lTime = 120000
    If Player.lDisplayTime < Player.lTime - 150 Then
      Player.lDisplayTime = Player.lDisplayTime + 150
    ElseIf Player.lDisplayTime > Player.lTime + 150 Then
      Player.lDisplayTime = Player.lDisplayTime - 150
    Else
      Player.lDisplayTime = Player.lTime
    End If
    If Player.lDisplayTime < 15000 Then
      sHand.SetColor 255, 255, 255, 55 + Abs(200 * Sine(T2.GetTimeMs / 5))
    Else
      sHand.SetColor 255, 255, 255, 255
    End If
    If Player.lTime < 10000 And (lSnd(7) = 0 Or Not pSound.SfxIsPlaying(lSnd(7))) Then
      lSnd(7) = pSound.SfxPlayExID("beep", LOOP_OFF, Settings.SfxVolume)
    End If
    If Player.lDisplayTime <= 0 And Not moving And Not dead Then bEnd = 5
    sHand.SetRotation -360 * (Player.lDisplayTime / 120000)
    sHand.Render
            
    'game status (level, score, nearest highscore, bombs)
    LedFont.DrawText Player.lScore, ReturnRECT(20, 34, 115, 64), RGBA(0, 0, 0, 255), DT_RIGHT
    y = -1
    For x = 9 To 0 Step -1
      If High(x).lScore > Player.lScore Then
        y = x
        Exit For
      End If
    Next
    If y = -1 Then
      LedFont.DrawText Player.lScore, ReturnRECT(20, 98, 115, 128), RGBA(0, 0, 0, 255), DT_RIGHT
    Else
      LedFont.DrawText High(y).lScore, ReturnRECT(20, 98, 115, 128), RGBA(0, 0, 0, 255), DT_RIGHT
    End If
    LedFont.DrawText Player.lBombs, ReturnRECT(20, 168, 115, 198), RGBA(0, 0, 0, 255), DT_RIGHT
    'TODO:
    If Player.lLevel < 10 Then
    LedFont.DrawText "Level:0" & Player.lLevel, ReturnRECT(0, 338, 115, 198), RGBA(0, 0, 0, 255), DT_RIGHT
    Else
    LedFont.DrawText "Level:" & Player.lLevel, ReturnRECT(0, 338, 115, 198), RGBA(0, 0, 0, 255), DT_RIGHT
    End If
    If LevelandMoveTimer.GetTimeMs > 1000 Then
        If NoMoreMoves = True Then
            NoMoreMovesRes = True
        Else
            NoMoreMovesRes = False
        End If
        LevelandMoveTimer.StartTime
    End If
    If NoMoreMovesRes = True Then
        LedFont.DrawText "No Moves", ReturnRECT(0, 370, 115, 198), RGBA(0, 0, 0, 255), DT_RIGHT
        LedFont.DrawText "Left", ReturnRECT(0, 402, 115, 198), RGBA(0, 0, 0, 255), DT_RIGHT
        LedFont.DrawText "Please", ReturnRECT(0, 434, 115, 198), RGBA(0, 0, 0, 255), DT_RIGHT
        LedFont.DrawText "Restart", ReturnRECT(0, 466, 115, 198), RGBA(0, 0, 0, 255), DT_RIGHT
    End If
    'END TODO:

    'special effects... sparks
    For x = 0 To NumSparks
      sSpark(x).Render
    Next
    'snow
    For x = 0 To NumSnow
      sSnow(x).Render
    Next
    'ghouls!
    For x = 0 To NumGhouls
      sGhoul(x).Render
    Next
    'rings
    For x = 0 To NumRings
      sRing(x).Render
    Next
    'smoke
    For x = 0 To NumSmoke
      sSmoke(x).Render
    Next
    'splats from green marble
    For x = 0 To NumSplats
      sSplat(x).Render
    Next
    'stars
    For x = 0 To NumStars
      sStar(x).Render
    Next
    'Fairy 'mud
    For x = 0 To NumMud
      sMud(x).Render
    Next
    'For x = 0 To NumMud
     ' With sMud(x)
      '  If .Active Then
          'rotate mud splats to they always face towards their moving angle
       '   .SetRotation 360 - GetAngle(0, 0, .GetAutoMovement.x, .GetAutoMovement.y)
        '  .Render
        'End If
      'End With
    'Next
    'explosions
    For x = 0 To NumExplosions
      sExplosion(x).Render
    Next
    'floating scores
    For x = 0 To NumScores
      sScores(x).Render
    Next
    
    Cleanup
    
    'game menu
    sLogo.Render
    For x = 0 To 3
      sText_Menu(x).Render
    Next
    For x = 0 To 2
      sText_GameMenu(x).Render
    Next

    sCursor.Render
    
    If bFps Then
      FontArial.DrawText pEngine.lFPS, ReturnRECT(0, 0, 100, 15), RGBA(0, 255, 0, 255), DT_LEFT
    End If
    
    pEngine.Render
    '###End of drawing sequence
    
    If DoEndGame And T2.GetTimeMs > 1500 Then bEnd = 5
    
  Loop Until bEnd
  
  Select Case bEnd
    Case 1
      Unload Me
    Case 2
      GameMenu
      If bEnd = 2 Then
        bEnd = 0
        sLogo.SetAutoPath 900, 100, 1000
        sText_GameMenu(0).SetAutoPath 900, 200, 1000
        sText_GameMenu(1).SetAutoPath 900, 275, 1000
        sText_GameMenu(2).SetAutoPath 900, 350, 1000
        sLogo.SetAutoFade 255, 255, 255, 0, 1000
        For x = 0 To 2
          sText_GameMenu(x).SetAutoFade 255, 255, 255, 0, 1000
        Next
      ElseIf bEnd = 6 Then
        bEnd = 0
        sLogo.SetAutoPath 900, 100, 1000
        sText_GameMenu(0).SetAutoPath 900, 200, 1000
        sText_GameMenu(1).SetAutoPath 900, 275, 1000
        sText_GameMenu(2).SetAutoPath 900, 350, 1000
        sLogo.SetAutoFade 255, 255, 255, 0, 1000
        For x = 0 To 2
          sText_GameMenu(x).SetAutoFade 255, 255, 255, 0, 1000
        Next
        DoEndGame = True
      End If
      GoTo gAgain
    Case 5
      EndGame
      'TODO:
      sField.CreateFromTexture "field"
      sField.bCenterScale = False
      sField.SetPosition 128, 0
      'END TODO:
  End Select
  
  
End Sub

Public Function CheckMatches() As Boolean
  Dim l As Long, l2 As Long, x As Long, y As Long
  'vertical
  For x = 0 To 7
    l = 0
    l2 = 0
    For y = 1 To 8
      If lGrid(x, y).lType = l Then
        l2 = l2 + 1
      Else
        l = lGrid(x, y).lType
        l2 = 1
      End If
      If l2 = 3 Then
        CheckMatches = True
        Exit Function
      End If
    Next
    If l2 = 3 Then
      CheckMatches = True
      Exit Function
    End If
  Next
  'horizontal
  For y = 1 To 8
    l = 0
    l2 = 0
    For x = 0 To 7
      If lGrid(x, y).lType = l Then
        l2 = l2 + 1
      Else
        l = lGrid(x, y).lType
        l2 = 1
      End If
      If l2 = 3 Then
        CheckMatches = True
        Exit Function
      End If
    Next
    If l2 = 3 Then
      CheckMatches = True
      Exit Function
    End If
  Next
End Function

Public Function InitSfx() As Boolean
  'This function loads all sound effects & music
  On Error GoTo ErrH
  InitSfx = True
  
  'Load sfx
  pSound.SfxLoad App.Path & "\data\soundfx\land01.wav", "land01"
  pSound.SfxLoad App.Path & "\data\soundfx\land02.wav", "land02"
  pSound.SfxLoad App.Path & "\data\soundfx\smash01.wav", "smash01"
  pSound.SfxLoad App.Path & "\data\soundfx\smash02.wav", "smash02"
  pSound.SfxLoad App.Path & "\data\soundfx\smash03.wav", "smash03"
  pSound.SfxLoad App.Path & "\data\soundfx\spark01.wav", "spark01"
  pSound.SfxLoad App.Path & "\data\soundfx\spark02.wav", "spark02"
  pSound.SfxLoad App.Path & "\data\soundfx\deselect.wav", "deselect"
  pSound.SfxLoad App.Path & "\data\soundfx\switch.wav", "switch"
  pSound.SfxLoad App.Path & "\data\soundfx\explosion01.wav", "explosion01"
  pSound.SfxLoad App.Path & "\data\soundfx\explosion02.wav", "explosion02"
  pSound.SfxLoad App.Path & "\data\soundfx\explosion03.wav", "explosion03"
  pSound.SfxLoad App.Path & "\data\soundfx\ghouls01.wav", "ghouls01"
  pSound.SfxLoad App.Path & "\data\soundfx\ghouls02.wav", "ghouls02"
  pSound.SfxLoad App.Path & "\data\soundfx\teleport.wav", "teleport"
  pSound.SfxLoad App.Path & "\data\soundfx\splat01.wav", "splat01"
  pSound.SfxLoad App.Path & "\data\soundfx\splat02.wav", "splat02"
  pSound.SfxLoad App.Path & "\data\soundfx\beep.wav", "beep"
  pSound.SfxLoad App.Path & "\data\soundfx\boop.wav", "boop"
  pSound.SfxLoad App.Path & "\data\soundfx\highscore.wav", "highscore"
  pSound.SfxLoad App.Path & "\data\soundfx\bonus.wav", "bonus"
  pSound.SfxLoad App.Path & "\data\soundfx\select.wav", "select"
  pSound.SfxLoad App.Path & "\data\soundfx\end.wav", "end"
  pSound.SfxLoad App.Path & "\data\soundfx\mud.wav", "mud"
  pSound.SfxLoad App.Path & "\data\soundfx\spec4.wav", "spec4"
  pSound.SfxLoad App.Path & "\data\soundfx\spec6.wav", "spec6"
  
  pSound.MusicLoad App.Path & "\data\music\00.mod", "0", False
  pSound.MusicLoad App.Path & "\data\music\01.mod", "1", False
  pSound.MusicLoad App.Path & "\data\music\02.mod", "2", False
  pSound.MusicLoad App.Path & "\data\music\03.mod", "3", False
  pSound.MusicLoad App.Path & "\data\music\04.mod", "4", False
  pSound.MusicLoad App.Path & "\data\music\05.mod", "5", False
  pSound.MusicLoad App.Path & "\data\music\06.mod", "6", False
  pSound.MusicLoad App.Path & "\data\music\07.mod", "7", False
  pSound.MusicLoad App.Path & "\data\music\08.mod", "8", False
  pSound.MusicLoad App.Path & "\data\music\09.mod", "9", False
  pSound.MusicLoad App.Path & "\data\music\10.mod", "10", False
  pSound.MusicLoad App.Path & "\data\music\11.mod", "11", False
  pSound.MusicLoad App.Path & "\data\music\12.mod", "12", False
  pSound.MusicLoad App.Path & "\data\music\13.mod", "13", False
  pSound.MusicLoad App.Path & "\data\music\14.mod", "14", False
  
  CurrentMusic = Int(Rnd * MaxMusic) + 1

  Exit Function
ErrH:
  MsgBox "Could not init game sound data. Please Reinstall.", vbCritical, "Error"
  InitSfx = False
End Function

Public Sub AddStar(ByVal x As Single, ByVal y As Single)
  Dim f As Long, num As Long
  
  For f = 0 To NumStars
    If sStar(f).Active = False Then
      Set sStar(f) = New pgeSprite
      num = f
      Exit For
    End If
  Next
  
  If num = 0 Then
    NumStars = NumStars + 1
    ReDim Preserve sStar(NumStars)
    num = NumStars
  End If
  
  With sStar(num)
    .CreateFromTexture "star"
    .SetColor 255, 255, 255, 200
    .SetPosition x, y
    '.SetAutoRotation 1 + Rnd * 3, 10
    .SetScale 0.1, 0.1
    .SetAutoScale 1, 1, 400
    .SetAutoFade 255, 255, 255, 0, 2500 + Int(Rnd * 2000), True
    '.SetAutoMovement -RotatePixel(angledir, 1 + Rnd * 1).x, -RotatePixel(angledir, 1 + Rnd * 1).y, 20
  End With

End Sub

Public Sub AddSpark(ByVal x As Single, ByVal y As Single, ByVal angledir As Single)
  Dim f As Long, num As Long
  
  For f = 0 To NumSparks
    If sSpark(f).Active = False Then
      Set sSpark(f) = New pgeSprite
      num = f
      Exit For
    End If
  Next
  
  If num = 0 Then
    NumSparks = NumSparks + 1
    ReDim Preserve sSpark(NumSparks)
    num = NumSparks
  End If
  
  With sSpark(num)
    .CreateFromTexture "spark2"
    .SetColor 255, 255, 255, 255
    .SetPosition x, y
    .SetAutoRotation 1 + Rnd * 3, 10
    .SetScale 0.1, 0.1
    .SetAutoScale 1, 1, 400
    .SetAutoFade 255, 255, 255, 0, 700, True
    .SetAutoMovement -RotatePixel(angledir, 1 + Rnd * 1).x, -RotatePixel(angledir, 1 + Rnd * 1).y, 20
  End With

End Sub

Public Sub AddMud(ByVal x As Single, ByVal y As Single, ByVal angledir As Single)
  Dim f As Single, num As Long
  
  For f = 0 To NumMud
    If sMud(f).Active = False Then
      Set sMud(f) = New pgeSprite
      num = f
      Exit For
    End If
  Next
  
  If num = 0 Then
    NumMud = NumMud + 1
    ReDim Preserve sMud(NumMud)
    num = NumMud
  End If
  
  With sMud(num)
    .CreateFromTexture "mud"
    .SetColor 255, 255, 255, 200
    .SetPosition x, y
    .SetRotation angledir
    .SetScale 0, 0
    .SetAutoScale 0.5, 0.5, 300
    .SetAutoFade 255, 255, 255, 0, 4000, True
    f = 2 + Rnd * 3
    .SetAutoMovement -RotatePixel(angledir, f).x, -RotatePixel(angledir, f).y, 20
  End With

End Sub

Public Sub AddSplat(ByVal x As Single, ByVal y As Single, ByVal angledir As Single)
  Dim f As Single, num As Long
  
  For f = 0 To NumSplats
    If sSplat(f).Active = False Then
      Set sSplat(f) = New pgeSprite
      num = f
      Exit For
    End If
  Next
  
  If num = 0 Then
    NumSplats = NumSplats + 1
    ReDim Preserve sSplat(NumSplats)
    num = NumSplats
  End If
  
  With sSplat(num)
    .CreateFromTexture "splat"
    .SetColor 255, 255, 255, 255
    .SetPosition x, y
    .SetRotation angledir
    .SetScale 0, 0
    .SetAutoScale 1, 1, 300
    .SetAutoFade 255, 255, 255, 0, 1000, True
    f = 2 + Rnd * 3
    .SetAutoMovement -RotatePixel(angledir, f).x, -RotatePixel(angledir, f).y, 20
  End With

End Sub

Public Sub AddSnow(ByVal x As Single, ByVal y As Single, ByVal angledir As Single)
  Dim f As Long, num As Long
  
  For f = 0 To NumSnow
    If sSnow(f).Active = False Then
      Set sSnow(f) = New pgeSprite
      num = f
      Exit For
    End If
  Next
  
  If num = 0 Then
    NumSnow = NumSnow + 1
    ReDim Preserve sSnow(NumSnow)
    num = NumSnow
  End If
  
  With sSnow(num)
    .CreateFromTexture "snow"
    .SetColor 255, 255, 255, 255
    .SetPosition x, y
    '.SetAutoRotation 1 + Rnd * 3, 10
    .SetScale 1, 1
    '.SetAutoScale 1, 1, 250
    .SetAutoFade 255, 255, 255, 0, 1000, True
    .SetAutoMovement -RotatePixel(angledir, 1 + Rnd * 1).x, -RotatePixel(angledir, 1 + Rnd * 1).y, 20
  End With

End Sub

Public Sub AddGhoul(ByVal x As Single, ByVal y As Single, ByVal angledir As Single)
  Dim f As Single, num As Long
  
  For f = 0 To NumGhouls
    If sGhoul(f).Active = False Then
      Set sGhoul(f) = New pgeSprite
      num = f
      Exit For
    End If
  Next
  
  If num = 0 Then
    NumGhouls = NumGhouls + 1
    ReDim Preserve sGhoul(NumGhouls)
    num = NumGhouls
  End If
  
  With sGhoul(num)
    .CreateFromTexture "ghoul"
    .SetColor 255, 255, 255, 200
    .SetPosition x, y
    .SetScale 0, 1
    .SetRotation angledir
    .SetAutoScale 1, 1 + Rnd * 1, 500
    .SetAutoFade 255, 255, 255, 0, 3000 + Int(Rnd * 1000), True
    f = 2 + Rnd * 3
    .SetAutoMovement -RotatePixel(angledir, f).x, -RotatePixel(angledir, f).y, 20
  End With

End Sub

Public Sub AddSmoke(ByVal x As Single, ByVal y As Single)
  Dim f As Long, num As Long
  
  For f = 0 To NumSmoke
    If sSmoke(f).Active = False Then
      Set sSmoke(f) = New pgeSprite
      num = f
      Exit For
    End If
  Next
  
  If num = 0 Then
    NumSmoke = NumSmoke + 1
    ReDim Preserve sSmoke(NumSmoke)
    num = NumSmoke
  End If
  
  With sSmoke(num)
    .CreateFromTexture "smoke"
    .SetColor 255, 255, 255, 100
    .SetPosition x, y
    .SetScale 0.1, 0.1
    .SetAutoScale 1.5, 1.5, 1000
    .SetAutoFade 255, 255, 255, 0, 2000 + Int(Rnd * 2000), True
    .SetAutoMovement -1 + Rnd * 2, -0.1 - Rnd * 2, 50
  End With

End Sub

Public Sub AddRing(ByVal x As Single, ByVal y As Single)
  Dim f As Long, num As Long
  
  For f = 0 To NumRings
    If sRing(f).Active = False Then
      Set sRing(f) = New pgeSprite
      num = f
      Exit For
    End If
  Next
  
  If num = 0 Then
    NumRings = NumRings + 1
    ReDim Preserve sRing(NumRings)
    num = NumRings
  End If
  
  With sRing(num)
    .CreateFromTexture "ring"
    .SetColor 255, 255, 255, 200
    .SetPosition x, y
    .SetScale 0, 0
    .SetAutoScale 5, 5, 3000
    .SetAutoFade 255, 255, 255, 0, 2000, True
  End With

End Sub

Public Sub AddIceExplosion(ByVal x As Single, ByVal y As Single, Optional ByVal scl As Single = 1)
  Dim f As Long, num As Long
  
  For f = 0 To NumExplosions
    If sExplosion(f).Active = False Then
      Set sExplosion(f) = New pgeSprite
      num = f
      Exit For
    End If
  Next
  
  If num = 0 Then
    NumExplosions = NumExplosions + 1
    ReDim Preserve sExplosion(NumExplosions)
    num = NumExplosions
  End If
  
  With sExplosion(num)
    .CreateFromTexture "explosion2", 1, 16, 75
    .bAnimStop = True
    .SetPosition x, y
    .SetScale scl, scl
    .SetColor 255, 255, 255, 200
    .SetAutoFade 255, 255, 255, 0, 1000, False
  End With
  
End Sub

Public Sub AddExplosion(ByVal x As Single, ByVal y As Single, Optional ByVal scl As Single = 1)
  Dim f As Long, num As Long
  
  For f = 0 To NumExplosions
    If sExplosion(f).Active = False Then
      Set sExplosion(f) = New pgeSprite
      num = f
      Exit For
    End If
  Next
  
  If num = 0 Then
    NumExplosions = NumExplosions + 1
    ReDim Preserve sExplosion(NumExplosions)
    num = NumExplosions
  End If
  
  With sExplosion(num)
    .CreateFromTexture "explosion", 1, 16, 75
    .bAnimStop = True
    .SetPosition x, y
    .SetScale scl, scl
    .SetColor 255, 255, 255, 200
    .SetAutoFade 255, 255, 255, 0, 1000, False
  End With
  
End Sub

Public Sub AddScore(ByVal x As Single, ByVal y As Single, ByVal score As String)
  Dim f As Long, num As Long
  
  For f = 0 To NumScores
    If sScores(f).Active = False Then
      Set sScores(f) = New pgeSprite
      num = f
      Exit For
    End If
  Next
  
  If num = 0 Then
    NumScores = NumScores + 1
    ReDim Preserve sScores(NumScores)
    num = NumScores
  End If
  
  With sScores(num)
    .CreateFromTexture score
    .SetPosition x, y
    .SetScale 1, 0
    .SetAutoScale 1, 1, 200
    .SetColor 255, 255, 255, 255
    .SetAutoFade 255, 255, 255, 0, 2000, True
    .SetAutoMovement 0, -0.1, 10
  End With
  
End Sub

Public Sub HandleScore(ByVal x As Long, ByVal y As Long, ByVal num As Long, ByVal bonus As Long)
  Dim T As Long, l As Long
  Select Case num
    Case 3
      AddScore lGrid(x, y).lX, lGrid(x, y).lY + 32, "30"
    Case 4
      AddScore lGrid(x, y).lX, lGrid(x, y).lY + 32, "40"
    Case 5
      AddScore lGrid(x, y).lX, lGrid(x, y).lY + 32, "50"
  End Select
  If bonus > 1 Then
    T = pSound.SfxPlayExID("bonus", LOOP_OFF, Settings.SfxVolume)
    pSound.SfxChangePlaying T, 21050 + bonus * 1000, -1, -1
  End If
  Select Case bonus
    Case 2
      AddScore lGrid(x, y).lX + 48, lGrid(x, y).lY + 32, "x2"
      Player.lTime = Player.lTime + 1500 * num
    Case 3
      AddScore lGrid(x, y).lX + 48, lGrid(x, y).lY + 32, "x3"
      Player.lTime = Player.lTime + 1500 * num
    Case 4
      AddScore lGrid(x, y).lX + 48, lGrid(x, y).lY + 32, "x4"
      Player.lTime = Player.lTime + 2000 * num
    Case 5
      AddScore lGrid(x, y).lX + 48, lGrid(x, y).lY + 32, "x5"
      Player.lTime = Player.lTime + 2000 * num
      Player.lBombs = Player.lBombs + 1
      AddScore 105, 198, "bomb"
    Case 6
      AddScore lGrid(x, y).lX + 48, lGrid(x, y).lY + 32, "x6"
      Player.lTime = Player.lTime + 2500 * num
    Case 7
      AddScore lGrid(x, y).lX + 48, lGrid(x, y).lY + 32, "x7"
      Player.lTime = Player.lTime + 2500 * num
    Case 8
      AddScore lGrid(x, y).lX + 48, lGrid(x, y).lY + 32, "x8"
      Player.lTime = Player.lTime + 3500 * num
      Player.lBombs = Player.lBombs + 1
      AddScore 105, 198, "bomb"
    Case 9
      AddScore lGrid(x, y).lX + 84, lGrid(x, y).lY + 32, "x9"
      Player.lTime = Player.lTime + 3500 * num
    Case 10
      AddScore lGrid(x, y).lX + 84, lGrid(x, y).lY + 32, "x10"
      Player.lTime = Player.lTime + 5000 * num
      Player.lBombs = Player.lBombs + 1
      AddScore 105, 198, "bomb"
  End Select
  
  If Player.lTime > 120000 Then
    Player.lScore = Player.lScore + 70 * bonus
  End If
End Sub

Public Sub GameMenu()
  'This is the in game menu
  
  Dim T As New pgeTimer
  Dim tRect As RECT
  Dim x As Long, y As Long
    
  T.StartTime
mAgain:
  bEnd = 0
    
  'setup graphics needed in this screen.
  sLogo.SetPosition 900, 100
  sText_GameMenu(0).SetPosition 900, 200
  sText_GameMenu(1).SetPosition 900, 275
  sText_GameMenu(2).SetPosition 900, 350
  sText_GameMenu(3).SetPosition 900, 425
  sLogo.SetAutoPath 384, 100, 1000
  sText_GameMenu(0).SetAutoPath 384, 200, 1000
  sText_GameMenu(1).SetAutoPath 384, 275, 1000
  sText_GameMenu(2).SetAutoPath 384, 350, 1000
   sText_GameMenu(3).SetAutoPath 384, 425, 1000
  sLogo.SetColor 255, 255, 255, 0
  sLogo.SetAutoFade 255, 255, 255, 255, 1000
  For x = 0 To 3
    sText_GameMenu(x).SetColor 255, 255, 255, 0
    sText_GameMenu(x).SetAutoFade 255, 255, 255, 255, 1000
  Next
  
  Do
    DoEvents
    
    'Poll & setup mouse pointer
    pMouse.Poll Me
    sCursor.SetPosition pMouse.g_cursorx, pMouse.g_cursory
    tRect.Left = pMouse.g_cursorx
    tRect.Top = pMouse.g_cursory
    tRect.Right = tRect.Left + 1
    tRect.bottom = tRect.Top + 1
    
    'Handle input, but only after menu items have appeared
    If T.GetTime > 1 Then
      'Check mouse vs menu options
      For x = 0 To 3
        If IntersectR(tRect, sText_GameMenu(x).GetDestRect) Then
          sText_GameMenu(x).SetColor 255, 255, 255, 100 + Abs(155 * Sine(T.GetTimeMs / 5))
          If pMouse.button1 Then
            Select Case x
              Case 0
                bEnd = 2
              Case 1
                bEnd = 3
                Case 2
                 Call CreditsAbout
              Case 3
                bEnd = 4
            End Select
          End If
        Else
          sText_GameMenu(x).SetColor 255, 255, 255, 255
        End If
      Next
      'Check keyboard
      If pKeyboard.KeyDown(DIK_ESCAPE) Then bEnd = 2
    End If
    
    
    '###Drawing sequence
    pEngine.Clear
    
    sBar.Render
    sField.Render
    sHand.Render
    
    For x = 0 To 7
      For y = 0 To 8
        With lGrid(x, y)
          If .lType > 0 Then
            sMarble(.lType - 1).SetPosition CSng(.lX) + 2, CSng(.lY) + 2
            sMarble(.lType - 1).SetScale 1, 1
            sMarble(.lType - 1).SetRotation 0
            sMarble(.lType - 1).SetColor 255, 255, 255, 255 - tob((T.GetTimeMs / 1000) * 255)
            sMarble(.lType - 1).Render
          End If
        End With
      Next
    Next
    
    LedFont.DrawText Player.lScore, ReturnRECT(20, 34, 115, 64), RGBA(0, 0, 0, 255), DT_RIGHT
    y = -1
    For x = 9 To 0 Step -1
      If High(x).lScore > Player.lScore Then
        y = x
        Exit For
      End If
    Next
    If y = -1 Then
      LedFont.DrawText Player.lScore, ReturnRECT(20, 98, 115, 128), RGBA(0, 0, 0, 255), DT_RIGHT
    Else
      LedFont.DrawText High(y).lScore, ReturnRECT(20, 98, 115, 128), RGBA(0, 0, 0, 255), DT_RIGHT
    End If
    LedFont.DrawText Player.lBombs, ReturnRECT(20, 168, 115, 198), RGBA(0, 0, 0, 255), DT_RIGHT
    
    sLogo.Render
    For x = 0 To 3
      sText_GameMenu(x).Render
    Next
    For x = 0 To 4
      sText_Settings(x).Render
    Next
    For x = 0 To 2
      sMeter(x).Render
      sMeterDrag(x).Render
    Next
    
    sCursor.Render
    
    If bFps Then
      FontArial.DrawText pEngine.lFPS, ReturnRECT(0, 0, 100, 15), RGBA(0, 255, 0, 255), DT_LEFT
    End If
    
    pEngine.Render
    '###End of drawing sequence
  Loop Until bEnd
  
  Select Case bEnd
    Case 1
      Unload Me
    Case 3
      sText_GameMenu(0).SetAutoPath 900, 200, 1000
      sText_GameMenu(1).SetAutoPath 900, 275, 1000
      sText_GameMenu(2).SetAutoPath 900, 350, 1000
      sText_GameMenu(3).SetAutoPath 900, 425, 1000
      For x = 0 To 2
        sText_GameMenu(x).SetAutoFade 255, 255, 255, 0, 1000
      Next
      Options
      GoTo mAgain
    Case 4
      sText_GameMenu(0).SetAutoPath 900, 200, 1000
      sText_GameMenu(1).SetAutoPath 900, 275, 1000
      sText_GameMenu(2).SetAutoPath 900, 350, 1000
       sText_GameMenu(3).SetAutoPath 900, 425, 1000
      For x = 0 To 3
        sText_GameMenu(x).SetAutoFade 255, 255, 255, 0, 1000
      Next
      bEnd = 6
  End Select

End Sub

Public Sub EndGame()
  'This sub end the current game
  bEnd = 0
  
  Dim T As New pgeTimer
  Dim T2 As New pgeTimer
  Dim tRect As RECT
  Dim x As Long, y As Long, z As Long
  
  pSound.MusicPlayID "0", Settings.MusicVolume
  Call pSound.SfxPlayExID("end", LOOP_OFF, Settings.SfxVolume)
    
  'setup graphics needed in this screen.
  sGameOver.SetPosition 384, -450
  sGameOver.SetAutoPath 384, 100, 1000
  sGameOver.SetColor 255, 255, 255, 0
  sGameOver.SetAutoFade 255, 255, 255, 255, 1000
  
  T.StartTime
  T2.StartTime
  Do
    DoEvents
    
    pMouse.Poll Me
    sCursor.SetPosition pMouse.g_cursorx, pMouse.g_cursory
    
    If T.GetTimeMs >= 4000 Then bEnd = 2
    
    '###Drawing sequence
    pEngine.Clear
    
    sBar.Render
    sField.Render
    
    sHand.SetColor 255, 255, 255, 255 - tob((T.GetTimeMs / 2000) * 255)
    sHand.Render
    
    If T2.GetTimeMs > 10 Then
      z = (T.GetTimeMs / 100) - 1
      If z > 7 Then z = 7
      For x = 0 To z
        For y = 1 To 8
          lGrid(x, y).lY = lGrid(x, y).lY + (T.GetTimeMs / 50) - z
        Next
      Next
      T2.StartTime
    End If
    
    For x = 0 To 7
      For y = 0 To 8
        With lGrid(x, y)
          If .lType > 0 Then
            sMarble(.lType - 1).SetPosition CSng(.lX) + 2, CSng(.lY) + 2
            sMarble(.lType - 1).SetScale 1, 1
            sMarble(.lType - 1).SetRotation 0
            sMarble(.lType - 1).Render
          End If
        End With
      Next
    Next
    
    LedFont.DrawText Player.lScore, ReturnRECT(20, 34, 115, 64), RGBA(0, 0, 0, 0 - tob((T.GetTimeMs / 2000) * 255)), DT_RIGHT
    y = -1
    For x = 9 To 0 Step -1
      If High(x).lScore > Player.lScore Then
        y = x
        Exit For
      End If
    Next
    If y = -1 Then
      LedFont.DrawText Player.lScore, ReturnRECT(20, 98, 115, 128), RGBA(0, 0, 0, 0 - tob((T.GetTimeMs / 2000) * 255)), DT_RIGHT
    Else
      LedFont.DrawText High(y).lScore, ReturnRECT(20, 98, 115, 128), RGBA(0, 0, 0, 0 - tob((T.GetTimeMs / 2000) * 255)), DT_RIGHT
    End If
    LedFont.DrawText Player.lBombs, ReturnRECT(20, 168, 115, 198), RGBA(0, 0, 0, 0 - tob((T.GetTimeMs / 2000) * 255)), DT_RIGHT
    
    sGameOver.Render
    
    sCursor.Render
    
    If bFps Then
      FontArial.DrawText pEngine.lFPS, ReturnRECT(0, 0, 100, 15), RGBA(0, 255, 0, 255), DT_LEFT
    End If
    
    pEngine.Render
    '###End of drawing sequence
  Loop Until bEnd
  
  Select Case bEnd
    Case 1
      Unload Me
    Case 2
      y = -1
      For x = 9 To 0 Step -1
        If High(x).lScore > Player.lScore Then
          y = x
          Exit For
        End If
      Next
      If y < 9 Then GotHighscore
  End Select
  
  Player.lTime = 0
  
  sGameOver.SetAutoPath 384, -450, 1000
  sGameOver.SetAutoFade 255, 255, 255, 0, 1000
  
  If y < 9 Then Highscore
  
End Sub

Public Sub GotHighscore()
  'This sub is the enter highscore screen
  bEnd = 0
  
  Dim T As New pgeTimer
  Dim tRect As RECT
  Dim x As Long, y As Long, z As Long
  pKeyboard.SetTimer 0.2
  pKeyboard.SetTimerEx DIK_LSHIFT, 0
  pKeyboard.SetTimerEx DIK_RSHIFT, 0
  pKeyboard.SetTimerEx DIK_BACKSPACE, 0.05
  
  'setup graphics needed in this screen.
  sQualify.SetPosition 384, 300
  sQualify.SetColor 255, 255, 255, 0
  sQualify.SetAutoFade 255, 255, 255, 255, 1000
  sQualify.SetScale 0, 0
  sQualify.SetAutoScale 1, 1, 2000
  
  Call pSound.SfxPlayExID("highscore", LOOP_OFF, Settings.SfxVolume)
  
  sText = ""
  bText = True 'get letters from Form_KeyPress
  T.StartTime
  Do
    DoEvents
    
    pMouse.Poll Me
    sCursor.SetPosition pMouse.g_cursorx, pMouse.g_cursory
    
    '###Handle keyboard
    With pKeyboard
      If .KeyDown(DIK_BACKSPACE) Then 'erase
        If Len(sText) > 0 Then
          sText = Left(sText, Len(sText) - 1)
        End If
      ElseIf .KeyDown(DIK_RETURN) Then 'done
        If Len(sText) > 0 Then bEnd = 2
      ElseIf .KeyDown(DIK_SPACE) Then 'space
        sText = sText & " "
      End If
    End With
    
    '###Drawing sequence
    pEngine.Clear
    
    sBar.Render
    sField.Render

    sGameOver.Render
    sQualify.Render
    
    MainFont.DrawText sText & "<", ReturnRECT(180, 415, 575, 500), RGBA(255, 255, 255, tob((T.GetTimeMs / 2000) * 255)), DT_CENTER Or DT_NOPREFIX
    
    sCursor.Render
    
    If bFps Then
      FontArial.DrawText pEngine.lFPS, ReturnRECT(0, 0, 100, 15), RGBA(0, 255, 0, 255), DT_LEFT
    End If
    
    pEngine.Render
    '###End of drawing sequence
  Loop Until bEnd
  
  bText = False
  
  pKeyboard.SetTimer 0.1
  
  For x = 9 To 0 Step -1
    If High(x).lScore > Player.lScore Then
      y = x + 1
      Exit For
    End If
  Next
  LatestHigh = y
  For x = 9 To y + 1 Step -1
    High(x).lScore = High(x - 1).lScore
    High(x).sName = High(x - 1).sName
  Next
  High(y).lScore = Player.lScore
  High(y).sName = sText
  
  sQualify.SetAutoFade 255, 255, 255, 0, 1000
  sQualify.SetAutoScale 0, 0, 2000
  
  Select Case bEnd
    Case 1
      Unload Me
  End Select
  
End Sub

Public Sub SaveScores()
  Dim file As New clsDatafile, x As Long
  file.fileName = App.Path & "\Config\HighScores.ini"
  For x = 0 To 9
    file.WriteStr High(x).sName
    file.WriteNumber High(x).lScore
  Next
End Sub

Public Sub SaveSettings()
  Dim file As New clsDatafile, x As Long
  file.fileName = App.Path & "\Config\Settings.ini"
    file.WriteNumber Settings.SfxVolume
  file.WriteNumber Settings.MusicVolume
  file.WriteNumber Settings.MouseSpeed * 10000
  file.WriteNumber Abs(bFps)
End Sub

Public Sub LoadSettings()
  Dim file As New clsDatafile, x As Long
  If FileExist(App.Path & "\Config\Settings.ini") Then
    file.fileName = App.Path & "\Config\Settings.ini"
    Settings.SfxVolume = file.ReadNumber
    Settings.MusicVolume = file.ReadNumber
    Settings.MouseSpeed = file.ReadNumber / 10000
    bFps = CBool(file.ReadNumber)
  Else
    Settings.SfxVolume = 200
    Settings.MusicVolume = 200
    Settings.MouseSpeed = 1.5
  End If
  pMouse.g_Sensitivity = Settings.MouseSpeed
End Sub

Public Sub Cleanup()
  'this cleans up the various special effect arrays,
  'because they slow down quite a bit when not used
  Dim x As Long
  
  'sparks
  For x = UBound(sSpark) To 1 Step -1
    If sSpark(x).Active = False Then
      ReDim Preserve sSpark(x - 1)
      NumSparks = x - 1
    Else
      Exit For
    End If
  Next
  
  'snow
  For x = UBound(sSnow) To 1 Step -1
    If sSnow(x).Active = False Then
      ReDim Preserve sSnow(x - 1)
      NumSnow = x - 1
    Else
      Exit For
    End If
  Next
  
  'explosions
  For x = UBound(sExplosion) To 1 Step -1
    If sExplosion(x).Active = False Then
      ReDim Preserve sExplosion(x - 1)
      NumExplosions = x - 1
    Else
      Exit For
    End If
  Next
  
  'ghouls!
  For x = UBound(sGhoul) To 1 Step -1
    If sGhoul(x).Active = False Then
      ReDim Preserve sGhoul(x - 1)
      NumGhouls = x - 1
    Else
      Exit For
    End If
  Next
  
  'smoke
  For x = UBound(sSmoke) To 1 Step -1
    If sSmoke(x).Active = False Then
      ReDim Preserve sSmoke(x - 1)
      NumSmoke = x - 1
    Else
      Exit For
    End If
  Next
  
  'splats from green marble
  For x = UBound(sSplat) To 1 Step -1
    If sSplat(x).Active = False Then
      ReDim Preserve sSplat(x - 1)
      NumSplats = x - 1
    Else
      Exit For
    End If
  Next
  
  'stars
  For x = UBound(sStar) To 1 Step -1
    If sStar(x).Active = False Then
      ReDim Preserve sStar(x - 1)
      NumStars = x - 1
    Else
      Exit For
    End If
  Next
  
  'rings
  For x = UBound(sRing) To 1 Step -1
    If sRing(x).Active = False Then
      ReDim Preserve sRing(x - 1)
      NumRings = x - 1
    Else
      Exit For
    End If
  Next
  
  'mud
  For x = UBound(sMud) To 1 Step -1
    If sMud(x).Active = False Then
      ReDim Preserve sMud(x - 1)
      NumMud = x - 1
    Else
      Exit For
    End If
  Next
  
  'floating scores
  For x = UBound(sScores) To 1 Step -1
    If sScores(x).Active = False Then
      ReDim Preserve sScores(x - 1)
      NumScores = x - 1
    Else
      Exit For
    End If
  Next
    
End Sub

Public Sub Credits()
  'This is the credits screen
  bEnd = 0
  
  Dim T As New pgeTimer
  Dim x As Long

  sCredits.SetPosition 900, 100
  sCredits.SetAutoPath 384, 100, 1000
  sCredits.SetColor 255, 255, 255, 0
  sCredits.SetAutoFade 255, 255, 255, 255, 1000
  
  sCredits_Detail.SetPosition 900, 320
  sCredits_Detail.SetAutoPath 384, 320, 1000
  sCredits_Detail.SetColor 255, 255, 255, 0
  sCredits_Detail.SetAutoFade 255, 255, 255, 255, 1000
  
  sLogo.SetAutoPath 900, 100, 1000
  sText_Menu(0).SetAutoPath 900, 200, 1000
  sText_Menu(1).SetAutoPath 900, 275, 1000
  sText_Menu(2).SetAutoPath 900, 350, 1000
  sText_Menu(3).SetAutoPath 900, 425, 1000
  sLogo.SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(0).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(1).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(2).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(3).SetAutoFade 255, 255, 255, 0, 1000

  T.StartTime
  Do
    DoEvents

    '###Drawing sequence
    pEngine.Clear
    
    sBar.Render
    sField.Render
    sLogo.Render
    For x = 0 To 3
      sText_Menu(x).Render
    Next
    sCredits.Render
    sCredits_Detail.Render
    
    If T.GetTimeMs >= 5000 Then bEnd = 1
    'If t.GetTimeMs > 1000 And pKeyboard.KeyDown(DIK_ESCAPE) Then bEnd = 1
    
    If bFps Then
      FontArial.DrawText pEngine.lFPS, ReturnRECT(0, 0, 100, 15), RGBA(0, 255, 0, 255), DT_LEFT
    End If
    
    pEngine.Render
    '###End of drawing sequence

  Loop Until bEnd

  Unload Me

End Sub
Public Sub CreditsAbout()
  'This is the credits screen
  bEnd = 0
  Dim tRect As RECT
  Dim x As Long, y As Long
  Dim T As New pgeTimer
  'Dim x As Long
  sCreditsinfotitle.SetPosition 900, 100
  sCreditsinfotitle.SetAutoPath 384, 100, 1000
  sCreditsinfotitle.SetColor 255, 255, 255, 0
  sCreditsinfotitle.SetAutoFade 255, 255, 255, 255, 1000
  
  sCredits_Info.SetPosition 900, 320
  sCredits_Info.SetAutoPath 384, 320, 1000
  sCredits_Info.SetColor 255, 255, 255, 0
  sCredits_Info.SetAutoFade 255, 255, 255, 255, 1000
  
  sLogo.SetAutoPath 900, 100, 1000
  sText_Menu(0).SetAutoPath 900, 200, 1000
  sText_Menu(1).SetAutoPath 900, 275, 1000
  sText_Menu(2).SetAutoPath 900, 350, 1000
  sText_Menu(3).SetAutoPath 900, 425, 1000
  sLogo.SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(0).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(1).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(2).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(3).SetAutoFade 255, 255, 255, 0, 1000

  T.StartTime
  Do
    DoEvents
pMouse.Poll Me
    sCursor.SetPosition pMouse.g_cursorx, pMouse.g_cursory
    tRect.Left = pMouse.g_cursorx
    tRect.Top = pMouse.g_cursory
    tRect.Right = tRect.Left + 1
    tRect.bottom = tRect.Top + 1

    '###Drawing sequence
    pEngine.Clear
     sCursor.Render
    sBar.Render
    sField.Render
    sLogo.Render
    For x = 0 To 4
      sText_Menu(x).Render
    Next
    sCreditsinfotitle.Render
    sCredits_Info.Render
    
   If T.GetTimeMs >= 1000000000 Then Call GameMenu
    'If t.GetTimeMs > 1000 And pKeyboard.KeyDown(DIK_ESCAPE) Then bEnd = 1
     sCursor.Render
    If bFps Then
      FontArial.DrawText pEngine.lFPS, ReturnRECT(0, 0, 100, 15), RGBA(0, 255, 0, 255), DT_LEFT
    End If
    
    pEngine.Render
    ' sText_Settings(4).CreateFromTexture "back"
    '###End of drawing sequence
If pKeyboard.KeyDown(DIK_ESCAPE) Then Call GameMenu
'Do
  Loop Until bEnd

 'Check keyboard
     ' If pKeyboard.KeyDown(DIK_ESCAPE) Then bEnd = 2

End Sub
Public Sub CreditsMenuAbout()
  'This is the credits screen
  bEnd = 0
  
  Dim T As New pgeTimer
  Dim x As Long

  sCredits.SetPosition 900, 100
  sCredits.SetAutoPath 384, 100, 1000
  sCredits.SetColor 255, 255, 255, 0
  sCredits.SetAutoFade 255, 255, 255, 255, 1000
  
  sCredits_Detail.SetPosition 900, 320
  sCredits_Detail.SetAutoPath 384, 320, 1000
  sCredits_Detail.SetColor 255, 255, 255, 0
  sCredits_Detail.SetAutoFade 255, 255, 255, 255, 1000
  
  sLogo.SetAutoPath 900, 100, 1000
  sText_Menu(0).SetAutoPath 900, 200, 1000
  sText_Menu(1).SetAutoPath 900, 275, 1000
  sText_Menu(2).SetAutoPath 900, 350, 1000
  sText_Menu(3).SetAutoPath 900, 425, 1000
  sLogo.SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(0).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(1).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(2).SetAutoFade 255, 255, 255, 0, 1000
  sText_Menu(3).SetAutoFade 255, 255, 255, 0, 1000

  T.StartTime
 ' Do
    DoEvents

    '###Drawing sequence
    pEngine.Clear
    
    sBar.Render
    sField.Render
    sLogo.Render
    For x = 0 To 3
      sText_Menu(x).Render
    Next
    sCredits.Render
    sCredits_Detail.Render
    
   ' If T.GetTimeMs >= 5000 Then bEnd = 2
    'If t.GetTimeMs > 1000 And pKeyboard.KeyDown(DIK_ESCAPE) Then bEnd = 1
    
    If bFps Then
      FontArial.DrawText pEngine.lFPS, ReturnRECT(0, 0, 100, 15), RGBA(0, 255, 0, 255), DT_LEFT
    End If
    
    pEngine.Render
    '###End of drawing sequence
'Check keyboard
      If pKeyboard.KeyDown(DIK_ESCAPE) Then
 ' Loop Until bEnd
      'Then bEnd = 1
      End If
End Sub
