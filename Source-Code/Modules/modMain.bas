Attribute VB_Name = "modMain"
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

Public Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Public Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long

Public pEngine As New pgeMain
Public pKeyboard As New pgeKeyboard
Public pTextures As New pgeTexturePool
Public pSound As New pgeSound
Public pMouse As New pgeMouse

Public FontArial As New pgeFont
Public MainFont As New pgeFont
Public LedFont As New pgeFont

Type tSettings
  SfxVolume As Byte
  MusicVolume As Byte
  MouseSpeed As Single
End Type

Type tHighScore
  lScore As Long
  sName As String
End Type

Type tPlayer
  lLevel As Long
  lScore As Long
  lTime As Long
  lDisplayTime As Long
  lBombs As Long
End Type

Type tGrd
  lType As Integer 'Marble type. 0 = empty
  lY As Single 'Y coordinate of marble
  lX As Single 'X coordinate of marble
  lDead As Long
  lFlag(3) As Long 'Used when dying
  bSpecial As Byte 'Special number. 0 = no special
End Type

Public CurrentMusic As Integer
Public Const MaxMusic As Integer = 14

Public LatestHigh As Long
Public bFps As Boolean 'Show fps on/off
Public lGrid(7, 8) As tGrd 'Playing field grid
Public Player As tPlayer 'Player status
Public Settings As tSettings 'Program settings
Public High(9) As tHighScore 'Highscore

Public Function FileExist(ByVal fileName As String) As Boolean
  FileExist = Not (Dir(fileName) = "")
End Function
