VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

'
'Texture sizes must be a power of 2!
'
Option Explicit

Private g_Texture As Direct3DTexture8
Private g_Info As D3DXIMAGE_INFO
Public SpriteWidth As Long
Public SpriteHeight As Long
Public Columns As Long
Public Rows As Long

Public Key As String

Friend Sub CreateFromFile(ByVal sFilename As String, Optional colorkey As Long = 0)
    Dim mipinfo As D3DSURFACE_DESC
    'create the texture
    
    Set g_Texture = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, sFilename, -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, colorkey, g_Info, ByVal 0)
    g_Texture.GetLevelDesc 0, mipinfo
    
    'set the size parameters
    g_Info.Width = mipinfo.Width
    g_Info.Height = mipinfo.Height
    SpriteWidth = mipinfo.Width
    SpriteHeight = mipinfo.Height
    Columns = 1
    Rows = 1
End Sub
Friend Function GetTexture() As Direct3DTexture8
    Set GetTexture = g_Texture
End Function
Friend Function GetWidth() As Long
    GetWidth = g_Info.Width
End Function
Friend Function GetHeight() As Long
    GetHeight = g_Info.Height
End Function
Friend Sub Destroy()
    Set g_Texture = Nothing
End Sub

