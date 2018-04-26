VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeMain"
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

Option Explicit

Dim lStart As Long, lCount As Long

Public lFPS As Long 'The number of calculated frames per second
Public lClearColor As Long 'The color to clear the backbuffer with

Public Sub Init(ByVal hwnd As Long, ByVal bWindowed As Boolean, Optional w As Long = -1, Optional h As Long = -1, Optional bVSync As Boolean = True)
  'This intializes the engine.
  'NOTES:
  '* The Default adapter is used.
  '* D3DDEVTYPE_HAL is used as device type
  '-
  '* w & h must be entered when using fullscreen
  
  Dim params As D3DPRESENT_PARAMETERS
  Dim dp As D3DDISPLAYMODE
  
  'Create Direct3D
  Set Direct3D = DirectX.Direct3DCreate
  
  'retrieve the display mode
  Direct3D.GetAdapterDisplayMode 0, dp

  If bWindowed Then 'If windowed
    'Gets render area size
    GetWindowRect hwnd, Target
    With Target
      w = .Right - .Left
      h = .bottom - .Top
      .Left = 0
      .Right = w
      .Top = 0
      .bottom = h
    End With
    
    'Sets up render parameters
    With params
      .BackBufferFormat = dp.format
      .EnableAutoDepthStencil = 0
      .Windowed = 1
      If bVSync Then
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
      Else
        .SwapEffect = D3DSWAPEFFECT_COPY
      End If
    End With
  Else 'If fullscreen
    'Setup screen size
    With Target
      Target.Left = 0
      Target.Right = w
      Target.Top = 0
      Target.bottom = h
    End With
    
    'Sets up render parameters
    With params
      .BackBufferFormat = dp.format
      .EnableAutoDepthStencil = 0
      .BackBufferWidth = w
      .BackBufferHeight = h
      .Windowed = 0
      .SwapEffect = D3DSWAPEFFECT_FLIP
    End With
  End If
  
  'Creates device
  Set Direct3DDevice = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, params)
  
  'Inits sprites
  Set Sprites = Direct3DX.CreateSprite(Direct3DDevice)
  
  'Init Input
  Set DirectInput = DirectX.DirectInputCreate
  
  'Turn off culling
  Direct3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
  'Turn off D3D lighting, since we are providing our own vertex colors
  Direct3DDevice.SetRenderState D3DRS_LIGHTING, 0
End Sub

Public Function GetTexturePool() As pgeTexturePool
  Set GetTexturePool = TPool
End Function

Public Sub Clear()
  'This clears the backbuffer
  Direct3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, lClearColor, 1, 0
  Direct3DDevice.BeginScene
  Sprites.Begin
End Sub

Public Sub Render(Optional hwnd As Long)
  'this renders the backbuffer to the render area
  On Error Resume Next
  
  'Calculate frames per second
  If lStart = 0 Or timeGetTime - lStart >= 1000 Then
    lStart = timeGetTime
    lFPS = lCount
    lCount = 0
  End If
  lCount = lCount + 1
  
  'presents the rendered image
  Sprites.End
  Direct3DDevice.EndScene
  Direct3DDevice.Present ByVal 0, ByVal 0, hwnd, ByVal 0
End Sub

Private Sub Class_Terminate()
  On Error Resume Next

  TPool.Clear
  
  Set Sprites = Nothing
  Set Direct3DDevice = Nothing
  Set Direct3D = Nothing
  Set DirectInput = Nothing
End Sub


