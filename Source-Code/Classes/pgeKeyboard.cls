VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeKeyboard"
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

Private Type KeyData
  KeyTimer As Long
  KeyInterval As Long
End Type

Public Suspended As Boolean
Private Key(255) As KeyData
Private EndHoldTimer As Long

Private DirectInputDevice As DirectInputDevice8

Private Sub Class_Terminate()
  DirectInputDevice.Unacquire
  Set DirectInputDevice = Nothing
End Sub

Public Sub Create(ByVal hwnd As Long)
  Set DirectInputDevice = DirectInput.CreateDevice("GUID_SysKeyboard")
  DirectInputDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
  DirectInputDevice.SetCooperativeLevel hwnd, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND
  Aquire
End Sub

Public Function KeyDown(ByVal KeyCode As CONST_DIKEYFLAGS) As Boolean
  On Error Resume Next
  
  If Not Suspended And timeGetTime >= EndHoldTimer Then
    Dim KeyState As DIKEYBOARDSTATE
    Call DirectInputDevice.GetDeviceStateKeyboard(KeyState)
    If KeyState.Key(KeyCode) And Key(KeyCode).KeyTimer - timeGetTime <= 0 Then
      Key(KeyCode).KeyTimer = timeGetTime + Key(KeyCode).KeyInterval
      KeyDown = True
    End If
  End If

ErrH:
  Aquire
End Function

Public Sub SetTimer(ByVal sec As Single)
  Dim x As Long
  sec = sec * 1000
  For x = 0 To 255
    Key(x).KeyInterval = sec
  Next
End Sub

Public Sub SetTimerEx(ByVal KeyCode As CONST_DIKEYFLAGS, ByVal sec As Single)
  Key(KeyCode).KeyInterval = sec * 1000
End Sub

Public Sub HoldTimer(ByVal sec As Single)
  EndHoldTimer = timeGetTime + (sec * 1000)
End Sub

Private Sub Aquire()
  On Error GoTo ErrH
  DirectInputDevice.Acquire
  Exit Sub
ErrH:
End Sub
