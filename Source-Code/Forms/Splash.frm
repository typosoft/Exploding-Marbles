VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Splash 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4680
   ClientLeft      =   4500
   ClientTop       =   3840
   ClientWidth     =   7485
   LinkTopic       =   "Form2"
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   4680
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar Xp_Pro 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4305
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   240
      Top             =   480
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ****************************************************************
'
'                  Exploding Marbles
'                  Version 2.0 - 5.0
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

Private Sub Form_Load()
Splash.Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen
 Dim AboutLoc As String

    On Error Resume Next

    AboutLoc = App.Path & "\Required\Required-File"
    If Dir$(AboutLoc) <> "" Then
        Set Me.Picture = LoadPicture(AboutLoc)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Splash = Nothing
Unload Me
frmMain.Show

End Sub


Private Sub Image1_Click()

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Timer1_Timer()
Xp_Pro.Value = Xp_Pro.Value + 1
If Xp_Pro.Value > 99 Then
Unload Me
End If
End Sub

