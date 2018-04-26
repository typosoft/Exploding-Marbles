VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrSplash 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar Xp_Pro 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8625
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
If App.PrevInstance Then
    MsgBox App.EXEName & " is already runing!", vbInformation + vbOKOnly, App.EXEName & ". Loading System."
   End
End If
If GDIAvailable = False Then MsgBox "ERROR LOADING GDI+", vbCritical, "U11D Checking System": End
MakeFormTop Me.hwnd, True
SetWinLng Me
MakePNG App.Path & "\SPLASH\Splash.png", Me, 240, False

CM.m(0, 0) = 1
CM.m(1, 1) = 1
CM.m(2, 2) = 1
CM.m(3, 3) = 1
End Sub

Private Sub TmrSplash_Timer()
Xp_Pro.Value = Xp_Pro.Value + 1
If Xp_Pro.Value > 99 Then
Unload Me
End If
    frmMain.Show
End Sub





Private Sub Xp_Pro_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub
