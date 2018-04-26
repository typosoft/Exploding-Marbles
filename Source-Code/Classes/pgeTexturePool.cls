VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeTexturePool"
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

'local variable to hold collection
Private mCol As Collection

Public Function AddFromFile(ByVal sFilename As String, sKey As String, Optional lColorKey As Long = 0) As pgeTexture
  'sFilename = Name of file. DirectX supports many file types, including type with an
  '            alpha channel, such as PNG.
  'sKey = Name/ID of texture.
  'lColorKey = Color value to set as transparent in this texture

  'create a new object
  Dim objNewMember As New pgeTexture

  'set the properties passed into the method
  objNewMember.Key = sKey
  objNewMember.CreateFromFile sFilename, lColorKey
  mCol.Add objNewMember, sKey

  'return the object created
  Set AddFromFile = objNewMember
  Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As pgeTexture
Attribute Item.VB_UserMemId = 0
  'used when referencing an element in the collection
  'vntIndexKey contains either the Index or Key to the collection,
  'this is why it is declared as a Variant
  'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
  'used when retrieving the number of elements in the
  'collection. Syntax: Debug.Print x.Count
  Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
  'used when removing an element from the collection
  'vntIndexKey contains either the Index or Key, which is why
  'it is declared as a Variant
  'Syntax: x.Remove(xyz)

  Me(vntIndexKey).Destroy
  mCol.Remove vntIndexKey
End Sub

Public Sub Clear()
  ' Used when removing all elements from the collection.
  Dim i As Long
  For i = Me.Count To 1 Step -1
      Me.Remove i
  Next i
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_MemberFlags = "40"
  'this property allows you to enumerate
  'this collection with the For...Each syntax
  Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
  'creates the collection when this class is created
  Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
  'destroys collection when this class is terminated
  Set mCol = Nothing
End Sub


