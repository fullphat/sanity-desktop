VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TListViewItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mIcon As mfxBitmap
Dim mContent As mfxView
Dim mUID As Long
Dim mState As Long
Dim mItem As TFile

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = CStr(mUID)

End Function

Private Function BTagItem_Value() As String

End Function

Public Function Init(ByRef Item As BTagItem) As Boolean

    If ISNULL(Item) Then _
        Exit Function

    ' /* create new id - we might switch this to using UUIDs in the future... */

    mUID = gUID
    gUID = gUID + 1

    ' /* assign the item */

    Set mItem = Item

    ' /* this causes SetState(0) to create the initial content */

    mState = -1

    ' /* figure out the icon to use - these should really be in a roster */

    If mItem.IsDirectory() Then
        Set mIcon = load_image_obj(g_MakePath(App.Path) & "icons\folder.png")
    
    ElseIf g_IsExecutable(Item.Name) Then
        Set mIcon = load_image_obj(g_MakePath(App.Path) & "icons\exe.png")

    Else
        Set mIcon = load_image_obj(g_MakePath(App.Path) & "icons\file.png")
    
    End If

    ' /* create the content */

    Me.SetItemState 0

End Function

Public Sub SetItemState(ByVal NewState As Long)


    ' /* if the item state hasn't changed we don't need to do anything */

    If NewState = mState Then _
        Exit Sub

    mState = NewState

Dim pb As mfxBitmap

    Set pb = mIcon.Duplicate

    ' /* change the icon based on its state flags */

    If (mState And CDIS_SELECTED) Then
        pb.ChangeBrightness -0.15

    Else

    End If

    Set mContent = New mfxView
    With mContent
        .SizeTo 48, 48
        .EnableSmoothing False
        .SetHighColour rgba(255, 255, 255)
        .FillRect .Bounds
        .DrawScaledImage pb, , new_BPoint(.Width, .Height)

    End With

'    Debug.Print "content was updated"

End Sub

Public Function Content() As mfxView

    Set Content = mContent

End Function

Public Function Item() As BTagItem

    Set Item = mItem

End Function

Public Function UID() As Long

    UID = mUID

End Function
