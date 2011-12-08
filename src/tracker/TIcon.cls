VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mFrame As BRect
Dim mUID As Long

Implements BTagItem
Implements BObject

Private Function BObject_Duplicate() As melon4.BObject
End Function

Private Function BObject_Init(Args As String) As melon4.B_STATUS
End Function

Private Function BObject_Name() As String
End Function

Private Function BObject_Ptr() As Long

    BObject_Ptr = mUID

End Function

Private Sub BObject_Tidyup()
End Sub

Private Function BObject_Type() As String
End Function

Private Function BTagItem_Name() As String
End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = CStr(mUID)

End Function

Public Sub SetFrame(ByRef Frame As BRect)

    Set mFrame = Frame

End Sub

Public Function Frame() As BRect

    Set Frame = mFrame

End Function

Public Sub SetUID(ByVal UID As Long)

    mUID = UID

End Sub

Public Function UID() As Long

    UID = mUID

End Function

