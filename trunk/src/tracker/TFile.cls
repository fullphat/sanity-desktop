VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mFileInfo As WIN32_FIND_DATA_API
Dim mName As String

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mName

End Function

Private Function BTagItem_Value() As String

End Function

Friend Sub SetTo(ByRef pwfd As WIN32_FIND_DATA_API)

    LSet mFileInfo = pwfd
    mName = g_TrimStr(pwfd.cFileName)

End Sub

Public Function IsDirectory() As Boolean

    IsDirectory = ((mFileInfo.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0)

End Function

Public Function Name() As String

    Name = mName

End Function
