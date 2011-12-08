VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00F4F4F4&
   Caption         =   "Filetracker concept 1/F (foxtrot)"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image3 
      Height          =   420
      Left            =   5700
      Picture         =   "Form1.frx":000C
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents theToolbar As TChildWindow
Attribute theToolbar.VB_VarHelpID = -1
Dim mBreadcrumbs As mfxView

Dim mhImgListIcon As Long
Dim mhImgListSmall As Long
Dim mPath As String

Dim WithEvents theListView As CListView
Attribute theListView.VB_VarHelpID = -1

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA_API) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA_API) As Long

Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    If NOTNULL(theListView) Then
        BWndProcSink_WndProc = theListView.HandleParentMessage(uMsg, wParam, lParam, ReturnValue)
        If BWndProcSink_WndProc = True Then _
            Exit Function

    End If

End Function

Private Sub Form_Load()

    gUID = &H200
    Set theToolbar = New TChildWindow

    window_subclass Me.hWnd, Me

    With theToolbar
        .Attach Me.hWnd
        .MoveTo 2, 0
        .SizeTo Me.ScaleWidth - 2, 28
        .Show

    End With

    Set theListView = New CListView
    With theListView
        .Create Me.hWnd, new_BPoint(2, 28), new_BPoint(Me.ScaleWidth - 4, Me.ScaleHeight - 28 - 2), 0

        mhImgListSmall = ImageList_Create(24, 24, ILC_COLOR24, 1, 1)
        .SetImageList LVSIL_SMALL, mhImgListSmall

        mhImgListIcon = ImageList_Create(48, 48, ILC_COLOR24, 1, 1)
        .SetImageList LVSIL_NORMAL, mhImgListIcon

    End With


    Me.Show

    uSetTo "c:\"

End Sub

Private Sub Form_Resize()

    theToolbar.SizeTo Me.ScaleWidth - 32 - (2 * 2), theToolbar.Bounds.Height

    If NOTNULL(theListView) Then _
        g_SizeWindow theListView.hWnd, Me.ScaleWidth - 4, Me.ScaleHeight - 28 - 2

    Image3.Move Me.ScaleWidth - Image3.Width - 2, 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    window_subclass Me.hWnd, Nothing

    theToolbar.Detach

    ImageList_Destroy mhImgListIcon
    ImageList_Destroy mhImgListSmall

End Sub

Private Sub theListView_Invoked(Item As TListViewItem)
Dim pf As TFile

    If ISNULL(Item) Then
        Debug.Print "whitespace"

    Else
        Set pf = Item.Item
        If pf.IsDirectory Then
            If pf.Name = ".." Then
                uSetTo g_GetPathParent(mPath)
            
            Else
                uSetTo mPath & pf.Name

            End If

        Else
            ShellExecute Me.hWnd, vbNullString, mPath & pf.Name, vbNullString, vbNullString, SW_SHOW

        End If
    End If

End Sub

Private Sub theListView_Menu(Item As TListViewItem)

    If NOTNULL(Item) Then
        Debug.Print "menu for '" & Item.Item.Name & "'"

    Else
        Debug.Print "menu for " & mPath

    End If

End Sub

Private Sub theToolbar_Draw(ByVal hDC As Long)

    If NOTNULL(mBreadcrumbs) Then _
        draw_view mBreadcrumbs, hDC

End Sub

Private Sub uSetTo(ByVal Path As String)

    mPath = g_MakePath(Path)
    uRebuildBreadcrumbs
    uGetFolderContent

End Sub

Private Sub uGetFolderContent()
Dim wfd As WIN32_FIND_DATA_API
Dim hFind As Long
Dim sz As String
Dim pf As TFile
Dim pi As TListViewItem

    theListView.Clear

    hFind = FindFirstFile(mPath & "*.*", wfd)
    If hFind <> INVALID_HANDLE_VALUE Then
        Do
            sz = g_TrimStr(wfd.cFileName)
            If (sz <> ".") Then
                ' /* create the TFile item */
                Set pf = New TFile
                pf.SetTo wfd

                ' /* create the TListViewItem */
                Set pi = New TListViewItem
                pi.Init pf                      ' // assign the file to the list item
                theListView.Add pi

            End If

        Loop While FindNextFile(hFind, wfd) <> 0

        FindClose hFind

    Else
        Debug.Print "CFolderContent2.SetTo(): bad path '" & mPath & "'"

    End If

End Sub

Private Sub uRebuildBreadcrumbs()

    If mPath = "" Then _
        Exit Sub

Dim sz() As String
Dim pr As BRect
Dim i As Long

    Set mBreadcrumbs = New mfxView
    With mBreadcrumbs
        .SizeTo theToolbar.Bounds.Width, theToolbar.Bounds.Height

        .EnableSmoothing False
        .SetHighColour rgba(244, 244, 244)
        .FillRect .Bounds

        .SetFont "Tahoma", 9
        .TextMode = MFX_TEXT_CLEARTYPE

        Set pr = new_BRect(0, 0, 48 - 1, .Bounds.Bottom)
        pr.InsetBy 0, 1

        ' /* increase the array size by one and move each entry along one */

        sz = Split(mPath, "\")
        ReDim Preserve sz(UBound(sz) + 1)

        For i = UBound(sz) - 1 To 1 Step -1
            Debug.Print CStr(i) & " = " & sz(i - 1)
            sz(i) = sz(i - 1)

        Next i

        ' /* insert "Places" entry at head */

        sz(0) = "Places"

        .EnableSmoothing True

        For i = 0 To UBound(sz) - 1

            pr.Right = pr.Left + MAX(24, (.StringWidth(sz(i)) + 10)) - 1

            .SetHighColour rgba(255, 255, 255)
            .FillRoundRect pr, 6, 6
            .SetHighColour rgba(196, 196, 196)
            .StrokeRoundRect pr, 6, 6

            .SetHighColour rgba(80, 80, 80, 200)
            .DrawString sz(i), pr, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

            .SetHighColour rgba(0, 0, 0, 0)
            .SetLowColour rgba(0, 0, 0, 24)
            .FillRoundRect pr.InsetByCopy(0, 1), 6, 6, MFX_VERT_GRADIENT

            pr.OffsetBy pr.Width + 2, 0

            If i = 0 Then _
                pr.OffsetBy 6, 0

        Next i

    End With

    theToolbar.Sync

End Sub

Private Sub theToolbar_Resized(ByVal Width As Long, ByVal Height As Long)

    uRebuildBreadcrumbs

End Sub
