VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TChildWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum B_CHILDWINDOW_FLAGS
    B_HORIZ_SCROLL = 1
    B_VERT_SCROLL = 2

End Enum

Public Enum B_SCROLL_DIRECTION
    B_SCROLLED_HORIZ = 0
    B_SCROLLED_VERT = 1

End Enum

Const SB_THUMBTRACK = 5
Const CLASS_NAME = "w>etinnae"

Dim mhWndParent As Long
Dim mhWnd As Long

Public Event Draw(ByVal hDC As Long)
Public Event Scroll(ByVal Direction As B_SCROLL_DIRECTION, ByVal Value As Long)

Public Event MouseDown(ByRef Position As BPoint)
Public Event Resized(ByVal Width As Long, ByVal Height As Long)

Implements BWndProcSink

Public Sub Attach(ByVal hWndParent As Long, Optional ByVal Flags As B_CHILDWINDOW_FLAGS)
Dim ws As Long

    EZRegisterClass CLASS_NAME, CS_HREDRAW Or CS_DBLCLKS

    ws = WS_CHILD Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS

    If (Flags And B_HORIZ_SCROLL) Then _
        ws = ws Or WS_HSCROLL

    If (Flags And B_VERT_SCROLL) Then _
        ws = ws Or WS_VSCROLL

    mhWnd = EZ4AddWindow(CLASS_NAME, Me, , ws, 0, hWndParent)

End Sub

Public Sub Detach()

    EZ4RemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

End Sub

Public Sub Show()

    ShowWindow mhWnd, SW_SHOW

End Sub

Public Sub Hide()

    ShowWindow mhWnd, SW_HIDE

End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static pps As PAINTSTRUCT
Static hDC As Long

    Select Case uMsg
    Case WM_ERASEBKGND
        ReturnValue = -1
        BWndProcSink_WndProc = True

    Case WM_PAINT
        hDC = BeginPaint(hWnd, pps)
        RaiseEvent Draw(hDC)
        EndPaint hWnd, pps
        ReturnValue = 0
        BWndProcSink_WndProc = True

    Case WM_VSCROLL
        If LoWord(wParam) = SB_THUMBTRACK Then _
            RaiseEvent Scroll(B_SCROLLED_VERT, HiWord(wParam))

    Case WM_HSCROLL
        If LoWord(wParam) = SB_THUMBTRACK Then _
            RaiseEvent Scroll(B_SCROLLED_HORIZ, HiWord(wParam))


    Case WM_LBUTTONDOWN
        RaiseEvent MouseDown(new_BPointFromInt32(lParam))

    Case WM_SIZE
        RaiseEvent Resized(LoWord(lParam), HiWord(lParam))

    End Select

End Function

Public Sub MoveTo(ByVal x As Long, ByVal y As Long)

    g_MoveWindow mhWnd, x, y

End Sub

Public Sub SizeTo(ByVal Width As Long, ByVal Height As Long)

    g_SizeWindow mhWnd, Width, Height

End Sub

Public Sub Sync()

    g_UpdateWindow mhWnd, False

End Sub

Public Function Bounds() As BRect
Dim rc As RECT

    GetClientRect mhWnd, rc
    Set Bounds = new_BRect(rc.Left, rc.Top, rc.Right - 1, rc.Bottom - 1)

End Function



