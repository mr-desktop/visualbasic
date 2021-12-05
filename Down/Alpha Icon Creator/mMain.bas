Attribute VB_Name = "mMain"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Declare Function GetSystemMetrics Lib "user32" ( _
      ByVal nIndex As Long _
   ) As Long

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12

Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
   
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" ( _
      ByVal hInst As Long, _
      ByVal lpsz As String, _
      ByVal uType As Long, _
      ByVal cxDesired As Long, _
      ByVal cyDesired As Long, _
      ByVal fuLoad As Long _
   ) As Long
   
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000&

Private Const IMAGE_ICON = 1

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
      ByVal hwnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, ByVal lParam As Long _
   ) As Long

Private Const WM_SETICON = &H80

Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4


' SetStyles sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
' - mods SPM http://vbAccelerator.com/

Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hwnd As Long
End Type

Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    style As Long
    lpszName As Long    ' String
    lpszClass As Long   ' String
    ExStyle As Long
End Type

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Const WH_CALLWNDPROC = 4

' Misc Windows messages
Private Const WM_CREATE = &H1
Private Const WM_DESTROY = &H2
Private Const WM_PARENTNOTIFY = &H210

' Window Styles
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_CHILD = &H40000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_DISABLED = &H8000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_VSCROLL = &H200000
Private Const WS_HSCROLL = &H100000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_GROUP = &H20000
Private Const WS_TABSTOP = &H10000

Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000

Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
'
'   Common Window Styles
'  /
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Private Const WS_CHILDWINDOW = (WS_CHILD)

' Extended Window Styles
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_TRANSPARENT = &H20&

Private Const WS_EX_APPWINDOW = &H40000
Private Const WS_EX_TOOLWINDOW = &H80&

Private Const GWL_EXSTYLE = (-20)
Private Const GWL_WNDPROC = (-4)

' VB5 & VB6 class names:
Private Const C_MDIFORMCLASS_IDE = "ThunderMDIForm"
Private Const C_MDIFORMCLASS_EXE = "ThunderRT6MDIForm"
Private Const C_MDIFORMCLASS5_IDE = "ThunderMDIForm"
Private Const C_MDIFORMCLASS5_EXE = "ThunderRT5MDIForm"
 
Private Const C_FORMCLASS_IDE_DC = "ThunderFormDC"
Private Const C_FORMCLASS_EXE_DC = "ThunderRT6FormDC"
Private Const C_FORMCLASS_IDE = "ThunderForm"
Private Const C_FORMCLASS_EXE = "ThunderRT6Form"
Private Const C_FORMCLASS5_IDE = "ThunderForm"
Private Const C_FORMCLASS5_EXE = "ThunderRT5Form"

Private m_hHook As Long
Private m_lHookWndProc As Long

Public Sub HookAttach()
   m_hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
   Debug.Assert m_hHook <> 0
End Sub
Public Sub HookDetach()
   If m_hHook <> 0 Then
      UnhookWindowsHookEx m_hHook
      m_hHook = 0
   End If
End Sub

Private Function AppHook(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim CWP As CWPSTRUCT
Dim k As Long, aClass As String
    
   If idHook >= 0 Then
      CopyMemory CWP, ByVal lParam, Len(CWP)
      Select Case CWP.message
      Case WM_CREATE
          aClass = Space$(128)
          k = GetClassName(CWP.hwnd, ByVal aClass, 128)
          aClass = left$(aClass, k)
          If IsIn(aClass, C_MDIFORMCLASS_IDE, C_MDIFORMCLASS_EXE, C_MDIFORMCLASS5_IDE, _
              C_MDIFORMCLASS5_EXE, C_FORMCLASS_IDE_DC, C_FORMCLASS_EXE_DC, C_FORMCLASS_IDE, _
              C_FORMCLASS_EXE, C_FORMCLASS5_IDE, C_FORMCLASS5_EXE) Then
             m_lHookWndProc = SetWindowLong(CWP.hwnd, GWL_WNDPROC, AddressOf Form_WndProc)
          End If
      End Select
   End If
   AppHook = CallNextHookEx(m_hHook, idHook, wParam, ByVal lParam)
End Function

Private Function Form_WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lSetStyleEX As Long
   ' SPM - specific wnd proc for a form.  Only called once for the WM_CREATE message.
   Select Case Msg
   Case WM_CREATE
      Dim tCS As CREATESTRUCT
      CopyMemory tCS, ByVal lParam, Len(tCS)
      lSetStyleEX = GetWindowLong(hwnd, GWL_EXSTYLE)
      lSetStyleEX = lSetStyleEX Or WS_EX_APPWINDOW
      lSetStyleEX = lSetStyleEX And (Not WS_EX_TOOLWINDOW)
      tCS.ExStyle = lSetStyleEX
      CopyMemory ByVal lParam, tCS, Len(tCS)
      SetWindowLong hwnd, GWL_WNDPROC, m_lHookWndProc
      SetWindowLong hwnd, GWL_EXSTYLE, tCS.ExStyle
   End Select
   Form_WndProc = CallWindowProc(m_lHookWndProc, hwnd, Msg, wParam, lParam)
End Function

Private Function IsIn(ByVal vComp As Variant, ParamArray vTo() As Variant) As Boolean
Dim i As Long, iL As Long, iU As Long
   On Error Resume Next
   iU = UBound(vTo)
   If Err.Number = 0 Then
      iL = LBound(vTo)
      For i = iL To iU
         If vComp = vTo(i) Then
            IsIn = True
            Exit Function
         End If
      Next i
   End If
End Function


Public Sub SetIcon( _
      ByVal hwnd As Long, _
      ByVal sIconResName As String, _
      Optional ByVal bSetAsAppIcon As Boolean = True _
   )
Dim lhWndTop As Long
Dim lhWnd As Long
Dim cx As Long
Dim cy As Long
Dim hIconLarge As Long
Dim hIconSmall As Long
      
   If (bSetAsAppIcon) Then
      ' Find VB's hidden parent window:
      lhWnd = hwnd
      lhWndTop = lhWnd
      Do While Not (lhWnd = 0)
         lhWnd = GetWindow(lhWnd, GW_OWNER)
         If Not (lhWnd = 0) Then
            lhWndTop = lhWnd
         End If
      Loop
   End If
   
   cx = GetSystemMetrics(SM_CXICON)
   cy = GetSystemMetrics(SM_CYICON)
   hIconLarge = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cx, cy, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
   End If
   SendMessageLong hwnd, WM_SETICON, ICON_BIG, hIconLarge
   
   cx = GetSystemMetrics(SM_CXSMICON)
   cy = GetSystemMetrics(SM_CYSMICON)
   hIconSmall = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cx, cy, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
   End If
   SendMessageLong hwnd, WM_SETICON, ICON_SMALL, hIconSmall
   
End Sub

Public Sub Main()
   
   InitCommonControls
      
   Dim f As New frmAlphaIconCreator
   f.Show
   
End Sub
