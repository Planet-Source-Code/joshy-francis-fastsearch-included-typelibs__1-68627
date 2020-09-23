Attribute VB_Name = "modShell"
Option Explicit
Public Type ITEM_VERB
    sVerb As String
    lCmdID As Long
End Type

Public Enum SpecialFolderConstants
   CSIDL_DESKTOP = &H0                 '{desktop}
   CSIDL_INTERNET = &H1                'Internet Explorer (icon on desktop)
   CSIDL_PROGRAMS = &H2                'Start Menu\Programs
   CSIDL_CONTROLS = &H3                'My Computer\Control Panel
   CSIDL_PRINTERS = &H4                'My Computer\Printers
   CSIDL_PERSONAL = &H5                'My Documents
   CSIDL_FAVORITES = &H6               '{user name}\Favorites
   CSIDL_STARTUP = &H7                 'Start Menu\Programs\Startup
   CSIDL_RECENT = &H8                  '{user name}\Recent
   CSIDL_SENDTO = &H9                  '{user name}\SendTo
   CSIDL_BITBUCKET = &HA               '{desktop}\Recycle Bin
   CSIDL_STARTMENU = &HB               '{user name}\Start Menu
   CSIDL_DESKTOPDIRECTORY = &H10       '{user name}\Desktop
   CSIDL_DRIVES = &H11                 'My Computer
   CSIDL_NETWORK = &H12                'Network Neighborhood
   CSIDL_NETHOOD = &H13                '{user name}\nethood
   CSIDL_FONTS = &H14                  'windows\fonts
   CSIDL_TEMPLATES = &H15
   CSIDL_COMMON_STARTMENU = &H16       'All Users\Start Menu
   CSIDL_COMMON_PROGRAMS = &H17        'All Users\Programs
   CSIDL_COMMON_STARTUP = &H18         'All Users\Startup
   CSIDL_COMMON_DESKTOPDIRECTORY = &H19 'All Users\Desktop
   CSIDL_APPDATA = &H1A                '{user name}\Application Data
   CSIDL_PRINTHOOD = &H1B              '{user name}\PrintHood
   CSIDL_LOCAL_APPDATA = &H1C          '{user name}\Local Settings\Application Data (non roaming)
   CSIDL_ALTSTARTUP = &H1D             'non localized startup
   CSIDL_COMMON_ALTSTARTUP = &H1E      'non localized common startup
   CSIDL_COMMON_FAVORITES = &H1F
   CSIDL_INTERNET_CACHE = &H20
   CSIDL_COOKIES = &H21
   CSIDL_HISTORY = &H22
   CSIDL_COMMON_APPDATA = &H23          'All Users\Application Data
   CSIDL_WINDOWS = &H24                 'GetWindowsDirectory()
   CSIDL_SYSTEM = &H25                  'GetSystemDirectory()
   CSIDL_PROGRAM_FILES = &H26           'C:\Program Files
   CSIDL_MYPICTURES = &H27              'C:\Program Files\My Pictures
   CSIDL_PROFILE = &H28                 'USERPROFILE
   CSIDL_SYSTEMX86 = &H29               'x86 system directory on RISC
   CSIDL_PROGRAM_FILESX86 = &H2A        'x86 C:\Program Files on RISC
   CSIDL_PROGRAM_FILES_COMMON = &H2B    'C:\Program Files\Common
   CSIDL_PROGRAM_FILES_COMMONX86 = &H2C 'x86 Program Files\Common on RISC
   CSIDL_COMMON_TEMPLATES = &H2D        'All Users\Templates
   CSIDL_COMMON_DOCUMENTS = &H2E        'All Users\Documents
   CSIDL_COMMON_ADMINTOOLS = &H2F       'All Users\Start Menu\Programs\Administrative Tools
   CSIDL_ADMINTOOLS = &H30              '{user name}\Start Menu\Programs\Administrative Tools
End Enum

Const MAX_PATH = 260

Public Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type


Public Const SHGFI_ICON = &H100
Public Const SHGFI_ATTRIBUTES = &H800
Public Const SHGFI_ICONLOCATION = &H1000
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_LINKOVERLAY = &H8000
Public Const SHGFI_SELECTED = &H10000
Public Const SHGFI_SMALLICON = &H1
Public Const SHGFI_OPENICON = &H2
Public Const SHGFI_PIDL = &H8
Public Const ICON_SHGFI_FLAGS = SHGFI_SYSICONINDEX Or SHGFI_ICON

Public Const SFGAO_HASSUBFOLDER = &H80000000
Public Const SFGAO_FOLDER = &H20000000
Public Const SFGAO_LINK = &H10000                                ' Is a shortcut (link)
Public Const SFGAO_SHARE = &H20000
Public Const SFGAO_CANRENAME = &H10&
Public Const SFGAO_CANDELETE = &H20&
Public Const SFGAO_FILESYSANCESTOR = &H10000000
Public Const SFGAO_GHOSTED = &H80000

Public Declare Function SHGetFileInfoPidl Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As Long, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function SHGetFileInfoStr Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'Really working horses!
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (pDest As Any, ByVal dwLength As Long, ByVal bFill As Byte)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'=========================================
'Special Shell staff (mostly undocumented)
'=========================================

'Shell window structure
Enum ShellViewWindows
    ShellEmbedding   'main class window
    Shelldef         'main window which receive notifications
    Syslistview      'listview
End Enum

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Const WC_SHVIEW = "SHELLDLL_DefView"
Public Const WC_LISTVIEW = "SysListView32"

'=============================================================
'You can send a number of command messages to WC_SHVIEW window
'like this:
'Call SendMessage(hShellDefView,WM_COMMAND,IDM_,ByVal 0&)
'IDM_ messages start from &H7000. All they are undocumented and
'was found through experiments:
'=============================================================

Public Const WM_COMMAND = &H111

Public Const IDM_SHVIEW_CREATELINK = &H7010
Public Const IDM_SHVIEW_DELETE = &H7011
Public Const IDM_SHVIEW_PROPERTIES = &H7013
Public Const IDM_SHVIEW_CUT = &H7018
Public Const IDM_SHVIEW_COPY = &H7019
Public Const IDM_SHVIEW_INSERT = &H701A
Public Const IDM_SHVIEW_UNDO = &H701B
Public Const IDM_SHVIEW_INSERTLINK = &H701C
Public Const IDM_SHVIEW_SELECTALL = &H7021
Public Const IDM_SHVIEW_INVERTSELECTION = &H7022

Public Const IDM_SHVIEW_LARGEICON = &H7029
Public Const IDM_SHVIEW_SMALLICON = &H702A
Public Const IDM_SHVIEW_LIST = &H702B
Public Const IDM_SHVIEW_REPORT = &H702C
Public Const IDM_SHVIEW_FOLDERVIEW = &H702F
Public Const IDM_SHVIEW_TOGGLEWEBVIEW = &H7030
Public Const IDM_SHVIEW_REFRESH = &H7103
Public Const IDM_SHVIEW_NEWFOLDER = &H7261
Public Const IDM_SHVIEW_NEWLINK = &H7262

Public Const IDM_SHVIEW_TOGGLECOL0 = &H7230
Public Const IDM_SHVIEW_TOGGLECOL1 = &H7231
Public Const IDM_SHVIEW_TOGGLECOL2 = &H7232
Public Const IDM_SHVIEW_TOGGLECOL3 = &H7233

'=========================
'All shell works with PIDLs. You can get them from IShellFolder
'interface. But shell32.dll itself have undoc functions for
'working with PIDLs. All them are colling by ordinal.
'Undoc shell32 functions:
'=========================

'Memory allocation
Public Declare Function SHAlloc Lib "shell32" Alias "#196" (ByVal cbSize As Long) As Long
Public Declare Sub SHFree Lib "shell32" Alias "#195" (ByVal pv As Long)
Public Declare Sub ILFree Lib "shell32" Alias "#155" (ByVal PIDL As Long)
Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

'From Path to PIDL
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal PIDL As Long, ByVal pszPath As String) As Long
Public Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Public Declare Function SHILFromPath Lib "shell32" Alias "#157" (ByVal szPath As String) As Long
Public Declare Function SHILCreateFromPath Lib "shell32" Alias "#28" (ByVal szPath As String, PIDL As Long, ByVal dwAttributes As Long) As Long

'Parsing PIDLs
Public Declare Function SHILGetNext Lib "shell32" Alias "#153" (ByVal PIDL As Long) As Long
Public Declare Function SHILFindLast Lib "shell32" Alias "#16" (ByVal PIDL As Long) As Long
Public Declare Function SHILFindChild Lib "shell32" Alias "#24" (ByVal pidlParen As Long, ByVal pidlChild As Long) As Long

'Copying and combining PIDLs
Public Declare Function SHILClone Lib "shell32" Alias "#18" (ByVal PIDL As Long) As Long
Public Declare Function SHILCloneFirst Lib "shell32" Alias "#19" (ByVal PIDL As Long) As Long
Public Declare Function SHILAppendID Lib "shell32" Alias "#154" (ByVal PIDL As Long, ByVal itemID As Long, ByVal bAddToEnd As Long) As Long
Public Declare Function SHILCombine Lib "shell32" Alias "#25" (ByVal pidl1 As Long, ByVal pidl2 As Long) As Long
Public Declare Function SHILRemoveLast Lib "shell32" Alias "#17" (ByVal PIDL As Long) As Boolean

'PIDLs comparison
Public Declare Function SHILIsEqual Lib "shell32" Alias "#21" (ByVal pidl1 As Long, ByVal pidl2 As Long) As Boolean
Public Declare Function SHILIsParent Lib "shell32" Alias "#23" (ByVal pidlParen As Long, ByVal pidlChild As Long, ByVal bImmediate As Boolean) As Boolean

'Others PIDL routines
Public Declare Function SHILGetSize Lib "shell32" Alias "#152" (ByVal PIDL As Long) As Long
Public Declare Function SHILGetDisplayName Lib "shell32" Alias "#15" (ByVal PIDL As Long, sName As String) As Boolean
'Some usefull Path routines
Public Declare Function PathFileExist Lib "shell32" Alias "#45" (ByVal sPath As String) As Boolean
Public Declare Function PathGetExtention Lib "shell32" Alias "#31" (ByVal sPath As String) As Long
Public Declare Function PathIsDirectory Lib "shell32" Alias "#159" (ByVal sPath As String) As Boolean
Public Declare Function CopyPointer2String Lib "kernel32" Alias "lstrcpyA" (ByVal NewString As String, ByVal OldString As Long) As Long


Public Declare Function FileIconInit Lib "shell32.dll" Alias "#660" (ByVal cmd As Boolean) As Boolean
Public Declare Function ImageList_Draw Lib "comctl32" (ByVal hIml As Long, ByVal I As Long, ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
Public Const ILD_TRANSPARENT = &H1

'=========================
' Context menu staff
'=========================

Public Type POINTAPI
    x  As Long
    y  As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400
Public Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal HMENU As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As Any) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal HMENU As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal HMENU As Long, ByVal nPos As Long) As Long
Public Const TPM_LEFTALIGN = &H0
Public Const TPM_RIGHTBUTTON = &H2
Public Const TPM_RETURNCMD = &H100
Public Declare Function DestroyMenu Lib "user32" (ByVal HMENU As Long) As Long
'=======Retrive Shell resources=======
Public Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function IsMenu Lib "user32" (ByVal HMENU As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal HMENU As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal HMENU As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal HMENU As Long) As Long
Public Declare Function GetMenuDefaultItem Lib "user32" (ByVal HMENU As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function SetMenuDefaultItem Lib "user32" (ByVal HMENU As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Public Const MF_SEPARATOR = &H800&
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal HMENU As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal HMENU As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Const MF_BITMAP = &H4&
Public Const MF_CHECKED = &H8&
Public Const MF_CONV = &H40000000
Public Const MF_DELETE = &H200&
Public Const MF_APPEND = &H100&
'Public Const MF_BYCOMMAND = &H0&
'Public Const MF_BYPOSITION = &H400&
Public Const MF_CALLBACKS = &H8000000
Public Const MF_CHANGE = &H80&
Public Const MF_DISABLED = &H2&
Public Const MF_ENABLED = &H0&
Public Const MF_END = &H80
Public Const MF_ERRORS = &H10000000
Public Const MF_GRAYED = &H1&
Public Const MF_HELP = &H4000&
Public Const MF_HILITE = &H80&
Public Const MF_HSZ_INFO = &H1000000
Public Const MF_INSERT = &H0&
Public Const MF_LINKS = &H20000000
Public Const MF_MASK = &HFF000000
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&
Public Const MF_MOUSESELECT = &H8000&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_POSTMSGS = &H4000000
Public Const MF_REMOVE = &H1000&
'Public Const MF_SEPARATOR = &H800&
Public Const MF_SENDMSGS = &H2000000
Public Const MF_STRING = &H0&
Public Const MF_SYSMENU = &H2000&
Public Const MF_UNCHECKED = &H0&
Public Const MF_UNHILITE = &H0&
Public Const MF_USECHECKBITMAPS = &H200&
Public Const WM_SETREDRAW = &HB
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
' Standard GDI draw icon function:
Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = &H3
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1
Public Declare Function DeleteMenu Lib "user32" (ByVal HMENU As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public DefaultItem As String
Public hID As Long
Function GetExe(ByVal sFile As String) As String
Dim Ret As Long, str As String, C As Long
    str = Space$(1024)
        Ret = FindExecutable(ByVal sFile, ByVal vbNullString, str)
            C = InStr(str, Chr(0))
                If C Then
                    str = Left$(str, C - 1)
                End If
            str = Trim$(str)
    If Ret > 32 Or str <> "" Then
        GetExe = str
    Else
        GetExe = ""
    End If
End Function
Sub SetAutoRedraw(ByVal hWnd As Long, ByVal b As Boolean)
SendMessage hWnd, WM_SETREDRAW, -CLng(b), ByVal 0
End Sub
Public Function GetFileIconIndexPIDL(PIDL As Long, uType As Long) As Long
  Dim sfi As SHFILEINFO
  If SHGetFileInfoPidl(ByVal PIDL, 0, sfi, Len(sfi), SHGFI_PIDL Or ICON_SHGFI_FLAGS Or uType) Then
    GetFileIconIndexPIDL = sfi.iIcon
  End If
End Function

Public Function INDEXTOOVERLAYMASK(iOverlay As Long) As Long
  '   INDEXTOOVERLAYMASK(i)   ((i) << 8)
  INDEXTOOVERLAYMASK = iOverlay * (2 ^ 8)
End Function

Public Function HasSubFolders(PIDL As Long) As Boolean
  Dim sfi As SHFILEINFO
  Call SHGetFileInfoPidl(ByVal PIDL, 0, sfi, Len(sfi), SHGFI_ATTRIBUTES Or SHGFI_PIDL)
  HasSubFolders = (sfi.dwAttributes And SFGAO_HASSUBFOLDER)
End Function

Public Function HasFiles(fld As Object) As Boolean
   Dim fldItem As Object
   For Each fldItem In fld.Items
       HasFiles = True
       Exit For
   Next
   Set fldItem = Nothing
End Function

'=======================
'Context menu staff
'=======================
'Return ITEM_VERB structure from context menu
'Use one more Undoc future - context menu handle can be obtain at offset
'32 from ItemVerbs collection pointer
Public Function Show_ContextMenu(ByVal hParent As Long, fldItem As Object) As ITEM_VERB
'Public Function Show_ContextMenu(ByVal hParent As Long, ByVal sPath As String, ByVal sFile As String) As ITEM_VERB
    Dim pt          As POINTAPI
    Dim fldItemVerbs  As FolderItemVerbs 'Object
    Dim HMENU       As Long
    Dim sCaption    As String
    Dim I           As Long
    Dim lngCmdId    As Long
    Dim Found As Boolean
'    Dim fi As FolderItem
'    Dim s As Shell
'        Set s = New Shell
'            Set fi = s.NameSpace(sPath).Items.Item(sFile)
    On Error GoTo Exit_ContextMenu
    Set fldItemVerbs = fldItem.Verbs
'    Set fldItemVerbs = fldItem
'    Call CopyMemory(hMenu, ByVal ObjPtr(fldItemVerbs) + 32, 4)
        I = 44
        I = 44
            HMENU = 0
            Found = False
GethMenu:
    Do Until IsMenu(HMENU) = 1
        Call CopyMemory(HMENU, ByVal ObjPtr(fldItemVerbs) + I, 4)
        Call GetCursorPos(pt)
            I = I + 1
        If I > 1024 Then
            Exit Do
        End If
    Loop
Dim C As Long, j As Long
Dim cMenu As Long
If Found = False Then
    C = GetMenuItemCount(HMENU)
    For j = 0 To C
           sCaption = String(64, 0)
           cMenu = GetMenuString(HMENU, j, sCaption, 64, MF_BYPOSITION)
           sCaption = Left(sCaption, cMenu)
        If InStr(sCaption, "Create &Shortcut") Or InStr(sCaption, "P&roperties") Or InStr(sCaption, "Cu&t") Then
                DeleteMenu HMENU, j + 2, MF_BYPOSITION Or MF_DELETE
            Found = True
            Exit For
        End If
    Next
        If Found = False Then
            HMENU = 0
            GoTo GethMenu
        End If
End If
        I = GetMenuDefaultItem(HMENU, 0, MF_BYPOSITION)
    DefaultItem = Space$(64)
       I = GetMenuString(HMENU, I, DefaultItem, 64, MF_BYCOMMAND)
       DefaultItem = Left(DefaultItem, I)
    
            cMenu = CreateMenu
        InsertMenu HMENU, 0, MF_BYPOSITION Or MF_SEPARATOR, cMenu, ByVal ""
            cMenu = CreateMenu
            hID = cMenu
        InsertMenu HMENU, 0, MF_BYPOSITION, cMenu, ByVal "Open Containing Folder"
            cMenu = CreateMenu
        InsertMenu HMENU, 1, MF_BYPOSITION, cMenu, ByVal "Find with Notepad"
            cMenu = CreateMenu
        InsertMenu HMENU, 1, MF_BYPOSITION, cMenu, ByVal "Find with Associated Program"
        
'            SetMenuDefaultItem hMenu, cMenu, 0
            SetMenuItemBitmaps HMENU, 0, MF_BYPOSITION Or MF_BITMAP Or MF_MASK, frmFastSearch.Picture1.Picture.handle, frmFastSearch.Picture1.Picture.handle
            SetMenuItemBitmaps HMENU, 1, MF_BYPOSITION Or MF_BITMAP Or MF_MASK, frmFastSearch.Pic.Picture.handle, frmFastSearch.Pic.Picture.handle
            SetMenuItemBitmaps HMENU, 2, MF_BYPOSITION Or MF_BITMAP Or MF_MASK, frmFastSearch.Picture3.Picture.handle, frmFastSearch.Picture3.Picture.handle
'        MsgBox IsMenu(hMenu), , i
    lngCmdId = TrackPopupMenu(HMENU, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, pt.x, pt.y, 0, hParent, ByVal 0&)
    If (lngCmdId > 0) Then
       sCaption = String(32, 0)
       I = GetMenuString(HMENU, lngCmdId, sCaption, 32, MF_BYCOMMAND)
       sCaption = Left(sCaption, I)
       Show_ContextMenu.sVerb = sCaption
       Show_ContextMenu.lCmdID = (lngCmdId - 1) Or &H7000
    End If
        
Exit_ContextMenu:
    Set fldItemVerbs = Nothing
'    Set fi = Nothing
'    Set s = Nothing
End Function

Function GetDefaultItem(fi As FolderItem) As String
    Dim fiv  As FolderItemVerbs 'Object
    Dim HMENU       As Long
    Dim I           As Long
    Dim Found As Boolean
On Error GoTo E
    Set fiv = fi.Verbs
        I = 44
            HMENU = 0
            Found = False
GethMenu:
    Do Until IsMenu(HMENU) = 1
        Call CopyMemory(HMENU, ByVal ObjPtr(fiv) + I, 4)
            I = I + 1
        If I > 1024 Then
            Exit Do
        End If
    Loop
        I = GetMenuDefaultItem(HMENU, 0, MF_BYPOSITION)
    DefaultItem = Space$(64)
       I = GetMenuString(HMENU, I, DefaultItem, 64, MF_BYCOMMAND)
       DefaultItem = Left(DefaultItem, I)
    GetDefaultItem = DefaultItem
        Set fiv = Nothing
E:
End Function
'Function ShowIContextMenu(ByVal sPath As String, ByVal hWnd As Long, _
'    Optional ByVal sp As String = "", Optional ByVal sf As String = "") As Boolean
'    On Error Resume Next
'Dim lpcm As olelib.IContextMenu
'Dim cmi As olelib.CMINVOKECOMMANDINFO
'Dim dwdwAttribs  As Long
'Dim idCmd As Long
'Dim HMENU As Long
'Dim hr As Long
'Dim lpsfParent As olelib.IShellFolder
'Dim lpi As Long
'Dim lppt  As POINTAPI
'Dim uuidCM As UUID
'Dim sFile As String
'Dim f As olelib.IShellFolder
'Dim uuidSF As UUID
'Dim rs As Long
'Dim EI As olelib.IEnumIDList
'Dim VO As olelib.IViewObject
'Dim uuidVo As UUID
'Dim N As Long
''sFile = "C:\Documents and Settings\Joshy Francis\Desktop\bliss.bmp"
'sFile = "bliss.bmp"
'Call CLSIDFromString(IIDSTR_IContextMenu, uuidCM)
'Call CLSIDFromString(IIDSTR_IShellFolder, uuidSF)
'Call CLSIDFromString(IIDSTR_IShellView, uuidVo)
'    GetCursorPos lppt
'
'Set lpsfParent = olelib.shell32.SHGetDesktopFolder
''Set f = olelib.Shell32.SHGetDesktopFolder
'
'Call lpsfParent.ParseDisplayName(hWnd, ByVal 0, StrPtr(sFile), Len(sFile), lpi, SFGAO_FILESYSTEM)
'    Dim C As Long
'        C = SHGetSpecialFolderLocation(hWnd, CSIDL_DRIVES)
'    Dim str As String, st As STRRET, si As SHFILEINFO
'    Dim k As Long, t As Long
'Re:
'        If N = 0 Then N = C
'        k = 0
'    If N Then
'        If Not f Is Nothing Then
'            Call f.BindToObject(N, 0, uuidSF, k)
'                Set f = Nothing
'        Else
'                Call lpsfParent.BindToObject(N, 0, uuidSF, k)
'        End If
'    End If
'If k Then CopyMemory f, k, 4
'    If Not f Is Nothing Then
''            sFile = sPath '"C:\t.txt"
''                    lpi = 0
''        Call f.ParseDisplayName(hWnd, ByVal 0, StrPtr(sFile), Len(sFile), lpi, SFGAO_FILESYSTEM)
'    Set EI = f.EnumObjects(hWnd, SHCONTF_NONFOLDERS Or SHCONTF_INCLUDEHIDDEN Or SHCONTF_FOLDERS)
'            N = 0: C = 0: k = 0
'        Do
'                rs = EI.Next(C, N)
'                str = Space$(255)
''                    str = StrConv(str, vbUnicode)
'            If N Then
'                        C = C + 1
''                    Dim st As STRRET
''                f.GetDisplayNameOf N, SHGDN_FORPARSING, st
'                f.GetDisplayNameOf N, SHGDN_FORPARSING, st
'                        str = StrConv(Space$(260), vbUnicode)
'                    StrRetToBuf VarPtr(st), N, str, 260
'                        t = InStr(str, Chr(0))
'                            If t Then
'                                str = Left$(str, t - 1)
'                            End If
'                    If LCase$(sPath) = LCase$(str) Then
'                        Exit Do
'                    Else
'                        If LCase$(Left$(str, Len(str))) = LCase$(Left$(sPath, Len(str))) Then
'                            GoTo Re
'                            Exit Do
'                        End If
'                    End If
''                Debug.Print str
'            End If
'                    If rs = 0 Then Exit Do
'        Loop
''            hr = f.GetUIObjectOf(hWnd, 1, lpi, uuidCM, ObjPtr(lpcm))
'            hr = f.GetUIObjectOf(hWnd, 1, N, uuidCM, ObjPtr(lpcm))
'
'    Else
'        'hr = lpsfParent.GetUIObjectOf(hWnd, 1, lpi, uuidCM, ObjPtr(lpcm))
''        hr = lpsfParent.GetUIObjectOf(hWnd, 1, lpi, uuidCM, ObjPtr(lpcm))
'        ShowIContextMenu = False
'        Exit Function
'    End If
'
'If hr Then
'        Dim nMnu As Long, nMnu2 As Long, nMnu3 As Long
'            nMnu = CreateMenu
'            nMnu2 = CreateMenu
'            nMnu3 = CreateMenu
'
'    CopyMemory lpcm, hr, 4
'        HMENU = CreatePopupMenu
''    Call lpcm.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_RESERVED Or CMF_INCLUDESTATIC Or CMF_NORMAL)
''    Call lpcm.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_CANRENAME Or CMF_NORMAL)
''    Call lpcm.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_NORMAL)
'    Call lpcm.QueryContextMenu(HMENU, 0, 1, &H7FFF, CMF_NORMAL Or CMF_INCLUDESTATIC Or CMF_NODEFAULT Or CMF_EXPLORE)
'            InsertMenu HMENU, 0, MF_STRING Or MF_BYPOSITION, nMnu, ByVal "Open Containing Folder"
'            SetMenuItemBitmaps HMENU, 0, MF_BYPOSITION Or MF_BITMAP Or MF_MASK, frmFastSearch.Picture1.Picture.handle, frmFastSearch.Picture1.Picture.handle
'        InsertMenu HMENU, 1, MF_STRING Or MF_BYPOSITION, nMnu2, ByVal "Find with Associated Program"
'        SetMenuItemBitmaps HMENU, 1, MF_BYPOSITION Or MF_BITMAP Or MF_MASK, frmFastSearch.Pic.Picture.handle, frmFastSearch.Pic.Picture.handle
'            InsertMenu HMENU, 2, MF_STRING Or MF_BYPOSITION, nMnu3, ByVal "Find with Notepad"
'            SetMenuItemBitmaps HMENU, 2, MF_BYPOSITION Or MF_BITMAP Or MF_MASK, frmFastSearch.Picture3.Picture.handle, frmFastSearch.Picture3.Picture.handle
'
'            InsertMenu HMENU, 3, MF_STRING Or MF_BYPOSITION Or MF_SEPARATOR, 0, ByVal ""
'
'        idCmd = TrackPopupMenu(HMENU, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, _
'            lppt.x, lppt.y, ByVal 0, hWnd, ByVal 0)
'    If idCmd Then
'        cmi.cbSize = Len(cmi)
'        cmi.fMask = 0
'        cmi.hWnd = hWnd
'        cmi.lpVerb = idCmd - 1 'MAKEINTRESOURCE(idCmd - 1)
'        cmi.lpParameters = 0
'        cmi.lpDirectory = 0
'        cmi.nShow = 1 ' SW_SHOWNORMAL
'        cmi.dwHotKey = 0
'        cmi.hIcon = 0
'        If idCmd = nMnu Then
'            Dim s As New Shell
'                s.Open s.NameSpace(sp)
'            Set s = Nothing
'        ElseIf idCmd = nMnu2 Then
'                ShellExecute hWnd, "open", sPath, _
'                         vbNullString, sp, SW_SHOWNORMAL
'
'        ElseIf idCmd = nMnu3 Then
'                Shell "Notepad " & sPath, vbNormalFocus
'            FindinWindow sPath, frmFastSearch.txtFInd
'
'
'        Else
'            On Error Resume Next
'            lpcm.InvokeCommand cmi
'        End If
'    End If
'        DestroyMenu HMENU
'        DestroyMenu nMnu
'        Set lpcm = Nothing
'            ShowIContextMenu = True
'Else
'    ShowIContextMenu = False
'End If
'    Set lpsfParent = Nothing
'End Function


