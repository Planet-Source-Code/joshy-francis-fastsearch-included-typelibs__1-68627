Attribute VB_Name = "mShell"
'=====================
' Main Shell staff
'=====================
Public oShell As Object
Public hSysILLarge As Long
Public hSysILSmall As Long
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
'Public Const SHGFI_ATTRIBUTES = &H800                   '  get attributes
Public Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Public Const SHGFI_EXETYPE = &H2000                     '  return exe type
'Public Const SHGFI_ICON = &H100                         '  get icon
'Public Const SHGFI_ICONLOCATION = &H1000                '  get icon location
Public Const SHGFI_LARGEICON = &H0                      '  get large icon
'Public Const SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
'Public Const SHGFI_OPENICON = &H2                       '  get open icon
'Public Const SHGFI_PIDL = &H8                           '  pszPath is a pidl
'Public Const SHGFI_SELECTED = &H10000                   '  show icon in selected state
Public Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
'Public Const SHGFI_SMALLICON = &H1                      '  get small icon
'Public Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Public Const SHGFI_TYPENAME = &H400                     '  get type name
Public Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute
Public Const SHGNLI_PIDL = &H1                          '  pszLinkTo is a pidl
Public Const SHGNLI_PREFIXNAME = &H2                    '  Make name "Shortcut to xxx"
Public Const SHIFT_PRESSED = &H10    '  the shift key is pressed.
Public Const SHIFTJIS_CHARSET = 128

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
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
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
'Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

'=========================
'All shell works with PIDLs. You can get them from IShellFolder
'interface. But shell32.dll itself have undoc functions for
'working with PIDLs. All them are colling by ordinal.
'Undoc shell32 functions:
'=========================

'Memory allocation
Public Declare Function SHAlloc Lib "Shell32" Alias "#196" (ByVal cbSize As Long) As Long
Public Declare Sub SHFree Lib "Shell32" Alias "#195" (ByVal pv As Long)
Public Declare Sub ILFree Lib "Shell32" Alias "#155" (ByVal pidl As Long)
Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

'From Path to PIDL
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHSimpleIDListFromPath Lib "Shell32" Alias "#162" (ByVal szPath As String) As Long
Public Declare Function SHILFromPath Lib "Shell32" Alias "#157" (ByVal szPath As String) As Long
Public Declare Function SHILCreateFromPath Lib "Shell32" Alias "#28" (ByVal szPath As String, pidl As Long, ByVal dwAttributes As Long) As Long

'Parsing PIDLs
Public Declare Function SHILGetNext Lib "Shell32" Alias "#153" (ByVal pidl As Long) As Long
Public Declare Function SHILFindLast Lib "Shell32" Alias "#16" (ByVal pidl As Long) As Long
Public Declare Function SHILFindChild Lib "Shell32" Alias "#24" (ByVal pidlParen As Long, ByVal pidlChild As Long) As Long

'Copying and combining PIDLs
Public Declare Function SHILClone Lib "Shell32" Alias "#18" (ByVal pidl As Long) As Long
Public Declare Function SHILCloneFirst Lib "Shell32" Alias "#19" (ByVal pidl As Long) As Long
Public Declare Function SHILAppendID Lib "Shell32" Alias "#154" (ByVal pidl As Long, ByVal itemID As Long, ByVal bAddToEnd As Long) As Long
Public Declare Function SHILCombine Lib "Shell32" Alias "#25" (ByVal pidl1 As Long, ByVal pidl2 As Long) As Long
Public Declare Function SHILRemoveLast Lib "Shell32" Alias "#17" (ByVal pidl As Long) As Boolean

'PIDLs comparison
Public Declare Function SHILIsEqual Lib "Shell32" Alias "#21" (ByVal pidl1 As Long, ByVal pidl2 As Long) As Boolean
Public Declare Function SHILIsParent Lib "Shell32" Alias "#23" (ByVal pidlParen As Long, ByVal pidlChild As Long, ByVal bImmediate As Boolean) As Boolean

'Others PIDL routines
Public Declare Function SHILGetSize Lib "Shell32" Alias "#152" (ByVal pidl As Long) As Long
Public Declare Function SHILGetDisplayName Lib "Shell32" Alias "#15" (ByVal pidl As Long, sName As String) As Boolean
'Some usefull Path routines
Public Declare Function PathFileExist Lib "Shell32" Alias "#45" (ByVal sPath As String) As Boolean
Public Declare Function PathGetExtention Lib "Shell32" Alias "#31" (ByVal sPath As String) As Long
Public Declare Function PathIsDirectory Lib "Shell32" Alias "#159" (ByVal sPath As String) As Boolean
Public Declare Function CopyPointer2String Lib "kernel32" Alias "lstrcpyA" (ByVal NewString As String, ByVal OldString As Long) As Long


Public Declare Function FileIconInit Lib "shell32.dll" Alias "#660" (ByVal cmd As Boolean) As Boolean
Public Declare Function ImageList_Draw Lib "comctl32" (ByVal hIml As Long, ByVal i As Long, ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Public Const ILD_TRANSPARENT = &H1

'=========================
' Context menu staff
'=========================

Public Type POINTAPI
    X  As Long
    Y  As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (LPPOINT As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, LPPOINT As POINTAPI) As Long
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400
Public Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As Any) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Const TPM_LEFTALIGN = &H0
Public Const TPM_RIGHTBUTTON = &H2
Public Const TPM_RETURNCMD = &H100
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
'=======Retrive Shell resources=======
Public Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Public Const MF_SEPARATOR = &H800&
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
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

Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
'Public Const SEM_FAILCRITICALERRORS = &H1
'Public Const SEM_NOGPFAULTERRORBOX = &H2
'Public Const SEM_NOOPENFILEERRORBOX = &H8000
Public Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
Public Const EXCEPTION_MAXIMUM_PARAMETERS = 15
Public Const EXCEPTION_DEBUG_EVENT = 1
 Public Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    pExceptionRecord As Long    ' Pointer to an EXCEPTION_RECORD structure
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type
Public Type CONTEXT
    FltF0 As Double
    FltF1 As Double
    FltF2 As Double
    FltF3 As Double
    FltF4 As Double
    FltF5 As Double
    FltF6 As Double
    FltF7 As Double
    FltF8 As Double
    FltF9 As Double
    FltF10 As Double
    FltF11 As Double
    FltF12 As Double
    FltF13 As Double
    FltF14 As Double
    FltF15 As Double
    FltF16 As Double
    FltF17 As Double
    FltF18 As Double
    FltF19 As Double
    FltF20 As Double
    FltF21 As Double
    FltF22 As Double
    FltF23 As Double
    FltF24 As Double
    FltF25 As Double
    FltF26 As Double
    FltF27 As Double
    FltF28 As Double
    FltF29 As Double
    FltF30 As Double
    FltF31 As Double

    IntV0 As Double
    IntT0 As Double
    IntT1 As Double
    IntT2 As Double
    IntT3 As Double
    IntT4 As Double
    IntT5 As Double
    IntT6 As Double
    IntT7 As Double
    IntS0 As Double
    IntS1 As Double
    IntS2 As Double
    IntS3 As Double
    IntS4 As Double
    IntS5 As Double
    IntFp As Double
    IntA0 As Double
    IntA1 As Double
    IntA2 As Double
    IntA3 As Double
    IntA4 As Double
    IntA5 As Double
    IntT8 As Double
    IntT9 As Double
    IntT10 As Double
    IntT11 As Double
    IntRa As Double
    IntT12 As Double
    IntAt As Double
    IntGp As Double
    IntSp As Double
    IntZero As Double

    Fpcr As Double
    SoftFpcr As Double

    Fir As Double
    Psr As Long

    ContextFlags As Long
    Fill(4) As Long
End Type
Public Type EXCEPTION_POINTERS
    pExceptionRecord As EXCEPTION_RECORD
    ContextRecord As CONTEXT
End Type
'#define EXCEPTION_EXECUTE_HANDLER       1
'#define EXCEPTION_CONTINUE_SEARCH       0
'#define EXCEPTION_CONTINUE_EXECUTION    -1
Public Const EXCEPTION_EXECUTE_HANDLER = 1
Public Const EXCEPTION_CONTINUE_SEARCH = 0
Public Const EXCEPTION_CONTINUE_EXECUTION = -1
'#define SEM_FAILCRITICALERRORS      0x0001
'#define SEM_NOGPFAULTERRORBOX       0x0002
'#define SEM_NOALIGNMENTFAULTEXCEPT  0x0004
'#define SEM_NOOPENFILEERRORBOX      0x8000
Public Const SEM_FAILCRITICALERRORS = &H1
Public Const SEM_NOGPFAULTERRORBOX = &H2
Public Const SEM_NOALIGNMENTFAULTEXCEPT = &H4
Public Const SEM_NOOPENFILEERRORBOX = &H8000
Public Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Const SW_SHOWNORMAL = 1

Public Type SHITEMID
cb As Long
abID As Byte
End Type

Public Type ITEMIDLIST
mkid As SHITEMID
End Type
'Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ppidl As ITEMIDLIST) As Long
Public Declare Function SHGetSpecialFolderLocationpidl Lib "shell32.dll" Alias "SHGetSpecialFolderLocation" (ByVal hwndOwner As Long, ByVal nFolder As Long, ppidl As Long) As Long

'Public Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrcpyA Lib "kernel32" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long

Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

'Public Type UUID
'    Data1 As Long
'    Data2 As Integer
'    Data3 As Integer
'    Data4(0 To 7) As Byte
'End Type
'        Public Declare Function CLSIDFromProgID _
            Lib "ole32.dll" (ByVal lpszProgID As Long, _
            pCLSID As UUID) As Long
'        Public Declare Function CLSIDFromString _
            Lib "ole32.dll" (ByVal lpszProgID As Long, _
            pCLSID As Guid) As Long
'        Public Declare Function CLSIDFromString _
            Lib "ole32.dll" (ByVal lpszProgID As Long, _
            pCLSID As UUID) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function WideCharToMultiByteW Lib "kernel32" Alias "WideCharToMultiByte" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Public Const CP_ACP = 0  '  default to ANSI code page
Public Const CP_NONE = 0                '  No clipping of output
Public Const CP_OEMCP = 1  '  default to OEM  code page
Public Const CP_RECTANGLE = 1           '  Output clipped to rects
Public Const CP_REGION = 2              '
Public Const CP_WINANSI = 1004  '  default codepage for windows old DDE convs.
Public Const CP_WINUNICODE = 1200
'Public Declare Function StrRetToBuf Lib "SHLWAPI.DLL" Alias "StrRetToBufA" (pstr As STRRET, ByVal pidl As Long, ByVal pszBuf As String, ByVal cchBuf As Long) As Long
'Public Declare Function StrRetToBuf Lib "SHLWAPI.DLL" Alias "StrRetToBufW" (pstr As STRRET, ByVal pidl As Long, ByVal pszBuf As String, ByVal cchBuf As Long) As Long

Public DefaultItem As String
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Public Const GW_CHILD = 5

Public Const GW_HWNDFIRST = 0

Public Const GW_HWNDLAST = 1

Public Const GW_HWNDNEXT = 2

Public Const GW_HWNDPREV = 3

Public Const GW_MAX = 5

Public Const GW_OWNER = 4
'Registry Key:
'
'
'   HKEY_CLASSES_ROOT\.lnk\ShellNew\Command
'
'
'Value for Command:
'
'   RunDLL32 AppWiz.Cpl,NewLinkHere %2
'
'
'NOTE: If the Windows Desktop Update component is not installed on your computer, the following is the value for Command:
'
'   RunDLL32 AppWiz.Cpl,NewLinkHere %1

Public Declare Sub NewLinkHere Lib "AppWiz.Cpl" (hWnd As Long, _
     sPath As String) 'As Boolean
Public Declare Function OpenAs_RunDLLA Lib "shell32.dll" (ByVal hWnd As Long, _
     ByVal sPath As String) As Boolean
'#ifndef _SHLOBJ_NO_PICKICONDLG
'WINSHELLAPI int   WINAPI PickIconDlg(HWND hwnd, LPWSTR pszIconPath, UINT cbIconPath, int *piIconIndex);
'#End If
Public Declare Function PickIconDlg Lib "shell32.dll" (ByVal hWnd As Long, _
     ByVal sIconPath As String, ByVal cbIconPath As Long, piIconIndex As Long) As Long

Sub CreateShortcut()
Dim sFile As String, hw As Long
    sFile = "C:\"
    On Error Resume Next
'Call NewLinkHere(ByVal hw, ByVal sFile)
'    or
Dim sCmd As String
    sCmd = "RUNDLL32 AppWiz.Cpl,NewLinkHere " & sFile
Shell sCmd, vbNormalFocus
    
End Sub
Sub OpenWithDlg()
Dim sFile As String, Ret As Long
    sFile = "C:\t.qqq"
Dim sCmd As String
    sCmd = "RUNDLL32 shell32.dll,OpenAs_RunDLL " & sFile
Shell sCmd, vbNormalFocus
'        or
'Ret = OpenAs_RunDLLA(0, StrConv(sFile, vbUnicode))

End Sub
Sub PickIcon()
Dim sFile As String
    sFile = StrConv("%sysdir%\shell32.dll" & Space$(1024 - Len(sFile)), vbUnicode)
        
Dim nIdx As Long, Ret As Long
    nIdx = 20
Ret = PickIconDlg(0, sFile, 1024, nIdx)
    If Ret = 1 Then
        sFile = StrConv(sFile, vbFromUnicode)
            Ret = InStr(sFile, Chr(0))
        If Ret Then
            sFile = Left(sFile, Ret - 1)
        End If
    End If
End Sub
Sub SetupControlPanel()
Shell "rundll32.exe shell32.dll,Control_FillCache_RunDLL", vbNormalFocus

End Sub
'LONG UnhandledExceptionFilter(
'  STRUCT _EXCEPTION_POINTERS *ExceptionInfo // address of exception info
');
Function UnhandledExceptionFilter(ExceptionInfo As EXCEPTION_POINTERS) As Long
'Return Values
'The function returns one of the following values:
'
'Value Meaning
'EXCEPTION_CONTINUE_SEARCH The process is being debugged, so the exception should be passed (as second chance) to the application's debugger.
'EXCEPTION_EXECUTE_HANDLER If the SEM_NOGPFAULTERRORBOX flag was specified in a previous call to SetErrorMode, no Application Error message box is displayed. The function returns control to the exception handler, which is free to take any appropriate action.
'
'
'Remarks
'If the process is not being debugged, the function displays an Application Error message box, depending on the current error mode. The default behavior is to display the dialog box, but this can be disabled by specifying SEM_NOGPFAULTERRORBOX in a call to the SetErrorMode function.
'
'The system uses UnhandledExceptionFilter internally to handle exceptions that occur during process and thread creation.
MsgBox "Error Occured at " & ExceptionInfo.pExceptionRecord.ExceptionAddress
UnhandledExceptionFilter = EXCEPTION_CONTINUE_SEARCH

End Function

Sub SetAutoRedraw(ByVal hWnd As Long, ByVal b As Boolean)
SendMessage hWnd, WM_SETREDRAW, -CLng(b), ByVal 0
End Sub
'Init Shell object and get system imagelists for large and small icons
Public Sub Init_Shell()
   Dim sfi As SHFILEINFO
   Call FileIconInit(True)
   Set oShell = CreateObject("Shell.Application")
   hSysILLarge = SHGetFileInfoStr("C:\", 0, sfi, Len(sfi), SHGFI_SYSICONINDEX)
   hSysILSmall = SHGetFileInfoStr("C:\", 0, sfi, Len(sfi), SHGFI_SYSICONINDEX Or SHGFI_SMALLICON)
End Sub

Public Function FolderFromVariant(varFolder As Variant) As Object
   Set FolderFromVariant = oShell.NameSpace(varFolder)
End Function
'One more undoc future - Folder object contain absolute pidl
'at offset = 56 from its pointer,while FolderItem object
'contain relative PIDL to its parent folder at offset = 44
'from its pointer
Public Function PidlFromFolder(fld As Object) As Long
   Dim fldPtr As Long
   Dim pidlParent As Long
   Dim pidlChild As Long
   Dim pidl As Long
   Dim fldParent As Object
   If TypeName(fld) = "Folder" Then
      fldPtr = ObjPtr(fld)
      CopyMemory pidl, ByVal fldPtr + 56, 4
      PidlFromFolder = SHILClone(pidl)
   ElseIf TypeName(fld) = "FolderItem" Then
      'Get full qualify parent folder pidl
      Set fldParent = fld.Parent
      fldPtr = ObjPtr(fldParent)
      CopyMemory pidlParent, ByVal fldPtr + 56, 4
      'Get relative child folderitem pidl
      fldPtr = ObjPtr(fld)
      CopyMemory pidlChild, ByVal fldPtr + 44, 4
      PidlFromFolder = SHILCombine(pidlParent, pidlChild)
   End If
End Function

'==================================
'Different functions to obtain info
'about folder/folderitem by pidl
'==================================
Public Function GetAttributes(pidl As Long) As Long
  Dim sfi As SHFILEINFO
  Call SHGetFileInfoPidl(ByVal pidl, 0, sfi, Len(sfi), SHGFI_ATTRIBUTES Or SHGFI_PIDL)
  GetAttributes = sfi.dwAttributes
End Function

Public Function GetFileIconIndexPIDL(pidl As Long, uType As Long) As Long
  Dim sfi As SHFILEINFO
  If SHGetFileInfoPidl(ByVal pidl, 0, sfi, Len(sfi), SHGFI_PIDL Or ICON_SHGFI_FLAGS Or uType) Then
    GetFileIconIndexPIDL = sfi.iIcon
  End If
End Function

Public Function INDEXTOOVERLAYMASK(iOverlay As Long) As Long
  '   INDEXTOOVERLAYMASK(i)   ((i) << 8)
  INDEXTOOVERLAYMASK = iOverlay * (2 ^ 8)
End Function

Public Function HasSubFolders(pidl As Long) As Boolean
  Dim sfi As SHFILEINFO
  Call SHGetFileInfoPidl(ByVal pidl, 0, sfi, Len(sfi), SHGFI_ATTRIBUTES Or SHGFI_PIDL)
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
'Public Function Show_ContextMenu(ByVal hParent As Long, fldItem As Object) As ITEM_VERB
Public Function Show_ContextMenu(ByVal hParent As Long, fldItem As Object) As ITEM_VERB
    Dim pt          As POINTAPI
    Dim fldItemVerbs  As FolderItemVerbs 'Object
    Dim hMenu       As Long
    Dim sCaption    As String
    Dim i           As Long
    Dim lngCmdId    As Long
    Dim Found As Boolean
    On Error GoTo Exit_ContextMenu
    Set fldItemVerbs = fldItem.Verbs
'    Set fldItemVerbs = fldItem
'    Call CopyMemory(hMenu, ByVal ObjPtr(fldItemVerbs) + 32, 4)
        i = 44
        i = 44
            hMenu = 0
            Found = False
GethMenu:
    Do Until IsMenu(hMenu) = 1
        Call CopyMemory(hMenu, ByVal ObjPtr(fldItemVerbs) + i, 4)
        Call GetCursorPos(pt)
            i = i + 1
        If i > 1024 Then
            Exit Do
        End If
    Loop
Dim c As Long, j As Long
Dim cMenu As Long
If Found = False Then
    c = GetMenuItemCount(hMenu)
    For j = 0 To c
           sCaption = String(64, 0)
           cMenu = GetMenuString(hMenu, j, sCaption, 64, MF_BYPOSITION)
           sCaption = Left(sCaption, cMenu)
        If InStr(sCaption, "Create &Shortcut") Or InStr(sCaption, "P&roperties") Or InStr(sCaption, "Cu&t") Then
            Found = True
            Exit For
        End If
    Next
        If Found = False Then
            hMenu = 0
            GoTo GethMenu
        End If
End If
        i = GetMenuDefaultItem(hMenu, 0, MF_BYPOSITION)
    DefaultItem = Space$(64)
       i = GetMenuString(hMenu, i, DefaultItem, 64, MF_BYCOMMAND)
       DefaultItem = Left(DefaultItem, i)
    
            cMenu = CreateMenu
        InsertMenu hMenu, 0, MF_BYPOSITION Or MF_SEPARATOR, cMenu, ByVal ""
            cMenu = CreateMenu
        InsertMenu hMenu, 0, MF_BYPOSITION, cMenu, ByVal "Select"
        
'            SetMenuDefaultItem hMenu, cMenu, 0
            SetMenuItemBitmaps hMenu, 0, MF_BYPOSITION Or MF_BITMAP Or MF_MASK, frmTest.Picture1.Picture.Handle, frmTest.Picture1.Picture.Handle
'        MsgBox IsMenu(hMenu), , i
    lngCmdId = TrackPopupMenu(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, pt.X, pt.Y, 0, hParent, ByVal 0&)
    If (lngCmdId > 0) Then
       sCaption = String(32, 0)
       i = GetMenuString(hMenu, lngCmdId, sCaption, 32, MF_BYCOMMAND)
       sCaption = Left(sCaption, i)
       Show_ContextMenu.sVerb = sCaption
       Show_ContextMenu.lCmdID = (lngCmdId - 1) Or &H7000
    End If
        
Exit_ContextMenu:
    Set fldItemVerbs = Nothing
End Function
Function ShowIContextMenu(ByVal sPath As String, ByVal hWnd As Long) As Boolean
Dim lpcm As olelib.IContextMenu
Dim cmi As olelib.CMINVOKECOMMANDINFO
Dim dwdwAttribs  As Long
Dim idCmd As Long
Dim hMenu As Long
Dim hr As Long
Dim lpsfParent As olelib.IShellFolder
Dim lpi As Long
Dim lppt  As POINTAPI
Dim uuidCM As UUID
Dim sFile As String
Dim f As olelib.IShellFolder
Dim uuidSF As UUID
Dim rs As Long
Dim EI As olelib.IEnumIDList
Dim VO As olelib.IViewObject
Dim uuidVo As UUID
Dim N As Long
'sFile = "C:\Documents and Settings\Joshy Francis\Desktop\bliss.bmp"
sFile = "bliss.bmp"
Call CLSIDFromString(IIDSTR_IContextMenu, uuidCM)
Call CLSIDFromString(IIDSTR_IShellFolder, uuidSF)
Call CLSIDFromString(IIDSTR_IShellView, uuidVo)
    GetCursorPos lppt
        
Set lpsfParent = olelib.Shell32.SHGetDesktopFolder
'Set f = olelib.Shell32.SHGetDesktopFolder

Call lpsfParent.ParseDisplayName(hWnd, ByVal 0, StrPtr(sFile), Len(sFile), lpi, SFGAO_FILESYSTEM)
    Dim c As Long
        c = SHGetSpecialFolderLocation(hWnd, CSIDL_DRIVES)
    Dim str As String, st As STRRET, si As SHFILEINFO
    Dim k As Long, t As Long
Re:
        If N = 0 Then N = c
        k = 0
    If N Then
        If Not f Is Nothing Then
            Call f.BindToObject(N, 0, uuidSF, k)
                Set f = Nothing
        Else
                Call lpsfParent.BindToObject(N, 0, uuidSF, k)
        End If
    End If
If k Then CopyMemory f, k, 4
    If Not f Is Nothing Then
'            sFile = sPath '"C:\t.txt"
'                    lpi = 0
'        Call f.ParseDisplayName(hWnd, ByVal 0, StrPtr(sFile), Len(sFile), lpi, SFGAO_FILESYSTEM)
    Set EI = f.EnumObjects(hWnd, SHCONTF_NONFOLDERS Or SHCONTF_INCLUDEHIDDEN Or SHCONTF_FOLDERS)
            N = 0: c = 0: k = 0
        Do
                rs = EI.Next(c, N)
                str = Space$(255)
'                    str = StrConv(str, vbUnicode)
            If N Then
                        c = c + 1
'                    Dim st As STRRET
'                f.GetDisplayNameOf N, SHGDN_FORPARSING, st
                f.GetDisplayNameOf N, SHGDN_FORPARSING, st
                        str = StrConv(Space$(260), vbUnicode)
                    StrRetToBuf VarPtr(st), N, str, 260
                        t = InStr(str, Chr(0))
                            If t Then
                                str = Left$(str, t - 1)
                            End If
                    If LCase$(sPath) = LCase$(str) Then
                        Exit Do
                    Else
                        If LCase$(Left$(str, Len(str))) = LCase$(Left$(sPath, Len(str))) Then
                            GoTo Re
                            Exit Do
                        End If
                    End If
'                Debug.Print str
            End If
                    If rs = 0 Then Exit Do
        Loop
'            hr = f.GetUIObjectOf(hWnd, 1, lpi, uuidCM, ObjPtr(lpcm))
            hr = f.GetUIObjectOf(hWnd, 1, N, uuidCM, ObjPtr(lpcm))
        
    Else
        'hr = lpsfParent.GetUIObjectOf(hWnd, 1, lpi, uuidCM, ObjPtr(lpcm))
'        hr = lpsfParent.GetUIObjectOf(hWnd, 1, lpi, uuidCM, ObjPtr(lpcm))
        ShowIContextMenu = False
        Exit Function
    End If

If hr Then
        Dim nMnu As Long
            nMnu = CreateMenu
                
    CopyMemory lpcm, hr, 4
        hMenu = CreatePopupMenu
'    Call lpcm.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_RESERVED Or CMF_INCLUDESTATIC Or CMF_NORMAL)
'    Call lpcm.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_CANRENAME Or CMF_NORMAL)
'    Call lpcm.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_NORMAL)
    Call lpcm.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_NORMAL Or CMF_INCLUDESTATIC Or CMF_NODEFAULT Or CMF_EXPLORE)
            InsertMenu hMenu, 0, MF_STRING Or MF_BYPOSITION, nMnu, ByVal "Test Item"
            SetMenuItemBitmaps hMenu, 0, MF_BYPOSITION Or MF_BITMAP Or MF_MASK, frmTest.Picture1.Picture.Handle, frmTest.Picture1.Picture.Handle
            InsertMenu hMenu, 1, MF_STRING Or MF_BYPOSITION Or MF_SEPARATOR, 0, ByVal ""
                
        idCmd = TrackPopupMenu(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, _
            lppt.X, lppt.Y, ByVal 0, hWnd, ByVal 0)
    If idCmd Then
        cmi.cbSize = Len(cmi)
        cmi.fMask = 0
        cmi.hWnd = hWnd
        cmi.lpVerb = idCmd - 1 'MAKEINTRESOURCE(idCmd - 1)
        cmi.lpParameters = 0
        cmi.lpDirectory = 0
        cmi.nShow = SW_SHOWNORMAL
        cmi.dwHotKey = 0
        cmi.hIcon = 0
        If idCmd = nMnu Then
            MsgBox "Test Item"
        Else
            On Error Resume Next
            lpcm.InvokeCommand cmi
        End If
    End If
        DestroyMenu hMenu
        DestroyMenu nMnu
        Set lpcm = Nothing
            ShowIContextMenu = True
Else
    ShowIContextMenu = False
End If
    Set lpsfParent = Nothing
End Function
Function GetDefaultItem(fi As FolderItem) As String
    Dim fiv  As FolderItemVerbs 'Object
    Dim hMenu       As Long
    Dim i           As Long
    Dim Found As Boolean
    Set fiv = fi.Verbs
        i = 44
            hMenu = 0
            Found = False
GethMenu:
    Do Until IsMenu(hMenu) = 1
        Call CopyMemory(hMenu, ByVal ObjPtr(fiv) + i, 4)
            i = i + 1
        If i > 1024 Then
            Exit Do
        End If
    Loop
        i = GetMenuDefaultItem(hMenu, 0, MF_BYPOSITION)
    DefaultItem = Space$(64)
       i = GetMenuString(hMenu, i, DefaultItem, 64, MF_BYCOMMAND)
       DefaultItem = Left(DefaultItem, i)
    GetDefaultItem = DefaultItem
        Set fiv = Nothing
End Function
'Show default shell menu for ViewWindow
'Return ITEM_VERB structure from menu.lCmdId contain one of the
'IDM_SHVIEW_ const
Public Function Show_ShellMenu(ByVal hParent As Long, Optional ByVal MenuID As Long = 215) As ITEM_VERB
   Dim pt          As POINTAPI
   Dim hModule As Long
   Dim hMenu       As Long
   Dim sCaption    As String
   Dim i           As Long
   Dim lngCmdId    As Long
   On Error GoTo Exit_ShellMenu
   hModule = LoadLibrary("shell32.dll")
   If hModule Then
'      hMenu = GetSubMenu(LoadMenu(hModule, 215), 0)
      hMenu = GetSubMenu(LoadMenu(hModule, MenuID), 0)
      Call FreeLibrary(hModule)
   End If
'    MsgBox IsMenu(hMenu)
   Call GetCursorPos(pt)
   lngCmdId = TrackPopupMenu(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, pt.X, pt.Y, 0, hParent, ByVal 0&)
   If (lngCmdId > 0) Then
      sCaption = String(32, 0)
      i = GetMenuString(hMenu, lngCmdId, sCaption, 32, MF_BYCOMMAND)
      sCaption = Left(sCaption, i)
      Show_ShellMenu.sVerb = sCaption
      Show_ShellMenu.lCmdID = lngCmdId
   End If
Exit_ShellMenu:
   DestroyMenu hMenu
End Function

'Invoke verb by Name. For some reason (bug?), doesn't allow late binding,i.e.
'you have to pass fldItem explicit As FolderItem, not As Object
'This is the only reason why project have reference on shell automation.
Public Function ShellInvokeVerb(fldItem As FolderItem, sVerb As String) As Boolean
   On Error Resume Next
   fldItem.InvokeVerb sVerb
   ShellInvokeVerb = (Err = 0)
End Function

'Return locale String for different shell messages
Public Function GetShellResourceString(idString As Long) As String
  Dim hModule As Long
  Dim sBuf As String * MAX_PATH
  Dim nChars As Long
  hModule = LoadLibrary("shell32.dll")
  If hModule Then
    nChars = LoadString(hModule, idString, sBuf, MAX_PATH)
    If nChars Then GetShellResourceString = Left$(sBuf, nChars)
    Call FreeLibrary(hModule)
  End If
End Function

'VB TreeView and ListView can sort items only by strings
'Here is process to compare items by date and size
Public Function CompareProc(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal lParamSort As Long) As Long
   Dim nColumn As Long
   Dim hr As Long
   Dim sCompare1 As String
   Dim sCompare2 As String
   Dim sFormatString As String
   Dim Item1 As Object, Item2 As Object
   Dim fldItem1 As Object, fldItem2 As Object
   On Error GoTo ErrCompare
   Set Item1 = GetItemFromlParam(lParam1)
   Set Item2 = GetItemFromlParam(lParam2)
   Set fldItem1 = Item1.Tag
   Set fldItem2 = Item2.Tag
   If (TypeName(fldItem1) = "Folder") Then Set fldItem1 = fldItem1.Items.Item
   If (TypeName(fldItem2) = "Folder") Then Set fldItem2 = fldItem2.Items.Item
   nColumn = lParamSort And Not SORT_DESCENDING
   If nColumn = 0 Then
      sCompare1 = Item1.Text
      sCompare2 = Item2.Text
   Else
      sCompare1 = Item1.SubItems(nColumn)
      sCompare2 = Item2.SubItems(nColumn)
   End If
   If (fldItem1.IsFolder And fldItem2.IsFolder) Or ((fldItem1.IsFolder = False) And (fldItem2.IsFolder = False)) Then
      If IsDate(sCompare1) And IsDate(sCompare2) Then
         hr = Sgn(CDate(sCompare1) - CDate(sCompare2))
      ElseIf IsSize(sCompare1) And IsSize(sCompare2) Then
         hr = Sgn(fldItem1.Size - fldItem2.Size)
      Else
         hr = StrComp(sCompare1, sCompare2)
      End If
      If (lParamSort And SORT_DESCENDING) Then hr = hr * (-1)
   Else
      If fldItem1.IsFolder Then hr = -1 Else hr = 1
   End If
ErrCompare:
    Set Item1 = Nothing
    Set Item2 = Nothing
    Set fldItem1 = Nothing
    Set fldItem2 = Nothing
    CompareProc = hr
End Function

'Determining if string can be size.
'Donno about all locales (espec. China, Japan etc)
'But usually size string looks like xxx KB, where
'xxx is numeric
Public Function IsSize(s As String) As Boolean
  On Error Resume Next
  IsSize = IsNumeric(Left(s, Len(s) - 3))
End Function

'Thanks to Brad Martinez for this trick. It's very usefull
'for comparing process

Public Function GetItemFromlParam(lParam As Long) As Object
  Dim pItem As Long
  Dim oItem As Object
  If lParam Then
    CopyMemory pItem, ByVal lParam + 8, 4
    If pItem Then
      CopyMemory oItem, pItem, 4&
      Set GetItemFromlParam = oItem
      FillMemory oItem, 4, 0
    End If
  End If
End Function

Public Function GetExtention(sPath As String) As String
   Dim ptr As Long, s As String
   ptr = PathGetExtention(ByVal sPath)
   s = String(255, Chr$(0))
   CopyPointer2String s, ptr
   s = Left(s, InStr(s, Chr$(0)) - 1)
'Remove dot from extention
   If s <> "" Then GetExtention = Mid(s, 2)
End Function
Sub ShellWndTest()
Dim Ret As Long, ShellWnd As Long 'SHELLDLL_DefView
'&H7000=28672
Dim str As String, c As Integer

ShellWnd = FindWindow("CabinetWClass", vbNullString)
        ShellWnd = GetWindow(ShellWnd, GW_CHILD)
    Do
        str = Space$(255)
            GetClassName ShellWnd, str, 255
        c = InStr(str, Chr(0))
            If c Then
                str = Left$(str, c - 1)
            End If
        If str = "SHELLDLL_DefView" Then
            Exit Do
        End If
            ShellWnd = GetWindow(ShellWnd, GW_HWNDNEXT)
    Loop Until IsWindow(ShellWnd) = 0
        If ShellWnd = 0 Then Stop
Dim X As Long
'28723=Choose Details Ret 0 =Ok
'28722=Properties
'28755=Toggle Hide Icons
'28756=Refresh Icons
'28771=Folder options
'28784=Help & Support Centre
'28689=Delete File
'28690=Renames Selected File
'28696=Cuts Selected File
'28697=Copies Selected File
'28697=Paste  Copied File
'28700=Paste Shortcut
'28702=Copy Items Dialog for Pasting the  Copied File
'28703=Move Items Dialog for Pasting the  Copied File
'28705=Select All
'28706=Invert Selection
'28707=Select None
'28714=View Icons
'28715=View List
'28716=View Details
'28717=View Thumbnails
'28718=View Tiles
'28719=View FolderView
'28756 = toggle Align to grid & Auto arrange
'28687 = Create Shortcut here
'    Ret = SendMessage(ShellWnd, WM_COMMAND, 28719, ByVal 0)
'    Ret = SendMessage(ShellWnd, WM_COMMAND, 28718, ByVal 0)
'    Ret = SendMessage(ShellWnd, WM_COMMAND, 28717, ByVal 0)
'    Ret = SendMessage(ShellWnd, WM_COMMAND, 28716, ByVal 0)
'    Ret = SendMessage(ShellWnd, WM_COMMAND, 28715, ByVal 0)
c = 0
'&H7000=28672
    For X = 28685 To 28790
'    If c = 48 Then Stop
'        If c <> 49 Then
'            Ret = SendMessage(ShellWnd, WM_COMMAND, X, ByVal 0)
'            Ret = SendMessage(ShellWnd, WM_COMMAND, 28755, ByVal 0)
        If X <> 28721 And X <> 28755 Then
'            Ret = SendMessage(ShellWnd, WM_COMMAND, X, ByVal "C:\Bliss.bmp")
            Ret = SendMessage(ShellWnd, WM_COMMAND, X, ByVal 0)
        End If
'        End If
    Debug.Print Ret & " = " & X & " , " & c
        c = c + 1
Next
End Sub
Function IVD(ByVal lParam As Long) As Long
    IVD = 1
End Function
