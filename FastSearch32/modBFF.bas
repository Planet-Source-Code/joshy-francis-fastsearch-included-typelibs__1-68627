Attribute VB_Name = "modBFF"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' modBrowseForFolder        © 2004 by Marco Wünschmann  ''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Beschreibung                                          ''
''    Zeigt den bekannten BrowseForFolder-Dialog an, mit ''
''    dem es möglich ist, ein beliebiges Verzeichnis     ''
''    auszuwählen. Dabei können zahlreiche Dialog-       ''
''    Parameter übergeben werden.                        ''
''                                                       ''
'' Aufruf                                                ''
''    Zum Anzeigen des Dialogs wird einfach die Funktion ''
''    BrowseForFolder mit den entsprechenden Parametern  ''
''    aufgerufen (Parametererklärung in der Funktion).   ''
''    Wird der Dialog mit OK geschlossen, gibt die Funk- ''
''    tion das ausgewählte Verzeichnis zurück, andern-   ''
''    falls einen Leerstring.                            ''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' == Dialog-Einstellungen ================================

' Text der vor dem aktuell ausgewählen Verzeichnis angezeigt wird,
' falls der ShowCurrentPath-Paramter True ist.
Private Const DialogCurrentSelectionText As String = "Auswahl: "


' == API-Deklarationen ===================================

Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type


Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
    Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
    Alias "SHGetPathFromIDListA" (ByVal PIDL As Long, _
    ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function ILCreateFromPath Lib "shell32" Alias "#157" _
    (ByVal Path As String) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hmem As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, _
    lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long


Private Const MAX_PATH = 260

Private Const WM_USER = &H400

Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED As Long = 2
Private Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Private Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const BIF_STATUSTEXT As Long = &H4

Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

Public Function BrowseForFolder(DialogText As String, DefaultPath As String, _
    OwnerhWnd As Long, Optional ShowCurrentPath As Boolean = True, _
    Optional RootPath As Variant, Optional NewDialogStyle As Boolean = False, _
    Optional IncludeFiles As Boolean = False) As String
    
    ' Zeigt den BrowseForFolder-Dialog an.
    
    ' Parameter:
    '    o DialogText        Dialogtext, der oben im Dialog angezeigt wird.
    '    o DefaultPath       Standardmäßig ausgewähltes Verzeichnis.
    '    o OwnerhWnd         hWnd des übergeordneten Fensters (in den meisten
    '                          Fällen Me.hWnd).
    '    o ShowCurrentPath   Legt fest, ob die aktuelle Verzeichnisauswahl
    '                          angezeigt werden soll.
    '    o RootPath          Root-Verzeichnis. Wird es angegeben, werden nur die
    '                          Ordner unterhalb dieses Verzeichnisses angezeigt.
    '    o NewDialogStyle    Legt fest, ob der Dialog in der neuen Darstellung
    '                          angezeigt werden soll (Dialog kann vergrößert/
    '                          verkleinert werden, es ist eine Schaltfläche zum
    '                          Anlegen eines neuen Ordners vorhanden, es können
    '                          Dateioperationen wie löschen etc. ausgeführt
    '                          werden, ...). Ist dieser Parameter True, hat der
    '                          Parameter ShowCurrentPath keine Wirkung.
    '                          Nicht unter Windowsversionen <= Win95 verfügbar.
    '    o IncludeFiles      Legt fest, ob auch Dateien im Dialog angezeigt und
    '                          ausgewählt werden können.
    
    Dim InfoBrowse As BROWSEINFO
    Dim lPIDL As Long
    Dim sBuffer As String
    Dim lpSelPath As Long
    Dim lpPathBuffer As Long

    With InfoBrowse
        ' Handle des übergeordneten Fensters
        .hOwner = OwnerhWnd
        
        ' PIDL des Rootordners
        If Not IsMissing(RootPath) Then .pidlRoot = PathToPIDL(RootPath)
        
        ' Dialogtext
        .lpszTitle = DialogText
        
        ' Stringbuffer für aktuell selektierten Pfad reservieren und
        ' Adresse zuweisen
        If ShowCurrentPath Then
            lpPathBuffer = LocalAlloc(LPTR, MAX_PATH)
            .pszDisplayName = lpPathBuffer
        End If
        
        ' Dialogeinstellungen
        .ulFlags = BIF_RETURNONLYFSDIRS + _
            IIf(ShowCurrentPath, BIF_STATUSTEXT, 0) + _
            IIf(NewDialogStyle, BIF_NEWDIALOGSTYLE, 0) + _
            IIf(IncludeFiles, BIF_BROWSEINCLUDEFILES, 0)
        
        ' Callbackfunktion-Adresse zuweisen
        .lpfnCallback = FARPROC(AddressOf CallbackString)
        
        ' Stringspeicher für vorselektierten Ordner reservieren
        lpSelPath = LocalAlloc(LPTR, Len(DefaultPath) + 1)
        
        ' Vorselektierten Ordnerpfad in den reservierten Speicherbereich
        ' kopieren
        CopyMemory ByVal lpSelPath, ByVal DefaultPath, Len(DefaultPath) + 1
        
        ' Adresse des vorselektierten Ordnerpfades zuweisen (wird im
        ' lpData-Parameter an die Callback-Funktion weitergeleitet)
        .lParam = lpSelPath
    End With

    ' BrowseForFolder-Dialog anzeigen
    lPIDL = SHBrowseForFolder(InfoBrowse)

    If lPIDL Then
        ' Stringspeicher reservieren
        sBuffer = Space$(MAX_PATH)
    
        ' Selektierten Pfad aus der zurückgegebenen PIDL ermitteln
        SHGetPathFromIDList lPIDL, sBuffer
        
        ' Nullterminierungszeichen des Strings entfernen
        sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        
        ' Selektierten Pfad zurückgeben
        BrowseForFolder = sBuffer
        
        ' Reservierten Task-Speicher wieder freigeben
        Call CoTaskMemFree(lPIDL)
    End If

    ' Stringspeicher wieder freigeben
    Call LocalFree(lpSelPath)
    Call LocalFree(lpPathBuffer)
End Function

Private Function CallbackString(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal lParam As Long, ByVal lpData As Long) As Long
    
    ' Callback-Funktion des BrowseForFolder-Dialogs. Wird bei eintretenden Ereignissen des Dialogs aufgerufen.
    
    Dim sBuffer As String
    
    ' Meldungen herausfiltern
    Select Case uMsg
    Case BFFM_INITIALIZED
        ' Dialog wurde initialisiert
        
        ' Zu selektierenden Pfad (dessen Pointer wurde in lpData übergeben) an den Dialog senden
        Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    Case BFFM_SELCHANGED
        ' Selektierung hat sich geändert
        
        ' Stringspeicher reservieren
        sBuffer = Space$(MAX_PATH)
        
        ' Aktuell selektierten Pfad ermitteln und diesen an den Dialog senden
        If SHGetPathFromIDList(lParam, sBuffer) Then Call SendMessage(hWnd, BFFM_SETSTATUSTEXTA, 0&, ByVal DialogCurrentSelectionText & sBuffer)
    End Select
End Function

Private Function FARPROC(pfn As Long) As Long
    ' Funktion wird benötigt, um Funktions-Adresse ermitteln zu können, dessen Adresse mit AddressOf übergeben und anschließend wieder zurückgegeben wird.
    
    FARPROC = pfn
End Function

Private Function PathToPIDL(ByVal Path As String) As Long
    ' Konvertiert einen Pfad in dessen PIDL.
    
    Dim lRet As Long
    
    lRet = ILCreateFromPath(Path)
    If lRet = 0 Then
        Path = StrConv(Path, VbStrConv.vbUnicode)
        lRet = ILCreateFromPath(Path)
    End If
    
    PathToPIDL = lRet
End Function
