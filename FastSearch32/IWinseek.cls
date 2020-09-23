VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFastSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements olelib.IContextMenu
Implements VBShellLib.IShellExtInit
'
Private Sub IContextMenu_GetCommandString(ByVal idCmd As Long, ByVal uType As olelib.GETCOMMANDSTRINGFLAGS, pwReserved As Long, ByVal pszName As Long, ByVal cchMax As Long)

End Sub

Private Sub IContextMenu_InvokeCommand(lpici As olelib.CMINVOKECOMMANDINFO)
Dim strExe As String
        strExe = App.Path
            If Right$(strExe, 1) <> "\" Then strExe = strExe & "\"
        strExe = strExe & App.EXEName & ".exe"
            MsgBox strExe
Shell strExe, vbNormalFocus
End Sub

Private Sub IContextMenu_QueryContextMenu(ByVal hMenu As Long, ByVal indexMenu As Long, ByVal idCmdFirst As Long, ByVal idCmdLast As Long, ByVal uFlags As olelib.QUERYCONTEXTMENUFLAGS)

End Sub

Private Sub IShellExtInit_Initialize(ByVal pidlFolder As VBShellLib.LPCITEMIDLIST, ByVal pDataObj As VBShellLib.IDataObject, ByVal hKeyProgID As VBShellLib.hkey)

End Sub
