VERSION 5.00
Begin VB.Form frmRegisterTypeLib 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RegisterTypeLib"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   LinkTopic       =   "RegisterTypeLib"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnRegister 
      Caption         =   "&Un Register"
      Height          =   420
      Left            =   5145
      TabIndex        =   4
      Top             =   720
      Width           =   2820
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&Register"
      Height          =   420
      Left            =   1020
      TabIndex        =   3
      Top             =   780
      Width           =   2820
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "...&."
      Height          =   390
      Left            =   8100
      TabIndex        =   2
      Top             =   150
      Width           =   585
   End
   Begin VB.TextBox txtFile 
      Height          =   360
      Left            =   960
      TabIndex        =   1
      Top             =   150
      Width           =   7005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Typelib"
      Height          =   195
      Left            =   285
      TabIndex        =   0
      Top             =   210
      Width           =   510
   End
End
Attribute VB_Name = "frmRegisterTypeLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================
' Name:     RegisterTypelib
' Author:   Joshy Francis (joshylogicss@yahoo.co.in)
' Date:     14 May 2007
'
' Requires: None
'
' Copyright © 2000-2007 Joshy Francis
' --------------------------------------------------------------------------------------
'The Typelib Registering Utility.
'you can freely use this code anywhere.But I wants you must include the Copyright Info

Private Type OpenFilename
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFilename) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OpenFilename) As Long
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000                         '  new look commdlg
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10

Function OpenFile(ByVal hwnd As Long) As String
Dim OF As OpenFilename, Ret As Long, str As String
Static FirstTime As Boolean
    With OF
                    .Flags = OFN_EXPLORER
        .hwndOwner = hwnd
        .lpstrFilter = "Typelib Files" & Chr(0) & "*.tlb" & Chr(0) & "All Files" & Chr(0) & "*.*" & Chr(0)
            If FirstTime = False Then
                .lpstrInitialDir = App.Path
                FirstTime = True
            End If
        .lStructSize = Len(OF)
        .lpstrFile = Space$(1023) & Chr(0)
        .nMaxFile = 1024
                    .lpstrTitle = "Select Typelib File for Register"
    End With
            Ret = GetOpenFileName(OF)
If Ret = 1 Then
    Dim C As Long
        C = InStrRev(OF.lpstrFile, Chr(0) & Chr(0))
        If C = 0 Then
            C = InStr(OF.lpstrFile, Chr(0))
        End If
    str = Left$(OF.lpstrFile, C - 1)
Else
    str = ""
End If
    OpenFile = str
End Function
Public Function IsDir(D) As Boolean
On Error GoTo E
        IsDir = False
    RmDir D
        MkDir D
        IsDir = True
Exit Function
E:
    If Err.Description = "Path not found" Then
        IsDir = False
    Else
        IsDir = True
    End If
End Function
Public Function IsFile(F) As Boolean
On Error GoTo E
    IsFile = False
        If InStr(F, ":") = 0 Then
            Exit Function
        End If
        If InStr(F, "\") = 0 Then
            Exit Function
        End If
Dim n As Integer
    n = FreeFile
        Open F For Input As n
        Close n
        Reset
    IsFile = True
Exit Function
E:
    Reset
IsFile = False
End Function
Public Function GetLastBackSlash(text As String) As String
    Dim I, pos As Integer
    Dim lastslash As Integer
    For I = 1 To Len(text)
        pos = InStr(I, text, "\", vbTextCompare)
        If pos <> 0 Then lastslash = pos
    Next I
    GetLastBackSlash = Right(text, Len(text) - lastslash)
End Function

Private Sub cmdBrowse_Click()
txtFile = OpenFile(hwnd)
End Sub

Private Sub cmdRegister_Click()
RegisterTypeLib_Confirm txtFile, True, True
End Sub

Private Sub cmdUnRegister_Click()
RegisterTypeLib_Confirm txtFile, False, True

End Sub
