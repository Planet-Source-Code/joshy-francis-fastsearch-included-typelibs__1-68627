VERSION 5.00
Begin VB.Form frmFastSearch 
   Caption         =   "FastSearch"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   Icon            =   "frmFastSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLogon 
      Caption         =   "Startup"
      Height          =   330
      Left            =   7875
      TabIndex        =   19
      Top             =   90
      Width           =   1110
   End
   Begin VB.PictureBox Picture3 
      Height          =   300
      Left            =   5550
      Picture         =   "frmFastSearch.frx":0442
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   18
      Top             =   5745
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SavePic"
      Height          =   705
      Left            =   1470
      TabIndex        =   17
      Top             =   6330
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   2415
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   16
      Top             =   5655
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3795
      Picture         =   "frmFastSearch.frx":0784
      ScaleHeight     =   480
      ScaleWidth      =   240
      TabIndex        =   15
      Top             =   5550
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton cmdBrowseStartDir 
      Caption         =   "..."
      Height          =   270
      Left            =   3270
      TabIndex        =   14
      Top             =   75
      Width           =   360
   End
   Begin VB.PictureBox Picture1 
      Height          =   300
      Left            =   4995
      Picture         =   "frmFastSearch.frx":0D0E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      Top             =   5715
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ComboBox txtPattern 
      Height          =   315
      Left            =   1185
      TabIndex        =   12
      Text            =   "*.*"
      Top             =   465
      Width           =   2040
   End
   Begin VB.CommandButton cmdClearR 
      Caption         =   "&Clear Recent"
      Height          =   375
      Left            =   5205
      TabIndex        =   11
      Top             =   15
      Width           =   1215
   End
   Begin VB.ComboBox txtFInd 
      Height          =   315
      Left            =   1485
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   900
      Width           =   6225
   End
   Begin VB.ComboBox txtStartDir 
      Height          =   315
      Left            =   1170
      TabIndex        =   9
      Top             =   60
      Width           =   2055
   End
   Begin VB.CommandButton cmdKill 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   6510
      TabIndex        =   8
      Top             =   45
      Width           =   1215
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   1500
      TabIndex        =   5
      Top             =   1260
      Width           =   6225
   End
   Begin VB.ListBox lstFiles 
      Height          =   3765
      Left            =   60
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   1635
      Width           =   7695
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   3825
      TabIndex        =   0
      Top             =   15
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected File"
      Height          =   195
      Left            =   75
      TabIndex        =   7
      Top             =   1260
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Containing Words"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   885
      Width           =   1260
   End
   Begin VB.Label lblNumFiles 
      AutoSize        =   -1  'True
      Caption         =   "0 Files, 0 K"
      Height          =   195
      Left            =   3315
      TabIndex        =   4
      Top             =   465
      Width           =   780
   End
   Begin VB.Label Label2 
      Caption         =   "Pattern"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   375
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Start directory"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   15
      Width           =   1095
   End
End
Attribute VB_Name = "frmFastSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ======================================================================================
' Name:     FastSearch
' Author:   Joshy Francis (joshylogicss@yahoo.co.in)
' Date:     14 May 2007
'
' Requires: None
'
' Copyright © 2000-2007 Joshy Francis
' --------------------------------------------------------------------------------------
'The Fast File Searcing Utility.
'   *********Features**********
'   1. Explorer Context Menu.
'   2. Drag-N-Drop to Explorer.
'   3. Search with Associated program.
'   4. Pattern Searching.
'   5. Fast File Cleaning.
'   6. Startmenu Shortcut-You can select this program from Startmenu-> Search(StartMenu must be in Classic mode)
'   7. Faster than Windows Search.
'you can freely use this code anywhere.But I wants you must include the Copyright Info
'some of the functions from PSC
' --------------------------------------------------------------------------------------
'Steps in Installation
'   1.Open Project -> Goto Reference-> Add olelib.tlb & vbshell.tlb to the Project.I got it from PSC.
'          It wiil appear as below
'      Edanmo's OLE interfaces for Implements v1.51
'      VB Shell Library
'   2. Goto FastSearch Project Properties->General Tab : Change Project Type to ActiveX Dll.
'      Change 'IFastSearch' Class module's Instancing to 5-MultiUse.
'      Make the Project.
'   3. Goto FastSearch Project Properties->General Tab : Change Project Type to Standard Exe.
'      Make & Run the Project.
' If You use classic Start menu then Goto Start Menu-> Search You can see   FastSearch,clicking on it will be run our program.
'
'
'Please..........Vote my program..........

Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Const ERROR_NO_MORE_FILES = 18&
Private Const INVALID_HANDLE_VALUE = -1
Private Const DDL_DIRECTORY = &H10

Private total_size As Double
Dim StopSearch As Boolean
Dim FoundPos As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long


' List all files below the directory that
' match the pattern.
Private Sub ListFiles(ByVal start_dir As String, ByVal pattern As String, ByVal lst As ListBox)
Const MAXDWORD = 2 ^ 32
Dim dir_names() As String
Dim num_dirs As Long
Dim I As Long
Dim fname As String
Dim attr As Integer
Dim search_handle As Long
Dim file_data As WIN32_FIND_DATA
Dim file_size As Double
On Error Resume Next
                            Dim ss As String
                If Right(start_dir, 1) <> "\" Then start_dir = start_dir & "\"
    ' Get the matching files in this directory.
    ' Get the first file.
    search_handle = FindFirstFile( _
        start_dir & pattern, file_data)
            If search_handle = INVALID_HANDLE_VALUE Then
                ss = pattern
                    If InStr(ss, "*") = 0 Then
                                ss = "*" & ss
                            search_handle = FindFirstFile( _
                                start_dir & ss, file_data)
                        If search_handle = INVALID_HANDLE_VALUE Then
'                            ss = "*.*"
                            If InStr(pattern, ".") Then
                                ss = "*." & Mid$(pattern, InStr(pattern, ".") + 1)
                            Else
                                ss = "*.*"
                            End If
                                search_handle = FindFirstFile( _
                                    start_dir & ss, file_data)
                        End If
                    End If
            End If
    If search_handle <> INVALID_HANDLE_VALUE Then
            If StopSearch = True Then Exit Sub
        ' Get the rest of the files.
        Do
                DoEvents
            fname = file_data.cFileName
            fname = Left$(fname, InStr(fname, Chr$(0)) - 1)
            file_size = (file_data.nFileSizeHigh * MAXDWORD) + file_data.nFileSizeLow
            If file_size > 0 Then
                If isFileContains(start_dir & fname, txtFInd) = True Then
                    lst.AddItem start_dir & fname & " ( FoundPos #" & FoundPos & "# FileSize #" & Format$(file_size) & "# )"
                    total_size = total_size + file_size
                End If
            Else
'                    lst.AddItem start_dir & fname
            End If
            ' Get the next file.
            If FindNextFile(search_handle, file_data) = 0 Then Exit Do
        Loop
        
        ' Close the file search hanlde.
        FindClose search_handle
    End If

    ' Get the list of subdirectories.
    search_handle = FindFirstFile( _
        start_dir & "*.*", file_data)
    If search_handle <> INVALID_HANDLE_VALUE Then
        ' Get the rest of the files.
        Do
            If StopSearch = True Then Exit Sub
'        DoEvents
            If file_data.dwFileAttributes And DDL_DIRECTORY Then
                fname = file_data.cFileName
                fname = Left$(fname, InStr(fname, Chr$(0)) - 1)
                If fname <> "." And fname <> ".." Then
                    num_dirs = num_dirs + 1
                    ReDim Preserve dir_names(1 To num_dirs)
                    dir_names(num_dirs) = fname
                                ss = Trim$(txtPattern)
'                                lblNumFiles.Caption = start_dir & fname
                    If InStr(LCase$(fname), LCase$(ss)) Then
                        lst.AddItem start_dir & fname
                    End If
                    If Trim(txtFInd) <> "" Then
                        If InStr(LCase$(fname), LCase$(txtFInd)) Then
                            lst.AddItem start_dir & fname
                        End If
                    End If
                End If
            End If
            ' Get the next file.
            If FindNextFile(search_handle, file_data) = 0 Then Exit Do
        Loop
        ' Close the file search handle.
        FindClose search_handle
    End If
    ' Search the subdirectories.
    For I = 1 To num_dirs
        ListFiles start_dir & dir_names(I) & "\", pattern, lst
    Next I
End Sub
Function isFileContains(ByVal sFile As String, Optional ByVal str As String = "", _
    Optional ByVal UseLikeStatement As Boolean = False) As Boolean
On Error GoTo E
Dim f As Integer
Dim Buf As String
Dim BufLen As Long
'            lblNumFiles.Caption = sFile
        If sFile <> "" And str = "" Then
                    str = txtPattern
                        str = Replace$(str, "*", "")
'                        str = Replace$(str, ".", "")
                FoundPos = InStr(LCase$(sFile), LCase$(str))
            If FoundPos Then
                isFileContains = True
            Else
                isFileContains = False
            End If
                Exit Function
        End If
                
'    isFileContains = True
'    isFileContains = False
        str = Trim$(str)
    If str = "" Then Exit Function
        isFileContains = False
    If Trim$(sFile) = "" Then Exit Function
            FoundPos = 0
        str = LCase$(str)
f = FreeFile
    Open sFile For Binary Access Read As f
    BufLen = LOF(f) 'FileLen(sFile)
    If BufLen = 0 Then
        GoTo E
    End If
        If UseLikeStatement = True Then
                If Right$(str, 1) <> "*" Then str = str & "*"
                If Left$(str, 1) <> "*" Then str = "*" & str
        End If
'    If BufLen > 32767 Then BufLen = 32767    'Exit Function
'Dim Pointer As Long 'position of pointer in file
'Dim x As Long 'used in for...next loop
'Dim Whole As Long
'Dim Part As Long
'    Whole = LOF(f) \ BufLen 'number of whole 20k chunks
'Part = LOF(f) Mod BufLen 'remainder at the end
''Buf = String$(2000, 0) 'buffer
'Buf = Space$(BufLen) 'buffer
'Pointer = 1 'start at position 1
'    For x = 1 To Whole
'           DoEvents
'        Get #1, Pointer, Buf 'get data
'            If LCase$(Buf) Like LCase$(str) Then    ' "SPD") Then
'                isFileContains = True
'                Exit For
'                GoTo E
'            Else
'                isFileContains = False
'            End If
'        Pointer = Pointer + BufLen 'put pointer 20k later
'    Next x
''Buf = String$(Part, 0)  'copy the last bit
'Buf = Space$(Part)    'copy the last bit
'Get #1, Pointer, Buf        'get the remaining bytes at the end
    Buf = Space$(BufLen)
        Screen.MousePointer = 11
        Get #f, , Buf
'If InStr(LCase(Buf), str) Then ' "SPD") Then
        If UseLikeStatement = True Then
            If LCase$(Buf) Like LCase$(str) Then    ' "SPD") Then
                isFileContains = True
                GoTo E
            Else
                isFileContains = False
            End If
        Else
                FoundPos = InStr(LCase$(Buf), LCase$(str))
            If FoundPos Then
                isFileContains = True
                GoTo E
            Else
                isFileContains = False
            End If
        End If
'    Buf = ""
'Exit Function
E:
    Close f
    Buf = ""
        Screen.MousePointer = 0
End Function

Private Sub chkLogon_Click()
If chkLogon.Value = 1 Then
    Winlogon False, True
Else
    Winlogon False, False
End If
End Sub

Private Sub cmdBrowseStartDir_Click()
On Error Resume Next
Dim s As String
    s = BrowseForFolder("Select start directory", txtStartDir, hWnd, True, "", True, False)
        If s <> "" Then txtStartDir.Text = s
End Sub

Private Sub cmdClearR_Click()
DelReg "S"
DelReg "P"
DelReg "F"
End Sub

Private Sub cmdKill_Click()
Dim x As Long, VCount As Long, file_name As String
If MsgBox("Are u sure?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete All Files") = vbNo Then Exit Sub
For x = lstFiles.ListCount - 1 To 0 Step -1
        file_name = lstFiles.List(x)
    If DeleteFile(file_name) <> 0 Then
            lstFiles.RemoveItem x
                VCount = VCount + 1
    End If
Next
    If VCount Then
        MsgBox VCount & " File(s) deleted", vbInformation
    Else
        MsgBox "No File found", vbInformation
    End If
End Sub

Private Sub cmdSearch_Click()
Dim start_dir As String
Dim pattern As String
Dim file_list As String
Dim StartTime As Date
Dim EndTime As Date
If cmdSearch.Caption = "&Search" Then
    StartTime = Now
    lblNumFiles.Caption = "Searching..."
    Caption = "FastSearch"
            If Trim(txtPattern) <> "" Then
                AddReg "P", txtPattern.ListCount, txtPattern
                txtPattern.AddItem txtPattern.Text
            End If
            If Trim(txtStartDir) <> "" Then
                AddReg "S", txtStartDir.ListCount, txtStartDir
                txtStartDir.AddItem txtStartDir
            End If
            If Trim(txtFInd) <> "" Then
                AddReg "F", txtFInd.ListCount, txtFInd
                txtFInd.AddItem txtFInd
            End If
    txtFileName = ""
        Screen.MousePointer = vbHourglass
    cmdSearch.Caption = "&Stop"
    StopSearch = False
'    SendMessage lstFiles.hWnd, WM_SETREDRAW, 0, 0
'    LockWindowUpdate lstFiles.hWnd
        lstFiles.Clear
    '    DoEvents
        total_size = 0
        start_dir = Trim$(txtStartDir.Text)
        If Right$(start_dir, 1) <> "\" Then _
            start_dir = start_dir & "\"
        pattern = Trim$(txtPattern.Text)
        ListFiles start_dir, pattern, lstFiles
        lblNumFiles.Caption = _
            Format$(lstFiles.ListCount) & " Files, " & _
            Format$(total_size / 1024, "0.000") & " KB " & _
            Format$(total_size / 1024 / 1024, "0.0") & " MB " & _
            Format$(total_size / 1024 / 1024 / 1024, "0.00") & " GB"
    cmdSearch.Caption = "&Search"
'            RefreshAll
'    LockWindowUpdate 0
'    SendMessage lstFiles.hWnd, WM_SETREDRAW, 1, 0
        Screen.MousePointer = vbDefault
Else
    Screen.MousePointer = vbDefault
    StopSearch = True
    cmdSearch.Caption = "&Search"
End If
    EndTime = Now
Caption = "FastSearch - Second(s) taken : " & DateDiff("s", StartTime, EndTime)
'    SendMessage lstFiles.hwnd, WM_SETREDRAW, 1, 0
End Sub

Private Sub Command1_Click()
With Pic
    .BorderStyle = 0
    .ScaleMode = 3
    .AutoRedraw = True
    .Width = 210 '240
    .Height = 210 ' 240
End With
Pic.PaintPicture Picture2.Picture, 0, 0, 16, 16, 0, 0

SavePicture Pic.Image, App.Path & "\t1.bmp"
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
    RefreshAll
End If
End Sub
Sub RefreshAll()
    FillDrives txtStartDir, , , False
        FillReg "S", txtStartDir, False, True
        
    FillReg "P", txtPattern
    FillReg "F", txtFInd
End Sub
Private Sub Form_Load()
Dim s As String, C As Long
    txtStartDir.SelStart = Len(txtStartDir.Text)
        RefreshAll
    s = Trim$(Replace$(Command$, Chr(34), ""))
If s <> "" Then
        C = InStrRev(s, "\")
    If C Then
        txtStartDir = Mid$(s, 1, C)
        txtPattern = Mid$(s, C + 1)
    Else
'        MsgBox s
        s = LCase$(s)
            If InStr(s, "@shell@") Then
                Winlogon True
            Else
                MsgBox s
            End If
    End If
End If
    chkLogon.Value = IIf(IsIamShell = True, 1, 0)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'MsgBox Button
End Sub

Private Sub Form_Resize()
Dim hgt As Single
    hgt = ScaleHeight - lstFiles.Top
    If hgt < 120 Then hgt = 120
    lstFiles.Move 0, lstFiles.Top, ScaleWidth, hgt
End Sub

Private Sub lstFiles_Click()
Dim file_name As String
Dim pos As Long
    ' Remove the file's size.
    file_name = lstFiles.Text
    pos = InStrRev(file_name, " (")
    If pos > 0 Then file_name = Left$(file_name, pos - 1)
txtFileName = file_name
    SetMenuBMP file_name
End Sub
Sub SetMenuBMP(ByVal sFile As String)
    Dim ExeFile As String, hIcon As Long
        ExeFile = GetExe(sFile)
            If ExeFile = "" Then ExeFile = "%sysdir%\shell32.dll"
                
    hIcon = ExtractIcon(0, ExeFile, 0)
    
    With Pic
        .Picture = LoadPicture
        .Cls
        .BorderStyle = 0
        .ScaleMode = 3
        .AutoRedraw = True
        .Width = 210 '240
        .Height = 210 ' 240
    End With
        If hIcon Then
            DrawIconEx Pic.hdc, 0, 0, hIcon, 16, 16, 0, 0, DI_NORMAL
            Pic.PaintPicture Pic.Image, 0, 0, 16, 16, 0, 0
                Set Pic.Picture = Pic.Image
            SavePicture Pic.Image, App.Path & "\t1.bmp"
        End If
End Sub
' Open the double-clicked file.
Private Sub lstFiles_DblClick()
Dim file_name As String
Dim pos As Long
    ' Remove the file's size.
    file_name = lstFiles.Text
    pos = InStrRev(file_name, " (")
    If pos > 0 Then file_name = Left$(file_name, pos - 1)
    Dim sp As String, sf As String, st As String
        st = lstFiles.Text
    If st = "" Then Exit Sub
        sp = Mid$(st, 1, InStrRev(st, "\"))
        sf = Mid$(st, InStrRev(st, "\") + 1)
    Dim V As ITEM_VERB
    Dim s As Shell, f As FolderItem
        Set s = New Shell
            Set f = s.NameSpace(sp).Items.item(sf)
'        V = Show_ContextMenu(hWnd, f)
    On Error Resume Next
    st = GetDefaultItem(f)
        If st <> "" Then
            f.InvokeVerb ' V.sVerb
        Else
    ShellExecute Me.hWnd, "open", file_name, _
        vbNullString, vbNullString, SW_SHOWNORMAL
'            If V.sVerb = "Open Containing Folder" Then
'                s.Open s.NameSpace(sp)
'            End If
        End If
        Set s = Nothing
        Set f = Nothing
End Sub

Private Sub lstFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Dim file_name As String, sName As String
    Dim pos As Long
        ' Remove the file's size.
        file_name = lstFiles.Text
        pos = InStrRev(file_name, " (")
        If pos > 0 Then file_name = Left$(file_name, pos - 1)
    Dim sp As String, sf As String, st As String
        st = file_name
            pos = InStrRev(file_name, "\")
        If pos Then
            sName = Mid$(file_name, pos + 1)
        End If
'            pos = InStr(sName, ".")
'        If pos Then
'            sName = Left$(sName, pos - 1)
'        End If
            SetMenuBMP file_name
    If st = "" Then Exit Sub
                sp = Mid$(st, 1, InStrRev(st, "\"))
                sf = Mid$(st, InStrRev(st, "\") + 1)
'                    ShowIContextMenu st, hWnd, sp, sf
                    
            Dim V As ITEM_VERB
            Dim s As Shell, f As FolderItem
                Set s = New Shell
                    Set f = s.NameSpace(sp).Items.item(sf)
                V = Show_ContextMenu(hWnd, f)
                If V.sVerb <> "" Then
                    If V.sVerb = "Rena&me" Then
                        Dim nsf As String
                            nsf = InputBox("Enter New Name", "Rename File", sf)
                        Name file_name As Replace$(file_name, sf, nsf)
                            lstFiles.List(lstFiles.ListIndex) = Replace$(lstFiles.List(lstFiles.ListIndex), sf, nsf)
                    Else
                        f.InvokeVerb V.sVerb
                    End If
                End If
                    If V.sVerb = "Open Containing Folder" Then
        '                ShellExecute Me.hWnd, "open", sp, _
        '                    vbNullString, vbNullString, SW_SHOWNORMAL
                        s.Open s.NameSpace(sp)
                            For pos = 0 To 1000
                                DoEvents
                            Next
                            SendKeys sName
                    End If
                        Call GetDefaultItem(f)
                    If V.sVerb = "Find with Notepad" Then
                        Shell "Notepad " & file_name, vbNormalFocus
                        FindinWindow file_name, txtFInd
                    End If
                    If V.sVerb = "Find with Associated Program" Then
                            f.InvokeVerb DefaultItem
                        FindinWindow file_name, txtFInd
                    End If

                Set s = Nothing
                Set f = Nothing
End If
End Sub

Private Sub lstFiles_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    lstFiles.OLEDrag
End If
End Sub

Private Sub lstFiles_OLESetData(Data As DataObject, DataFormat As Integer)
DataFormat = vbCFFiles
End Sub

Private Sub lstFiles_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Dim file_name As String
Dim pos As Long
    ' Remove the file's size.
    file_name = lstFiles.Text
    pos = InStrRev(file_name, " (")
    If pos > 0 Then file_name = Left$(file_name, pos - 1)
If file_name <> "" Then
    AllowedEffects = vbDropEffectCopy
    Data.Clear
    Data.Files.Add file_name
    Data.SetData , vbCFFiles
End If
End Sub

Private Sub txtFInd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
If txtFInd.Text = "" Then Exit Sub
    DelReg "F", txtFInd.Text
    RefreshAll
End If
End Sub

Private Sub txtPattern_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
If txtPattern.Text = "" Then Exit Sub
    DelReg "P", txtPattern.Text
    RefreshAll
End If
End Sub

Private Sub txtStartDir_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
If txtStartDir.Text = "" Then Exit Sub
    DelReg "S", txtStartDir.Text
    RefreshAll
End If
End Sub
