Attribute VB_Name = "mEWmain"
Option Explicit
  Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" _
          (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, _
          ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
  Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" _
          (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, _
          ByVal wMsgFilterMax As Long) As Long
  Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
  Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" _
          (lpMsg As Msg) As Long
  Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
  Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
  Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
          (lpVersionInformation As OSVERSIONINFO) As Long
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
          (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
          lParam As Any) As Long
  Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
          ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
          ByVal cy As Long, ByVal wFlags As Long) As Long
  
  Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
  End Type
  Private Type POINTAPI
    X As Long
    Y As Long
  End Type
  Private Type Msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    Time As Long
    point As POINTAPI
  End Type
  
  Private Const PM_NOREMOVE = &H0
  Private Const PM_REMOVE = &H1
  Private Const WM_QUIT = &H12
  Private Const LVM_FIRST As Long = &H1000
  Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
  Private Const LVSCW_AUTOSIZE As Long = -1
  Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2 'Note: On last column, its width fills remaining width
  Private Const SWP_NOSIZE = &H1
  Private Const SWP_NOACTIVATE = &H10
  Private Const SWP_NOMOVE = &H2
  Private Const HWND_TOPMOST = -1
  Private Const HWND_NOTOPMOST = -2
  
  Public sLogFile As String
  
  Dim Message As Msg
  Dim strAppLog As String
  
Public Sub Main()
  'EWtip.Show
End Sub

Public Function API_DoEvents()
  '*** use the PeekMessage version if you want to use 100% CPU
  '*** frex to do background processing that doesn't rely on
  '*** windows messages
  If PeekMessage(Message, 0&, 0&, 0&, PM_REMOVE) Then
    Call TranslateMessage(Message)
    Call DispatchMessage(Message)
  End If

  '*** use the GetMessage version if you only want to do
  '*** processing if there's a message
  'If GetMessage(Message, 0&, 0&, 0&) Then
  '  Call TranslateMessage(Message)
  '  Call DispatchMessage(Message)
  'End If

  '*** the 'pure vb' poor way to do this
  'DoEvents
End Function

Public Function GetVersionString() As String
  Dim OSinfo                      As OSVERSIONINFO
  Dim retvalue                    As Integer
  OSinfo.dwOSVersionInfoSize = 148
  OSinfo.szCSDVersion = Space$(128)
  retvalue = GetVersionExA(OSinfo) 'Get Version
  With OSinfo
    Select Case .dwPlatformId
      Case 1 'If Platform is 9x/Me
        Select Case .dwMinorVersion 'Depends on Minor Version now.
          Case 0
            GetVersionString = "Windows 95"
          Case 10
            GetVersionString = "Windows 98"
          Case 90
            GetVersionString = "Windows Mellinnium"
        End Select
      Case 2 'NT Based
        Select Case .dwMajorVersion 'Depends on Major version this time.
          Case 3
            GetVersionString = "Windows NT 3.51"
          Case 4
            GetVersionString = "Windows NT 4.0"
          Case 5
            If .dwMinorVersion = 0 Then
            GetVersionString = "Windows 2000"
          Else
            GetVersionString = "Windows XP"
          End If
        End Select
      Case Else
        GetVersionString = "Failed" 'Don't think this should ever happen unless you are using a new windows I haven't heard of yet =)
    End Select
  End With
  If Not GetVersionString = "Failed" Then
    GetVersionString = GetVersionString & " Ver." & Str$(OSinfo.dwMajorVersion) _
                       + "." + LTrim(Str(OSinfo.dwMinorVersion)) & _
                       ", Build : " & Str(OSinfo.dwBuildNumber)
  End If
End Function

'Remove all trailing Chr$(0)'s
Public Function StripTerminator(sInput As String) As String
  Dim ZeroPos As Long
  ZeroPos = InStr(1, sInput, Chr$(0))
  If ZeroPos > 0 Then
    StripTerminator = Left$(sInput, ZeroPos - 1)
  Else
    StripTerminator = sInput
  End If
End Function

Public Function AutoSizeListView(lvList As ListView, lvCol As Integer)
  LockWindowUpdate lvList.hwnd        ' Lock update of ListView. Prevents ghostly text
                                      ' from appearing. I have seen it happen in other
                                      ' projects, but not this one. Always a good idea
                                      ' to use nonetheless.
  SendMessage lvList.hwnd, LVM_SETCOLUMNWIDTH, lvCol, LVSCW_AUTOSIZE_USEHEADER ' The magic of auotosize
  LockWindowUpdate 0
End Function

Public Function CreateNestedFolders(ByVal strPath As String) As Boolean
  Dim strNestedFolder As String
  Dim FSO As New FileSystemObject
  Dim Z As Long
  Dim AA() As String
  Dim strSplit As String

  On Error GoTo Error_CreateNestedFolders
  If Right$(strPath, 1) = "\" Then
    strPath = Left$(strPath, Len(strPath) - 1)
  Else
    strPath = strPath
  End If
  AA = Split(strPath, "\")
  For Z = LBound(AA) To UBound(AA)
    strSplit = AA(Z)
    If Z < UBound(AA) Then
      strNestedFolder = strNestedFolder & AA(Z) & "\"
    Else
      strNestedFolder = strNestedFolder & AA(Z)
    End If
    If Z > LBound(AA) Then
      If FSO.FolderExists(strNestedFolder) = False Then
        FSO.CreateFolder (strNestedFolder)
      End If
    End If
  Next Z
  CreateNestedFolders = True
  strNestedFolder = ""
  Exit Function

Error_CreateNestedFolders:
  MsgBox "Error: " & CStr(Err.Number) & "  " & Err.Description & _
         vbCrLf & vbCrLf & "An error occured while trying to create " & _
         strNestedFolder, vbInformation + vbOKOnly, "Error creating folder"
  CreateNestedFolders = False
End Function

Public Sub WriteALogFile(strCurrentLog As String)
  ' If Log directory is not there
  ' =============================
  If Dir(App.Path & "\Log", vbDirectory) = "" Then
    Call CreateNestedFolders(App.Path & "\Log")
  End If
  
  If Dir(App.Path & "\Log\" & (Format(Date, "dd-mm-yyyy")) & ".Log") = "" Then
    strAppLog = "ActiveX EXE/DLL/OCX/TLB/OLB Registrar ver." & App.Major & _
    "." & App.Minor & "." & App.Revision & vbCrLf & _
    "Log File Created On " & Date & " At " & Time & vbCrLf & _
    "========================================================="
    OpenFileForReadWrite App.Path & "\Log\" & (Format(Date, "dd-mm-yyyy")) & ".Log", "FileWrite", strAppLog
  ElseIf OpenFileForReadWrite(App.Path & "\Log\" & (Format(Date, "dd-mm-yyyy")) & ".Log", "FileRead") = "" Then
    strAppLog = "ActiveX EXE/DLL/OCX/TLB/OLB Registrar ver." & App.Major & _
    "." & App.Minor & "." & App.Revision & vbCrLf & _
    "Log File Created On " & Date & " At " & Time & vbCrLf & _
    "========================================================="
    OpenFileForReadWrite App.Path & "\Log\" & (Format(Date, "dd-mm-yyyy")) & ".Log", "FileWrite", strAppLog
  End If
  strAppLog = OpenFileForReadWrite(App.Path & "\Log\" & (Format(Date, "dd-mm-yyyy")) & ".Log", "FileRead")
  strAppLog = strAppLog & vbCrLf & strCurrentLog
  OpenFileForReadWrite App.Path & "\Log\" & (Format(Date, "dd-mm-yyyy")) & ".Log", "FileWrite", strAppLog
  sLogFile = App.Path & "\Log\" & (Format(Date, "dd-mm-yyyy")) & ".Log"
End Sub

Function OpenFileForReadWrite(strFileName, ReadOrWrite, Optional strWriteLine)
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
  Dim fs, F, ts, s
  Set fs = CreateObject("Scripting.FileSystemObject")
  
  If strFileName = "" Then
    If Dir(App.Path & "\Dummy.Txt") = "" Then
      fs.CreateTextFile App.Path & "\Dummy.Txt"  'Create a file
    End If
    Set F = fs.GetFile(App.Path & "\Dummy.Txt")
  Else
    If Dir(strFileName) = "" Then
      fs.CreateTextFile strFileName  'Create a file
    End If
    Set F = fs.GetFile(strFileName)
  End If
  
  If ReadOrWrite = "FileRead" Then
    Set ts = F.OpenAsTextStream(ForReading, TristateUseDefault)
    s = ts.ReadAll
    ts.Close
    OpenFileForReadWrite = s
  ElseIf ReadOrWrite = "FileWrite" Then
    Set ts = F.OpenAsTextStream(ForWriting, TristateUseDefault)
    If strWriteLine <> "" Then
      ts.Write strWriteLine
    End If
    ts.Close
  ElseIf ReadOrWrite = "FileAppend" Then
    Set ts = F.OpenAsTextStream(ForAppending, TristateUseDefault)
    If strWriteLine <> "" Then
      ts.Write strWriteLine
    End If
    ts.Close
  Else
    Exit Function
  End If
End Function

Public Function OnTop(SForm As Form)
  'Keep form on top. Note that this is switched
  'off if form is minimized, so place in resize event
  'as well.
  Const wFlags = SWP_NOMOVE Or SWP_NOSIZE
  SetWindowPos SForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, wFlags    'Window will stay on top
  DoEvents
End Function

Public Function RemoveTop(RForm As Form)
  Const wFlags = SWP_NOMOVE Or SWP_NOSIZE
  SetWindowPos RForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, wFlags    'Window will stay on top
  DoEvents
End Function

