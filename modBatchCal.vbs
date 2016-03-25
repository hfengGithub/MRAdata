'===== modBatchCal Pwd=MRARAW

Option Explicit
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = -1&
Private Const BATCHCAL_LOG As String = "BatchCal_Log"
Private Const SETTINGS As String = "Settings"


Public Function BatchCal(sRunID As String, sDescription As String, sIn As String, sOut As String, sParameter As String, sCriteria As String, sCat As String, sNewFile As String) As String

  On Error GoTo ErrorHandler
  
  Dim oFSO As FileSystemObject
  
  Dim oOut As File
  
  Dim osOut As TextStream
  Dim osIn As TextStream
  
  Dim sTempFileName As String
  Dim sTempBatFile As String
  Dim sTempMainBatFile As String
  Dim sCmdLine As String
  Dim sRunCmd As String
  Dim sLine As String
  Dim lStatus As Double
  
  
  Dim sTempDir As String
  
  If Trim(sCriteria) <> "" Then
    If MsgBox(sCriteria, vbYesNo, "Confirmation") = vbNo Then
        BatchCal = "run Skipped"
        Exit Function
    End If
  End If
    
  sTempDir = GetAppSettings("Working Dir")
  
  If sTempDir = "" Then
    sTempDir = "c:\temp"
  End If
  
  If sNewFile <> "" Then
    DeleteFile sOut, sNewFile
  End If
  
  sTempBatFile = sTempDir & "\pptemp_" & sCat & sRunID & ".bat"
  sTempMainBatFile = sTempDir & "\ppmain_" & sCat & sRunID & ".bat"
  
  sTempFileName = sTempDir & "\ppmain_" & sCat & sRunID & ".log"
  
  Set oFSO = New FileSystemObject
  If sIn <> "" Then
    If Not oFSO.FileExists(sIn) Then
        MsgBox sIn & " file not exists, please check the file location. ", vbCritical
        BatchCal = "Error: " & sIn & " file not exists, please check the file location. "
        Exit Function
    End If
  End If
  
  Set osIn = oFSO.CreateTextFile(sTempBatFile, True)

  
  If Not oFSO.FolderExists(Mid(sOut, 1, InStrRev(sOut, "\") - 1)) Then
      oFSO.CreateFolder Mid(sOut, 1, InStrRev(sOut, "\") - 1)
  End If
  
  sCmdLine = GetAppSettings("BatchCal EXE") & " " & sParameter
  If Trim(sIn) <> "" Then
    sCmdLine = sCmdLine & " """ & sIn & """ "
  End If
  If Trim(sOut) <> "" Then
    sCmdLine = sCmdLine & " """ & sOut & """ "
  End If
  
  sRunCmd = sCmdLine

  osIn.WriteLine sCmdLine
  osIn.WriteLine " @ echo off"
  osIn.WriteLine " if %ERRORLEVEL% == 0 GOTO SUCCESS "
  osIn.WriteLine "  echo 'Error Return: %ERRORLEVEL%'"
  osIn.WriteLine "  exit -1 "
  osIn.WriteLine " :SUCCESS"
  osIn.WriteLine "  exit 0"
  osIn.Close
  
  Set osIn = oFSO.CreateTextFile(sTempMainBatFile, True)
  sCmdLine = sTempBatFile & " > " & sTempFileName & " 2>&1"
  osIn.WriteLine sCmdLine
  osIn.Close
  
  Dim sWinStyle As String
  sWinStyle = GetAppSettings("CMD Window")
 ' frmWait.Label1 = "BatchCal is running, please wait..."
 ' frmWait.lblCmdLine = "Started at: " & Now() & vbCrLf & sRunCmd
  
 ' frmWait.Show
  
  If sWinStyle = "hidden" Then
    ShellAndWait sTempMainBatFile, vbHide
  Else
    ShellAndWait sTempMainBatFile, vbNormalFocus
  End If
 ' Unload frmWait
  
  
  Set osOut = oFSO.OpenTextFile(sTempFileName, ForReading)

  While osOut.AtEndOfStream <> True
    sLine = Mid(osOut.ReadLine, 1, 1024)
    WriteLog sRunID, sLine
  Wend
  osOut.Close
  
  On Error Resume Next
  oFSO.DeleteFile sTempBatFile
  oFSO.DeleteFile sTempMainBatFile
  oFSO.DeleteFile sTempFileName
    
  BatchCal = "Finished"
  
  Exit Function
  
ErrorHandler:
  BatchCal = "Error: " & Err.Number & ": " & Err.Description & " " & sCmdLine
End Function

Public Function WriteLog(sRunID As String, sLine As String)

  frmRun.lstStatus.AddItem sLine
  
End Function
Public Function IsFileExist(sFile As String) As Boolean

  On Error Resume Next
  
  Dim oFSO As FileSystemObject
  
  
  IsFileExist = False
  
  Set oFSO = New FileSystemObject
  
  If oFSO.FileExists(sFile) Then
      IsFileExist = True
   End If
  
  Set oFSO = Nothing

  
 
End Function


Public Sub DeleteFile(sFile As String, sNewFile As String)
   Dim oFSO As FileSystemObject
   Dim i As Integer
   
   If sNewFile = "1" Then
     Dim sFiles() As String
     Set oFSO = New FileSystemObject
     sFiles = Split(sFile, ",")
     For i = 0 To UBound(sFiles)
      If oFSO.FileExists(sFiles(i)) Then
       oFSO.DeleteFile sFiles(i)
      End If
     Next
     Set oFSO = Nothing
   End If
   

End Sub


Public Function GetAppSettings(sKey As String) As String
  Dim sTmp As String
  sTmp = ""
  If UCase(sKey) = "BATCHCAL EXE" Then
     sTmp = "\\w2k3applyp01\polypaths\BatchCal.exe"
  ElseIf UCase(sKey) = UCase("CMD Window") Then
     sTmp = IIf(frmRun.chkShowRunWin.Value, "", "hidden")
  ElseIf UCase(sKey) = UCase("Working Dir") Then
     sTmp = "d:\temp\test"
  ElseIf UCase(sKey) = UCase("MR Dir") Then
     sTmp = "K:\MRA\Polypaths\DailyInputs\PolyTalkOutput\"
  End If
  
  GetAppSettings = sTmp
  
End Function

Public Sub ShellAndWait(ByVal program_name As String, _
    ByVal window_style As VbAppWinStyle)
Dim process_id As Long
Dim process_handle As Long

    ' Start the program.
    On Error GoTo ShellError
    process_id = Shell(program_name, window_style)
    On Error GoTo 0

    DoEvents

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If

    Exit Sub

ShellError:
    MsgBox "Error starting task " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Sub


Public Function ExtractRows(sRunID As String, sDescription As String, sIn As String, sOut As String, iSkip As Integer, iTotal As Long, sCriteria As String, sNewFile As String) As String

  On Error GoTo ErrorHandler
  
  Dim oFSO As FileSystemObject
  
  Dim oIn As File
  Dim oOut As File
  
  Dim osIn As TextStream
  Dim osOut As TextStream
  
  Dim sLine As String
  
  Dim i As Integer
  Dim j As Long
  Dim iCons As Integer
  Dim sCrit() As String
  Dim sCols() As String
  Dim sTemps() As String
  Dim sCon() As String
  Dim sHeaders() As String
  Dim oHeaderIndex As New Collection
  Dim sReturn As String
  
  Dim bOutput As Boolean
  
  If Trim(sOut) = "" Then
    bOutput = False
  Else
    bOutput = True
  End If
  
  If Trim(sCriteria) <> "" Then
    sCrit = Split(sCriteria, ",")
    i = 0
    iCons = 0
    
    ReDim sCols(UBound(sCrit)) As String
    ReDim sCon(UBound(sCrit)) As String
    While i <= UBound(sCrit)
      If sCrit(i) <> "" And InStr(sCrit(i), "=") > 0 Then
        sTemps = Split(sCrit(i), "=")
        sCols(i) = UCase(sTemps(0))
        sCon(i) = sTemps(1)
        iCons = iCons + 1
      End If
      i = i + 1
    Wend
  
  End If
  
  Set oFSO = New FileSystemObject
  
  If Not oFSO.FileExists(sIn) Then
    MsgBox sIn & " file not exists, please check the file location. ", vbCritical
    ExtractRows = "Error: " & sIn & " file not exists, please check the file location. "
    Exit Function
  End If
  
  
  Set osIn = oFSO.OpenTextFile(sIn, ForReading)
  If bOutput Then
    If Not oFSO.FileExists(sOut) Or sNewFile = "1" Then
      If Not oFSO.FolderExists(Mid(sOut, 1, InStrRev(sOut, "\") - 1)) Then
          oFSO.CreateFolder Mid(sOut, 1, InStrRev(sOut, "\") - 1)
      End If
      Set osOut = oFSO.CreateTextFile(sOut, True)
    Else
      Set osOut = oFSO.OpenTextFile(sOut, ForAppending)
    End If
  End If
  
  sLine = osIn.ReadLine
  sHeaders = Split(sLine, ",")

  For i = 0 To UBound(sHeaders)
    sHeaders(i) = UCase(sHeaders(i))
    oHeaderIndex.Add i, sHeaders(i)
  Next i
  
    'check if all columns exists
  
  Dim C
  Dim bFind As Boolean
  
  For i = 0 To iCons - 1
    bFind = False
    For Each C In sHeaders
      If C = sCols(i) Then
        bFind = True
        GoTo FindCol
      End If
    Next
    If Not bFind Then
      sReturn = "Error. Cannot find column : " & sCols(i) & " in input file: " & sIn
      MsgBox "Cannot find column : " & sCols(i) & " in input file: " & sIn, vbCritical
      GoTo ExitLoop
    End If
FindCol:
    bFind = True
  Next i
  
  
  i = 0
  Do While osIn.AtEndOfLine <> True And i < (iSkip - 1)
      osIn.SkipLine
      i = i + 1
  Loop
  
  If iSkip <= 0 Then
    If bOutput Then
        osOut.WriteLine sLine
    End If
    j = 1
  Else
    j = 0
  End If
  
  If iTotal <> -1 And j >= iTotal Then
      GoTo ExitLoop
  End If
  
  Do While osIn.AtEndOfStream <> True
    sLine = osIn.ReadLine
    If Trim(sLine) <> "" Then
        If Trim(sCriteria) <> "" Then
            sTemps = Split(cleanComma(sLine), ",")
            i = 0
            While i < iCons
              If sCon(i) = "NotNull" Then
                 If Trim(sTemps(oHeaderIndex(sCols(i)))) <> "" Then
                   If bOutput Then
                    osOut.WriteLine sLine
                   End If
                   j = j + 1
                 End If
              ElseIf sTemps(oHeaderIndex(sCols(i))) Like sCon(i) Then
                 If bOutput Then
                    osOut.WriteLine sLine
                 End If
                 j = j + 1
              End If
             i = i + 1
            Wend
        Else
            If bOutput Then
                osOut.WriteLine sLine
            End If
             j = j + 1
        End If
    End If
    If iTotal <> -1 And j >= iTotal Then
      GoTo ExitLoop
    End If
  Loop
  sReturn = j & " counts extracted"
  
ExitLoop:
  
  If bOutput Then
    osOut.Close
  End If
  osIn.Close
  
  ExtractRows = sReturn
  
  Exit Function
  
ErrorHandler:
 ExtractRows = "Error: " & Err.Number & ": " & Err.Description
End Function

Public Function cleanComma(sLine As String, Optional iBegin As Long = 1)
  Dim iStart As Long
  Dim iEnd As Long
  Dim iLen As Long
  Dim sTemp As String
  
  If InStr(iBegin, sLine, """") <= 0 Then
    cleanComma = sLine
    Exit Function
  Else
    While InStr(iBegin, sLine, """") > 0
        iStart = InStr(iBegin, sLine, """")
        iEnd = InStr(iStart + 1, sLine, """")
        sTemp = Replace(Mid(sLine, iStart, iEnd - iStart + 1), ",", "")
        
        sLine = Mid(sLine, 1, iStart - 1) & sTemp & Mid(sLine, iEnd + 1)
        iBegin = iStart + Len(sTemp) + 1
    Wend
    
  End If
  
  cleanComma = sLine

End Function
