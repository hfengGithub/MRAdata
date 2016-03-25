'=== frmRun


Private Sub cmdclearlog_Click()
  lstStatus.Clear
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdOutputDir_Click()
 Dim sFile As String
 Dim sName As String
 
 cdlgFile.ShowSave
 sFile = getDir(cdlgFile.Filename)
 
 If cboStartDate.Text <> "" And cboEndDate.Text <> "" Then
  sName = Format(cboStartDate.Text, "YYYYMMDD") & "_" & Format(cboEndDate.Text, "YYYYMMDD")
 Else
  sName = Format(Now(), "YYYYMMDD")
 End If
 
 txtOutputDir.Text = sFile & "\output_" & sName & ".csv"
  
End Sub

Private Sub cmdRun_Click()

  Dim sFileName As String
  
  Dim sStartDate As Date
  
  Dim sEndDate As Date
  
  Dim sDate As Date
  
  Dim sMkt As String
  Dim lDay As Integer
  Dim sTempDir As String
  Dim iSkip As Integer
  Dim sOut As String
  Dim sDefPara As String
  
  On Error GoTo ErrorHandler
  
  If Trim(txtFileName.Text) = "" Then
      MsgBox "Please specify the input portfolio file!", vbCritical
      txtFileName.SetFocus
      Exit Sub
  End If
  
  If Not IsDate(cboStartDate.Text) Or (Not IsDate(cboEndDate.Text)) Then
     MsgBox "Start or End Date is not a valid date!", vbCritical
     Exit Sub
     
  End If
  
  If Not IsFileExist(txtFileName.Text) Then
    MsgBox "Input Portfolio file: " & txtFileName.Text & " doesn't exist!", vbCritical
    txtFileName.SetFocus
    Exit Sub
  End If
  
  Me.MousePointer = fmMousePointerHourGlass
  
  DoEvents
  
  sTempDir = GetAppSettings("Working Dir")
  
  sDefPara = " /ungroup  /ppconfig ""PP_SAVE_YC_IN_CSV=N"" /ppconfig ""PP_SAVE_SUMMARY_IN_CSV = N"" " & IIf(txtTemplate.Text = "", "", " /template """ & txtTemplate.Text & """ ")
  
  sStartDate = CDate(cboStartDate.Text)
  sEndDate = CDate(cboEndDate.Text)
  
  sOut = txtOutputDir.Text
  
  sDate = sStartDate
  ldate = 1
  While sDate <= sEndDate
     
     sMkt = GetAppSettings("MR Dir") & Format(sDate, "YYYY-MM") & "\" & Format(sDate, "YYYY-MM-DD") & "\Market_Rates_" & Format(sDate, "YYYY-MM-DD") & ".mr"
     
     If IsFileExist(sMkt) Then
       WriteLog "2", BatchCal(CStr(ldate) & Format(Now, "ss"), "Run Batch", txtFileName.Text, sTempDir & "\tmprun_" & ldate & ".csv", "/mr """ & sMkt & """ " & sDefPara & " " & txtOptParams.Text, "", "1", "1")
       If ldate = 1 Then
        iSkip = 0
       Else
        iSkip = 1
       End If
       
       If chkAppend.Value Then iSkip = 1
       
       If chkAppend.Value Then
         ExtractRows CStr(ldate), "merging file", sTempDir & "\tmprun_" & ldate & ".csv", sOut, iSkip, -1, "", ""
       ElseIf ldate = 1 Then
         ExtractRows CStr(ldate), "merging file", sTempDir & "\tmprun_" & ldate & ".csv", sOut, iSkip, -1, "", "1"
       Else
         ExtractRows CStr(ldate), "merging file", sTempDir & "\tmprun_" & ldate & ".csv", sOut, iSkip, -1, "", ""
       End If
       ldate = ldate + 1
     Else
       WriteLog "2", Format(sDate, "YYYY-MM-DD") & " mr file not exists!"
     End If
     
     sDate = DateAdd("d", 1, sDate)
     
  Wend
  
  Me.MousePointer = fmMousePointerDefault

  DoEvents
  
  MsgBox " Run is finished!", vbInformation
  
  Exit Sub
  
ErrorHandler:

  Me.MousePointer = fmMousePointerDefault
  MsgBox "Error Occurs: " & Err.Description, vbCritical

End Sub


Private Sub cmdSelectFile_Click()
 Dim sFile As String
 Dim sName As String

 cdlgFile.Filter = "PF Files (*.pf)|*.pf|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
 cdlgFile.ShowOpen
 sFile = cdlgFile.Filename
 
 txtFileName.Text = sFile
 
 If txtOutputDir.Text = "" Then
 
  If cboStartDate.Text <> "" And cboEndDate.Text <> "" Then
    sName = Format(cboStartDate.Text, "YYYYMMDD") & "_" & Format(cboEndDate.Text, "YYYYMMDD")
  Else
    sName = Format(Now(), "YYYYMMDD")
  End If
 
   txtOutputDir.Text = getDir(txtFileName.Text) & "output_" & sName & ".csv"
  
 End If
 

 
End Sub

Private Function getDir(sFile As String)

  getDir = Mid(sFile, 1, InStrRev(sFile, "\"))

End Function


Private Sub cmdSelectTemplate_Click()

 cdlgFile.Filter = "CSV Files (*.csv)|*.csv"
 cdlgFile.ShowOpen
 txtTemplate.Text = cdlgFile.Filename
 
End Sub

Private Sub cmdShowLog_Click()
If frmRun.Height >= 390.25 Then
  frmRun.Height = 215.25
  cmdShowLog.Caption = "Show Log >>"
Else
  frmRun.Height = 391.25
  cmdShowLog.Caption = "Hide Log <<"
End If

End Sub

Private Sub UserForm_Initialize()
  cboStartDate.Text = Format(DateAdd("d", -7, Now()), "MM/DD/YYYY")
  cboEndDate.Text = Format(Now, "MM/DD/YYYY")
  txtOptParams.Text = ""  ' /ungroup  /ppconfig ""PP_SAVE_YC_IN_CSV=N"" /ppconfig ""PP_SAVE_SUMMARY_IN_CSV = N"" "
  frmRun.Height = 215.25
  cmdShowLog.Caption = "Show Log >>"
  
  cmdSelectFile.SetFocus
  
  '/template "\\prodfs\rcgroupp\MRA\Polypaths\DailyInputs\PolyTalkOutput\derheaders.csv"

End Sub
