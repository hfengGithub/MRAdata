

Private Sub Load_Data_Click()
On Error GoTo Err_Load_Data_Click
    With Form_LoadData
    'MsgBox ListProc.Value & "-" & tEndDate.Value
    Call mainProc.GetData(ListProc.Value, tStartDate.Value, tEndDate.Value)
    

    'DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
    End With

Exit_Load_Data_Click:
    Exit Sub

Err_Load_Data_Click:
    MsgBox Err.Description
    Resume Exit_Load_Data_Click
    
End Sub


Option Compare Database

'---------------------- The Main driver

Public Sub GetDataOrig()
    Dim iStartYr, iStartMth, iStartDate As Integer
    Dim dAsOf As Date
    Dim sProc As String
    
    'sProc = "Attribution"  -- for previous months
    'sProc = "AttrCur"   ---- for current month
    'sProc = "OAS"
    'sProc = "CalcOAS"
    'sProc = "SA"
    'sProc = "RawData"
    'sProc = "MrdExt"
    'sProc = "Polybal"
    
    sProc = "IDCPx"
    iStartYr = 2009
    iStartMth = 12
    iStartDate = 10
    If sProc = "Attribution" Then
        Call transferAttribution(iStartYr, iStartMth, iStartDate)
    ElseIf sProc = "IDCPx" Then
        Call transferIDCPx("12/10/2009", "1/10/2010")
    ElseIf sProc = "AttrCur" Then
        Call transAttrCur(iStartYr, iStartMth, iStartDate)
    ElseIf sProc = "MrdExt" Then
        Call transferMrdExtOrig(iStartYr, iStartMth, iStartDate)
    ElseIf sProc = "Polybal" Then
        Call transPolybal(iStartYr, iStartMth, iStartDate)
    ElseIf sProc = "summitMP" Then
        Call transSummit
    End If
    
End Sub


Public Sub GetData(ByVal sProc As String, ByVal dStartDate As Date, ByVal dEndDate As Date)
    Dim iStartYr, iStartMth, iStartDate As Integer
    Dim dAsOf As Date
    
    If sProc = "IDCPx" Then
        Call transferIDCPx(dStartDate, dEndDate)
    ElseIf sProc = "MasterRaw" Or sProc = "SA" Then
        Call transferRpt(dStartDate, dEndDate, sProc)
    ElseIf sProc = "MrdExt" Then
        Call transferMrdExt(dStartDate, dEndDate)
    ElseIf sProc = "transfer1" Then
        Call transfer1
    ElseIf sProc = "Polybal" Then
        Call transPolybal(iStartYr, iStartMth, iStartDate)
    ElseIf sProc = "summitMP" Then
        Call transSummit
    End If
    
End Sub



Public Sub transferRawData(ByVal dStartDate As Date, ByVal dEndDate As Date)
On Error GoTo ErrProc
    Const iYears = 5
    Const iMonths = iYears * 12
    Dim i, j, iYr, iMth, jDay As Integer
    Dim sDir, aMthDirs(iMonths) As String
    Dim dFileDate As Date
    
    Dim sCurFile, sRootDir, sDayDir, sFile, sName As String
    sRootDir = "K:\MRA\Polypaths\DailyReports\"
    
    'K:\MRA\Polypaths\DailyReports\2008-03\20080331
    'sFile = "PolyBal-2008-03-31.xls" or MasterRawData-2008-03-11.csv
    
    i = 0
    sDir = Dir(sRootDir & "20*", vbDirectory)
    Debug.Print aMthDirs(i)
    
    Do While sDir <> ""
        'MsgBox (aMthDirs(i) & i)
        
        If IsNumeric(Left(sDir, 4)) And IsNumeric(Right(sDir, 2)) Then
            iYr = CInt(Left(sDir, 4))
            If iYr > Year(dStartDate) Or (iYr = Year(dStartDate) And CInt(Right(sDir, 2)) >= Month(dStartDate)) Then
                aMthDirs(i) = sDir
                i = i + 1
            End If
        End If
        sDir = Dir
        
    Loop
    iMth = i
        
    'MsgBox (Left(aMthDirs(iMonth - 1), 4))
    Debug.Print iMth
    
    '======= get daily dir and file
    i = 0
    Do While i < iMth
        sDayDir = Dir(sRootDir & aMthDirs(i) & "\20*", vbDirectory)
        Do While sDayDir <> ""
           'sFile = "PolyBal-" & aMthDirs(i) & "-" & Right(sDayDir, 2) & ".xls"
           sFile = "MasterRawdata-" & aMthDirs(i) & "-" & Right(sDayDir, 2) & ".csv"
           sCurFile = sRootDir & aMthDirs(i) & "\" & sDayDir & "\" & sFile
           dFileDate = DateValue(Mid(sDayDir, 5, 2) & "/" & Right(sDayDir, 2) & "/" & Left(sDayDir, 4))
           If dFileDate >= dStartDate And dFileDate <= dEndDate Then
              On Error Resume Next
              Call dropTable("MasterRawData")
              On Error GoTo ErrProc
              '---- Call ClearTable("dbo_tmpMRD", " where pricingDate=" & sDayDir)
              DoCmd.TransferText acLinkDelim, "MasterRawDataImportSpec", "MasterRawData", sCurFile, -1
              Call AppendRawData(sDayDir)
           End If
           sDayDir = Dir
           
           'MsgBox (sCurFile)
        Loop
        i = i + 1
        'If i > 1 Then
        '    GoTo ExitLoop
        'End If
    Loop

    MsgBox "transferRawData: Done!"
    Exit Sub

ErrProc:
    MsgBox "ERROR " & Err.Number & " - " & Err.Description, vbCritical, "transferRawData"
End Sub


Public Sub transferMrdExt(ByVal dStartDate As Date, ByVal dEndDate As Date)
On Error GoTo ErrProc
    Dim sDate, aDateParts() As String
    Dim dFileDate As Date
    
    Dim sCurFile, sRootDir, sFile, sName As String
    ' sRootDir = "K:\MRA\Polypaths\DailyInputs\PolyTalkOutput\w2k3applyp01_output\ReportsArchive\2010-04\"
    sRootDir = "\\w2k3applyp01\output\Reports\"
    ' sRootDir = "d:\temp\"
    ' Daily-MasterRawData-Addendum-2009-07-28.csv
    
    sName = "Daily-MasterRawData-Addendum-"
    
    '======= get daily dir and file
        sFile = Dir(sRootDir & sName & "????-??-??.csv")
        'MsgBox (sFile)
        Do While sFile <> ""
           sDate = Left(Right(sFile, 14), 10)
           aDateParts = Split(sDate, "-")
           If IsNumeric(aDateParts(0)) And IsNumeric(aDateParts(1)) And IsNumeric(aDateParts(2)) And aDateParts(0) * aDateParts(1) * aDateParts(2) > 0 Then
              dFileDate = DateValue(aDateParts(1) & "/" & aDateParts(2) & "/" & aDateParts(0))
              If dFileDate >= dStartDate And dFileDate <= dEndDate Then
                 On Error Resume Next
                 Call dropTable("MrdExt")
                 On Error GoTo ErrProc
                 sCurFile = sRootDir & sFile
                 DoCmd.TransferText acLinkDelim, "specMrdExt", "MrdExt", sCurFile, -1
                 Call AppendMrdExt(aDateParts(0) & aDateParts(1) & aDateParts(2))
              End If
           End If
           
           sFile = Dir
        Loop
    MsgBox "transferMrdExt: Finished " & sCurFile
    Exit Sub
ErrProc:
    MsgBox "ERROR " & Err.Number & " - " & Err.Description, vbCritical, "transferMrdExt"
End Sub


Public Sub transferMrdExtOrig(iStartYr, iStartMth, Optional iStartDate As Integer = 1)
On Error GoTo ErrProc
    Const iYears = 5
    Const iMonths = iYears * 12
    Dim i, j, iYr, iMth, jDay As Integer
    Dim sDate, aDateParts() As String
    Dim dStartDate, dFileDate, dEndDate
    
    Dim sCurFile, sRootDir, sDayDir, sFile, sName As String
    sRootDir = "\\w2k3applyp01\output\Reports\"
    ' sRootDir = "d:\temp\"
    dStartDate = DateValue(iStartMth & "/" & iStartDate & "/" & iStartYr)
    dEndDate = DateValue("9/9/2009")
    ' Daily-MasterRawData-Addendum-2009-07-28.csv
    
    sName = "Daily-MasterRawData-Addendum-"
    
    'MsgBox (Left(aMthDirs(iMonth - 1), 4))
    
    '======= get daily dir and file
        sFile = Dir(sRootDir & sName & "????-??-??.csv")
        'MsgBox (sFile)
        Do While sFile <> ""
           sDate = Left(Right(sFile, 14), 10)
           aDateParts = Split(sDate, "-")
           If IsNumeric(aDateParts(0)) And IsNumeric(aDateParts(1)) And IsNumeric(aDateParts(2)) And aDateParts(0) * aDateParts(1) * aDateParts(2) > 0 Then
              dFileDate = DateValue(aDateParts(1) & "/" & aDateParts(2) & "/" & aDateParts(0))
              If dFileDate >= dStartDate And dFileDate <= dEndDate Then
                 On Error Resume Next
                 Call dropTable("MrdExt")
                 On Error GoTo ErrProc
                 sCurFile = sRootDir & sFile
                 DoCmd.TransferText acLinkDelim, "specMrdExt", "MrdExt", sCurFile, -1
                 Call AppendMrdExt(aDateParts(0) & aDateParts(1) & aDateParts(2))
              End If
           End If
           
           sFile = Dir
        Loop

ErrProc:
    If Err.Number <> 0 Then
        MsgBox "ERROR " & Err.Number & " - " & Err.Description, vbCritical, "transferRawData"
    Else
        MsgBox "transferRawData: Done!"
    End If
    
    'Resume ExitProc
    
End Sub

Public Sub transferIDCPx(ByVal dStartDate As Date, ByVal dEndDate As Date)
On Error GoTo ErrProc
    Const iYears = 5
    Const iMonths = iYears * 12
    Dim i, iYrIndex, iYr, iMth, iStartYr, iEndYr  As Integer
    Dim aMthDirs(iMonths), aDateParts() As String
    Dim dFileDate As Date
    
    Dim sCurFile, sMRO, sRootDir, sDir, sFile, sStartYr, sEndYr As String
    ' K:\MRO\Archive\2010\2010-09\2010-09-29\IDC-Pricing\PolyPaths_Inputs_IDC_Prices.csv
    ' sRootDir = "K:\MRA\Polypaths\DailyInputs\PolyTalkOutput\"
    sMRO = "K:\MRO\Archive\"
    sFile = "\IDC-Pricing\PolyPaths_Inputs_IDC_Prices.csv"
    iStartYr = Year(dStartDate)
    iEndYr = Year(dEndDate)
    iYrIndex = iStartYr
    
    Do While iYrIndex <= iEndYr
    
        sRootDir = sMRO & iYrIndex & "\"
        ' sFile = "\1 Input CSV files\PolyPaths_Inputs_IDC_Prices.csv"
        
        ' old K:\MRA\Polypaths\DailyInputs\PolyTalkOutput\2009-09\2009-09-18\Enterprise Input\PolyPaths_Inputs_IDC_Prices.csv
        i = 0
        sDir = Dir(sRootDir & "20*", vbDirectory)
        Debug.Print sDir
        
        Do While sDir <> ""
            If Len(sDir) = 7 And IsNumeric(Left(sDir, 4)) And IsNumeric(Right(sDir, 2)) Then
                iYr = CInt(Left(sDir, 4))
                iMth = CInt(Right(sDir, 2))
                If ((100 * iYr + iMth >= 100 * Year(dStartDate) + Month(dStartDate)) And _
                    (100 * iYr + iMth <= 100 * Year(dEndDate) + Month(dEndDate))) Then
                    aMthDirs(i) = sDir
                    i = i + 1
                End If
            End If
            sDir = Dir
        Loop
        iMth = i
            
        Debug.Print iMth
        
        '======= get daily dir and file
        i = 0
        Do While i < iMth
            sDir = Dir(sRootDir & aMthDirs(i) & "\20*", vbDirectory)
            Do While sDir <> ""

               aDateParts = Split(sDir, "-")
               If IsNumeric(aDateParts(0)) And IsNumeric(aDateParts(1)) And IsNumeric(aDateParts(2)) And aDateParts(0) * aDateParts(1) * aDateParts(2) > 0 Then
                  dFileDate = DateValue(aDateParts(1) & "/" & aDateParts(2) & "/" & aDateParts(0))
                  If dFileDate >= dStartDate And dFileDate <= dEndDate Then
                     On Error Resume Next
'                     MsgBox (sRootDir & aMthDirs(i) & "\" & sDir & sFile)
                     Call dropTable("IDCPrice")
                     On Error GoTo ErrProc
                     sCurFile = sRootDir & aMthDirs(i) & "\" & sDir & sFile
                     DoCmd.TransferText acLinkDelim, , "IDCPrice", sCurFile, -1
                     Call AppendIDCPx(dFileDate)
                  End If
               End If
               sDir = Dir
                         
               'MsgBox (sCurFile)
            Loop
            i = i + 1
        Loop
        
        iYrIndex = iYrIndex + 1
    Loop

    MsgBox "transferIDCPx: Finished " & sCurFile
ExitLoop:
    Exit Sub
ErrProc:
    MsgBox "ERROR-transferIDCPx: " & Err.Number & " - " & Err.Description, vbCritical, "transferIDCPx"
    'Resume ExitProc

ExitProc:
    
End Sub
Public Sub transfer1()
    Dim sCurFile, sRootDir, sDayDir, sFile, sName As String
    sFile = "K:\MRA\Polypaths\DailyReports\2008-03\20080331\RSR_HAUS Price 2008-03-31.xls"
    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel9, "SA", sFile, 1, "S-A!A1:R1000"
    MsgBox "Done"
End Sub




Public Sub transferRpt(ByVal dStartDate As Date, ByVal dEndDate As Date, ByVal sProc As String)
On Error GoTo ErrProc
    Const iMonths = 60
    Dim i, j, iYr, iMth As Integer
    Dim sTable, sDir, aMthDirs(iMonths) As String
    Dim dFileDate As Date
    Dim sCurFile, sRootDir, sDayDir, sFile  As String
    sRootDir = "K:\MRA\Polypaths\DailyReports\"
    
    'K:\MRA\Polypaths\DailyReports\2008-03\20080331
    'sFile = "PolyBal-2008-03-31.xls" or MasterRawData-2008-03-11.csv or RSR_HAUS Price 2009-09-23.xls
    
    i = 0
    sDir = Dir(sRootDir & "20*", vbDirectory)
    Debug.Print aMthDirs(i)
    
    Do While sDir <> ""
        If IsNumeric(Left(sDir, 4)) And IsNumeric(Right(sDir, 2)) Then
            iYr = CInt(Left(sDir, 4))
            If iYr > Year(dStartDate) Or (iYr = Year(dStartDate) And CInt(Right(sDir, 2)) >= Month(dStartDate)) Then
                aMthDirs(i) = sDir
                i = i + 1
            End If
        End If
        sDir = Dir
        
    Loop
    iMth = i
            
    '======= get daily dir and file
    i = 0
    Do While i < iMth
        sDayDir = Dir(sRootDir & aMthDirs(i) & "\20*", vbDirectory)
        Do While sDayDir <> ""
           dFileDate = DateValue(Mid(sDayDir, 5, 2) & "/" & Right(sDayDir, 2) & "/" & Left(sDayDir, 4))
           If dFileDate >= dStartDate And dFileDate <= dEndDate Then
              'sFile = "PolyBal-" & aMthDirs(i) & "-" & Right(sDayDir, 2) & ".xls"
              If sProc = "SA" Then
                 sFile = "RSR_HAUS Price " & aMthDirs(i) & "-" & Right(sDayDir, 2) & ".xls"
                 sTable = "linkSA"
              ElseIf sProc = "MasterRaw" Then
                 sFile = "MasterRawdata-" & aMthDirs(i) & "-" & Right(sDayDir, 2) & ".csv"
                 sTable = "MasterRawData"
              End If
              
              sCurFile = sRootDir & aMthDirs(i) & "\" & sDayDir & "\" & sFile
              On Error Resume Next
              Call dropTable(sTable)
              On Error GoTo ErrProc
              '---- Call ClearTable("dbo_tmpMRD", " where pricingDate=" & sDayDir)
              If sProc = "SA" Then
                 DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel9, sTable, sCurFile, 1, "S-A!A1:R1000"
                 Call AppendSA(sDayDir)
              ElseIf sProc = "MasterRaw" Then
                 DoCmd.TransferText acLinkDelim, "MasterRawDataImportSpec", sTable, sCurFile, -1
                 Call AppendRawData(sDayDir)
              End If
           
           End If
           sDayDir = Dir
           
           'MsgBox (sCurFile)
        Loop
        i = i + 1
        'If i > 1 Then
        '    GoTo ExitLoop
        'End If
    Loop

    MsgBox "Finished transfer: " & sCurFile
    Exit Sub

ErrProc:
    MsgBox "ERROR " & Err.Number & " - " & Err.Description, vbCritical, "transferRpt"
End Sub




Public Sub transferAttribution(iStartYr, iStartMth, Optional iStartDate As Integer = 1)
On Error GoTo ErrProc
    Const iYears = 10
    Const iMonths = iYears * 12
    Dim i, j, iYr, iMth, jDay As Integer
    Dim sDir, aMthDirs(iMonths), aDateParts() As String
    Dim dStartDate, dFileDate
    
    Dim sCurFile, sRootDir, sDayDir, sFile, sName As String
    'sRootDir = "K:\MRA\Polypaths\DailyInputs\PolyTalkOutput\w2k3applyp01_output\DailyArchive\"
    sRootDir = "Q:\Daily\"
    dStartDate = DateValue(iStartMth & "/" & iStartDate & "/" & iStartYr)
    
    '\\W2k3applyp01\Output\Daily\2008-10-10\AttributionReport-2008-10-10.csv (current month)
    'Q:\Daily\2008-09\2008-09-30\AttributionReport-2008-09-30.csv  in Daily
    ' archived: K:\MRA\Polypaths\DailyInputs\PolyTalkOutput\w2k3applyp01_output\DailyArchive\2007-12\2007-12-10\AttributionReport-2007-12-10.csv
    
    i = 0
    sDir = Dir(sRootDir & "20*", vbDirectory)
    aMthDirs(i) = ""
    
    Do While sDir <> ""
        'MsgBox (aMthDirs(i) & i)
        If IsNumeric(Left(sDir, 4)) And IsNumeric(Right(sDir, 2)) Then
            iYr = CInt(Left(sDir, 4))
            If iYr > iStartYr Or (iYr = iStartYr And CInt(Right(sDir, 2)) >= iStartMth) Then
                aMthDirs(i) = sDir
                i = i + 1
            End If
        End If
        sDir = Dir
    Loop
    iMth = i
    
    'MsgBox (Left(aMthDirs(iMonth - 1), 4))
    Debug.Print iMth
    
    '======= get daily dir and file
    i = 0
    Do While i < iMth
        sDayDir = Dir(sRootDir & aMthDirs(i) & "\20*", vbDirectory)
        Do While sDayDir <> ""
           'sFile = "AttributionReport-" & sYr & "-" & sMth & "-" & sDay & ".csv"
           aDateParts = Split(sDayDir, "-")
           If IsNumeric(aDateParts(0)) And IsNumeric(aDateParts(1)) And IsNumeric(aDateParts(2)) Then
                sFile = "AttributionReport-" & sDayDir & ".csv"
                sCurFile = sRootDir & aMthDirs(i) & "\" & sDayDir & "\" & sFile
                dFileDate = DateValue(aDateParts(1) & "/" & aDateParts(2) & "/" & aDateParts(0))
                If dFileDate >= dStartDate Then
                    On Error Resume Next
                    Call dropTable("AttributionReport")
                    On Error GoTo ErrProc
                    DoCmd.TransferText acLinkDelim, "attributionReportSpec", "AttributionReport", sCurFile, -1
                    Call AppendAttribution(aDateParts(0) & aDateParts(1) & aDateParts(2))
                End If
           End If
           sDayDir = Dir
           
           'MsgBox (sCurFile)
        Loop
        i = i + 1
    Loop
    
    MsgBox ("End of proc, i=" & i)
    Exit Sub
    
ExitLoop:
ErrProc:
    MsgBox "ERROR- transferAttribution: " & Err.Number & " - " & Err.Description, vbCritical, "transferAttribution"
    'Resume ExitProc
    
End Sub


Public Sub transAttrCur(iStartYr, iStartMth, Optional iStartDate As Integer = 1)
On Error GoTo ErrProc
    Const iYears = 10
    Const iMonths = iYears * 12
    Dim i, j, iYr, iMth, jDay As Integer
    Dim sDir, aMthDirs(iMonths), aDateParts() As String
    Dim dStartDate, dFileDate
    
    Dim sCurFile, sRootDir, sDayDir, sFile, sName As String
    'sRootDir = "K:\MRA\Polypaths\DailyInputs\PolyTalkOutput\w2k3applyp01_output\DailyArchive\"
    sRootDir = "\\W2k3applyp01\Output\Daily\"
    dStartDate = DateValue(iStartMth & "/" & iStartDate & "/" & iStartYr)
    
    '\\W2k3applyp01\Output\Daily\2008-10-10\AttributionReport-2008-10-10.csv (current month)
    'Q:\Daily\2008-09\2008-09-30\AttributionReport-2008-09-30.csv  in Daily
    ' archived: K:\MRA\Polypaths\DailyInputs\PolyTalkOutput\w2k3applyp01_output\DailyArchive\2007-12\2007-12-10\AttributionReport-2007-12-10.csv
    
    i = 0
        sDayDir = Dir(sRootDir & iStartYr & "-" & Right("0" & iStartMth, 2) & "-*", vbDirectory)
        Do While sDayDir <> ""
           'sFile = "AttributionReport-" & sYr & "-" & sMth & "-" & sDay & ".csv"
           aDateParts = Split(sDayDir, "-")
           If IsNumeric(aDateParts(0)) And IsNumeric(aDateParts(1)) And IsNumeric(aDateParts(2)) Then
                sFile = "AttributionReport-" & sDayDir & ".csv"
                sCurFile = sRootDir & sDayDir & "\" & sFile
                dFileDate = DateValue(aDateParts(1) & "/" & aDateParts(2) & "/" & aDateParts(0))
                If dFileDate >= dStartDate Then
                    'MsgBox "File: " & sCurFile
                    On Error Resume Next
                    Call dropTable("AttributionReport")
                    On Error GoTo ErrProc
                    DoCmd.TransferText acLinkDelim, "attributionReportSpec", "AttributionReport", sCurFile, -1
                    Call AppendAttribution(aDateParts(0) & aDateParts(1) & aDateParts(2))
                End If
           End If
           sDayDir = Dir
           i = i + 1
        Loop
    
    MsgBox ("End of proc, i=" & i)
    Exit Sub
    
ExitLoop:
ErrProc:
    MsgBox "ERROR- transAttrCur: " & Err.Number & " - " & Err.Description, vbCritical, "transAttrCur"
    'Resume ExitProc
    
End Sub


Public Sub transPolybal(iStartYr, iStartMth, Optional iStartDate As Integer = 1)
On Error GoTo ErrProc
    Const iYears = 10
    Const iMonths = iYears * 12
    Dim i, j, iYr, iMth, jDay As Integer
    Dim sDir, aMthDirs(iMonths), aDateParts() As String
    Dim dStartDate, dFileDate
    
    Dim sCurFile, sRootDir, sDayDir, sDate, sName As String
    'sRootDir = "K:\MRA\Polypaths\DailyInputs\PolyTalkOutput\w2k3applyp01_output\DailyArchive\"
    sRootDir = "\\W2k3applyp01\Output\reports\"
    dStartDate = DateValue(iStartMth & "/" & iStartDate & "/" & iStartYr)
    sName = "Daily_rpt_polybal_"
    
    ' sRootDir = "\\W2k3applyp01\Output\reports\Daily_rpt_polybal_20080711.csv"
    
    i = 0
        sDayDir = Dir(sRootDir & sName & "????????.csv")
        Do While sDayDir <> ""
           sDate = Left(Right(sDayDir, 12), 8)
           If IsNumeric(sDate) Then
                dFileDate = DateValue(Mid(sDate, 5, 2) & "/" & Right(sDate, 2) & "/" & Left(sDate, 4))
                sCurFile = sRootDir & sDayDir
                If dFileDate >= dStartDate Then
                    'MsgBox "File: " & sCurFile
                    On Error Resume Next
                    Call dropTable("polybal")
                    On Error GoTo ErrProc
                    DoCmd.TransferText acLinkDelim, "polybalSpec", "polybal", sCurFile, -1
                    Call AppendPolybal
                End If
           End If
           
           sDayDir = Dir
           i = i + 1
        Loop
    
    MsgBox ("End of proc, i=" & i)
    Exit Sub
    
ExitLoop:
ErrProc:
    MsgBox "ERROR- transPolybal: " & Err.Number & " - " & Err.Description, vbCritical, "transPolybal"
    'Resume ExitProc
    
End Sub



Public Sub transSummit()
On Error GoTo ErrProc
    Const iYears = 10
    Const iMonths = iYears * 12
    Dim i, j, iYr, iMth, jDay As Integer
    Dim sDir, aMthDirs(iMonths), aDateParts() As String
    Dim dStartDate, dFileDate
    
    Dim sCurFile, sRootDir, sDayDir, sDate, sName As String
    'sRootDir = "K:\MRA\Polypaths\DailyInputs\PolyTalkOutput\w2k3applyp01_output\DailyArchive\"
    sRootDir = "S:\share\Summit\MRA\SUMMIT +_- 25BP\"
    sName = "SPL_"
    
    ' sRootDir = "\\W2k3applyp01\Output\reports\Daily_rpt_polybal_20080711.csv"
    
    i = 0
        sDayDir = Dir(sRootDir & sName & "*.txt")
        Do While sDayDir <> ""
           sDate = Left(Right(sDayDir, 12), 4)
           If IsNumeric(sDate) Then
                If CInt(sDate) > 1000 Then
                    dFileDate = DateValue(Left(sDate, 2) & "/" & Right(sDate, 2) & "/2008")
                Else
                    dFileDate = DateValue(Left(sDate, 2) & "/" & Right(sDate, 2) & "/2009")
                End If
                
                sCurFile = sRootDir & sDayDir
                On Error Resume Next
                Call dropTable("summitMP25")
                On Error GoTo ErrProc
                DoCmd.TransferText acLinkFixed, "summitMP25Spec", "summitMP25", sCurFile, -1
                Call AppendSummitMP(dFileDate)
           End If
           
           sDayDir = Dir
           i = i + 1
        Loop
    
    MsgBox ("End of proc, i=" & i)
    Exit Sub
    
ExitLoop:
ErrProc:
    MsgBox "ERROR- transPolybal: " & Err.Number & " - " & Err.Description, vbCritical, "transPolybal"
    'Resume ExitProc
    
End Sub

