'======== modMDB

Dim oCnn As Connection

Sub RunDR()
'
' RunDR Macro
' Macro recorded 7/14/2008 by wyao
'
' Keyboard Shortcut: Ctrl+Shift+M
'
  frmGetData.Show
End Sub
Sub RunPort()
'
' RunPort Macro
' Macro recorded 7/14/2008 by wyao
'
' Keyboard Shortcut: Ctrl+Shift+R
'
  frmRun.Show
  
End Sub

Public Function MDBX(sDate As Date, AccountName As String, sFldLst As String) As Variant

    Dim oRs As ADODB.Recordset
    
    Dim i As Integer
    Dim j As Integer
    
    Dim sVal() As Variant
    Dim iYear As Integer
    Dim iQuater As Integer
    Dim sTableName As String
    Dim sCrit As String
    Dim sFlds() As String
    Dim sFields As String
    
    On Error GoTo ErrorHandler
    
    If Trim(sFldLst) = "" Or (Not IsDate(sDate)) Then
       ReDim sVal(1 To 1)
      sVal(1) = "date and fieldname are both required!"
      MDBX = sVal
      Exit Function
    End If
    
  '  If oCnn Is Nothing Then
      Set oCnn = New ADODB.Connection
      oCnn.ConnectionString = "Provider=SQLOLEDB;Data Source=FHCDBGENMP04\MRAP1;Initial Catalog=mraTest;Trusted_Connection=yes;"
      oCnn.Open    'Create your recordset
   ' End If
    
    sTableName = "dbo.MasterRawData"
    
    Set oRs = New ADODB.Recordset
    If InStr(AccountName, "%") > 0 Then
      sCrit = " where accountname like '" & AccountName & "' and [PricingDate] = '" & Format(sDate, "YYYYMMDD") & "'"
    Else
      sCrit = " where accountname = '" & AccountName & "' and [PricingDate] = '" & Format(sDate, "YYYYMMDD") & "'"
    End If


    sFlds = Split(sFldLst, ",")
    
    For i = 0 To UBound(sFlds)
        If InStr(sFlds(i), "[") <= 0 Then
          sFields = IIf(sFields = "", "[" & Trim(sFlds(i)) & "]", sFields & "," & "[" & Trim(sFlds(i)) & "]")
        End If
    Next
    
    oRs.Open "SELECT " & sFields & " FROM " & sTableName & sCrit, oCnn, adOpenKeyset, adLockReadOnly, adCmdText

    Dim iRow As Integer
    
    'Add to your current workbook and add the field names as column headers (optional)
    If Not oRs.EOF Then
        iRow = 0
        While Not oRs.EOF
          iRow = iRow + 1
          oRs.MoveNext
        Wend
        
        oRs.MoveFirst
        ReDim sVal(1 To iRow + 1, 1 To oRs.Fields.Count)
        For i = 0 To oRs.Fields.Count - 1
           sVal(1, i + 1) = oRs.Fields(i).Name
        Next
        i = 2
        While Not oRs.EOF
          For j = 1 To oRs.Fields.Count
            sVal(i, j) = oRs.Fields(j - 1).Value
          Next
          oRs.MoveNext
          i = i + 1
        Wend
    End If
    
    
    'Clean up ADO Objects
    oRs.Close
    Set oRs = Nothing
    oCnn.Close
    Set oCnn = Nothing
    
    MDBX = sVal
    Exit Function
    
ErrorHandler:

    MDBX = "Error Occurred. " & Err.Description
End Function


Public Function MDB(sDate As Date, cusip As String, sFld As String, Optional custid As String = "") As Variant

    Dim oRs As ADODB.Recordset
    Dim oCnn As ADODB.Connection
    Dim i As Integer
    Dim j As Integer
    
    Dim sVal() As Variant
    Dim iYear As Integer
    Dim iQuater As Integer
    Dim sTableName As String
    Dim sCrit As String
    Dim sFlds() As String
    Dim sFields As String
    
    On Error GoTo ErrorHandler
    
    If Trim(cusip) = "" Or Trim(sFld) = "" Or (Not IsDate(sDate)) Then
     ' sVal = "cusip, date and fieldname are both required!"
      MDB = sVal
      Exit Function
    End If
    
  '  If oCnn Is Nothing Then
      Set oCnn = New ADODB.Connection
      oCnn.ConnectionString = "Provider=SQLOLEDB;Data Source=FHCDBGENMP04\MRAP1;Initial Catalog=mraTest;Trusted_Connection=yes;"
      oCnn.Open    'Create your recordset
   ' End If
        
    sTableName = "dbo.MasterRawData"
    
    Set oRs = New ADODB.Recordset
    
    If sCustID = "" Then
          sCrit = " where cusip='" & cusip & "' and [PricingDate] = '" & Format(sDate, "YYYYMMDD") & "'"
    Else
          sCrit = " where cusip='" & cusip & "' and [Custid] = '" & custid & "' and [PricingDate] = '" & sDate & "'"
    End If
    

    sFlds = Split(sFld, ",")

    
    For i = 0 To UBound(sFlds)
        If InStr(sFlds(i), "[") <= 0 Then
          sFields = IIf(sFields = "", "[" & Trim(sFlds(i)) & "]", sFields & "," & "[" & Trim(sFlds(i)) & "]")
        End If
    Next
    
    oRs.Open "SELECT " & sFields & " FROM " & sTableName & sCrit, oCnn, adOpenKeyset, adLockReadOnly, adCmdText

    Dim iRow As Integer
    
    If Not oRs.EOF Then

        iRow = 0
        While Not oRs.EOF
          iRow = iRow + 1
          oRs.MoveNext
        Wend
        
        oRs.MoveFirst
        ReDim sVal(1 To iRow, 1 To oRs.Fields.Count)

        i = 1
        While Not oRs.EOF
          For j = 1 To oRs.Fields.Count
            sVal(i, j) = oRs.Fields(j - 1).Value
          Next
          oRs.MoveNext
          i = i + 1
        Wend
        
    End If
    
    
    'Clean up ADO Objects
    oRs.Close
    Set oRs = Nothing
    oCnn.Close
    Set oCnn = Nothing

    MDB = sVal
    Exit Function
    
ErrorHandler:

    MDB = "Error Occurred. " & Err.Description
    
End Function

Public Function OneDayAgo(sDate As Date) As Date
  Dim bdates As Variant
  
  Dim sTemp As String
  Dim i As Long
  bdates = BusinessCal
  ReDim Preserve bdates(2000) As String
  sTemp = Format(DateAdd("d", -1, sDate), "YYYYMMDD")
  i = 1
  While bdates(i) <= sTemp
    i = i + 1
  Wend
  
  sTemp = bdates(i - 1)
  
  OneDayAgo = Mid(sTemp, 5, 2) & "/" & Mid(sTemp, 7, 2) & "/" & Mid(sTemp, 1, 4)

End Function


Public Function OneWeekAgo(sDate As Date) As Date
  Dim bdates As Variant
  
  Dim sTemp As String
  Dim i As Long
  bdates = BusinessCal
  ReDim Preserve bdates(2000) As String
  sTemp = Format(DateAdd("d", -7, sDate), "YYYYMMDD")
  i = 1
  While bdates(i) <= sTemp
    i = i + 1
  Wend
  
  sTemp = bdates(i - 1)
  
  OneWeekAgo = Mid(sTemp, 5, 2) & "/" & Mid(sTemp, 7, 2) & "/" & Mid(sTemp, 1, 4)

End Function
Public Function OneMonthAgo(sDate As Date) As Date
  Dim bdates As Variant
  
  Dim sTemp As String
  Dim i As Long
  bdates = BusinessCal
  ReDim Preserve bdates(2000) As String
  sTemp = Format(DateAdd("M", -1, sDate), "YYYYMMDD")
  
  i = 1
  While bdates(i) <= sTemp
    i = i + 1
  Wend
  
  sTemp = bdates(i - 1)
  
  OneMonthAgo = Mid(sTemp, 5, 2) & "/" & Mid(sTemp, 7, 2) & "/" & Mid(sTemp, 1, 4)

End Function

Public Function OneYearAgo(sDate As Date) As Date
  Dim bdates As Variant
  
  Dim sTemp As String
  Dim i As Long
  bdates = BusinessCal
  ReDim Preserve bdates(2000) As String
 
  sTemp = Format(DateAdd("yyyy", -1, sDate), "YYYYMMDD")
  i = 1
  While bdates(i) <= sTemp
    i = i + 1
  Wend
  
  sTemp = bdates(i - 1)
  
  OneYearAgo = Mid(sTemp, 5, 2) & "/" & Mid(sTemp, 7, 2) & "/" & Mid(sTemp, 1, 4)

End Function

Private Function BusinessCal() As Variant
   Dim bdates(2000) As String
   Dim sTemp As String
   Dim i As String
   
   i = 2
   sTemp = Worksheets("BusinessDate").Cells(i, 1).Value
   
   While sTemp <> ""
     bdates(i - 1) = sTemp
     i = i + 1
     sTemp = Worksheets("BusinessDate").Cells(i, 1).Value
   Wend
   
   BusinessCal = bdates
   
End Function

