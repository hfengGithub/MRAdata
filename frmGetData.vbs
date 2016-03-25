'=========== frmGetData

Dim oCnn As ADODB.Connection
Dim bAllGroup As Boolean
Const gDBLocation As String = "K:\MRA\PolyPaths\MRADB\masterrawdata"
'Const gDBMasterTable As String = "UV_rawdata"
Const gDBConnectStr = "Provider=SQLOLEDB;Data Source=w2k3dbmrap1;Initial Catalog=mratest;Trusted_Connection=yes;"
Const gDBMasterTable As String = "MasterRawData"
'Const gDBConnectStr = "Provider=SQLOLEDB;Data Source=w2k3dbmrap1;Initial Catalog=mradb;Trusted_Connection=yes;"

Private Sub btnGetMarket_Click()

    Dim sStartDate As String
    Dim sEndDate As String
    Dim sSQL As String
    Dim sSection As String
    Dim sCode As String
    Dim i As Long
    Dim sprevDate As String
    Dim iDividor As Integer
    Dim sTitle As String
    Dim sSheetName As String
    Dim sCellName As String
    
    On Error GoTo ErrorHandler
    iDividor = 1
    
    If Trim(txtSelectiveDates_mkt.Text) = "" Then
        If cboStartDate_mkt.Text = "" And cboEndDate_mkt.Text = "" Then
          MsgBox "You must enter start and end date", vbCritical
          Exit Sub
        End If
        sStartDate = IIf(cboStartDate_mkt.Value <> "", CDate(cboStartDate_mkt.Value), Now())
        sEndDate = IIf(cboEndDate_mkt.Value <> "", CDate(cboEndDate_mkt.Value), Now())
    End If
    
    sCurvetype = Trim(cboCurveType.SelText)
    
    If sCurvetype = "" Or (sCurvetype = "CAP_VOLS" And Trim(txtTenor.Text) = "") Or (sCurvetype = "SWAPTION_VOLS" And Trim(txtTenor.Text) = "") Then
      MsgBox "Must Select Curve Type! and for Vol, Tenor is also needed!"
      Exit Sub
    End If
    
    setStatus "Begin Data Retrieval, please wait ....."
    
    Dim sSQLSection
    
    Select Case sCurvetype
      Case "SWAP_CURVE"
        sSQL = "select distinct b.pricing_date, a.term_in_month as terms,a.coupon as rate from mr_point a inner join market_rate b on a.mr_key = b.mr_key where section = '" & cboCurveType.SelText & "' "
      Case "UST_CURVE"
        sSQL = "select  distinct b.pricing_date, a.term_in_month as terms,a.coupon as rate from mr_point a inner join market_rate b on a.mr_key = b.mr_key where section = '" & cboCurveType.SelText & "' "
      Case "AGENCY"
        sSQL = "select  distinct  b.pricing_date, a.term_in_month as terms,a.coupon as rate from mr_point a inner join market_rate b on a.mr_key = b.mr_key where section = '" & cboCurveType.SelText & "' and Code='AGENCY_YIELD_CURVE' "
        iDividor = CInt(txtDividor.Text)
      Case "Index"
        sSQL = "select  distinct  b.pricing_date, Code as IndexName,a.coupon as rate from mr_point a inner join market_rate b on a.mr_key = b.mr_key where section = '" & cboCurveType.SelText & "' "
      Case "SWAPTION_VOLS"
        sSQL = "select  distinct  b.pricing_date, a.term_in_month as terms,a.coupon as rate from mr_point a inner join market_rate b on a.mr_key = b.mr_key where section = '" & cboCurveType.SelText & "' and cast(Code as numeric(5,2)) =" & Trim(txtTenor.Text) & " "
      Case "CAP_VOLS"
          sSQL = "select  distinct  b.pricing_date,a.term_in_month as terms,a.coupon as rate from mr_point a inner join market_rate b on a.mr_key = b.mr_key where section = '" & cboCurveType.SelText & "' and cast(Code as numeric(5,2)) =" & Trim(txtTenor.Text) & " "
    End Select
    
    If txtSelectiveDates_mkt.Text = "" Then
      sSQL = sSQL & " and pricing_date >='" & sStartDate & "' and pricing_date <='" & sEndDate & "' and b.mr_type ='E' order by pricing_date"
    Else
      sSQL = sSQL & " and pricing_date in ('" & Replace(Trim(txtSelectiveDates_mkt.Text), ",", "','") & "') and b.mr_type ='E' order by pricing_date"
    End If
    
    sSheetName = Trim(txtSheetName_mkt.Text)
    sCellName = Trim(txtCellName_mkt.Text)
    sTitle = sCurvetype & " " & IIf(Trim(txtTenor.Text) = "", "", "Tenor:" & txtTenor.Text)
    
    
    GetData_Market sSheetName, sCellName, sSQL, iDividor, sTitle, chkBreakDate.Value
    
    setStatus "Data Retrieval is completed!"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error Occurs: " & Err.Description
    setStatus Err.Description
    
End Sub



Public Function GetData_Market(sSheetName As String, sCellName As String, sSQL As String, iDividor As Integer, sTitle As String, bBreakDate As Boolean) As String

    
    Dim oRs As ADODB.Recordset
    
    On Error GoTo ErrorHandler
   
    Set oRs = GetRS(sSQL)
    
    Dim iStartRow As Integer
    Dim iStartCol As Integer
    
    iStartRow = Sheets(sSheetName).Range(sCellName).Row
    iStartCol = Sheets(sSheetName).Range(sCellName).Column
    
    Sheets(sSheetName).Cells(iStartRow, iStartCol).Value = sTitle
    iStartRow = iStartRow + 1

    
    If Not oRs.EOF Then

        Sheets(sSheetName).Cells(iStartRow, iStartCol).Value = "Pricing Date"
        
        j = iStartRow + 1
        i = iStartCol + 1
        sprevDate = CStr(oRs.Fields("pricing_date").Value)
        Sheets(sSheetName).Cells(iStartRow + 1, iStartCol).Value = CStr(oRs.Fields("pricing_date").Value)
        While Not oRs.EOF
           If i Mod iDividor = 0 Then
             Sheets(sSheetName).Cells(j, CInt(i / iDividor)).Value = oRs.Fields("rate").Value
          
             If (j = iStartRow + 1) And chkBreakDate.Value = False Then
               Sheets(sSheetName).Cells(j - 1, CInt(i / iDividor)).Value = oRs.Fields(1).Value
               Sheets(sSheetName).Cells(j - 1, CInt(i / iDividor)).Font.Bold = True
               Sheets(sSheetName).Cells(j - 1, CInt(i / iDividor)).Interior.ColorIndex = 19
             ElseIf chkBreakDate.Value = True Then
               Sheets(sSheetName).Cells(j - 1, CInt(i / iDividor) + iStartCol).Value = oRs.Fields(1).Value
               Sheets(sSheetName).Cells(j - 1, CInt(i / iDividor) + iStartCol).Font.Bold = True
               Sheets(sSheetName).Cells(j - 1, CInt(i / iDividor) + iStartCol).Interior.ColorIndex = 19
             End If
           End If
           i = i + 1

           oRs.MoveNext
           If Not oRs.EOF Then
             If CStr(oRs.Fields("pricing_date").Value) <> sprevDate Then
               
               If bBreakDate Then
                  iStartRow = j + 1
                  Sheets(sSheetName).Cells(iStartRow, iStartCol).Value = sTitle
                  j = j + 2
               End If
            
               i = iStartCol + 1
               j = j + 1
               sprevDate = CStr(oRs.Fields("pricing_date").Value)
              
             Sheets(sSheetName).Cells(j, iStartCol).Value = oRs.Fields("pricing_date").Value
             End If
           End If
        Wend
        
        'Clean up ADO Objects
        oRs.Close
        Set oRs = Nothing
    End If
    
    GetData_Market = "OK"
    Exit Function
    
ErrorHandler:
    MsgBox "Error Occurs: " & Err.Description
    GetData_Market = Err.Description
    
    setStatus Err.Description


End Function



Private Sub chkAllGroup_Click()
  bAllGroup = Not chkAllGroup.Value
  chkAllGroup.Value = Not chkAllGroup.Value
  FillGroupBox
End Sub

Private Sub cmdClose_Click()
  Me.Hide
End Sub

Private Sub cmdClose_mkt_Click()
 Me.Hide
End Sub

Private Sub cmdClose_port_Click()
 Me.Hide
End Sub

Private Sub cmdClose_prc_Click()
 Me.Hide
End Sub

Private Sub cmdGetPort_Click()

    Dim sStartDate As String
    Dim sEndDate As String
    Dim sSQL As String
    Dim sSection As String
    Dim sCode As String
    Dim i As Long
    Dim sprevDate As String
    Dim iDividor As Integer
    
    On Error GoTo ErrorHandler
    
    iDividor = 1
    If Trim(txtSelectiveDates_port.Text) = "" Then
        If cboStartDate_port.Text = "" And cboEndDate_port.Text = "" Then
          MsgBox "You must enter start and end date", vbCritical
          Exit Sub
        End If
        sStartDate = IIf(cboStartDate_port.Value <> "", CDate(cboStartDate_port.Value), Now())
        sEndDate = IIf(cboEndDate_port.Value <> "", CDate(cboEndDate_port.Value), Now())
    End If
    
    
    setStatus "Begin Data Retrieval, please wait ....."
    
    sSQL = "Select * from PortsPnL where "
    
    If txtSelectiveDates_mkt.Text = "" Then
      sSQL = sSQL & " pricing_date >='" & sStartDate & "' and pricing_date <='" & sEndDate & "' order by pricing_date"
    Else
      sSQL = sSQL & " pricing_date in ('" & Replace(Trim(txtSelectiveDates_mkt.Text), ",", "','") & "') order by pricing_date"
    End If
    

    Dim sSheetName As String
    Dim sCellName As String
    
    sSheetName = Trim(txtSheetName_port.Text)
    sCellName = Trim(txtCellName_port.Text)
    
    GetData_MTMPorts sSheetName, sCellName, sSQL
    
    
    setStatus "Data Retrieval is completed!"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error Occurs: " & Err.Description
    setStatus Err.Description
End Sub


Public Function GetData_MTMPorts(sSheetName As String, sCellName As String, sSQL As String) As String
    Dim oRs As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    Set oRs = GetRS(sSQL)
    
    Dim iStartRow As Integer
    Dim iStartCol As Integer

    If Not oRs.EOF Then

        iStartRow = Sheets(sSheetName).Range(sCellName).Row
        iStartCol = Sheets(sSheetName).Range(sCellName).Column
        For i = 0 To oRs.Fields.Count - 1
           Sheets(sSheetName).Cells(iStartRow, iStartCol + i).Value = oRs.Fields(i).Name
           Sheets(sSheetName).Cells(iStartRow, iStartCol + i).Font.Bold = True
           Sheets(sSheetName).Cells(iStartRow, iStartCol + i).Interior.ColorIndex = 19
        Next
        Sheets(sSheetName).Range(sCellName).Font.Bold = True
        Sheets(sSheetName).Cells(iStartRow + 1, iStartCol).CopyFromRecordset oRs
        
        'Clean up ADO Objects
        oRs.Close
        Set oRs = Nothing
    End If
    
    GetData_MTMPorts = "OK"
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error Occurs: " & Err.Description
    setStatus Err.Description
    
    GetData_MTMPorts = Err.Description
    
End Function



Private Sub cmdGetPrice_Click()
    Dim oRs As ADODB.Recordset
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sSQL As String
    Dim sSection As String
    Dim sCode As String
    Dim i As Long
    Dim sprevDate As String
    Dim iDividor As Integer
    Dim sChosenFields As String
    Dim sCusipLst As String
    Dim k As Integer
    Dim j As Integer
    
    On Error GoTo ErrorHandler
    
    iDividor = 1
    If Trim(txtSelectiveDates_Prc.Text) = "" Then
        If cboStartDate_prc.Text = "" And cboEndDate_prc.Text = "" Then
          MsgBox "You must enter start and end date", vbCritical
          Exit Sub
        End If
        sStartDate = IIf(cboStartDate_prc.Value <> "", CDate(cboStartDate_prc.Value), Now())
        sEndDate = IIf(cboEndDate_prc.Value <> "", CDate(cboEndDate_prc.Value), Now())
    End If
    
    setStatus "Begin Data Retrieval, please wait ....."
    
    Dim sSQLSection
    
    Dim iStartRow As Integer
    Dim iStartCol As Integer
    
    Dim sSheetName As String
    Dim sCellName As String
    
    sSheetName = Trim(txtSheetName_prc.Text)
    sCellName = Trim(txtCellName_prc.Text)
    iStartRow = Sheets(sSheetName).Range(sCellName).Row
    iStartCol = Sheets(sSheetName).Range(sCellName).Column
    
    sCusipLst = Trim(txtCusip_prc.Text)
    For k = 0 To lstPrcSource.ListCount - 1
        If lstPrcSource.Selected(k) Then
              
              sChosenFields = lstPrcSource.List(k, 0)
              Dim sDateCol As String
              
              If sChosenFields = "Risk" Then
                sSQL = "select distinct pricingdate as pricing_date, cusip, price from dbo.masterrawdata "
                sDateCol = "pricingdate"
                sStartDate = Format(sStartDate, "YYYYMMDD")
                sEndDate = Format(sEndDate, "YYYYMMDD")
              ElseIf sChosenFields = "IDC" Then
                sSQL = "select pricing_date, cusip, price from dbo.idc_prices "
                sDateCol = "pricing_date"
              End If
              
            If sCusipLst <> "" Then
                 sSQL = sSQL & " where cusip in ('" & Replace(sCusipLst, ",", "','") & "') and "
            Else
                 sSQL = sSQL & " where "
            End If
            
            If txtSelectiveDates_Prc.Text = "" Then
               sSQL = sSQL & sDateCol & "  >='" & sStartDate & "' and " & sDateCol & " <='" & sEndDate & "'  order by " & sDateCol
            Else
               sSQL = sSQL & sDateCol & " in ('" & Replace(Trim(txtSelectiveDates_Prc.Text), ",", "','") & "')  order by " & sDateCol
            End If
            Set oRs = GetRS(sSQL)
            If Not oRs.EOF Then

               ' Sheets(sSheetName).Cells(iStartRow, iStartCol).Value = "Pricing Date"
                
                j = iStartRow
                i = iStartCol
    
                
                While Not oRs.EOF
                   
                     Sheets(sSheetName).Cells(j, i).Value = CStr(oRs.Fields("pricing_date").Value)
                     Sheets(sSheetName).Cells(j, i + 1).Value = oRs.Fields("cusip").Value
                     Sheets(sSheetName).Cells(j, i + 2).Value = oRs.Fields("price").Value
                  
                     If j = iStartRow Then
                       Sheets(sSheetName).Cells(j, i).Value = "PricingDate"
                       Sheets(sSheetName).Cells(j, i).Font.Bold = True
                       Sheets(sSheetName).Cells(j, i).Interior.ColorIndex = 19
                       Sheets(sSheetName).Cells(j, i + 1).Value = "Cusip"
                       Sheets(sSheetName).Cells(j, i + 1).Font.Bold = True
                       Sheets(sSheetName).Cells(j, i + 1).Interior.ColorIndex = 19
                       Sheets(sSheetName).Cells(j, i + 2).Value = "Price"
                       Sheets(sSheetName).Cells(j, i + 2).Font.Bold = True
                       Sheets(sSheetName).Cells(j, i + 2).Interior.ColorIndex = 19
                     End If
                   
                   oRs.MoveNext
                   j = j + 1
                   
                Wend
                
                
                'Clean up ADO Objects
                oRs.Close
                Set oRs = Nothing
            End If
             iStartCol = iStartCol + 3
            
        End If
       
    Next k
    
    setStatus "Data Retrieval is completed!"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error Occurs: " & Err.Description
    setStatus Err.Description
End Sub

Private Sub cmdRefreshColName_Click()
        Dim oRs As ADODB.Recordset
        Set oRs = GetRS(buildSQL("*", "-1=0"))
        If Not (oRs Is Nothing) Then
            lstDataFields.Clear
            lstDataFields.AddItem "ALL", 0
            lstDataFields.Selected(0) = True
            For i = 0 To oRs.Fields.Count - 1
             lstDataFields.AddItem CStr(oRs.Fields(i).Name), i + 1
            Next
            
            oRs.Close
            
            Set oRs = Nothing
        End If
        FillGroupBox
End Sub

Private Sub cmdRefreshCustID_Click()
    Dim oRs As ADODB.Recordset
    setStatus "Begin Data Retrieval, please wait ....."
        
    Set oRs = GetRS(buildSQL("distinct [CustID]"))
    If Not (oRs Is Nothing) Then
     cboCustID.Text = ""
     cboCustID.Clear
     While Not oRs.EOF
       cboCustID.AddItem oRs.Fields(0).Value
       oRs.MoveNext
     Wend
     
     'Clean up ADO Objects
     oRs.Close
     
     Set oRs = Nothing
    End If

    
    setStatus "Data Retrieval is Completed!"
    
End Sub

'Public Function GetDBName() As String
'
'  Dim sStartDate As Date
'
'  sStartDate = IIf(cboStartDate.Value <> "", CDate(cboStartDate.Value), "")
'  gDBYear = Year(sStartDate)
'  GetDBName = gDBLocation & Year(sStartDate) & ".mdb"
'
'
'End Function


Public Function GetRS(sSQL As String) As ADODB.Recordset
  On Error GoTo ErrorHandler
    txtSQL.Text = txtSQL.Text & vbCr & sSQL
    If oCnn Is Nothing Then
      Set oCnn = GetConn()
      
    End If
    oCnn.CommandTimeout = 180
    Set GetRS = oCnn.Execute(sSQL)
    Exit Function
ErrorHandler:
    setStatus "Error Occurs:" & Err.Description
    MsgBox Err.Description, vbCritical
    Set GetRS = Nothing
End Function


Private Function formatDatesList(sInput As String, Optional sFormat As String = "YYYYMMDD") As String
    
    Dim i As Integer
    Dim sSel() As String
    Dim sSelDates As String
    
    sSelDates = ""
    
    sSel = Split(sInput, ",")
    For i = 0 To UBound(sSel)
       sSelDates = IIf(sSelDates = "", Format(sSel(i), sFormat), sSelDates & "," & Format(sSel(i), sFormat))
    Next
    
    formatDatesList = sSelDates
    
End Function

Private Function buildSQL(Optional sFieldList As String = "*", Optional sCriteria As String = "") As String
  Dim sStartDate As Date
  Dim sEndDate As Date
  Dim sAccountName As String
  Dim sCusip As String
  Dim sCustID As String
  Dim iQuater As Integer
  Dim iYear As Integer
  Dim sSQL As String
  Dim sCrit As String
  Dim sChosenFields As String
  Dim oGroupby As New Dictionary
  Dim sGroupby As String
  Dim sSel() As String
  Dim sSelDates As String
  
  If txtSelectiveDates.Text = "" Then
        sStartDate = IIf(cboStartDate.Value <> "", CDate(cboStartDate.Value), Now())
        sEndDate = IIf(cboEndDate.Value <> "", CDate(cboEndDate.Value), Now())
  Else
       sSelDates = formatDatesList(txtSelectiveDates.Text)
  End If

  sAccountName = IIf(cboAccount.Text <> "", cboAccount.Text, "")
  sCusip = IIf(cboCusip.Text <> "", cboCusip.Text, "")
  sCustID = IIf(cboCustID.Text <> "", cboCustID.Text, "")

  iYear = Year(sStartDate)
  iQuater = Month(sStartDate) / 4 + 1
  
  'sTableName = "MasterRawData" & IIf(iQuater = 1, "0101", IIf(iQuater = 2, "0401", IIf(iQuater = 3, "0701", "1001"))) & iYear
  'sTableName = sTableName & "_" & IIf(iQuater = 1, "0331", IIf(iQuater = 2, "0630", IIf(iQuater = 3, "0930", "1231"))) & iYear
   
   sTableName = gDBMasterTable ' "dbo.MasterRawData"
   
   
  For i = 0 To lstGroupBy.ListCount - 1
      If lstGroupBy.Selected(i) = True Then
        oGroupby.Add lstGroupBy.List(i, 0), lstGroupBy.List(i, 0)
        sGroupby = IIf(sGroupby = "", "", sGroupby & ",") & "[" & lstGroupBy.List(i, 0) & "]"
      End If
  Next i
  
  sCrit = ""
  
  
  For i = 0 To lstDataFields.ListCount - 1
  
    If lstDataFields.Selected(i) Then
     If sGroupby = "" Then
          sChosenFields = IIf(sChosenFields = "", "", sChosenFields & ",") & "[" & lstDataFields.List(i, 0) & "]"
     ElseIf Not (oGroupby.Exists(lstDataFields.List(i))) Then
          sChosenFields = IIf(sChosenFields = "", "", sChosenFields & ",") & "sum([" & lstDataFields.List(i, 0) & "]) as " & lstDataFields.List(i, 0) & "_sum "
     End If
    End If
  Next i
  
  If InStr(sChosenFields, "ALL") > 0 Or sChosenFields = "" Then
    sChosenFields = "*"
  End If
  If Trim(txtSelectiveDates.Text) <> "" Then
     sSQL = "select " & IIf(sFieldList = "", IIf(sGroupby = "", sChosenFields, sGroupby & "," & sChosenFields), sFieldList) & " from " & sTableName & " where [PricingDate] in ('" & Replace(Trim(sSelDates), ",", "','") & "') "
  Else
     sSQL = "select " & IIf(sFieldList = "", IIf(sGroupby = "", sChosenFields, sGroupby & "," & sChosenFields), sFieldList) & " from " & sTableName & " where [PricingDate] >= '" & Format(sStartDate, "YYYYMMDD") & "' and [PricingDate] <='" & Format(sEndDate, "YYYYMMDD") & "' "
  End If
  If sCusip <> "" Then
    If InStr(sCusip, "%") > 0 Then
      sCrit = " cusip like '" & sCusip & "' "
    Else
      sCrit = " cusip = '" & sCusip & "' "
    End If
  End If
  
  If sCustID <> "" Then
    If InStr(sCustID, "%") > 0 Then
      sCrit = IIf(sCrit <> "", sCrit & " and ", "") & " [CustID] like '" & sCustID & "'"
    Else
      sCrit = IIf(sCrit <> "", sCrit & " and ", "") & " [CustID] = '" & sCustID & "'"
    End If
  End If
  
  If sAccountName <> "" Then
    If InStr(sAccountName, "%") > 0 Then
       sCrit = IIf(sCrit <> "", sCrit & " and ", "") & " [AccountName] like '" & sAccountName & "'"
    Else
      sCrit = IIf(sCrit <> "", sCrit & " and ", "") & " [AccountName] = '" & sAccountName & "'"
    End If
  End If
   
  If Trim(txtCustomCriteria.Text) <> "" Then
     sCrit = IIf(sCrit <> "", sCrit & " " & txtCustomCriteria.Text & " ", " " & txtCustomCriteria.Text & " ")
  End If
    
  buildSQL = sSQL & IIf(sCrit <> "", " and " & sCrit, "") & IIf(sCriteria = "", "", " and " & sCriteria) & IIf(sGroupby = "", "", "  group by " & sGroupby) & IIf(InStr(sSQL, "AccountName") > 0, " order by [AccountName]", "")
  
  i = 0
  
End Function

Private Sub cmdGetData_Click()

  Dim i As Integer

  Dim sSheetName As String
  Dim sCellName As String
  Dim sSQL As String
  
  sSheetName = Trim(txtSheetName.Text)
  sCellName = Trim(txtCellName.Text)

  If cboAccount.Text = "" And cboCusip.Text = "" And cboCustID.Text = "" Then
    MsgBox "You must enter at least one of the followings: Account, Cusip or Cust ID", vbCritical
    Exit Sub
  End If
  
  setStatus "Begin Data Retrieval, please wait ....."

  sSQL = buildSQL("")
    
  GetData_Portfolio sSheetName, sCellName, sSQL
      
  setStatus "Data Retrieval is Completed!"
        
End Sub


Public Function GetData_Portfolio(sSheetName As String, sCellName As String, sSQL As String) As String
    Dim oRs As ADODB.Recordset
    
    Dim i As Integer

    Set oRs = GetRS(sSQL)
    
    Dim iStartRow As Integer
    Dim iStartCol As Integer

    If Not (oRs Is Nothing) Then
      If (Not oRs.EOF) Then

        iStartRow = Sheets(sSheetName).Range(sCellName).Row
        iStartCol = Sheets(sSheetName).Range(sCellName).Column
        For i = 0 To oRs.Fields.Count - 1
           Sheets(sSheetName).Cells(iStartRow, iStartCol + i).Value = oRs.Fields(i).Name
           Sheets(sSheetName).Cells(iStartRow, iStartCol + i).Font.Bold = True
           Sheets(sSheetName).Cells(iStartRow, iStartCol + i).Interior.ColorIndex = 19
        Next
        Sheets(sSheetName).Range(sCellName).Font.Bold = True
        Sheets(sSheetName).Cells(iStartRow + 1, iStartCol).CopyFromRecordset oRs
        
        'Clean up ADO Objects
        oRs.Close
        Set oRs = Nothing
      End If
    End If
    GetData_Portfolio = "OK"
    
End Function
Private Function buildDateSQL(Optional sFieldList As String = "*", Optional sCriteria As String = "") As String
  Dim sStartDate As Date
  Dim sEndDate As Date
  Dim sAccountName As String
  Dim sCusip As String
  Dim sCustID As String
  Dim iQuater As Integer
  Dim iYear As Integer
  Dim sSQL As String
  Dim sCrit As String
  Dim sChosenFields As String
  Dim oGroupby As New Dictionary
  Dim sGroupby As String
  
  
  sStartDate = IIf(cboStartDate.Value <> "", CDate(cboStartDate.Value), Now())
  sEndDate = IIf(cboEndDate.Value <> "", CDate(cboEndDate.Value), Now())

  sAccountName = IIf(cboAccount.Text <> "", cboAccount.Text, "")
  sCusip = IIf(cboCusip.Text <> "", cboCusip.Text, "")
  sCustID = IIf(cboCustID.Text <> "", cboCustID.Text, "")

  iYear = Year(sStartDate)
  iQuater = Month(sStartDate) / 4 + 1
  
  'sTableName = "MasterRawData" & IIf(iQuater = 1, "0101", IIf(iQuater = 2, "0401", IIf(iQuater = 3, "0701", "1001"))) & iYear
  'sTableName = sTableName & "_" & IIf(iQuater = 1, "0331", IIf(iQuater = 2, "0630", IIf(iQuater = 3, "0930", "1231"))) & iYear
   
   sTableName = gDBMasterTable ' "dbo.MasterRawData"
   
   
  For i = 0 To lstGroupBy.ListCount - 1
      If lstGroupBy.Selected(i) = True Then
        oGroupby.Add lstGroupBy.List(i, 0), lstGroupBy.List(i, 0)
        sGroupby = IIf(sGroupby = "", "", sGroupby & ",") & "[" & lstGroupBy.List(i, 0) & "]"
      End If
  Next i
  
  sCrit = ""
  
  
  
  For i = 0 To lstDataFields.ListCount - 1
  
    If lstDataFields.Selected(i) Then
     If sGroupby = "" Then
          sChosenFields = IIf(sChosenFields = "", "", sChosenFields & ",") & "[" & lstDataFields.List(i, 0) & "]"
     ElseIf Not (oGroupby.Exists(lstDataFields.List(i))) Then
     '     sChosenFields = IIf(sChosenFields = "", "", sChosenFields & ",") & "[" & lstDataFields.List(i, 0) & "]"
     'Else
          sChosenFields = IIf(sChosenFields = "", "", sChosenFields & ",") & "sum([" & lstDataFields.List(i, 0) & "]) as " & lstDataFields.List(i, 0) & "_sum "
     End If
    'ElseIf oGroupby.Exists(lstDataFields.List(i)) Then
    '    sChosenFields = IIf(sChosenFields = "", "", sChosenFields & ",") & "[" & lstDataFields.List(i, 0) & "]"
    End If
  Next i
  
  If InStr(sChosenFields, "ALL") > 0 Or sChosenFields = "" Then
    sChosenFields = "*"
  End If
  
  sSQL = "select " & IIf(sFieldList = "", IIf(sGroupby = "", sChosenFields, sGroupby & "," & sChosenFields), sFieldList) & " from " & sTableName & " where [PricingDate] >= '" & Format(sStartDate, "YYYYMMDD") & "' and [PricingDate] <='" & Format(sEndDate, "YYYYMMDD") & "' "
  If sCusip <> "" Then
    If InStr(sCusip, "%") > 0 Then
      sCrit = " cusip like '" & sCusip & "' "
    Else
      sCrit = " cusip = '" & sCusip & "' "
    End If
  End If
  
  If sCustID <> "" Then
    If InStr(sCustID, "%") > 0 Then
      sCrit = IIf(sCrit <> "", sCrit & " and ", "") & " [CustID] like '" & sCustID & "'"
    Else
      sCrit = IIf(sCrit <> "", sCrit & " and ", "") & " [CustID] = '" & sCustID & "'"
    End If
  End If
  
  If sAccountName <> "" Then
    If InStr(sAccountName, "%") > 0 Then
       sCrit = IIf(sCrit <> "", sCrit & " and ", "") & " [AccountName] like '" & sAccountName & "'"
    Else
      sCrit = IIf(sCrit <> "", sCrit & " and ", "") & " [AccountName] = '" & sAccountName & "'"
    End If
  End If
   
   
  
   
  buildSQL = sSQL & IIf(sCrit <> "", " and " & sCrit, "") & IIf(sCriteria = "", "", " and " & sCriteria) & IIf(sGroupby = "", "", "  group by " & sGroupby) & IIf(InStr(sSQL, "AccountName") > 0, " order by [AccountName]", "")
  
  i = 0
  
End Function



Private Sub setStatus(sStatus As String)
  Me.txtStatus.Text = sStatus
  DoEvents
End Sub
Private Sub cmdRefAcc_Click()
    Dim oRs As ADODB.Recordset

    
    setStatus "Begin Data Retrieval, please wait ....."
    DoEvents
    If cboStartDate.Value <> "" Then
        Set oRs = GetRS(buildSQL("distinct [AccountName]", "1=1"))
        cboAccount.Clear
        cboAccount.Text = ""
        If Not (oRs Is Nothing) Then
        
        
         While Not oRs.EOF
           cboAccount.AddItem oRs.Fields(0).Value
           oRs.MoveNext
         Wend
         
         oRs.Close
        
         Set oRs = Nothing
        End If
    End If
    setStatus "Data Retrieval is Completed!"
    DoEvents
   
End Sub



Public Function GetConn() As ADODB.Connection

    Set oCnn = New ADODB.Connection
    
     oCnn.ConnectionString = gDBConnectStr ' "Provider=SQLOLEDB;Data Source=w2k3dbmrap1;Initial Catalog=mradb;Trusted_Connection=yes;"
     'Integrated Security=SSPI;"
    ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GetDBName() & ";User Id=admin;Password=;"
    oCnn.Open    'Create your recordset
    Set GetConn = oCnn
    
End Function

Private Sub cmdRefreshCusip_Click()
   Dim oRs As ADODB.Recordset
    
    setStatus "Begin Data Retrieval, please wait ....."
    
 '   If Trim(cboAccount.Text) <> "" Then
      
        Set oRs = GetRS(buildSQL("distinct [cusip]"))
        If Not (oRs Is Nothing) Then
            cboCusip.Clear
              cboCusip.Text = ""
            While Not oRs.EOF
              cboCusip.AddItem oRs.Fields(0).Value
              oRs.MoveNext
            Wend
            oRs.Close
            Set oRs = Nothing
        End If
  '  End If
    setStatus "Data Retrieval is Completed!"
    
End Sub



Private Sub CommandButton1_Click()
   RunBatchGetData Trim(txtSelectiveDates.Text)
End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub UserForm_Activate()

        
End Sub

Private Sub FillGroupBox()
   
   
   
   lstGroupBy.Clear
   If bAllGroup Then
    Dim oRs As ADODB.Recordset
    Set oRs = GetRS(buildSQL("*", "-1=0"))
    If Not (oRs Is Nothing) Then
     lstGroupBy.Clear
    ' lstGroupBy.AddItem "ALL", 0
    ' lstGroupBy.Selected(0) = True
     j = 0
     For i = 0 To oRs.Fields.Count - 1
       If oRs.Fields(i).Type = adChar Or oRs.Fields(i).Type = adVarChar Or oRs.Fields(i).Type = adVarWChar Or oRs.Fields(i).Type = adWChar Then
         lstGroupBy.AddItem CStr(oRs.Fields(i).Name), j
         j = j + 1
       End If
     Next
     
     oRs.Close
     
     Set oRs = Nothing
    End If
   Else
       With lstGroupBy
         .AddItem "PricingDate", 0
         .AddItem "AccountName", 1
         .AddItem "BSAccount", 2
         .AddItem "SubActI", 3
         .AddItem "SubActII", 4
       End With
   End If
   
    lstPrcSource.Clear
    With lstPrcSource
         .AddItem "Risk", 0
         .AddItem "IDC", 1
         .AddItem "FASB", 2
         .AddItem "EJV", 3
         .AddItem "Summit", 4
    End With
   
   Dim gCurveType As Variant
   gCurveType = Array("SWAP_CURVE", "AGENCY", "UST_CURVE", "Index", "SWAPTION_VOLS", "CAP_VOLS")
   cboCurveType.Clear
   For i = 0 To UBound(gCurveType)
     cboCurveType.AddItem gCurveType(i), i
   Next
   
End Sub

Private Sub UserForm_Initialize()
        lstDataFields.Clear
        lstGroupBy.Clear
        lstDataFields.AddItem "ALL", 0
        lstDataFields.Selected(0) = True
        bAllGroup = False
        chkAllGroup.Value = False
        'cboStartDate.Text = "01/01/2008"
        'cboEndDate.Text = "02/02/2008"
        cboStartDate.Text = Format(DateAdd("w", -1, Now()), "MM/DD/YYYY")
        cboEndDate.Text = Format(Now(), "MM/DD/YYYY")
        'If Not IsNull(Sheets("RawPortData")) Then
           txtSheetName.Text = "RawPortData"
       ' Else
       '    txtSheetName.Text = "Sheet1"
       ' End If
        
        txtCellName.Text = "A1"
       ' If Not IsNull(Sheets("RawMarketData")) Then
           txtSheetName_mkt.Text = "RawMarketData"
       ' Else
        '   txtSheetName_mkt.Text = "Sheet2"
        'End If
        
        txtCellName_mkt.Text = "A1"
        
        txtSheetName_port.Text = "Sheet3"
        txtCellName_port.Text = "A1"
        
        txtSheetName_prc.Text = "Sheet3"
        txtCellName_prc.Text = "A1"
        
        'If Not IsNull(Range("DataDates")) Then
        '   txtSelectiveDates.Text = Range("DataDates").Value
        '   txtSelectiveDates_mkt.Text = Range("DataDates").Value
        'End If
        
        FillGroupBox
End Sub



'-- SubActI='SS' Or SubActI ='OR' or SubActI='SS fwd' Or SubActI ='OR fwd' or SubActI='OC' Or SubActI ='OP' Or SubActI ='OP fwd' Or SubActI='ED' OR  cusip in ('DC31SS2003305','DC31SS2003305.5','DC32SS2003154.5') 
