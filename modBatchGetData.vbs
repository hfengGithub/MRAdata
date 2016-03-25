'====== modBatchGetData
' Hua 20090908 added order by terms and order by code
' Hua 20091020 changed query for stand alone info - PortsData
' Hua 20091112 Added KRD & KRC in queries
' Hua 20091117 Added cusip level data -- dataByCusip
' Hua 20100105 Change bsAccount; s.custID join with m.ocusip
' Yuanyuan, Hua 20100108 Added secType and filter to eliminate the "Hypothetical" accounts
' Hua 20100129 Filter out the "%Unwound%" accounts
' ????? "Non-Compliance"
' Hua 20100510 use uv_standAlone instead of standAlone table

Public Sub RunBatchGetData(sDateList As String)

Dim sSQL As String
Dim sSheetName As String
Dim sCellName As String
Dim sDates As String

  sDates = formatDatesList(sDateList)
  
  '-- Summary
  sSheetName = "RawPortData"
  
  sCellName = "D1"
  sSQL = " Select [AccountName],replace(bsAccount, 'Derivatives', 'Derivative') as BSAccount,[PricingDate],secType,[SubActI],[SubActII],[ActL1],[ActL2],[ActL3], " & _
         "    sum([DlrCvx]) as DlrCvx_sum, sum([DlrDur]) as DlrDur_sum ,sum([Holding]) as Holding_sum ,sum([MarketValue]) as MarketValue_sum, " & _
         "    sum([MarketValue_noacc]) as MarketValue_noacc_sum ,sum([MtgeSprdDlrDur]) as MtgeSprdDlrDur_sum, sum([OpeningMarketValue]) as OpeningMarketValue_sum , " & _
         "    sum([PrepayDlrDur]) as PrepayDlrDur_sum ,sum([TotalOAS01]) as TotalOAS01_sum, sum([VolDlrDur]) as VolDlrDur_sum, " & _
         "    sum(KRDur1M) as KRDur1M_sum ,sum(KRDur3M) as KRDur3M_sum ,sum(KRDur6M) as KRDur6M_sum, sum(KRDur1Y) as KRDur1Y_sum, sum(KRDur2Y) as KRDur2Y_sum, " & _
         "    sum(KRDur3Y) as KRDur3Y_sum ,sum(KRDur4Y) as KRDur4Y_sum ,sum(KRDur5Y) as KRDur5Y_sum, sum(KRDur7Y) as KRDur7Y_sum, sum(KRDur10Y) as KRDur10Y_sum, " & _
         "    sum(KRDur12Y) as KRDur12Y_sum ,sum(KRDur15Y) as KRDur15Y_sum ,sum(KRDur20Y) as KRDur20Y_sum, sum(KRDur25Y) as KRDur25Y_sum, sum(KRDur30Y) as KRDur30Y_sum, " & _
         "    sum(KRCvx1M) as KRCvx1M_sum ,sum(KRCvx3M) as KRCvx3M_sum ,sum(KRCvx6M) as KRCvx6M_sum, sum(KRCvx1Y) as KRCvx1Y_sum, sum(KRCvx2Y) as KRCvx2Y_sum, " & _
         "    sum(KRCvx3Y) as KRCvx3Y_sum ,sum(KRCvx4Y) as KRCvx4Y_sum, sum(KRCvx5Y) as KRCvx5Y_sum, sum(KRCvx7Y) as KRCvx7Y_sum, sum(KRCvx10Y) as KRCvx10Y_sum , " & _
         "    sum(KRCvx12Y) as KRCvx12Y_sum ,sum(KRCvx15Y) as KRCvx15Y_sum, sum(KRCvx20Y) as KRCvx20Y_sum, sum(KRCvx25Y) as KRCvx25Y_sum, sum(KRCvx30Y) as KRCvx30Y_sum  " & _
         " FROM   dbo.MasterRawData  " & _
         " where PricingDate in (" & sDates & ") and accountName not like '%Hypothetical%' and accountName not like '%Unwound%' " & _
         " group by AccountName,BSAccount,PricingDate,secType,SubActI,SubActII,ActL1,ActL2,ActL3 order by AccountName "

  frmGetData.GetData_Portfolio sSheetName, sCellName, sSQL
  
  '-- cusip level info
  sSheetName = "dataByCusip"
  
  sCellName = "B1"
  sSQL = " Select [AccountName],replace(bsAccount, 'Derivatives', 'Derivative') as BSAccount,[PricingDate],[SubActI],[SubActII],[ActL1],[ActL2],[ActL3], " & _
         "    [DlrCvx], [DlrDur],[Holding],[MarketValue], [MarketValue_noacc],[MtgeSprdDlrDur], [OpeningMarketValue], " & _
         "    [PrepayDlrDur],[TotalOAS01], [VolDlrDur], KRDur1M,KRDur3M,KRDur6M, KRDur1Y, KRDur2Y, " & _
         "    KRDur3Y,KRDur4Y,KRDur5Y, KRDur7Y, KRDur10Y, KRDur12Y,KRDur15Y,KRDur20Y, KRDur25Y, KRDur30Y, " & _
         "    KRCvx1M,KRCvx3M,KRCvx6M, KRCvx1Y, KRCvx2Y, KRCvx3Y ,KRCvx4Y, KRCvx5Y, KRCvx7Y, KRCvx10Y, " & _
         "    KRCvx12Y,KRCvx15Y, KRCvx20Y, KRCvx25Y, KRCvx30Y, cusip, custID, [description]  " & _
         " FROM   dbo.MasterRawData  " & _
         " where PricingDate in (" & sDates & ") and accountName not like '%Hypothetical%' and accountName not like '%Unwound%' "

  frmGetData.GetData_Portfolio sSheetName, sCellName, sSQL
  
  
  '-- Market data
  sSheetName = "RawMarketData"

  sCellName = "B1"
  sDates = "'" & Replace(sDateList, ",", "','") & "'"
  sSQL = "select b.pricing_date, a.term_in_month as terms,a.coupon as rate from polyMr_point a inner join polyMarket_rate b on a.mr_key = b.mr_key where a.currency='USD' and section = 'SWAP_CURVE'  and pricing_date in (" & sDates & ") and b.mr_type ='E' order by pricing_date, terms"

  frmGetData.GetData_Market sSheetName, sCellName, sSQL, 1, "Swap Curve", False
  
  sSheetName = "RawMarketData"

  sCellName = "B10"
  sDates = "'" & Replace(sDateList, ",", "','") & "'"
  sSQL = "select b.pricing_date, a.term_in_month as terms,a.coupon as rate from polyMr_point a inner join polyMarket_rate b on a.mr_key = b.mr_key where a.currency='USD' and section = 'AGENCY' and Code='AGENCY_YIELD_CURVE'  and pricing_date in (" & sDates & ") and b.mr_type ='E' order by pricing_date, terms"

  frmGetData.GetData_Market sSheetName, sCellName, sSQL, 1, "CO Curve", True
  
  sSheetName = "RawMarketData"

  sCellName = "B28"
  sDates = "'" & Replace(sDateList, ",", "','") & "'"
  sSQL = "select b.pricing_date, a.term_in_month as terms,a.coupon as rate from polyMr_point a inner join polyMarket_rate b on a.mr_key = b.mr_key where a.currency='USD' and section = 'SWAPTION_VOLS' and cast(Code as numeric(5,2)) =10  and pricing_date in (" & sDates & ") and b.mr_type ='E' order by pricing_date, terms"

  frmGetData.GetData_Market sSheetName, sCellName, sSQL, 1, "Swaption VOL:10Y Tenor", False
  
  sSheetName = "RawMarketData"

  sCellName = "B48"
  sDates = "'" & Replace(sDateList, ",", "','") & "'"
  sSQL = "select b.pricing_date, Code as IndexName,a.coupon as rate from polyMr_point a inner join polyMarket_rate b on a.mr_key = b.mr_key where a.currency='USD' and section = 'Index'  and pricing_date in (" & sDates & ") and b.mr_type ='E' order by pricing_date, code"

  frmGetData.GetData_Market sSheetName, sCellName, sSQL, 1, "Index", True
  
  '-- standAlone 20100510 Hua: use view instead of StandAlone table.
  sSheetName = "PortsData"
  sCellName = "B1"
  sDates = "'" & Replace(sDateList, ",", "','") & "'"
  sSQL = " Select s.PricingDate, subAct=m.ActL2 ,s.CustID ,  " & _
         "    portfolio=Case when s.Ports in (191,195) then 'StandAlone' WHEN s.Ports IS NULL THEN 'Port'+substring(s.cusip,3,2) else 'Port'+ cast(Ports as varchar(20)) end, " & _
         "    m.SecType, MV=m.MarketValue, m.MarketValue_noacc, m.DlrCvx, m.DlrDur, m.VolDlrDur, " & _
         "    KRDur1M, KRDur3M  ,KRDur6M  , KRDur1Y  , KRDur2Y  , " & _
         "    KRDur3Y, KRDur4Y  ,KRDur5Y  , KRDur7Y  , KRDur10Y , " & _
         "    KRDur12Y,KRDur15Y, KRDur20Y , KRDur25Y , KRDur30Y , " & _
         "    KRCvx1M, KRCvx3M  ,KRCvx6M  , KRCvx1Y  , KRCvx2Y  , " & _
         "    KRCvx3Y, KRCvx4Y , KRCvx5Y  , KRCvx7Y  , KRCvx10Y  , " & _
         "    KRCvx12Y,KRCvx15Y, KRCvx20Y , KRCvx25Y , KRCvx30Y   " & _
         " FROM   uv_StandAlone as s INNER JOIN masterRawData as m ON s.PricingDate = m.PricingDate AND " & _
         "    s.cusip=m.cusip AND s.custID=M.custID  " & _
         " where m.pDate in (" & sDates & ") "


  frmGetData.GetData_MTMPorts sSheetName, sCellName, sSQL
  
  MsgBox "Batch is done!"
  
End Sub

Private Function formatDatesList(sInput As String, Optional sFormat As String = "YYYYMMDD") As String
    
    Dim i As Integer
    Dim sSel() As String
    Dim sSelDates As String
    
    sSelDates = ""
    
    sSel = Split(sInput, ",")
    For i = 0 To UBound(sSel)
       sSelDates = IIf(sSelDates = "", Format(sSel(i), sFormat), sSelDates & "," & Format(sSel(i), sFormat))
    Next
    
    formatDatesList = "'" & Replace(sSelDates, ",", "','") & "'"
    
End Function
