'====== modBatchGetData

Public Sub RunBatchGetData(sDateList As String)

Dim sSQL As String
Dim sSheetName As String
Dim sCellName As String
Dim sDates As String

  sDates = formatDatesList(sDateList)
  
  sSheetName = "RawPortData"
  
  sCellName = "D1"
  sSQL = "select [AccountName],[BSAccount],[PricingDate],[SubActI],[SubActII],[ActL1],[ActL2],[ActL3],sum([DlrCvx]) as DlrCvx_sum ,sum([DlrDur]) as DlrDur_sum ,sum([Holding]) as Holding_sum ,sum([MarketValue]) as MarketValue_sum ,sum([MarketValue_noacc]) as MarketValue_noacc_sum ,sum([MtgeSprdDlrDur]) as MtgeSprdDlrDur_sum ,sum([OpeningMarketValue]) as OpeningMarketValue_sum ,sum([PrepayDlrDur]) as PrepayDlrDur_sum ,sum([TotalOAS01]) as TotalOAS01_sum ,sum([VolDlrDur]) as VolDlrDur_sum  from dbo.MasterRawData where [PricingDate] in ( " & sDates & " )  group by [AccountName],[BSAccount],[PricingDate],[SubActI],[SubActII],[ActL1],[ActL2],[ActL3] order by [AccountName]"

  frmGetData.GetData_Portfolio sSheetName, sCellName, sSQL
  
  sSheetName = "RawMarketData"

  sCellName = "B1"
  sDates = "'" & Replace(sDateList, ",", "','") & "'"
  sSQL = "select b.pricing_date, a.term_in_month as terms,a.coupon as rate from mr_point a inner join market_rate b on a.mr_key = b.mr_key where section = 'SWAP_CURVE'  and pricing_date in (" & sDates & ") and b.mr_type ='E' order by pricing_date"

  frmGetData.GetData_Market sSheetName, sCellName, sSQL, 1, "Swap Curve", False
  
  sSheetName = "RawMarketData"

  sCellName = "B10"
  sDates = "'" & Replace(sDateList, ",", "','") & "'"
  sSQL = "select b.pricing_date, a.term_in_month as terms,a.coupon as rate from mr_point a inner join market_rate b on a.mr_key = b.mr_key where section = 'AGENCY' and Code='AGENCY_YIELD_CURVE'  and pricing_date in (" & sDates & ") and b.mr_type ='E' order by pricing_date"

  frmGetData.GetData_Market sSheetName, sCellName, sSQL, 2, "CO Curve", True
  
  sSheetName = "RawMarketData"

  sCellName = "B28"
  sDates = "'" & Replace(sDateList, ",", "','") & "'"
  sSQL = "select b.pricing_date, a.term_in_month as terms,a.coupon as rate from mr_point a inner join market_rate b on a.mr_key = b.mr_key where section = 'SWAPTION_VOLS' and cast(Code as numeric(5,2)) =10  and pricing_date in (" & sDates & ") and b.mr_type ='E' order by pricing_date"

  frmGetData.GetData_Market sSheetName, sCellName, sSQL, 1, "Swaption VOL:10Y Tenor", False
  
  sSheetName = "RawMarketData"

  sCellName = "B48"
  sDates = "'" & Replace(sDateList, ",", "','") & "'"
  sSQL = "select b.pricing_date, Code as IndexName,a.coupon as rate from mr_point a inner join market_rate b on a.mr_key = b.mr_key where section = 'Index'  and pricing_date in (" & sDates & ") and b.mr_type ='E' order by pricing_date"

  frmGetData.GetData_Market sSheetName, sCellName, sSQL, 1, "Index", True
  
  sSheetName = "PortsData"
  sCellName = "D1"
  sDates = "'" & Replace(sDateList, ",", "','") & "'"
  sSQL = "select * from dbo.MasterRawData where [PricingDate] in (" & sDates & ") and (SubActI='SS' Or SubActI ='OR' or SubActI='SS fwd' Or SubActI ='OR fwd' or SubActI='OC' Or SubActI ='OP' Or SubActI ='OP fwd' Or SubActI='ED' OR  cusip in ('DC31SS2003305','DC31SS2003305.5','DC32SS2003154.5') ) "


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
