# Readme..
	# This script will generate LDAP IP/Filters/TimeWaits summary Excel pivot tables from Directory Service's event 1644 EventLogs using LogParser and Excel via COM objects in 2 steps.
    #    1. Script calls LogParser to scans all event 1644 evtx in input directory, exact event data from event 1644, export to CSV.
    #    2. Script calls into Excel to import resulting CSV, create pivot tables for common ldap workload analysis. Delete CSV afterward.
	#
	# LogParserWrapper_1644.ps1 v0.7 11/22(split set-PivotPageRows)
	#		Steps: 
	#   	1. Install LogParser 2.2 from https://www.microsoft.com/en-us/download/details.aspx?id=24659
	#     	Note: More about LogParser2.2 https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-xp/bb878032(v=technet.10)?redirectedfrom=MSDN
	#   	2. Copy Directory Service EVTX from target DC(s) to same directory as this script.
	#     		Tip: When copying Directory Service EVTX, filter on event 1644 to reduce EVTX size for quicker transfer. 
	#					Note: Script will process all *.EVTX in script directory when run.
	#   	3. Run script

#------Script variables block, modify to fit your needs ---------------------------------------------------------------------
$g_StartTime = '3/19/2010 1:11:49 AM' # Earliest 1644 event to export, in the form of M/d/yyyy H:m:s tt' example: '3/19/2010 1:11:49 AM'. Use this to filter events after changes.
$g_LookBackDays = 0 #2080             # 0 means script start list events after $g_StartTime, when set to an interger, script will list events in last $g_LookBackDays days. For examle: 1 will list events occurs in last 24 hours. Use this to filter events after changes.
$g_MaxExports = 99999                 # Max number of 1644 events to export per each EVTX. Use this for quicker spot checks.
$g_ColorBar   = $True                 # Can set to $false to speed up excel import & reduce memory requirement. 
$g_ColorScale = $True                 # Can set to $false to speed up excel import & reduce memory requirement. Color Scale requires '$g_ColorBar = $True' for color index. 
$ErrorActionPreference = "SilentlyContinue"
function Set-PivotField { param (
  $PivotField = $null, $Orientation = $null, $NumberFormat = $null, $Function = $null, $Calculation = $null, $BaseField = $null, $Name = $null, $Position = $null, $Group = $null
  )
    if ($null -ne $Orientation) {$PivotField.Orientation = $Orientation}
    if ($null -ne $NumberFormat) {$PivotField.NumberFormat = $NumberFormat}
    if ($null -ne $Function) {$PivotField.Function = $Function}
    if ($null -ne $Calculation) {$PivotField.Calculation = $Calculation}
    if ($null -ne $BaseField) {$PivotField.BaseField = $BaseField}
    if ($null -ne $Name) {$PivotField.Name = $Name}
    if ($null -ne $Position) {$PivotField.Position = $Position}
    if ($null -ne $Group) {($PivotField.DataRange.Item($group)).group($true,$true,1,($false, $true, $true, $true, $false, $false, $false)) | Out-Null}
}
function Set-PivotPageRows { param (
    $Sheet = $null, $PivotTable = $null, $Page = $null, $Rows = $null
  )
  $xlRowField   = 1 #XlPivotFieldOrientation 
  $xlPageField  = 3 #XlPivotFieldOrientation 
  Set-PivotField -PivotField $Sheet.PivotTables("$PivotTable").PivotFields("$Page") -Orientation $xlPageField
  $i=0
  ($Rows).foreach({
    $i++
    If ($i -lt ($Rows).count) {Set-PivotField -PivotField $Sheet.PivotTables("$PivotTable").PivotFields("$_") -Orientation $xlRowField}
    else {Set-PivotField -PivotField $Sheet.PivotTables("$PivotTable").PivotFields("$_") -Orientation $xlRowField -Group $i}
  })
}
function Set-TableFormats { param (
  $Sheet = $null, $Table = $null, $ColumnWidth = $null, $label = $null, $Name = $null, $ColorScale = $null, $ColorBar = $null, $SortColumn = $null, $Hide = $null, $ColumnHiLite = $null
  )
  $Sheet.PivotTables("$Table").HasAutoFormat = $False
    $Column = 1
    $ColumnWidth.foreach({ $Sheet.columns.item($Column).columnwidth = $_
      $Column++
    })
    $Sheet.Application.ActiveWindow.SplitRow = 3
    $Sheet.Application.ActiveWindow.SplitColumn = 2
    $Sheet.Application.ActiveWindow.FreezePanes = $true
    $Sheet.Cells.Item(3,1) = $label
    $Sheet.Name = $Name
    if ($null -ne $SortColumn) {$null = $Sheet.Cells.Item($SortColumn,4).Sort($Sheet.Cells.Item($SortColumn,4),2)}
    if ($null -ne $Hide) {$Hide.foreach({($Sheet.PivotTables("$Table").PivotFields($_)).ShowDetail = $false})}
    if ($null -ne $ColumnHiLite) {
      $Sheet.Range("A4:"+[char]($sheet.UsedRange.Cells.Columns.count+64)+[string](($Sheet.UsedRange.Cells).Rows.count-1)).interior.Color = 16056319
      $ColumnHiLite.ForEach({$sheet.Range(($_+"3")).interior.ColorIndex = 37})
    }
    if (($null -ne $ColorBar) -and ($g_ColorBar -eq $true)) {
      $ColorRange='$'+$ColorBar+'$4:$'+$ColorBar+'$'+(($Sheet.UsedRange.Cells).Rows.Count-1)
      $null = $Sheet.Range($ColorRange).FormatConditions.AddDatabar()
    }
    if (($null -ne $ColorScale) -and ($g_ColorScale -eq $true)) {
      $ColorRange='$'+$ColorScale+'$4:$'+$ColorScale+'$'+(($Sheet.UsedRange.Cells).Rows.Count-1)
      $null = $Sheet.Range($ColorRange).FormatConditions.AddColorScale(3)
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(1).type = 1
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(1).FormatColor.Color = 8109667
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(2).FormatColor.Color = 8711167
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(3).type = 2 
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(3).FormatColor.Color = 7039480
    }
}
#------Main---------------------------------
if ($g_LookBackDays -ne 0) { $g_StartTime = (Get-Date).AddDays(0-$g_LookBackDays) -f 'M/d/yyyy H:m:s tt' }
  $ScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
  $TotalSteps = ((Get-ChildItem -Path $ScriptPath -Filter '*.evtx').count)+9
    $Step=1
  $TimeStamp = "{0:yyyy-MM-dd_hh-mm-ss_tt}" -f (Get-Date)
  (Get-ChildItem -Path $ScriptPath -Filter '*.evtx').foreach({
      $InFile = $ScriptPath+'\'+$_
    $InputFormat = New-Object -ComObject MSUtil.LogQuery.EventLogInputFormat
    $OutputFormat = New-Object -ComObject MSUtil.LogQuery.CSVOutputFormat
    $OutTitle = 'Temp1644-'+$_.BaseName
    $OutFile = "$ScriptPath\$TimeStamp-$OutTitle.csv"
    $Query = @"
      SELECT TOP $g_MaxExports
        ComputerName as LDAPServer,
        TimeGenerated as TimeGenerated,
        EXTRACT_TOKEN ( Strings,0,'|') as StartingNode,		
        REPLACE_STR(REPLACE_STR(Strings,STRCAT(EXTRACT_PREFIX(Strings,0,'|'),'|'),''),STRCAT('|',EXTRACT_SUFFIX(Strings,12,'|')),'') as Filter,
        EXTRACT_PREFIX(EXTRACT_SUFFIX(Strings,12,'|'),0,'|') as VisitedEntries,
        EXTRACT_PREFIX(EXTRACT_SUFFIX(Strings,11,'|'),0,'|')  as ReturnedEntries,
        EXTRACT_PREFIX(EXTRACT_SUFFIX(Strings,10,'|'),0,':') as ClientIP,
        EXTRACT_SUFFIX(EXTRACT_PREFIX(EXTRACT_SUFFIX(Strings,10,'|'),0,'|'),0,':') as ClientPort,
        EXTRACT_PREFIX(EXTRACT_SUFFIX(Strings, 9,'|'),0,'|') as SearchScope,
        EXTRACT_PREFIX(EXTRACT_SUFFIX(Strings, 8,'|'),0,'|') as AttributeSelection,
        EXTRACT_PREFIX(EXTRACT_SUFFIX(Strings, 7,'|'),0,'|') as ServerControls,
        EXTRACT_PREFIX(EXTRACT_SUFFIX(Strings, 6,'|'),0,'|') as UsedIndexes,
        EXTRACT_PREFIX(EXTRACT_SUFFIX(Strings, 5,'|'),0,'|') as PagesReferenced,
        EXTRACT_PREFIX(EXTRACT_SUFFIX(Strings, 4,'|'),0,'|') as PagesReadFromDisk,
        EXTRACT_PREFIX(EXTRACT_SUFFIX(Strings, 3,'|'),0,'|') as PagesPreReadFromDisk,
        EXTRACT_PREFIX(EXTRACT_SUFFIX(Strings, 2,'|'),0,'|') as CleanPagesModified,
        EXTRACT_PREFIX(EXTRACT_SUFFIX(Strings, 1,'|'),0,'|') as DirtyPagesModified,
        EXTRACT_SUFFIX (Strings, 0,'|') as SearchTimeMS
      INTO $OutFile
      FROM $InFile
      WHERE
        EventID = 1644 and TimeGenerated > TO_TIMESTAMP('$g_StartTime','M/d/yyyy H:m:s tt') 
"@
    Write-Progress -Activity "Generating $OutTitle report" -PercentComplete (($Step++/$TotalSteps)*100)
    while ((get-service eventlog).status -ne 'Running') { 
      try { Start-Service EventLog -ErrorAction stop }
      catch {  Write-Host 'Unable to start EventLog service, possible persmission issue, try restart script under admin powershell prompt.' -BackgroundColor Red
      Start-Sleep 5 }
    }
    $LPQuery = New-Object -ComObject MSUtil.LogQuery
    $null = $LPQuery.ExecuteBatch($Query,$InputFormat,$OutputFormat) | Out-Null
  })
  while ((get-service eventlog).status -ne 'Running') { 
    try { Start-Service EventLog -ErrorAction stop }
    catch {  Write-Host 'Unable to start EventLog service, possible persmission issue, try restart script under admin powershell prompt.' -BackgroundColor Red
      Start-Sleep 5 }
  }
  $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($InputFormat) 
  $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OutputFormat)
  $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($LPQuery)
#-----Combine CSV(s) into one for faster Excel import
  $OutTitle1 = 'LDAP-1644-Report'
  $OutFile1 = "$ScriptPath\$TimeStamp-$OutTitle1.csv"
  Write-Progress -Activity "Generating $OutTitle report" -PercentComplete (($Step++/$TotalSteps)*100)
    Get-ChildItem -Path $ScriptPath -Filter "$TimeStamp-Temp1644-*.csv" | Select-Object -ExpandProperty FullName | Import-Csv | Export-Csv  $OutFile1 -NoTypeInformation -Append
    $null = Get-ChildItem -Path $ScriptPath -Filter "$TimeStamp-Temp1644-*.csv" | Remove-Item
#----Excel COM variables-------------------------------------------------------------------
  $fmtNumber  = "###,###,###,###,###"
  $fmtPercent = "#0.00%"
  $xlDataField  = 4 #XlPivotFieldOrientation 
  $xlAverage    = -4106 #XlConsolidationFunction
  $xlSum        = -4157 #XlConsolidationFunction 
  $xlPercentOfTotal = 8 #XlPivotFieldCalculation 
#-------Import to Excel
If (Test-Path $OutFile1) { 
  $Excel = New-Object -ComObject excel.application
  Write-Progress -Activity "Import to Excel $OutTitle report" -PercentComplete (($Step++/$TotalSteps)*100)
    # $Excel.visible = $true
    $Excel.Workbooks.OpenText("$OutFile1")
    $Sheet0 = $Excel.Workbooks[1].Worksheets[1]
      $null = $Sheet0.Range("A1").AutoFilter()
      $Sheet0.Application.ActiveWindow.SplitRow=1  
      $Sheet0.Application.ActiveWindow.FreezePanes = $true
      $null = $Sheet0.Columns.AutoFit() 
      $Sheet0.Columns.Item(3).columnwidth = $Sheet0.Columns.Item(4).columnwidth = $Sheet0.Columns.Item(10).columnwidth = $Sheet0.Columns.Item(11).columnwidth = $Sheet0.Columns.Item(12).columnwidth = 70
      $Sheet0.Columns.Item('E').numberformat = $Sheet0.Columns.Item('F').numberformat = $Sheet0.Columns.Item('M').numberformat = $Sheet0.Columns.Item('N').numberformat = $Sheet0.Columns.Item('O').numberformat = $Sheet0.Columns.Item('p').numberformat = $Sheet0.Columns.Item('Q').numberformat = $Sheet0.Columns.Item('R').numberformat = "###,###,###,###,###"
      $Sheet0.Name = $OutTitle1
      $null = $Sheet0.ListObjects.Add(1, $Sheet0.Application.ActiveCell.CurrentRegion, $null ,0)
    #----Pivot Table 1-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopCount StartingNode Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet1 = $Excel.Workbooks[1].Worksheets.add()
      $PivotTable1 = $Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5) 
      $PivotTable1.CreatePivotTable("Sheet1!R1C1") | Out-Null
      Set-PivotPageRows -Sheet $sheet1 -PivotTable "PivotTable1" -Page "LDAPServer" -Rows ("StartingNode","Filter","ClientIP","TimeGenerated")
      Set-PivotField -PivotField $Sheet1.PivotTables("PivotTable1").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" -Position 1
      Set-PivotField -PivotField $Sheet1.PivotTables("PivotTable1").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlAverage -Name "AvgSearchTime" -Position 2 
      Set-PivotField -PivotField $Sheet1.PivotTables("PivotTable1").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal" -Position 3
        Set-TableFormats -Sheet $Sheet1 -Table "PivotTable1" -ColumnWidth (60,12,14,12,14) -label 'StartingNode grouping' -Name '1.TopCount StartingNode' -SortColumn 4 -Hide ('ClientIP','Filter','StartingNode') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D'
    #----Pivot Table 2-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopCount IP Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet2 = $Excel.Workbooks[1].Worksheets.add()
      $PivotTable2 = $Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5) 
      $PivotTable2.CreatePivotTable("Sheet2!R1C1") | Out-Null
      Set-PivotPageRows -Sheet $sheet2 -PivotTable "PivotTable2" -Page "LDAPServer" -Rows ("ClientIP","Filter","TimeGenerated")
      Set-PivotField -PivotField $Sheet2.PivotTables("PivotTable2").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" -Position 1
      Set-PivotField -PivotField $Sheet2.PivotTables("PivotTable2").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlAverage -Name "AvgSearchTime" -Position 2 
      Set-PivotField -PivotField $Sheet2.PivotTables("PivotTable2").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal" -Position 3
        Set-TableFormats -Sheet $Sheet2 -Table "PivotTable2" -ColumnWidth (60,12,19,12) -label 'IP grouping' -Name '2.TopCount IP' -SortColumn 4 -Hide ('ClientIP','Filter') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D'
    #----Pivot Table 3-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopCount Filters Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet3 = $Excel.Workbooks[1].Worksheets.add()
      $PivotTable3 = $Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5) 
      $PivotTable3.CreatePivotTable("Sheet3!R1C1") | Out-Null
      Set-PivotPageRows -Sheet $sheet3 -PivotTable "PivotTable3" -Page "LDAPServer" -Rows ("Filter","ClientIP","TimeGenerated")
      Set-PivotField -PivotField $Sheet3.PivotTables("PivotTable3").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" -Position 1
      Set-PivotField -PivotField $Sheet3.PivotTables("PivotTable3").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlAverage -Name "AvgSearchTime" -Position 2 
      Set-PivotField -PivotField $Sheet3.PivotTables("PivotTable3").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal" -Position 3
        Set-TableFormats -Sheet $Sheet3 -Table "PivotTable3" -ColumnWidth (70,12,19,12) -label 'Filter grouping' -Name '3.TopCount Filters' -SortColumn 4 -Hide ('ClientIP','Filter') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D'
    #----Pivot Table 4-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopTime IP Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet4 = $Excel.Workbooks[1].Worksheets.add()
      $PivotTable4 = $Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5) 
      $PivotTable4.CreatePivotTable("Sheet4!R1C1") | Out-Null
      Set-PivotPageRows -Sheet $sheet4 -PivotTable "PivotTable4" -Page "LDAPServer" -Rows ("ClientIP","Filter","TimeGenerated")
      Set-PivotField -PivotField $Sheet4.PivotTables("PivotTable4").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlSum -Name "Total SearchTime" -Position 1
      Set-PivotField -PivotField $Sheet4.PivotTables("PivotTable4").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" -Position 2 
      Set-PivotField -PivotField $Sheet4.PivotTables("PivotTable4").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal" -Position 3
        Set-TableFormats -Sheet $Sheet4 -Table "PivotTable4" -ColumnWidth (50,21,12,19) -label 'IP grouping' -Name '4.TopTime IP' -SortColumn 4 -Hide ('ClientIP','Filter') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D'
    #----Pivot Table 5-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopTime Filter Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet5 = $Excel.Workbooks[1].Worksheets.add()
      $PivotTable5 = $Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5) 
      $PivotTable5.CreatePivotTable("Sheet5!R1C1") | Out-Null
      Set-PivotPageRows -Sheet $sheet5 -PivotTable "PivotTable5" -Page "LDAPServer" -Rows ("Filter","ClientIP","TimeGenerated")
      Set-PivotField -PivotField $Sheet5.PivotTables("PivotTable5").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlSum -Name "Total SearchTime" -Position 1
      Set-PivotField -PivotField $Sheet5.PivotTables("PivotTable5").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" -Position 2 
      Set-PivotField -PivotField $Sheet5.PivotTables("PivotTable5").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal" -Position 3
        Set-TableFormats -Sheet $Sheet5 -Table "PivotTable5" -ColumnWidth (70,21,12,19) -label 'IP grouping' -Name '5.TopTime Filter' -SortColumn 4 -Hide ('ClientIP','Filter') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D'
    #---General Tab Operations-------------------------------------------------------------------
    $Sheet1.Tab.ColorIndex = $Sheet2.Tab.ColorIndex = $Sheet3.Tab.ColorIndex = 35
    $Sheet4.Tab.ColorIndex = $Sheet5.Tab.ColorIndex = 36
      $WorkSheetNames = New-Object System.Collections.ArrayList  #---Sort by sheetName-
      foreach($WorkSheet in $Excel.Workbooks[1].Worksheets) { $null = $WorkSheetNames.add($WorkSheet.Name) }
        $null = $WorkSheetNames.Sort()
        For ($i=0; $i -lt $WorkSheetNames.Count-1; $i++){ ($Excel.Workbooks[1].Worksheets.Item($WorkSheetNames[$i])).Move($Excel.Workbooks[1].Worksheets.Item($i+1)) }
    $Sheet1.Activate()
    #----Save and clean up
      $Excel.Workbooks[1].SaveAs($ScriptPath+'\'+$TimeStamp+'-'+$OutTitle1,51)
        $iCSV = "$ScriptPath\$TimeStamp-$OutTitle1.csv"
        Remove-Item $iCSV
      $Excel.visible = $true
      $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
      # Stop-process -Name Excel 
} else {
	Write-Host 'No LogParser CSV found. Please confirm evtx contain event 1644.' -ForegroundColor Red
}
