$serviceNowIncidentExcelFileObject = New-Object -ComObject excel.application
$serviceNowIncidentExcelFileObject.Visible = $true
$serviceNowIncidentExcelFileObject.DisplayAlerts=$false
$serviceNowIncidentExcelWorkbook = $serviceNowIncidentExcelFileObject.Workbooks.Open("C:\Users\smadhra\Desktop\Dupoint\Service Now Incident.xlsx")
$serviceNowIncidentExcelWorkbook.activate()
$serviceNowIncidentExcelWorksheet = $serviceNowIncidentExcelWorkbook.Worksheets.Item(1)
$serviceNowIncidentExcelWorksheet.Activate()

$serviceNowIncidentCount = $serviceNowIncidentExcelFileObject.WorksheetFunction.CountIf($serviceNowIncidentExcelWorksheet.Range("A1:" + "A" + $serviceNowIncidentExcelWorksheet.Rows.Count), "<>") - 1
$serviceNowIncidentCount
$serviceNowIncidentExcelWorksheetRange = $serviceNowIncidentExcelWorksheet.Range("A2:M$serviceNowIncidentCount")
$serviceNowIncidentExcelWorksheetRange.copy()

$serviceNowMasterSheetIncidentExcelFileObject = New-Object -ComObject excel.application
$serviceNowMasterSheetIncidentExcelFileObject.Visible = $true
$serviceNowMasterSheetIncidentExcelFileObject.DisplayAlerts=$false
$serviceNowMasterSheetIncidentExcelWorkbook = $serviceNowMasterSheetIncidentExcelFileObject.Workbooks.Open("C:\Users\smadhra\Desktop\Dupoint\Master Sheet Excel - Copy.xlsx")
$serviceNowMasterSheetIncidentExcelWorkbook.activate()
$serviceNowIncidentMasterSheetExcelWorksheet = $serviceNowMasterSheetIncidentExcelWorkbook.Worksheets.Item("DATA")
$serviceNowIncidentMasterSheetExcelWorksheet.Activate()

$serviceNowIncidentMasterSheetExcelWorksheetRange = $serviceNowIncidentMasterSheetExcelWorksheet.Range("A2:N1048576")
$serviceNowIncidentMasterSheetExcelWorksheetRange.clear()

$serviceNowIncidentMasterSheetExcelWorksheetRange = $serviceNowIncidentMasterSheetExcelWorksheet.Range("B2")
$serviceNowIncidentMasterSheetExcelWorksheet.Paste($serviceNowIncidentMasterSheetExcelWorksheetRange)
$serviceNowIncidentMasterSheetExcelWorksheet.Cells.Item(2,1) = 'DUP-CSC-Logistics Apps'
$serviceNowIncidentMasterSheetExcelWorksheetRange = $serviceNowIncidentMasterSheetExcelWorksheet.Range("A2")
$serviceNowIncidentMasterSheetExcelWorksheetRange.copy()
$serviceNowIncidentMasterSheetExcelWorksheet_ = $serviceNowIncidentMasterSheetExcelWorksheet.Range("A3:A$serviceNowIncidentCount")
$serviceNowIncidentMasterSheetExcelWorksheet.Paste($serviceNowIncidentMasterSheetExcelWorksheet_)
$serviceNowIncidentMasterSheetExcelWorksheet.Columns.Item(8).NumberFormat = "DD-MM-YYYY HH:MM:SS"
$serviceNowIncidentMasterSheetExcelWorksheet.Range("O2:R$serviceNowIncidentCount").Formula = $serviceNowIncidentMasterSheetExcelWorksheet.Range("O2:R2").Formula
$serviceNowMasterSheetIncidentExcelWorkbook.Save()
$serviceNowIncidentExcelFileObject.Quit()


$serviceNowIncidentMasterSheetExcelWorksheet.Sort.SortFields.Clear()
$serviceNowIncidentMasterSheetExcelWorksheetSortRange = $serviceNowIncidentMasterSheetExcelWorksheet.Range("H2:H$serviceNowIncidentCount")
$serviceNowIncidentMasterSheetExcelWorksheetFullRange = $serviceNowIncidentMasterSheetExcelWorksheet.Range("A1:R$serviceNowIncidentCount")
$serviceNowIncidentMasterSheetExcelWorksheet.Sort.SortFields.Add2($serviceNowIncidentMasterSheetExcelWorksheetSortRange, 0,  1, 0)
$serviceNowIncidentMasterSheetExcelWorksheet.Sort.SetRange($serviceNowIncidentMasterSheetExcelWorksheetFullRange)
$serviceNowIncidentMasterSheetExcelWorksheet.Sort.Header = 1
$serviceNowIncidentMasterSheetExcelWorksheet.Sort.MatchCase = $False
$serviceNowIncidentMasterSheetExcelWorksheet.Sort.Orientation = 1
$serviceNowIncidentMasterSheetExcelWorksheet.Sort.SortMethod = 1
$serviceNowIncidentMasterSheetExcelWorksheet.Sort.Apply()
$serviceNowMasterSheetIncidentExcelWorkbook.Save()

Start-Sleep -m 3


$lastIncidentCreationDate  = $serviceNowIncidentMasterSheetExcelWorksheet.Cells.Item($serviceNowIncidentCount,8).Text

$lastIncidentCreationDate = $lastIncidentCreationDate.Substring(0,10)

$lastIncidentCreationDateFormat = [datetime]::ParseExact($lastIncidentCreationDate.Trim(),”dd-MM-yyyy”,$null)

$i = 0
$creationArray=@{}
while ($i -lt 6){

$date_ = $lastIncidentCreationDateFormat.AddMonths(-$i);
$year = $date_.Year
$month = $date_.Month
 if ($Month -lt 10){
     $setLargetMonthYear = $year.ToString() + "/" + "0" + $month.ToString()
     
     }else{
     $setLargetMonthYear = $year.ToString()  + "/" + $month.ToString()
     
     }

     $creationArray.Add($setLargetMonthYear,$setLargetMonthYear);
     
    $i++
   

}


$serviceNowTocusIncidentMasterSheetExcelWorksheet = $serviceNowMasterSheetIncidentExcelWorkbook.Worksheets.Item("Tocus")
$serviceNowTocusIncidentMasterSheetExcelWorksheet.Activate()

$pivotTable1 = $serviceNowTocusIncidentMasterSheetExcelWorksheet.PivotTables("PivotTable5")

$pivotTable1.SourceData ="DATA!R1C1:R" + $serviceNowIncidentCount + "C18"
$pivotTable1.RefreshTable()
$pivotTable1.Update()

$pivotTable = $serviceNowTocusIncidentMasterSheetExcelWorksheet.PivotTables("PivotTable6")
$pivotTable.SourceData ="DATA!R1C1:R" + $serviceNowIncidentCount + "C18"
$pivotTable.RefreshTable()
$pivotTable.Update()
$pivotTable = $serviceNowTocusIncidentMasterSheetExcelWorksheet.PivotTables("PivotTable7")
$pivotTable.SourceData ="DATA!R1C1:R" + $serviceNowIncidentCount + "C18"
$pivotTable.RefreshTable()
$pivotTable.Update()
$serviceNowMasterSheetIncidentExcelWorkbook.RefreshAll



$PivotFields1 = $pivotTable1.PivotFields("Creation")
$PivotFields1.EnableMultiplePageItems = $True
     foreach($item in $PivotFields1.PivotItems() ){
 
     
     if ($creationArray.ContainsKey($item.Name))
     {
     $PivotFields1.PivotItems($item.Name).Visible = $True
        }
     
     else
     {
     $PivotFields1.PivotItems($item.Name).Visible = $False
     }
     }

     $PivotFields1.PivotItems("2016/05").Visible = $False

#$pivotTable = $serviceNowTocusIncidentMasterSheetExcelWorksheet.PivotTables("PivotTable5")
#$pivotTable.SourceData ="DATA!R1C1:R" + $serviceNowIncidentCount + "C18"
#$pivotTable.RefreshTable()
#$pivotTable.Update()
$serviceNowMasterSheetIncidentExcelWorkbook.Save()


$pivotTable = $serviceNowTocusIncidentMasterSheetExcelWorksheet.PivotTables("PivotTable5")
$pivotTable.TableRange2.Offset(1, 0).Copy
$pivotTable.PivotSelect( "", 0, $True)
$serviceNowMasterSheetIncidentExcelFileObject.Selection.Copy()
$word = New-Object -ComObject word.application
$word.Visible = $true
$doc = $word.Documents.Open("C:\Users\smadhra\Desktop\Dupoint\Word Template.doc")
$selection = $word.Selection
Start-Sleep -s 3
$default = [Type]::Missing
$selection.GoTo(1,2,9,2);
Start-Sleep -m 3
$selection.TypeText("SUMMARY FOR TOCUS");
$selection.TypeParagraph()
$selection.TypeParagraph()
$Selection.ParagraphFormat.LeftIndent =  $word.InchesToPoints(0)
$Selection.ParagraphFormat.SpaceBeforeAuto = $False
$Selection.ParagraphFormat.SpaceAfterAuto = $False
Start-Sleep -m 3

$selection.PasteSpecial($default, $false, 0, $false, 9, $default,$default)
$selection.EndKey(6)
$selection.TypeParagraph()
$selection.TypeParagraph()
$Selection.ParagraphFormat.LeftIndent =  $word.InchesToPoints(0)
$Selection.ParagraphFormat.SpaceBeforeAuto = $False
$Selection.ParagraphFormat.SpaceAfterAuto = $False
Start-Sleep -m 3

$serviceNowMasterSheetIncidentExcelWorkbook.activate()
$serviceNowTocusIncidentMasterSheetExcelWorksheet.Activate()
Start-Sleep -m 3

$pivotTable = $serviceNowTocusIncidentMasterSheetExcelWorksheet.PivotTables("PivotTable6")

$pivotTable.TableRange2.Offset(1, 0).Copy
$pivotTable.PivotSelect( "", 0, $True)
$serviceNowMasterSheetIncidentExcelFileObject.Selection.Copy()
$selection.PasteSpecial($default, $false, 0, $false, 9, $default,$default)
$selection.EndKey(6)
$selection.TypeParagraph()
$selection.TypeParagraph()
$Selection.ParagraphFormat.LeftIndent =  $word.InchesToPoints(0)
$Selection.ParagraphFormat.SpaceBeforeAuto = $False
$Selection.ParagraphFormat.SpaceAfterAuto = $False

$serviceNowMasterSheetIncidentExcelWorkbook.activate()
$serviceNowTocusIncidentMasterSheetExcelWorksheet.Activate()
Start-Sleep -m 3

$pivotTable = $serviceNowTocusIncidentMasterSheetExcelWorksheet.PivotTables("PivotTable7")
$pivotTable.TableRange2.Offset(1, 0).Copy
$pivotTable.PivotSelect( "", 0, $True)
$serviceNowMasterSheetIncidentExcelFileObject.Selection.Copy()
$selection.PasteSpecial($default, $false, 0, $false, 9, $default,$default)
$doc.Save()  #"C:\Users\smadhra\Desktop\Dupoint\suresh.docx")
$pdfPath ="C:\Users\smadhra\Desktop\Dupoint\Bionix data1.pdf"
$wdExportFormatPDF = 17
# Export the PDF file and close without saving a Word document
$doc.ExportAsFixedFormat($pdfPath,$wdExportFormatPDF) 
$doc.Close()
$word.Quit()
$serviceNowMasterSheetIncidentExcelFileObject.Quit()
