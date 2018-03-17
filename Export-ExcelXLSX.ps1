#Author: Adán Bucio

Add-Type -Path ([System.IO.Path]::Combine($PSScriptRoot, 'FastMember.Signed.dll'));
Add-Type -Path ([System.IO.Path]::Combine($PSScriptRoot, 'DocumentFormat.OpenXml.dll'));
Add-Type -Path ([System.IO.Path]::Combine($PSScriptRoot, 'ExcelNumberFormat.dll'));
Add-Type -Path ([System.IO.Path]::Combine($PSScriptRoot, 'ClosedXML.dll'));

function Export-ExcelXLSX
{ 
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)]
            [PSObject[]]$InputObject,
        [Parameter(Mandatory=$true)]
            [string]$Path,
        [Parameter(Mandatory=$false)]
            [string]$SheetName='PsOutput',
        [Parameter(Mandatory=$false)]
            [switch]$AppendSheet
    )
 
    begin 
    {
        [ClosedXML.Excel.XLWorkbook]$workBook = $null;
        [ClosedXML.Excel.IXLWorksheet]$workSheet = $null;

        try{
            if([System.IO.Path]::GetExtension($Path) -ne '.xlsx') {
                throw New-Object System.IO.FileFormatException -ArgumentList 'File not supported'
            }
            $workBook = if([System.IO.File]::Exists($Path) -and $AppendSheet) {
                            New-Object ClosedXML.Excel.XLWorkbook -ArgumentList $Path;
                        } else { New-Object ClosedXML.Excel.XLWorkbook };
        }
        catch {
            $errorID = 'ClosedXML';
            $errorCategory = [Management.Automation.ErrorCategory]::NotSpecified;
            $target = 'ClosedXML.Excel.XLWorkbook';
            $errorRecord = New-Object Management.Automation.ErrorRecord $_.Exception, $errorID, $errorCategory, $target;
            $PSCmdlet.ThrowTerminatingError($errorRecord);
        }
        if([string]::IsNullOrWhiteSpace($SheetName)){
            $SheetName = 'PsOutput'
        }
        
        $rowCount = 0;
        $cellCount = 1;
    } 

    process 
    {
        foreach ($object in $InputObject) 
        { 
            #If this is an array of datatables...
            if($object.GetType().ToString() -eq 'System.Data.DataTable') {
                
                $TmpSheetName = ($object.DataSet.DataSetName, $object.TableName, $SheetName -ne '')[0];
                $sheetCount = ($workbook.Worksheets | ? {
                        $_.Name.StartsWith($TmpSheetName, [System.StringComparison]::InvariantCultureIgnoreCase)
                    }).Count;
                $TmpSheetName += if($sheetCount -gt 0) {"`_$($sheetCount + 1)" }
                $workbook.Worksheets.Add($object, $TmpSheetName) >> $null;
            } 
            else {
                $cix = 0;

                if(++$rowCount -eq 1) {
                    $workSheet = $workBook.Worksheets.Add($SheetName);
                    
                    foreach($p in $object.psobject.Properties) {
                        $workSheet.Cell(1, ++$cix).Value = $p.Name;
                    }
                    $cellCount = $cix;
                    $cix = 0;
                    $rowCount++;
                }

                foreach($p in $object.psobject.Properties) {
                    $cell = $workSheet.Cell($rowCount, ++$cix);
                    $cell.Value = $p.Value;
                    $cell.Style.Alignment.WrapText = $false;
                }
            }
        }
    }
    
    end {
        if($workbook -ne $null) {
            if($workSheet -ne $null) {
                $rng = $workSheet.Range(1, 1, $rowCount, $cellCount);
                $tbl = $rng.CreateTable();
            }
            $workBook.SaveAs($Path);
            $workbook.Dispose();
        }
    }
}