#Excel Encrypt (xenc)
#June 14 2021
#Usage: .\xdec <EXCEL File to be obstrusted> <COLUMN to be muddled>
param([String]$xlsxInput="file.xlsx", [int32]$ktxt="0") 


#hmm in powershell it is tougher to overwrite existing variables, It is applicable but not recommended 
#dynamic variables/overload here
If(!(test-path "$PSScriptRoot\\settings.conf"))
{
    Write-Warning "Settings file does not exist, creating default."
    Write-Output "keep=1" > $PSScriptRoot\settings.conf
}
Write-Output "=====================Loading dynamic variables..."
Get-Content "$PSScriptRoot\\settings.conf" | Foreach-Object{
    if ($_ -Match "#REM") {
Write-Output "Ignore"
    }
     else 
     {
 $var = $_.Split('=')

   New-Variable -Name $var[0] -Value $var[1] -Force
   Write-Output "       Variable Name" $var[0]
   Write-Output "       Variable Value" $var[1]
    }
  
}
Write-Output "------------------------Done loading dynamic variables!"

if (Test-Path ShutDownWatcher) {
    Write-Warning "xenc did not exit cleanly, performing backup cleanup measures"
    Write-Warning "Stopping all Excel processes now"
    Get-Process | Where-Object {$_.Path -like "*excel*"}| Stop-Process -whatif
    Get-Process | Where-Object {$_.Path -like "*excel*"}| Stop-Process
}

#first time startup check
If(!(test-path "$PSScriptRoot\out"))
{
      New-Item -ItemType Directory -Force -Path "$PSScriptRoot\out"
}


Write-Output $null >> ShutDownWatcher #create a file at the begining




Write-Warning   "TEST"

#call banner script here
& .\pban.ps1 "xdec"


 Get-Content "$PSScriptRoot\temp.txt"
Remove-Item "$PSScriptRoot\out\tempFile_dec.xlsx"
$xFile="$PSScriptRoot\out\tempFile_dec.xlsx"
 $ExcelWB = new-object -comobject excel.application
 Write-Output "Converting to xlsx"
 Get-ChildItem -Path "$PSScriptRoot\out" -Filter "*.xlsx" | ForEach-Object{
     Write-Output "$_"
     $Workbook = $ExcelWB.Workbooks.Open($_.Fullname) 
     $worksheet = $Workbook.sheets.item("Sheet1")
     $countUsed = $worksheet.UsedRange.Rows.Count
     Write-Output "INTEROP counts this many rows: $countUsed"
     $countColumns = $worksheet.UsedRange.Columns.Count
     Write-Output "INTEROP counts this many columns: $countColumns"
   # Write-Output "">key.txt
     for ($i = 1; $i -le $countUsed; $i++) {
         $uid=New-Guid


         #if ( $ws.Cells.Item(4, $i).Value2 -eq "sort" ) {
         #$tc= $worksheet.Cells.Item($i, 4).Value2 #get the tc#
		 $tc = (Get-Content "$PSScriptRoot\key.txt")[$i]
if ($tc -notmatch '........-....-....-....-............')
{
	if ( $tc -ne $null)
	{
	   $uuid=$worksheet.Cells.Item($i, 4).Value2
	   $worksheet.Cells.Item($i, 4).Value2="$tc"
	}




   # echo "$i" "$tc"
	}
	else {
		if ($tc -ne $null)
		{
			$uuid=$worksheet.Cells.Item($i, 4).Value2
			$WorkSheet.Columns.Replace("$uuid","$tc") | out-null

		}

	}
}

     $Workbook.SaveAs("$xFile")
     #$Workbook.SaveAs("$xFile", 6) #6 is for xlsx
     $Workbook.Close($false)
 }


 $ExcelWB.quit()


 #clean exit, we don't want no excel processes building up and potentially crashing
 [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook)
 [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelWB)
 [GC]::Collect()


##wrap up and finish with excel file here
#Write-Output ""
#$wb.SaveAs($xlsxInput+"_new",6)
#$wb.Close($false)
# $ExcelWB.quit()
#clean exit, we don't want no excel processes building up and potentially crashing
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
#[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
###[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelWB)
[GC]::Collect()

Remove-Item ShutDownWatcher #remove the watch file
#Remove-Item $destinationExcel\\*.xlsx
#Rename-Item -Path ($xlsxInput+"_new") -NewName "$xlsxInput"


Write-Output "xlsx file created"
exit 0