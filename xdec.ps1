#Excel Encrypt (xenc)
#June 14 2021
#Usage: .\xdec <EXCEL File to be obstrusted> <COLUMN to be muddled>
param([int32]$colval="4", [String]$xlsxInput="file.xlsx", [int32]$ktxt="0") 


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

###############


# Get-Content "$PSScriptRoot\temp.txt"
#Remove-Item "$PSScriptRoot\out\tempFile_dec.xlsx"
#$xFile="$PSScriptRoot\out\tempFile_dec.xlsx"
#####################
$ExcelWB = new-object -comobject excel.application
Write-Output "Converting to xlsx"
$sel=Get-ChildItem -Path $PSScriptRoot -Filter "*.xlsx" 
$sela=$sel.Name
if (($null -eq $sela) -or !(test-path "$PSScriptRoot\keys\$sela.$colval.exf"))
{
	Remove-Item ShutDownWatcher #remove the watch file
	Write-Warning "No valid xlsx files found in directory: $pwd"
	exit 1
}
else {
Get-ChildItem -Path $PSScriptRoot -Filter "*.xlsx" | ForEach-Object{
	Write-Output "$_"
	$valName=$_
	$Workbook = $ExcelWB.Workbooks.Open($_.Fullname) 
	for ($i = 1; $i -le $Workbook.Sheets.Count; $i++)
	   {
		  $worksheet = $Workbook.Sheets.Item($i)
			$countUsed = $worksheet.UsedRange.Rows.Count
			Write-Output "INTEROP counts this many rows: $countUsed"
			$countColumns = $worksheet.UsedRange.Columns.Count
			Write-Output "INTEROP counts this many columns: $countColumns"
		   #Write-Output "">key.txt
		   Write-Output "Extract column for faster processing"
		   $testCol=$WorkSheet.Columns($colval)
		   $testV=($testCol[1].Value2 -split '\r?\n').Trim()
		   #$ft=$testV | Group-Object -AsHashTable -AsString
			
			for ($i = 0; $i -le $countUsed; $i++) {
				$uid=$testV[$i]
				$tc = (Get-Content "$PSScriptRoot\keys\$sela.$colval.exf")[$i]
				#$tc=$testV[$i]
				#echo "THI " $tc
			   # echo "THIS IS" $testV[$i]
				if ( $tc -ne $null)
				{
					if ($tc -ne "")
					{
						if ($uid -ne "")
						{
						 #  Write-Output "$tc">>key.txt
						   $testV[$i]="$tc"
						   #$df=$ft["$tc"].Count
						   $WorkSheet.Columns.Replace("$uid","$tc") | out-null
						   $te = $te + 1
						}
					} 
				
				
				}
			
			
			
			
			   # echo "$i" "$tc"
		   ##	Write-Host -NoNewline "."
				}
				#$testCol.Value2=$testV
				Write-Host ""
				Write-Host -NoNewline "Toss column back in,"
				for ($t=1; $t -le $countUsed; $t++)
				{
				   $worksheet.Cells.Item($t, $colval).Value2=$testV[$t-1]
				}
				Write-Host -NoNewline " Done."
				Write-Host ""
	   #echo "FINAL RESULT" $testV "END"
			#	 $WorkSheet.Columns(4)= $testCol.Value2
	   }

	$Workbook.SaveAs("$PSScriptRoot\$valName.temp")
	#$Workbook.SaveAs("$xFile", 6) #6 is for xlsx
	$Workbook.Close($false)

}


$ExcelWB.quit()
}

#clean exit, we don't want no excel processes building up and potentially crashing
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($testCol)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelWB)
[GC]::Collect()
Remove-Item "keys\\$valName.$colval.exf"
Remove-Item "$PSScriptRoot\$valName"
Move-Item -Path "$PSScriptRoot\$valName.temp" -Destination "$PSScriptRoot\$valName"
##wrap up and finish with excel file here
#Write-Output ""
#$wb.SaveAs($xlsxInput+"_new",6)
#$wb.Close($false)
# $ExcelWB.quit()
#clean exit, we don't want no excel processes building up and potentially crashing

#[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
###[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelWB)
[GC]::Collect()

Remove-Item ShutDownWatcher #remove the watch file
#Remove-Item $destinationExcel\\*.xlsx
#Rename-Item -Path ($xlsxInput+"_new") -NewName "$xlsxInput"


Write-Output "xlsx file created"
exit 0