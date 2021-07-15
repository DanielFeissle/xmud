#Excel Encrypt (xenc)
#May 27 2021
#Usage: .\xenc <EXCEL File to be obstrusted> <COLUMN to be muddled>
param([String]$xlsxInput="file.xlsx",[int32]$colval="4",[int32]$cpos="14", [int32]$ktxt="0") 
$ErrorActionPreference = "Stop"
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


If(!(test-path "$PSScriptRoot\\keys"))
{
      New-Item -ItemType Directory -Force -Path "$PSScriptRoot\\keys"
}



#call banner script here
& .\pban.ps1 "xenc"




#################
 $ExcelWB = new-object -comobject excel.application
 Write-Output "Converting to xlsx"
 $sel=Get-ChildItem -Path $PSScriptRoot -Filter "$xlsxInput" 
 if ($null -eq $sel )
{
	Remove-Item ShutDownWatcher #remove the watch file
	Write-Warning "No valid xlsx files found in directory: $pwd"
	Write-Warning "File may not be present or file may already be encrypted"

}
else {
	Get-ChildItem -Path $PSScriptRoot -Filter "$xlsxInput" | ForEach-Object{
		$valName=$_
		If((test-path "$PSScriptRoot\\$valName.temp"))
		   {
			   Remove-Item "$PSScriptRoot\\$valName.temp"
		   }
   
   
	   $Workbook = $ExcelWB.Workbooks.Open($_.Fullname) 
	   for ($i = 1; $i -le $Workbook.Sheets.Count; $i++)
		  {
			  $errorVal=0
			 $worksheet = $Workbook.Sheets.Item($i)
			   $countUsed = $worksheet.UsedRange.Rows.Count
			   Write-Output "INTEROP counts this many rows: $countUsed"
			   $countColumns = $worksheet.UsedRange.Columns.Count
			   Write-Output "INTEROP counts this many columns: $countColumns"
			   #op-prdat
				# for ($zx = 1; $zx -le $countColumns; $zx++)
				# {
				# 	$worksheet.Cells.Item($countUsed+1, $zx).Value2="x"
				# }
				# $countUsed=$countUsed+1
				#eop
			   $xx=$countUsed+$countColumns
			   $PM=Get-Random -Maximum $xx -SetSeed $colval
			   $colRef=($PM + $colval) * ($countUsed + $countColumns+$cpos)
			   if (!(test-path "$PSScriptRoot\keys\$sel.$colRef.exf"))
			   {
	
			  Write-Output "">"keys\\$valName.$colRef.exf"
			  Remove-Item "keys\\$valName.$colRef.exf"
			#fix potential issues if seperated by line breaks in column
			  Write-Output "STANDBY..."
			  	for ($zx = 1; $zx -lt $countUsed; $zx++)
				{
					if ($worksheet.Cells.Item($zx, $colval).Value2 -ne $null)
					{
						$fud=$worksheet.Cells.Item($zx, $colval).Value2
					
						if ($fud -ne "")
						{
							if ($fud.Contains("`n"))
							{
								$fud=$fud -replace "`n","_AWK523DRFBREAKER14_" #| Out-Null
								#write back now
								$worksheet.Cells.Item($zx, $colval).Value2=$fud
							}
						}
					}

				}

			  Write-Output "Extract column for faster processing"
			  $testCol=$WorkSheet.Columns($colval)
			  $testV=($testCol[1].Value2 -split '\r?\n').Trim()
			  #$ft=$testV | Group-Object -AsHashTable -AsString
			 $order= 0..$countUsed | Get-Random -SetSeed ($xx+$colval) -Count $testV.Length
			# $order
			# $array = New-Object 'object[]' $countUsed
			   for ($z = 0; $z -le $countUsed; $z++) {
				   $uid=New-Guid
				   $tc=$testV[$order[$z]]
				   
				  # echo "THIS IS" $testV[$i]
				  # if ( $tc -ne $null)
				   #{
					   #if ($tc -ne "")
					   #{
						#to keep values hidden
						if ($tc -eq "")
						{
							$tc="bean$z"
						}


						
						$ff="$tc" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
						#end
						$tc=$ff

						   Write-Output "$tc">>"keys\\$valName.$colRef.exf"
						 $drf=$order[$z]
					#	 echo "WHAT IS THIS $drf act val $tc  num in sheet $z     te $te"
						# $array[$drf]=$tc
						  $testV[$drf]="$uid"
						  #$df=$ft["$tc"].Count
						  #7-1-2021-disabled due to issues with xedc combined
						###  $WorkSheet.Columns.Replace("$tc","$uid") | out-null
						  $te = $te + 1
					   #} 
				   
				   
				   #}
			   
				  # echo "$i" "$tc"
			  ##	Write-Host -NoNewline "."
				   }
				  # $array
				  # $array | Out-File -Append "keys\\$valName.$colRef.exf"
				   #$testCol.Value2=$testV
				   Write-Host ""
				   Write-Host -NoNewline "Toss column back in,"
				   for ($t=1; $t -le $countUsed; $t++)
				   {
					  $worksheet.Cells.Item($t, $colval).Value2=$testV[$t-1]
				   }
				   Write-Host -NoNewline " Done."
				   Write-Host ""
				}
				else {
					$errorVal=1
					break
				}
		  #echo "FINAL RESULT" $testV "END"
			   #	 $WorkSheet.Columns(4)= $testCol.Value2
		  }

		  if ( $errorVal -eq 1 )
		  {
			Write-Error "File may already be encrypted. Aborting encryption. Refer to the keys folder."
			$ExcelWB.quit()
			[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($testCol)
			[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
			[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook)
			[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelWB)
			[GC]::Collect()
			Remove-Item ShutDownWatcher #remove the watch file
			exit 1
		  }
		$randHour=  Get-Random -Maximum 23
		$randMin=Get-Random -Maximum 59
		$randSec=Get-Random -Maximum 59
		$yestDay=(get-date (get-date).addDays(-1) -UFormat "%m/%d/%y")
		(Get-Item "keys\\$valName.$colRef.exf").LastWriteTime = "$yestDay ${randHour}:${randMin}:${randSec}"
	   $Workbook.SaveAs("$PSScriptRoot\$valName.temp")
	   #$Workbook.SaveAs("$xFile", 6) #6 is for xlsx
	   $Workbook.Close($false)
   
   
   
   
	}
	$ExcelWB.quit()


	#clean exit, we don't want no excel processes building up and potentially crashing
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($testCol)
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook)
	[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelWB)
	[GC]::Collect()

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
}



