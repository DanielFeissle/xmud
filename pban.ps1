#pban
#powershell banner
#6-1-2021 v1
#call this powershell .\pban <txt> <keepfile>
#changes
#6-10-2021 v1.1
#Fix display for banner

param([String]$ban="aaa", [int32]$ktxt="0") 
#file to be kept is the temp file, 0 deletes and 1 keeps
#$ban=$args[0]
$pu=0
$ban | ForEach-Object { $ay=$_.ToCharArray() }
for ($i = 0; $i -le 6; $i++) {
	$pu=0
	$ay  | ForEach-Object { 
    $h= $_
    If(!(test-path "$PSScriptRoot\dat\banner\$h.txt"))
    {
        Write-Warning "FILE NOT FOUND: $PSScriptRoot\dat\banner\$h.txt"
        Write-Warning "ATTEMPTED TEXT: $ban"
    }
    else {
		$a=Get-Content "$PSScriptRoot\dat\banner\$h.txt"
		$test=$test+$a[$i]
		$a=Get-Content "$PSScriptRoot\dat\banner\$h.txt"
		if ($ban.Length-1 -eq $pu) {
			#escape the single line here
			$test=$test+"`n"
		}
    }
$pu=$pu+1
	}
}
Write-Output $test
if ($ktxt -eq 1)
{
    Write-Output "$test" > "$PSScriptRoot\temp.txt"
}