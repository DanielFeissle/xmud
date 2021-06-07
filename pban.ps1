#pban
#powershell banner
#6-1-2021 v1
#call this powershell .\pban <txt> <keepfile>
param([String]$ban="aaa", [int32]$ktxt="0") 
#file to be kept is the temp file, 0 deletes and 1 keeps
#$ban=$args[0]
$pu=0
Write-Output "|" "|" "|" "|" "|" "|" "|"> "$PSScriptRoot\temp.txt"
$ban | ForEach-Object { $ay=$_.ToCharArray() }
$ay  | ForEach-Object { 
    #$du=$_
    $h= $_
    If(!(test-path "$PSScriptRoot\dat\banner\$h.txt"))
    {
        Write-Warning "FILE NOT FOUND: $PSScriptRoot\dat\banner\$h.txt"
        Write-Warning "ATTEMPTED TEXT: $ban"

    }
    else {
        $a=Get-Content "$PSScriptRoot\dat\banner\$h.txt"
        $fileContent = Get-Content "$PSScriptRoot\temp.txt"
        for ($i = 0; $i -le 6; $i++) {
        
            $fileContent[$i] += $a[$i]
        
        
        }
        $fileContent | Set-Content "$PSScriptRoot\temp.txt"

    }
$pu=$pu+1
}
Get-Content "$PSScriptRoot\temp.txt"
if ($ktxt -eq 0)
{
    Remove-Item "$PSScriptRoot\temp.txt"
}