$ko = "KO33915"
$log = "C:\temp\KO-Automation.csv"
$rundate = Get-Date -Format "yy-MM-dd//HH:mm:ss" 
If (!(Test-Path $log))
{
	new-Item $log -force
}

add-content "$env:computername $ko $rundate" -path $log

