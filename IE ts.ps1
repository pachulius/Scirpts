#closing Internet Explorer in case it is running
If (get-process iexplore -ErrorAction SilentlyContinue)
{
	$ieclose = Read-Host "Internet Explorer is running, Is that ok to close it now (Y/N)"
	while ("Y", "N" -notcontains $ieclose)
	{
		$ieclose = Read-Host "Internet Explorer is running, Is that ok to close it now (Y/N)"
	}
	if ($ieclose -eq "N")
	{
		Exit
	}
	if ($ieclose -eq "Y")
	{
		Stop-Process -name iexplore -force
	}
	
}

#Creating logs
$LogFileName = (Get-Date -Format "yyyy-MM-dd HHmm") + ".log"
$logFilePath = "c:\DRV\tools\logs\KO Automation\KO12339"

if (!(Test-Path -path $logFilePath)) { New-Item -ItemType directory -Path $LogFilePath -Force }

#Display Menu Options
function Show-Menu
{
	
	Write-Host "=========Internet Explorer Troubleshooting==============" -foregroundColor Green
	
	Write-Host "1. Fully Clear IE, Set Connection settings, Disable popup blocker and 3rd party addons"
	Write-Host "2. Complete IE Reset (use in case 1st option won't help)"
	Write-Host "3. Exit"
}


do
{
	Show-Menu
	$input = Read-Host "Please make a selection"
	switch ($input)
	{
		
		"1" {
			# Delete internet temporary files, history, cookies and addon data
            Write-Output "Clearing IE..."            
			RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 4351 -wait
			Start-Sleep -Milliseconds 4000
			
            #IE connection settings
            Write-Output "Configuring your Internet Settings..." 
			
			$key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections'
			$data = (Get-ItemProperty -Path $key -Name DefaultConnectionSettings).DefaultConnectionSettings
			$data[8] = 13
			Set-ItemProperty -Path $key -Name DefaultConnectionSettings -Value $data
			Start-Sleep -Milliseconds 4000
			Add-content $LogfilePath$logFileName "This information is logged"
             
            $utp = (Get-ItemProperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -Name AutoConfigURL -ea 0).AutoConfigURL
            if ($utp -ne 'http://config.intranet.mckinsey.com/scripts/autoproxy.pac' )
			{ Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -Name AutoConfigURL -Value 'http://config.intranet.mckinsey.com/scripts/autoproxy.pac'}
			


			
			
			#Clear ssl State
            Write-Output "Clearing SSL State..."
			add-type -AssemblyName microsoft.VisualBasic
			add-type -AssemblyName System.Windows.Forms
			
			$rundll = Start-Process -FilePath "C:\Windows\system32\rundll32.exe" -argumentlist "shell32.dll,Control_RunDLL inetcpl.cpl,,3" -PassThru -windowstyle Hidden
			Start-Sleep 1
			[Microsoft.VisualBasic.Interaction]::AppActivate("Internet Properties")
			[System.Windows.Forms.SendKeys]::SendWait({ ~ })
			Start-Sleep -Milliseconds 500
			If (get-process -name rundll32) { Stop-Process -Id $rundll.id }
			
			#popup blocker
            Write-Output "Disabling popup blocker..."
			$PopupMgr = Get-ItemProperty "Registry::HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\New Windows" "PopupMgr"
			if ($PopupMgr.PopupMgr -eq "no")
			{
				Write-host "Pop-up blocker is disabled already" -ForegroundColor Green
			}
			else
			{
				Set-ItemProperty -Path "Registry::HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\New Windows" -Name PopupMgr -Value "no"
			}
			write-host "Complete" -ForegroundColor Green
			pause
			
            #Disabling 3rd party addons
			
			
		}
		
		"2" {
			
            #IE reset
            Write-Output "Reseting IE..."
			add-type -AssemblyName microsoft.VisualBasic
			add-type -AssemblyName System.Windows.Forms
			
			$res = Start-Process -FilePath "C:\Windows\system32\rundll32.exe" -argumentlist "InetCpl.cpl,ResetIEtoDefaults" -PassThru
			start-sleep 1
			[Microsoft.VisualBasic.Interaction]::AppActivate("Reset Internet Explorer Settings")
			[System.Windows.Forms.SendKeys]::SendWait("{tab}")
			[Microsoft.VisualBasic.Interaction]::AppActivate("Reset Internet Explorer Settings")
			[System.Windows.Forms.SendKeys]::SendWait(" ")
			[Microsoft.VisualBasic.Interaction]::AppActivate("Reset Internet Explorer Settings")
			[System.Windows.Forms.SendKeys]::SendWait("r")
			Start-Sleep 1
			
			While ($res)
			{
			[Microsoft.VisualBasic.Interaction]::AppActivate("Reset Internet Explorer Settings") 
            [System.Windows.Forms.SendKeys]::SendWait("c") 
            Start-Sleep 1 
            $res = Get-Process -id $res.Id -EA SilentlyContinue
			}
            
            #Enabling Vital adons

			Write-host "Complete" -ForegroundColor Green
            pause    
		}
		
		
		
		"3" { Exit }
		
	}
	
	
}
until ($input -eq "3")




