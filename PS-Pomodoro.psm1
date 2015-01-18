# This script depends on the Lync 2013 SDK. It can be
# downloaded from here:
#	http://www.microsoft.com/en-nz/download/details.aspx?id=36824

Set-Variable AvailableAvailabilityId -Option Constant -Value 3000
Set-Variable DoNotDisturbAvailabilityId -Option Constant -Value 9000 # Not quite over 9000
	
Function Import-RequiredModules {
	if (-not (Get-Module -Name Microsoft.Lync.Model)) {
		try {
			$module = (Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath "Microsoft Office 2013\LyncSDK\Assemblies\Desktop\Microsoft.Lync.Model.dll")
			
			Write-Host "Importing Microsoft.Lync.Model from $module"
			Import-Module -Name $module -ErrorAction Stop
			
		} catch {
			Write-Warning "Microsoft.Lync.Model not available. Download and install it from http://www.microsoft.com/en-nz/download/details.aspx?id=36824"
			
			break
		}
	}
}

Function Complete-Pomodoro {
	if (-Not ($script:activePomodoro)) {
		Write-Host "[Pomodoro] No active pomodoro to complete"
		
		return
	}
	
	Import-RequiredModules
	
	Write-Host "[Pomodoro] Completed"
	$script:activePomodoro.Enabled = $false
	$script:activePomodoro = $null
	
	$lync = [Microsoft.Lync.Model.LyncClient]::GetClient()
	$contactInfo = New-Object 'System.Collections.Generic.Dictionary[Microsoft.Lync.Model.PublishableContactInformationType, object]'
	$contactInfo.Add([Microsoft.Lync.Model.PublishableContactInformationType]::Availability, $AvailableAvailabilityId)
	$contactInfo.Add([Microsoft.Lync.Model.PublishableContactInformationType]::PersonalNote, "...")
	
	$publish = $lync.Self.BeginPublishContactInformation($contactInfo, $null, $null)
	$lync.Self.EndPublishContactInformation($publish)
	
	$notification = New-Object System.Media.SoundPlayer
	$notification.SoundLocation = "C:\Windows\Media\notify.wav"
	$notification.PlayLooping()
	
	1..5 | foreach {
		Sleep -S 1
	}
	
	$notification.Stop()
}

Function Start-Pomodoro {
	Import-RequiredModules
	
	$pomodoro = New-Object Timers.Timer

	$pomodoro.Interval	= 2500
	$pomodoro.AutoReset = $false
	$pomodoro.Enabled	= $true

	$script:activePomodoro = $pomodoro
	
	Write-Host "[Pomodoro] Started"

	$lync = [Microsoft.Lync.Model.LyncClient]::GetClient()
	$contactInfo = New-Object 'System.Collections.Generic.Dictionary[Microsoft.Lync.Model.PublishableContactInformationType, object]'
	$contactInfo.Add([Microsoft.Lync.Model.PublishableContactInformationType]::Availability, $DoNotDisturbAvailabilityId)
	$contactInfo.Add([Microsoft.Lync.Model.PublishableContactInformationType]::PersonalNote, "In a Pomodoro sprint")
	
	$publish = $lync.Self.BeginPublishContactInformation($contactInfo, $null, $null)
	$lync.Self.EndPublishContactInformation($publish)
	
	Register-ObjectEvent -InputObject $pomodoro -EventName Elapsed -Action { Complete-Pomodoro }
}

Export-ModuleMember -Function Start-Pomodoro
Export-ModuleMember -Function Complete-Pomodoro