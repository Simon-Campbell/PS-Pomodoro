# This script depends on the Lync 2013 SDK. It can be
# downloaded from here:
#	http://www.microsoft.com/en-nz/download/details.aspx?id=36824
Function global:Complete-Pomodoro { # Global so that it can be seen by action block
	$AvailableAvailabilityId = 3000
	
	Write-Host "[Pomodoro] Completed"
	
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
	$DoNotDisturbAvailabilityId = 9000 # Not quite over 9000
	
	$pomodoro = New-Object Timers.Timer

	$pomodoro.Interval	= 2500
	$pomodoro.AutoReset = $false
	$pomodoro.Enabled	= $true

	Write-Host "[Pomodoro] Started"

	$lync = [Microsoft.Lync.Model.LyncClient]::GetClient()
	$contactInfo = New-Object 'System.Collections.Generic.Dictionary[Microsoft.Lync.Model.PublishableContactInformationType, object]'
	$contactInfo.Add([Microsoft.Lync.Model.PublishableContactInformationType]::Availability, $DoNotDisturbAvailabilityId)
	$contactInfo.Add([Microsoft.Lync.Model.PublishableContactInformationType]::PersonalNote, "In a Pomodoro sprint")
	
	$publish = $lync.Self.BeginPublishContactInformation($contactInfo, $null, $null)
	$lync.Self.EndPublishContactInformation($publish)
	
	Register-ObjectEvent -InputObject $pomodoro -EventName Elapsed -Action { Complete-Pomodoro }
}

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

Import-RequiredModules

Start-Pomodoro