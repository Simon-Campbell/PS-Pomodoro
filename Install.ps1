$modulePath = "$Home\Documents\WindowsPowerShell\Modules\PS-Pomodoro\"

if (-Not (Test-Path $modulePath)) {
	Write-Host "[PS-Pomodoro] Creating module path for installation of PS-Pomodoro"
	Write-Host "[PS-Pomodoro] $modulePath"
	
	New-Item $modulePath -Type Directory
}

Write-Host "[PS-Pomodoro] Copying *.psm1 files to $modulePath"
Copy-Item *.psm1 $modulePath

if (Get-Module "PS-Pomodoro") {
	Write-Host "[PS-Pomodoro] Unloading older version of PS-Pomodoro module"

	Remove-Module "PS-Pomodoro"
}

Write-Host "[PS-Pomodoro] Loading PS-Pomodoro module"
Import-Module "PS-Pomodoro"