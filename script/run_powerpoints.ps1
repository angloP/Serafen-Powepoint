# Runnable script for running the powerpoint presentations
#powershell -ExecutionPolicy ByPass -File script.ps1

$global:file_extension=".pptx"
$global:default_ppt="Default"
$global:relative_powerpoint_folder="\..\powerpoint\"
$global:sleep_time=60
$global:old_ppt = ""

Function Open-Powerpoint() {
	$app = New-Object -ComObject powerpoint.application
	$app.visible = $TRUE
	return $app;
} 

Function Close-Powerpoint($app) {
	$app.quit()
	$app = $null
	Garbage-Collect
	Stop-Process -name "POWERPNT"
}

Function Open-Presentation($ppt, $app) {
	$presentation = $app.Presentations.open($ppt)
	$presentation.SlideShowSettings.Run()                         
	return $presentation
}

Function Close-Presentation($presentation) {
	$presentation.Close()
}

Function Garbage-Collect() {
	[GC]::collect()
	[GC]::WaitForPendingFinalizers()
}

Function Date(){
	return (Get-Date -UFormat "%Y-%m-%d")
}

Function Powerpoint-Path($path, $date){
	return ($path + $date + $global:file_extension)
}

Function Get-Powerpoint-Folder($script_path){
	return ($script_path + $global:relative_powerpoint_folder)
}




Write-Host "*********Serafens Powerpoint Runner*********"
Write-Host "Running..."
$app = Open-Powerpoint
$script_path=Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$path=Get-Powerpoint-Folder $script_path

$previous_date=""

$isOpen = $FALSE
$isDefaultOpen = $FALSE

while($TRUE) {
	$date=Date
	
	IF(!($previous_date -eq $date)) {
		# There is a new date
		$previous_date = $date
		# Need to close old presentation and open new
		$ppt = Powerpoint-Path $path $date
		IF(Test-Path $ppt) {
			# Todays presentation existed
			IF($isOpen) {
				# Need to close the open presentation
				Write-Host ("Closing " + $global:old_ppt)
				Close-Powerpoint $app
				$isOpen = $FALSE
			}
			# Open new presentation
			$app = Open-Powerpoint
			$pres = Open-Presentation $ppt $app
			Write-Host ("Opening " + $ppt)
			$isOpen = $TRUE
			$isDefaultOpen = $FALSE
		} else {
			IF($isDefaultOpen) {
				# Do nothing
				Write-Host ("Still no presentation exists for today")
			} else {
				IF($isOpen) {
					# Need to close the open presentation
					Write-Host ("Closing " + $global:old_ppt)
					Close-Powerpoint $app
					$isOpen = $FALSE
				}
				# Open default presentation
				Write-Host ("No presentation exists for today")
				$app = Open-Powerpoint
				$ppt = Powerpoint-Path $path $global:default_ppt
				$pres = Open-Presentation $ppt $app
				Write-Host ("Opening " + $ppt)
				$isDefaultOpen = $TRUE
				$isOpen = $TRUE
			}
		}
		$global:old_ppt = $ppt
	}
	sleep $global:sleep_time
}


