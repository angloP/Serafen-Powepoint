# Installation Script

Write-Host "Installing....."

$dir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

$install_dir = $dir + "\.."

$icons_dir = $install_dir + "\icons"

$w = New-Object -ComObject WScript.Shell

$desktop = [system.environment]::GetFolderPath("Desktop")

$desk_link = $w.CreateShortcut("$desktop\Inedal_Powerpoint_Menu.lnk")
$desk_link.TargetPath = 'powershell.exe' 
$desk_link.arguments = ' -ExecutionPolicy ByPass -file ' + $dir + "\run_powerpoints.ps1"
$desk_link.workingDirectory = $dir
$desk_link.IconLocation = $icons_dir + "\serafen.ico"
$desk_link.save() > $null

$link = $w.CreateShortcut("$install_dir\Inedal_Powerpoint_Menu.lnk")
$link.TargetPath = 'powershell.exe' 
$link.arguments = ' -ExecutionPolicy ByPass -file ' + $dir + "\run_powerpoints.ps1"
$link.workingDirectory = $dir
$link.IconLocation = $icons_dir + "\serafen.ico"
$link.save() > $null