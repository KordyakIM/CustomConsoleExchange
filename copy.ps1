###########################################################
# Разработчик: Кордяк Иван Михайлович kordyakim@gmail.com #
###########################################################
#--------------------------------------------------#
$FolderBrowsers = Get-Process -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Path | Where-Object -FilterScript {$_ -like "*ExConsole.exe*"}
$FolderBrowser = $FolderBrowsers -replace "ExConsole.exe",""
$pathvERSION = "\\SMBPath\vERSION.v"
$version = Get-Content $pathvERSION
	if($FolderBrowser -ne $null) 
	{
		while(Get-Process -Name "ExConsole"){
			Stop-Process -Name "ExConsole" -Force -ErrorAction SilentlyContinue
			sleep 1
		}
		Copy-Item "\\SMBPath\ExConsole.exe" -Destination $FolderBrowser
		Copy-Item "\\SMBPath\ExConsole.exe.config" -Destination $FolderBrowser
        $output = [System.Windows.Forms.MessageBox]::Show("Программа ExConsole обновлена до версии $version")
		Start-Process -filepath $folderbrowser"ExConsole.exe" -ErrorAction SilentlyContinue
		Stop-Process -Name copy -Force -ErrorAction SilentlyContinue
    }