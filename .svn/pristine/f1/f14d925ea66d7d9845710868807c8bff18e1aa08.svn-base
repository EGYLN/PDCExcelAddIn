param(
[string] $DeploymentDirectory,
[string] $buildnumber,
[string] $Stage,
[string] $language,
[string] $WorkingDirectory,
[string] $ApplicationId
)
$WorkingFolder = $Stage + '_' + $language

$DeployFolder = Join-Path $DeploymentDirectory $buildnumber
if (!(Test-Path $DeployFolder)) {
	New-Item -ItemType Directory $DeployFolder
}

$stagepath = (Join-Path $WorkingDirectory $WorkingFolder)
	
New-Item -ItemType Directory -Path $stagepath

$LanguageDir = Join-Path $stagepath 'de-DE'
$ResourceDir = Join-Path $stagepath 'Resources'

New-Item -ItemType Directory -Path $LanguageDir
New-Item -ItemType Directory -Path $ResourceDir

Copy-Item PDCLib\libs\*.dll -Destination $stagepath
Copy-Item PDCLib\libs\AutoUpdater.* $stagepath
Copy-Item PDCLib\libs\autoupdater_instructions.xml $stagepath
Copy-Item PDCLib\bin\Release\PDCLib.dll -Destination $stagepath
Copy-Item PDCLib\bin\Release\de-DE\PDCLib.resources.dll -Destination $LanguageDir
Copy-Item PDCSwitcher\PDCSwitcher\bin\Release\PDCSwitcher.exe -Destination $stagepath
Copy-Item PDCSwitcher\PDCSwitcher\bin\Release\PDCSwitcher.exe.config -Destination $stagepath
Copy-Item PDCUpdater\bin\Release\PdcUpdater.dll -Destination $stagepath
Copy-Item PDCVersion\bin\Release\PDCVersion.dll -Destination $stagepath
Copy-Item OpenLib\bin\Release\Openlib.dll -Destination $stagepath
Copy-Item OpenLib\bin\Release\Openlib.tlb -Destination $stagepath
Copy-Item Molfile2Clipboard\bin\Release\*.exe* -Destination $stagepath
Copy-Item PDCExcelAddIn\bin\Release\PDCExcelAddIn.dll* -Destination $stagepath
Copy-Item PDCExcelAddIn\bin\Release\PDCExcelAddIn.tlb -Destination $stagepath
Copy-Item PDCExcelAddIn\bin\Release\PDCExcelAddIn.vsto -Destination $stagepath
Copy-Item PDCExcelAddIn\bin\Release\de-DE\PDCExcelAddIn.resources.dll -Destination $LanguageDir
Copy-Item PDCExcelAddIn\bin\Release\PDC.log4net -Destination $stagepath
Copy-Item PDCExcelAddIn\Resources\confidential.png -Destination $ResourceDir
Copy-Item SetSecurity\bin\Release\SetSecurity.dll -Destination $stagepath
Copy-Item PDCExcelAddInSetup\*.cer -Destination $stagepath
Copy-Item PDCExcelAddInSetup\CustomAtomColor.properties $stagepath
Copy-Item PDCExcelAddInSetup\isislib.dll $stagepath

$PropertyFileName = 'PDCExcelAddInSetup\Config\pdcconfig.properties.' + $Stage 
$TargetPropertyFileName = Join-Path $stagepath 'pdcconfig.properties';
Copy-Item $PropertyFileName -Destination $TargetPropertyFileName
$NewLanguageProperty = 'Language=' + $language
((Get-Content $TargetPropertyFileName) -replace'^Language=.*', $NewLanguageProperty) | Set-Content $TargetPropertyFileName

$AutoUpdaterIni  = Join-Path $stagepath 'AutoUpdater.Settings.ini'
((Get-Content $AutoUpdaterIni) -replace'^ApplicationID:.*', ('ApplicationID:' + $ApplicationID))  | Set-Content $AutoUpdaterIni

$ZipFileName = 'updates_' + $Stage + '.zip';

$TargetDirectory = Join-Path $DeployFolder ($stage + '_' + $language)


New-Item -ItemType Directory -Path $TargetDirectory
Add-Type -assembly "system.io.compression.filesystem"
[io.compression.zipfile]::CreateFromDirectory($stagepath, (Join-Path $TargetDirectory $ZipFileName))
Copy-Item info.xml -Destination $TargetDirectory

Copy-Item PDCExcelAddInSetup\pdc_vsto*.reg $TargetDirectory
Copy-Item PDCExcelAddInSetup\unregister*.reg $TargetDirectory
Copy-Item PDCExcelAddInSetup\InstallationReadme.txt $TargetDirectory