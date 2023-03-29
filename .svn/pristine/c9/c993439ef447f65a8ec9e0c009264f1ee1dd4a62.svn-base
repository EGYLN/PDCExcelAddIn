param(
[string] $Version,
[string] $BuildDirectory
)
cd 
(Get-Content .\PDCVersion\Properties\AssemblyInfo.cs).Replace('"\d+.\d+.\d+.\d+"', '"' + $Version +'"') | Set-Content .\PDCVersion\Properties\AssemblyInfo.cs
(Get-Content .\info.xml).Replace('\d+.\d+.\d+.\d+', $Version) | Set-Content $BuildDirectory

