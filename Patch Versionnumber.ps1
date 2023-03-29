param(
[string] $Version
)
$QuotedVersion = '"' + $Version + '"'

((Get-Content .\PDCVersion\Properties\AssemblyInfo.cs) -replace '"\d+.\d+.\d+.\d+"', $QuotedVersion) | Set-Content .\PDCVersion\Properties\AssemblyInfo.cs
((Get-Content info.xml) -replace '\d+.\d+.\d+.\d+', $Version) | Set-Content info.xml
