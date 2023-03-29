param(
[string] $PdcBaseUrlQA,
[string] $PDCServiceUrlQA
)
$Replacement = $PdcBaseUrlQA + '<'
$ReplacementCs = $PdcBaseUrlQA + '"'
(((Get-Content PDCLib\Properties\Settings.settings) -replace 'http(s?)://.+/PDCService', $PDCServiceUrlQA) -replace 'http(s?)://.+/pdc<', $Replacement) | Set-Content PDCLib\Properties\Settings.settings
(((Get-Content PDCLib\Properties\Settings.Designer.cs) -replace 'http(s?)://.+/PDCService', $PDCServiceUrlQA) -replace 'http(s?)://.+/pdc"', $ReplacementCs) | Set-Content PDCLib\Properties\Settings.Designer.cs
(((Get-Content PDCLib\app.config) -replace 'http(s?)://.+/PDCService<', $PDCServiceUrlQA) -replace 'http(s?)://.+/pdc<', $Replacement) | Set-Content PDCLib\app.config
