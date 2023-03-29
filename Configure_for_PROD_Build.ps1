param(
[string] $PdcBaseUrlProd,
[string] $PDCServiceUrlProd
)
$Replacement = $PdcBaseUrlProd + '<'
$ReplacementCs = $PdcBaseUrlProd + '"'
(((Get-Content PDCLib\Properties\Settings.settings) -replace 'http(s?)://.+/PDCService', $PDCServiceUrlProd) -replace 'http(s?)://.+/pdc<', $Replacement) | Set-Content PDCLib\Properties\Settings.settings
(((Get-Content PDCLib\Properties\Settings.Designer.cs) -replace 'http(s?)://.+/PDCService', $PDCServiceUrlProd) -replace 'http(s?)://.+/pdc"', $ReplacementCs) | Set-Content PDCLib\Properties\Settings.Designer.cs

(((Get-Content PDCLib\app.config) -replace 'http(s?)://.+/PDCService<', $PDCServiceUrlProd) -replace 'http(s?)://.+/pdc<', $Replacement) | Set-Content PDCLib\app.config
Copy-Item PDCLib\Properties\Settings.settings (Join-Path %DeploymentDirectory% prod.settings)
