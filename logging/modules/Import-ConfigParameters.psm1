function Import-ConfigParameters {
    param (
        [Parameter(Mandatory = $true)][string]$ConfigPath,
        [Parameter(Mandatory = $true)][object]$BoundParameters
    )
    if (Test-Path $ConfigPath) {
        $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        foreach ($property in $config.PSObject.Properties) {
            if (-not $BoundParameters.ContainsKey($property.Name)) {
                Set-Variable -Name $property.Name -Value $property.Value -Scope 1
            }
        }
    }
}
Export-ModuleMember -Function Import-ConfigParameters