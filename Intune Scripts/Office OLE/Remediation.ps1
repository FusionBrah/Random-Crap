# Registry settings to enforce
$RegistrySettings = @(
    @{ Path = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"; Name = "PackagerPrompt"; Value = 2 },
    @{ Path = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security"; Name = "PackagerPrompt"; Value = 2 },
    @{ Path = "HKCU:\Software\Microsoft\Office\16.0\Word\Security"; Name = "PackagerPrompt"; Value = 2 }
)

foreach ($Setting in $RegistrySettings) {
    # Create key if missing
    if (-Not (Test-Path $Setting.Path)) {
        New-Item -Path $Setting.Path -Force | Out-Null
        Write-Host "Created missing registry path: $($Setting.Path)"
    }

    # Set the value
    New-ItemProperty -Path $Setting.Path -Name $Setting.Name -Value $Setting.Value -PropertyType DWORD -Force | Out-Null
    Write-Host "Set $($Setting.Name) to $($Setting.Value) in $($Setting.Path)"
}

Write-Host "Remediation complete."
exit 0
