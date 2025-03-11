# Paths to check
$RegistryChecks = @(
    @{ Path = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"; Name = "PackagerPrompt"; ExpectedValue = 2 },
    @{ Path = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security"; Name = "PackagerPrompt"; ExpectedValue = 2 },
    @{ Path = "HKCU:\Software\Microsoft\Office\16.0\Word\Security"; Name = "PackagerPrompt"; ExpectedValue = 2 }
)

$Compliant = $true

foreach ($Check in $RegistryChecks) {
    if (Test-Path $Check.Path) {
        $Value = Get-ItemProperty -Path $Check.Path -Name $Check.Name -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $Check.Name -ErrorAction SilentlyContinue
        if ($Value -ne $Check.ExpectedValue) {
            Write-Host "Non-compliant value found at $($Check.Path)\$($Check.Name): $Value"
            $Compliant = $false
        }
    }
    else {
        Write-Host "Non-compliant: Path $($Check.Path) does not exist"
        $Compliant = $false
    }
}

if ($Compliant) {
    Write-Host "Compliant"
    exit 0
}
else {
    Write-Host "Non-Compliant"
    exit 1
}
