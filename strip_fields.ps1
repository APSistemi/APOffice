$lines = Get-Content "c:\ApProject\APOffice\frmGesVariazioni.cs"
$outLines = @()
$skip = $false
foreach ($line in $lines) {
    if ($line -match "private IContainer components = null;") {
        $skip = $true
    }
    if ($line -match "public frmGesVariazioni\(\)") {
        $skip = $false
    }
    if (-not $skip) {
        $outLines += $line
    }
}
$outLines | Set-Content "c:\ApProject\APOffice\frmGesVariazioni.cs"
