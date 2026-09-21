$lines = Get-Content "c:\ApProject\APOffice\TempDecompiledFull\APOffice\frmGesVariazioni.cs"
$outLines = @()
foreach ($line in $lines) {
    if ($line -match "protected override void Dispose") {
        break
    }
    $outLines += $line
}
$outLines = $outLines -replace "public class frmGesVariazioni : Form", "public partial class frmGesVariazioni : Form"
$outLines += "}"
$outLines | Set-Content "c:\ApProject\APOffice\frmGesVariazioni.cs"
