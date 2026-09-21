$files = Get-ChildItem -Path "C:\ApProject\APOffice" -Filter "*.Designer.cs" -Recurse

foreach ($file in $files) {
    # Skip backup files
    if ($file.FullName -like "*.bak") { continue }

    $content = Get-Content $file.FullName -Raw
    $original = $content

    # Safe replace strategy: match exactly 'new System.Windows.Forms.DataGridView()'
    $content = $content -replace "new System\.Windows\.Forms\.DataGridView\(\)", "new APOffice.APDataGridView()"
    
    # Safe replace strategy: match exactly variable declarations
    $content = $content -replace "private System\.Windows\.Forms\.DataGridView (dgv\w+|[a-zA-Z0-9_]+);", "private APOffice.APDataGridView `$1;"

    if ($original -ne $content) {
        Set-Content $file.FullName $content -Encoding UTF8
        Write-Host "Updated $($file.Name)"
    }
}
Write-Host "Global upgrade complete."
