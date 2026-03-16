
try {
    $obj = New-Object -ComObject "QSBar.WpsAddIn"
    Write-Host "SUCCESS: COM object created successfully!"
    $obj = $null
} catch {
    Write-Host "ERROR: Failed to create COM object."
    Write-Host $_.Exception.Message
    Write-Host $_.Exception.InnerException.Message
}
