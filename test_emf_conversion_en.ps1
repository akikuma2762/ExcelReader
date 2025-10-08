# Test EMF to PNG conversion functionality
Write-Host "Starting EMF to PNG conversion test..." -ForegroundColor Green

# Read EMF Base64 data
$emfBase64 = Get-Content "圖片編碼.txt" -Raw
Write-Host "EMF Base64 data length: $($emfBase64.Length) characters" -ForegroundColor Yellow

# Convert to binary to verify EMF format
try {
    $emfBytes = [System.Convert]::FromBase64String($emfBase64.Trim())
    Write-Host "EMF binary data length: $($emfBytes.Length) bytes" -ForegroundColor Yellow
    
    # Check EMF format signature
    if ($emfBytes.Length -ge 44) {
        $emfSignature = [System.BitConverter]::ToString($emfBytes[40..43])
        Write-Host "EMF format signature (bytes 40-43): $emfSignature" -ForegroundColor Cyan
        
        $isEmf = ($emfBytes[40] -eq 0x20) -and ($emfBytes[41] -eq 0x45) -and ($emfBytes[42] -eq 0x4D) -and ($emfBytes[43] -eq 0x46)
        Write-Host "Is EMF format: $isEmf" -ForegroundColor $(if($isEmf){"Green"}else{"Red"})
    }
} catch {
    Write-Host "Base64 decode failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Now you can use this EMF data to test image conversion in Excel" -ForegroundColor Green
Write-Host "Please upload an Excel file containing this EMF image to the frontend application" -ForegroundColor Yellow