# 測試EMF轉PNG功能
Write-Host "開始測試EMF轉PNG功能..." -ForegroundColor Green

# 讀取EMF Base64數據
$emfBase64 = Get-Content "圖片編碼.txt" -Raw
Write-Host "EMF Base64數據長度: $($emfBase64.Length) 字符" -ForegroundColor Yellow

# 轉換為二進制以驗證EMF格式
try {
    $emfBytes = [System.Convert]::FromBase64String($emfBase64.Trim())
    Write-Host "EMF二進制數據長度: $($emfBytes.Length) bytes" -ForegroundColor Yellow
    
    # 檢查EMF格式標識
    if ($emfBytes.Length -ge 44) {
        $emfSignature = [System.BitConverter]::ToString($emfBytes[40..43])
        Write-Host "EMF格式標識 (40-43字節): $emfSignature" -ForegroundColor Cyan
        
        $isEmf = ($emfBytes[40] -eq 0x20) -and ($emfBytes[41] -eq 0x45) -and ($emfBytes[42] -eq 0x4D) -and ($emfBytes[43] -eq 0x46)
        Write-Host "是否為EMF格式: $isEmf" -ForegroundColor $(if($isEmf){"Green"}else{"Red"})
    }
} catch {
    Write-Host "Base64解碼失敗: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n現在可以使用此EMF數據測試Excel中的圖片轉換功能" -ForegroundColor Green
Write-Host "請上傳包含此EMF圖片的Excel檔案到前端應用程序" -ForegroundColor Yellow