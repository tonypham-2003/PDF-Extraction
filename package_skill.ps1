# Package freight-invoice-to-excel skill
$srcPath   = "C:\Users\v1411\AppData\Roaming\Claude\local-agent-mode-sessions\645568c0-af0c-4feb-8cd2-30040d0c76aa\56f45cff-a91d-48ee-a4df-f234fd438df4\local_4250bc65-56ca-4c52-9711-b016805e862d\outputs\freight-invoice-to-excel"
$destSkill = "C:\Users\v1411\AppData\Roaming\Claude\local-agent-mode-sessions\645568c0-af0c-4feb-8cd2-30040d0c76aa\56f45cff-a91d-48ee-a4df-f234fd438df4\local_4250bc65-56ca-4c52-9711-b016805e862d\outputs\freight-invoice-to-excel.skill"

if (-not (Test-Path $srcPath)) {
    Write-Host "ERROR: Not found: $srcPath" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

if (Test-Path $destSkill) { Remove-Item $destSkill -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($srcPath, $destSkill)

Write-Host "Done! File saved to:" -ForegroundColor Green
Write-Host "  $destSkill" -ForegroundColor Cyan
Read-Host "Press Enter to exit"
