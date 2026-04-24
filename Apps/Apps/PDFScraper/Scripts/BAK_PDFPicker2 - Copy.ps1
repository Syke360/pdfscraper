Add-Type -AssemblyName System.Windows.Forms
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

$baseDir = "C:\Apps\PDFScraper"
$outputBase = Join-Path $baseDir "Output"
$tempDir = Join-Path $baseDir "Temp"

# --- STAT TRACKERS ---
$totalFound = 0
$totalDownloaded = 0
$totalConverted = 0

# --- STEP 1: NUCLEAR WORD RESET ---
$msg = "This tool requires Microsoft Word to be closed. Is it okay to FORCE CLOSE all Word applications now? (Save your work first!)"
$choice = [System.Windows.Forms.MessageBox]::Show($msg, "Word Reset Required", "YesNo", "Warning")

if ($choice -eq "No") { exit }

Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

# --- STEP 2: SETUP ---
@($outputBase, $tempDir) | ForEach-Object { if (!(Test-Path $_)) { New-Item -ItemType Directory -Path $_ -Force | Out-Null } }

$fb = New-Object System.Windows.Forms.OpenFileDialog
$fb.InitialDirectory = Join-Path $env:USERPROFILE "Downloads"
$fb.Filter = "PDF Files (*.pdf)|*.pdf"

if ($fb.ShowDialog() -eq "OK") {
    $pdfPath = $fb.FileName
    $pdfBaseName = [System.IO.Path]::GetFileNameWithoutExtension($pdfPath)
    $out = Join-Path $outputBase "$($pdfBaseName)_$(Get-Date -Format 'yyyyMMdd_HHmm')"
    New-Item -ItemType Directory -Path $out -Force | Out-Null
    
    $logFile = Join-Path $out "Processing_Log.txt"
    function Write-Log($msg) { 
        $timestamp = Get-Date -Format "HH:mm:ss"
        "[$timestamp] $msg" | Out-File -FilePath $logFile -Append 
        Write-Host "[$timestamp] $msg"
    }

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    try {
        $doc = $word.Documents.Open($pdfPath, $false, $true)
        $wordPid = (Get-Process -Name "WINWORD" | Sort-Object StartTime -Descending | Select-Object -First 1).Id
        
        $allFound = New-Object System.Collections.Generic.List[string]
        foreach ($h in $doc.Hyperlinks) { $allFound.Add($h.Address) }
        $regex = [regex]'https?://[^\s"]+essl\.co\.uk[^\s"]+'
        foreach ($m in $regex.Matches($doc.Content.Text)) { $allFound.Add($m.Value.Trim().TrimEnd('.')) }
        $uniqueLinks = $allFound | Where-Object { $_ -like "*essl.co.uk*" } | Select-Object -Unique
        $totalFound = $uniqueLinks.Count

        if ($totalFound -eq 0) { Write-Log "❌ No links found." }
        else {
            foreach ($url in $uniqueLinks) {
                try {
                    $headers = @{ "User-Agent" = "Mozilla/5.0"; "Referer" = "https://live.inmotion.essl.co.uk/" }
                    $id = if ($url -match "documents/([^/]+)") { $matches[1].Substring(0,8) } else { "File_$(Get-Random)" }
                    $tempFile = Join-Path $tempDir "dl_$id"
                    
                    $response = Invoke-WebRequest -Uri $url -OutFile $tempFile -Headers $headers -UseBasicParsing -ErrorAction Stop -PassThru
                    $contentType = $response.Headers["Content-Type"].ToLower()
                    $targetPdf = [System.IO.Path]::GetFullPath((Join-Path $out "Doc_$($id).pdf"))

                    if ($contentType -match "image") {
                        Write-Log "🖼️ CONVERTING: $id"
                        $newDoc = $word.Documents.Add()
                        $newDoc.InlineShapes.AddPicture($tempFile) | Out-Null
                        $newDoc.ExportAsFixedFormat($targetPdf, 17)
                        $newDoc.Close(0)
                        $totalConverted++
                        $totalDownloaded++
                    } 
                    else {
                        Move-Item $tempFile $targetPdf -Force
                        $totalDownloaded++
                        Write-Log "🆗 Direct PDF: $id"
                    }
                } catch {
                    Write-Log "❌ Link $id failed: $($_.Exception.Message)"
                }
            }
            Invoke-Item $out
        }
    } finally {
        if ($word) { $word.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null }
        if ($wordPid) { Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue }
    }
}

# --- JOB DONE SCOREBOARD ---
Write-Host "`n==============================" -ForegroundColor Cyan
Write-Host "         JOB DONE!" -ForegroundColor Green
Write-Host "==============================" -ForegroundColor Cyan
Write-Host "Links Found:      $totalFound"
Write-Host "Files Downloaded: $totalDownloaded"
Write-Host "Images Converted: $totalConverted"
Write-Host "==============================" -ForegroundColor Cyan
Write-Host "`nPress any key to continue..."
$null = [Console]::ReadKey()
