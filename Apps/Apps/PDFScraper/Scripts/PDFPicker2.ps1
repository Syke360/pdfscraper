Add-Type -AssemblyName System.Windows.Forms
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

$baseDir = "C:\Apps\PDFScraper"
$outputBase = Join-Path $baseDir "Output"
$tempDir = Join-Path $baseDir "Temp"

# Function to encapsulate the core processing logic
function Start-PDFProcessing {
    param([string]$pdfPath)

    $totalFound = 0
    $totalDownloaded = 0
    $totalConverted = 0

    @($outputBase, $tempDir) | ForEach-Object { if (!(Test-Path $_)) { New-Item -ItemType Directory -Path $_ -Force | Out-Null } }

    $pdfBaseName = [System.IO.Path]::GetFileNameWithoutExtension($pdfPath)
    $out = Join-Path $outputBase "$($pdfBaseName)_$(Get-Date -Format 'yyyyMMdd_HHmm')"
    New-Item -ItemType Directory -Path $out -Force | Out-Null
    
    $logFile = Join-Path $out "Processing_Log.txt"
    function Local-Write-Log($msg) { 
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

        if ($totalFound -eq 0) { Local-Write-Log "❌ No links found." }
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
                        Local-Write-Log "🖼️ CONVERTING: $id"
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
                        Local-Write-Log "🆗 Direct PDF: $id"
                    }
                } catch {
                    Local-Write-Log "❌ Link $id failed: $($_.Exception.Message)"
                }
            }
        }
    } finally {
        if ($word) { $word.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null }
        if ($wordPid) { Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue }
    }

    Write-Host "`n==============================" -ForegroundColor Cyan
    Write-Host "         FILE COMPLETE" -ForegroundColor Green
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host "Processed:        $pdfBaseName"
    Write-Host "Links Found:      $totalFound"
    Write-Host "Files Downloaded: $totalDownloaded"
    Write-Host "Images Converted: $totalConverted"
    Write-Host "==============================" -ForegroundColor Cyan
}

# --- STEP 1: INITIAL WORD RESET ---
$msg = "This tool requires Microsoft Word to be closed. Force close now?"
$choice = [System.Windows.Forms.MessageBox]::Show($msg, "Word Reset", "YesNo", "Warning")
if ($choice -eq "No") { exit }
Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force

# --- MAIN MENU LOOP ---
do {
    Clear-Host
    Write-Host "PDF SCRAPER INTERFACE" -ForegroundColor Cyan
    Write-Host "1. Process Single PDF"
    Write-Host "2. Batch Process Folder"
    Write-Host "3. Exit"
    $selection = Read-Host "`nSelect an option"

    switch ($selection) {
        "1" {
            $fb = New-Object System.Windows.Forms.OpenFileDialog
            $fb.InitialDirectory = Join-Path $env:USERPROFILE "Downloads"
            $fb.Filter = "PDF Files (*.pdf)|*.pdf"
            if ($fb.ShowDialog() -eq "OK") { 
                Start-PDFProcessing -pdfPath $fb.FileName 
                Invoke-Item $outputBase
            }
            Read-Host "`nPress Enter to return to menu..."
        }
        "2" {
            $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
            if ($fbd.ShowDialog() -eq "OK") {
                $files = Get-ChildItem -Path $fbd.SelectedPath -Filter "*.pdf"
                Write-Host "Found $($files.Count) PDFs. Starting batch..." -ForegroundColor Yellow
                foreach ($file in $files) { Start-PDFProcessing -pdfPath $file.FullName }
                Invoke-Item $outputBase
            }
            Read-Host "`nBatch complete. Press Enter to return to menu..."
        }
        "3" {
            Write-Host "Exiting..." -ForegroundColor Red
            $running = $false
        }
        Default {
            Write-Host "Invalid selection." -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        }
    }
} while ($selection -ne "3")