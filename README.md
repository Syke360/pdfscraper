# pdfscraper
Powershell Script to scan a PDF for Links and Pull the downloads from those links - Pictures and PDF Files. Good for Invoice Scrapping and Proof Of Delivery where one PDF contains links to multiple files. 

Designed for Windows 10/11 - Requires Microsoft Word to operate


PDF SCRAPER USER MANUAL 
================================================
1. Double-click CLICK_ME.bat.
2. Select your supplier PDF.
3. If Word asks to 'Convert', check 'Do not show again' and click OK.
4. Results appear in the \Output folder.

FAILURES & LOGS:
- If a download fails, it will retry 3 times automatically.
- Check 'Processing_Log.txt' inside the job folder for error details.
- 'Security Block': The file was not a PDF or Image.
- 'Size Block': The file was over 50MB.
