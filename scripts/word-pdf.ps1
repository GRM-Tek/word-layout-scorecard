# word-pdf.ps1 -DocsDir <folder of .docx> -OutDir <folder for .pdf> [-Log <file>]
# Microsoft Word prints every .docx in DocsDir to a PDF of the same name in OutDir, through "Microsoft Print to PDF"
# (saving as PDF freezes in this computer's Word). The answer key of a document set is made this way, and so is
# Word's print of an editor's saved copies for the "pages identical" measure. Nothing leaves this computer.
param([string]$DocsDir, [string]$OutDir, [string]$Log = "")
if (-not (Test-Path $DocsDir)) { throw "no such folder: $DocsDir" }
New-Item -ItemType Directory -Force $OutDir | Out-Null
if ($Log -eq "") { $Log = Join-Path $OutDir "word-pdf.log" }
$DocsDir = (Resolve-Path $DocsDir).Path; $OutDir = (Resolve-Path $OutDir).Path
function L($m) { $line = "{0:HH:mm:ss} {1}" -f (Get-Date), $m; Write-Output $line; Add-Content -Path $Log -Value $line }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  # Save-as-PDF hangs in this Word install; printing to Microsoft Print to PDF works. Set for this Word session only,
  # never as the Windows default printer.
  [System.__ComObject].InvokeMember('FilePrintSetup', [Reflection.BindingFlags]::InvokeMethod, $null, $word.WordBasic,
    [object[]]@('Microsoft Print to PDF', 1), $null, $null, [string[]]@('Printer', 'DoNotSetAsSysDefault')) | Out-Null
  L ("printer " + $word.ActivePrinter)
  Get-ChildItem -Path $DocsDir -Filter *.docx | ForEach-Object {
    $name = $_.BaseName
    try {
      $doc = $word.Documents.Open($_.FullName, $false, $true, $false)
      $pages = $doc.ComputeStatistics(2)
      $out = Join-Path $OutDir ($name + '.pdf')
      if (Test-Path $out) { Remove-Item $out }
      [System.__ComObject].InvokeMember('PrintOut', [Reflection.BindingFlags]::InvokeMethod, $null, $doc,
        [object[]]@($false, [string]$out, $true), $null, $null, [string[]]@('Background', 'OutputFileName', 'PrintToFile')) | Out-Null
      $doc.Close(0)
      $wait = 0; while (-not (Test-Path $out) -and $wait -lt 120) { Start-Sleep -Milliseconds 500; $wait++ }
      L "done $name pages=$pages"
    } catch { L ("error $name " + $_.Exception.Message) }
  }
} finally { $word.Quit() }
