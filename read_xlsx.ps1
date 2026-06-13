# PowerShell script to extract and read Excel (.xlsx) file using standard zip and XML processing

$xlsxPath = "Scan Ex.xlsx"
$tempDir = "temp_xlsx_extract"

if (Test-Path $tempDir) {
    Remove-Item -Recurse -Force $tempDir
}

# Expand the zip file
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($xlsxPath, $tempDir)

# Read shared strings
$sharedStrings = @()
$sharedStringsPath = "$tempDir\xl\sharedStrings.xml"
if (Test-Path $sharedStringsPath) {
    [xml]$ssXml = Get-Content $sharedStringsPath -Encoding UTF8
    # The namespace for spreadsheetml
    $ns = New-Object System.Xml.XmlNamespaceManager($ssXml.NameTable)
    $ns.AddNamespace("ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    
    $nodes = $ssXml.SelectNodes("//ns:t", $ns)
    foreach ($node in $nodes) {
        $sharedStrings += $node.InnerText
    }
}

# Read sheet1.xml
$sheetPath = "$tempDir\xl\worksheets\sheet1.xml"
if (Test-Path $sheetPath) {
    [xml]$sheetXml = Get-Content $sheetPath -Encoding UTF8
    $ns = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
    $ns.AddNamespace("ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    
    $rows = $sheetXml.SelectNodes("//ns:row", $ns)
    Write-Output "Total rows found: $($rows.Count)"
    
    # Print the first 25 rows
    $count = 0
    foreach ($row in $rows) {
        if ($count -ge 25) { break }
        $rowData = @()
        $cells = $row.SelectNodes("ns:c", $ns)
        foreach ($cell in $cells) {
            $valNode = $cell.SelectSingleNode("ns:v", $ns)
            $val = $null
            if ($valNode) {
                $val = $valNode.InnerText
                $t = $cell.getAttribute("t")
                if ($t -eq "s") {
                    $val = $sharedStrings[[int]$val]
                }
            }
            $rowData += "$($cell.getAttribute('r')):$val"
        }
        Write-Output "Row $($row.getAttribute('r')) [Level $($row.getAttribute('outlineLevel'))]: $($rowData -join ', ')"
        $count++
    }
}

# Clean up
if (Test-Path $tempDir) {
    Remove-Item -Recurse -Force $tempDir
}
