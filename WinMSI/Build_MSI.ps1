# Define location of WiX Toolset (Adjust if you installed elsewhere)
$WixPath = "${env:ProgramFiles(x86)}\WiX Toolset v3.14\bin"

if (-not (Test-Path $WixPath)) {
    Write-Error "WiX Toolset not found at $WixPath.`nPlease install WiX Toolset v3.14 from https://wixtoolset.org/"
    exit
}

$Candle = Join-Path $WixPath "candle.exe"
$Light = Join-Path $WixPath "light.exe"

$SourceFile = Join-Path $PSScriptRoot "Product.wxs"
$ObjectFile = Join-Path $PSScriptRoot "Product.wixobj"
$MsiFile    = Join-Path $PSScriptRoot "ConvertQuoteMarks.msi"

Write-Host "Compiling WiX source..."
& $Candle $SourceFile -out $ObjectFile

if (Test-Path $ObjectFile) {
    Write-Host "Linking MSI..."
    # -spdb suppresses the creation of a .pdb debug file
    & $Light $ObjectFile -out $MsiFile -spdb
    
    Write-Host "Success! MSI created: $MsiFile" -ForegroundColor Green
}