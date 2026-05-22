param (
    [Parameter(Mandatory=$true)]
    [string]$InputFilePath,
    [Parameter(Mandatory=$false)]
    [string]$OutputFilePath
)

$InputFile = Get-Item $InputFilePath

if (-not $InputFile.Exists) {
    Write-Error "Input file not found: $($InputFilePath)"
    exit 1
}

# Read content as UTF8
$content = Get-Content -Path $InputFilePath -Encoding UTF8 -Raw

# Replace slanted quotes and apostrophes
# Left single quotation mark U+2018
# Right single quotation mark U+2019
# Left double quotation mark U+201C
# Right double quotation mark U+201D
# Prime U+2032 (often used as apostrophe or single quote)
# Double Prime U+2033 (often used as double quote)

$content = $content -replace '[\u2018\u2019\u2032\u02BC\u0060\u00B4\u201a\u201b\uFF07]', "'" # Slanted single quotes, primes, and accents to straight apostrophe
$content = $content -replace '[\u201c\u201d\u2033\u201e\u201f\uFF02]', '"' # Slanted double quotes and double prime to straight double quote

# Construct output file path
if ([string]::IsNullOrWhiteSpace($OutputFilePath)) {
    $OutputFilePath = Join-Path -Path $InputFile.DirectoryName -ChildPath "$($InputFile.BaseName)_ansi.txt"
}

# Write content to output file with ANSI encoding
$content | Set-Content -Path $OutputFilePath -Encoding Default

Write-Host "Successfully converted '$($InputFilePath)' to ANSI with straight quotes."
Write-Host "Output file: '$($OutputFilePath)'"
