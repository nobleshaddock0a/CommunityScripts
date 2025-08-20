<#
    HOW TO USE THIS SCRIPT

    1. Export your media from StashDB:
       - In your StashDB instance, go to **Scenes** (list view).
       - Select **all scenes**.
       - Click the **three dots (...)** → **Export** (or "Export all").

    2. Extract the downloaded archive:
       - The extracted folder will look like: `export{yyyy}{mm}{dd}-{time}`

    3. Configure this script:
       - Edit **line 27**: set `$InputDir` to the path of your extracted `scenes` folder.
       - (Optional) Edit **line 28**: set `$OutputXlsx` to your preferred save location.
         By default, the Excel file will be saved to your Desktop.

    4. Run the script:
       - Open PowerShell.
       - Run the script. It will generate a single Excel file containing your exported data.

#>


# Requires: PowerShell 5.1+ (Windows) or PowerShell 7+, and the ImportExcel module (auto-installed below)

# --- User-configurable paths ---
$InputDir = Join-Path $env:USERPROFILE 'Downloads\export20250819-200555\scenes'
$OutputXlsx = Join-Path $env:USERPROFILE 'Desktop\Scenes_Export.xlsx'

# --- Ensure ImportExcel is available ---
try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "ImportExcel module not found. Installing for current user..."
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module ImportExcel -ErrorAction Stop
} catch {
    Write-Error "Failed to install/load ImportExcel: $($_.Exception.Message)"
    exit 1
}

# --- Validate source directory ---
if (-not (Test-Path -LiteralPath $InputDir)) {
    Write-Error "Input directory not found: $InputDir"
    exit 1
}

# --- Helper: join an array safely into comma-separated string ---
function Join-CSV([object]$value) {
    if ($null -eq $value) { return $null }
    if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
        $clean = $value | Where-Object { $_ -ne $null -and $_.ToString().Trim() -ne '' } | ForEach-Object { $_.ToString().Trim() }
        if ($clean.Count -eq 0) { return $null }
        return [string]::Join(', ', $clean)
    }
    # Scalar
    $s = $value.ToString().Trim()
    if ($s -eq '') { return $null }
    return $s
}

# --- Read JSON files and build rows ---
$rows = New-Object System.Collections.Generic.List[object]

$files = Get-ChildItem -LiteralPath $InputDir -Filter *.json -File -ErrorAction Stop | Sort-Object Name
if ($files.Count -eq 0) {
    Write-Error "No *.json files found in: $InputDir"
    exit 1
}

foreach ($file in $files) {
    try {
        $jsonText = Get-Content -LiteralPath $file.FullName -Raw -ErrorAction Stop
        $obj = $jsonText | ConvertFrom-Json -ErrorAction Stop

        # Build a record with required fields
        $record = [PSCustomObject]@{
            Title       = ($obj.title       | ForEach-Object { $_ })              # passthrough or null
            Studio      = ($obj.studio      | ForEach-Object { $_ })
            URL         = Join-CSV $obj.urls
            Date        = ($obj.date        | ForEach-Object { $_ })
            Organized   = $obj.organized
            Details     = ($obj.details     | ForEach-Object { $_ })
            Performers  = Join-CSV $obj.performers
            Tags        = Join-CSV $obj.tags
            Created_At  = ($obj.created_at  | ForEach-Object { $_ })
        }

        $rows.Add($record) | Out-Null
    } catch {
        Write-Warning "Skipping '$($file.Name)': $($_.Exception.Message)"
    }
}

if ($rows.Count -eq 0) {
    Write-Error "No JSON files parsed into rows. Check the directory and file contents."
    exit 1
}

# --- Export to Excel ---
try {
    # -NoNumberConversion prevents Excel from auto-converting e.g., tag values like "2"
    $rows | Export-Excel `
        -Path $OutputXlsx `
        -WorksheetName 'Scenes' `
        -TableName 'ScenesTable' `
        -ClearSheet `
        -AutoSize `
        -AutoFilter `
        -FreezeTopRow `
        -BoldTopRow `
        -NoNumberConversion 'URL','Performers','Tags'

    Write-Host "Done. Excel saved to: $OutputXlsx"
} catch {
    Write-Error "Failed to write Excel: $($_.Exception.Message)"
    exit 1
}
