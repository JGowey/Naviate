<#
    Script: ITM Weight Calculator (lbs)
    Purpose:
        Calculates dry weight of MEP Fabrication ITMs and RFAs in pounds (lbs).
        - If ITM is a fishmouth: gets branch size and pipe schedule (STD, SCH40, SCH10) and looks up lbs/ft.
        - If not a fishmouth: uses "Weight" parameter in kg and converts to lbs.
        - If RFA: reads "CP_Weight" parameter from instance or type.
		- If Null Return 0
        Always returns a numeric value rounded to 3 decimal places in pounds (lbs).
#>

# Handle RFAs (non-ITMs)
if ($naviateElement -isnot [Autodesk.Revit.DB.FabricationPart]) {
    $cpWeightParam = $naviateElement.LookupParameter("CP_Weight")
    if (-not $cpWeightParam -or -not $cpWeightParam.HasValue) {
        $typeId = $naviateElement.GetTypeId()
        $typeElement = $naviateElement.Document.GetElement($typeId)
        if ($typeElement) {
            $cpWeightParam = $typeElement.LookupParameter("CP_Weight")
        }
    }

    if ($cpWeightParam -and $cpWeightParam.HasValue) {
        $rawWeight = $cpWeightParam.AsDouble()
        return [Math]::Round($rawWeight, 3)
    } else {
        return 0
    }
}

# --- Weight-per-foot tables ---
$pipeWeight_STD = @{
    0.125=0.24; 0.25=0.42; 0.375=0.57; 0.5=0.85; 0.75=1.13; 1.0=1.68; 1.25=2.27; 1.5=2.72;
    2.0=3.66; 2.5=5.80; 3.0=7.58; 3.5=9.12; 4.0=10.79; 5.0=14.62; 6.0=18.97; 8.0=28.55;
    10.0=40.48; 12.0=49.56; 14.0=57.93; 16.0=67.54; 18.0=77.52; 20.0=88.22; 24.0=110.62;
    30.0=138.04; 36.0=169.27; 42.0=201.58; 48.0=233.63
}
$pipeWeight_SCH40 = @{
    0.125=0.24; 0.25=0.41; 0.375=0.54; 0.5=0.85; 0.75=1.13; 1.0=1.68; 1.25=2.27; 1.5=2.72;
    2.0=2.638; 2.5=3.652; 3.0=5.618; 3.5=7.093; 4.0=7.568; 5.0=10.79; 6.0=18.97; 8.0=28.58;
    10.0=40.48; 12.0=49.56; 14.0=62.58; 16.0=75.58; 18.0=89.68; 20.0=105.1; 24.0=138.7;
    30.0=172.0; 36.0=208.4; 42.0=248.3; 48.0=288.5
}
$pipeWeight_SCH10 = @{
    0.125=0.20; 0.25=0.33; 0.375=0.48; 0.5=0.72; 0.75=0.97; 1.0=1.30; 1.25=1.70; 1.5=2.10;
    2.0=2.09; 2.5=2.72; 3.0=3.66; 3.5=4.52; 4.0=5.62; 5.0=7.77; 6.0=9.29; 8.0=13.60;
    10.0=17.81; 12.0=22.32; 14.0=24.00; 16.0=27.17; 18.0=30.30; 20.0=33.84; 24.0=39.69;
    30.0=47.5; 36.0=56.0; 42.0=64.0; 48.0=72.0
}

# Get pipe schedule
$spec = $naviateElement.ProductSpecificationDescription
if ($spec -match "SCH ?10") { $schedule = "SCH10" }
elseif ($spec -match "SCH ?40") { $schedule = "SCH40" }
elseif ($spec -match "STD") { $schedule = "STD" }
else { $schedule = "STD" }

# Read Product Entry (e.g., "4x2") to detect fishmouth
$productEntry = ""
$entryParam = $naviateElement.LookupParameter("Product Entry")
if ($entryParam -and $entryParam.HasValue) {
    $productEntry = $entryParam.AsString()
}

$isFish = $false
$branchSize = $null
if ($productEntry -match "([0-9 ]+)(x|X)([0-9 \/\.]+)") {
    function Convert-ToDecimal {
        param([string]$text)
        $text = $text.Trim()
        if ($text -match "^(\d+)\s+(\d+)/(\d+)$") {
            return [double]$matches[1] + ([double]$matches[2] / [double]$matches[3])
        } elseif ($text -match "^(\d+)/(\d+)$") {
            return [double]$matches[1] / [double]$matches[2]
        } else {
            return [double]$text
        }
    }
    $branchSize = Convert-ToDecimal $matches[3]
    $isFish = $true
}

# Helper: Closest match in lookup
function Get-ClosestKey {
    param($lookup, $target)
    $closest = $null
    $minDiff = [double]::MaxValue
    foreach ($k in $lookup.Keys) {
        $diff = [math]::Abs($k - $target)
        if ($diff -lt $minDiff) {
            $minDiff = $diff
            $closest = $k
        }
    }
    return $closest
}

# Fishmouth: Use lookup table
if ($isFish -and $branchSize -ne $null) {
    switch ($schedule) {
        "SCH10" { $lookup = $pipeWeight_SCH10 }
        "SCH40" { $lookup = $pipeWeight_SCH40 }
        default { $lookup = $pipeWeight_STD }
    }

    $nomSize = Get-ClosestKey $lookup $branchSize
    $weightPerFoot = $lookup[$nomSize]

    $lengthParam = $naviateElement.LookupParameter("Length")
    $lengthFeet = 0
    if ($lengthParam -and $lengthParam.HasValue) {
        $lengthFeet = $lengthParam.AsDouble()
    }

    $weight = $weightPerFoot * $lengthFeet
}
else {
    # Standard ITM: Use "Weight" in kg and convert
    $weightParam = $naviateElement.LookupParameter("Weight")
    if ($weightParam -and $weightParam.HasValue) {
        $weight = $weightParam.AsDouble() * 2.20462
    } else {
        $weight = 0
    }
}

# Round to 3 decimal places
return [Math]::Round($weight, 3)
