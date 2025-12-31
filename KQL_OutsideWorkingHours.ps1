param(
    [bool]$ExcludeWeekends = $false,
    [bool]$IncludeNextMorning = $true
)

# Prompt for Month/Year
$mmYYYY = Read-Host "Enter month and year (MM-YYYY), or press Enter to input separately"

[int]$Year = 0
[int]$Month = 0

if ([string]::IsNullOrWhiteSpace($mmYYYY)) {
    $Year  = [int](Read-Host "Enter Year (e.g., 2025)")
    $Month = [int](Read-Host "Enter Month (1-12)")
}
else {
    $parts = $mmYYYY.Trim().Split('-')
    if ($parts.Count -ne 2) {
        Write-Error "Invalid format. Use MM-YYYY, e.g., 10-2025."
        exit 1
    }

    if (-not [int]::TryParse($parts[0], [ref]$Month) -or
        -not [int]::TryParse($parts[1], [ref]$Year)) {
        Write-Error "Month/Year must be numeric. Example: 10-2025."
        exit 1
    }
}

# Validate ranges
if ($Year -lt 2000 -or $Year -gt 2100) {
    Write-Error "Year must be between 2000 and 2100."
    exit 1
}

if ($Month -lt 1 -or $Month -gt 12) {
    Write-Error "Month must be between 1 and 12."
    exit 1
}

# London time zone (BST/GMT aware)
$tz = [TimeZoneInfo]::FindSystemTimeZoneById('GMT Standard Time')

function New-LocalDate {
    param(
        [datetime]$date,
        [int]$hour,
        [int]$minute = 0,
        [int]$second = 0
    )
    $d = [datetime]::new($date.Year, $date.Month, $date.Day, $hour, $minute, $second)
    return [datetime]::SpecifyKind($d, [System.DateTimeKind]::Unspecified)
}

# Month boundaries
$startLocal = [datetime]::SpecifyKind([datetime]::new($Year, $Month, 1, 0, 0, 0), 'Unspecified')

$nextMonthLocal = if ($Month -eq 12) {
    [datetime]::new($Year + 1, 1, 1, 0, 0, 0)
} else {
    [datetime]::new($Year, $Month + 1, 1, 0, 0, 0)
}
$nextMonthLocal = [datetime]::SpecifyKind($nextMonthLocal, 'Unspecified')

# Working hours: 09:00–17:00 local
$WORK_START = 9
$WORK_END   = 17

$blocks = New-Object System.Collections.Generic.List[string]

$curLocal = $startLocal

while ($curLocal -lt $nextMonthLocal) {

    $isWeekday = ($curLocal.DayOfWeek -ge [DayOfWeek]::Monday -and
                  $curLocal.DayOfWeek -le [DayOfWeek]::Friday)

    if (-not $ExcludeWeekends -or $isWeekday) {

        $nextDayLocal = $curLocal.AddDays(1)

        #
        # Morning block: 00:00 → 09:00
        #
        $outsideMorningStartLocal = New-LocalDate $curLocal 0
        $outsideMorningEndLocal   = New-LocalDate $curLocal $WORK_START

        $outsideMorningStartUtc = [TimeZoneInfo]::ConvertTimeToUtc($outsideMorningStartLocal, $tz)
        $outsideMorningEndUtc   = [TimeZoneInfo]::ConvertTimeToUtc($outsideMorningEndLocal, $tz)

        $mStartStr = $outsideMorningStartUtc.ToString("yyyy-MM-ddTHH:mm:ss'Z'")
        $mEndStr   = $outsideMorningEndUtc.ToString("yyyy-MM-ddTHH:mm:ss'Z'")

        $blocks.Add("(Sent>=$mStartStr AND Sent<$mEndStr)")

        #
        # Evening block: 17:00 → 00:00 next day
        #
        $outsideEveningStartLocal = New-LocalDate $curLocal $WORK_END
        $nextMidnightLocal        = New-LocalDate $nextDayLocal 0

        $outsideEveningStartUtc = [TimeZoneInfo]::ConvertTimeToUtc($outsideEveningStartLocal, $tz)
        $nextMidnightUtc        = [TimeZoneInfo]::ConvertTimeToUtc($nextMidnightLocal, $tz)

        $eStartStr = $outsideEveningStartUtc.ToString("yyyy-MM-ddTHH:mm:ss'Z'")
        $eEndStr   = $nextMidnightUtc.ToString("yyyy-MM-ddTHH:mm:ss'Z'")

        $blocks.Add("(Sent>=$eStartStr AND Sent<$eEndStr)")
    }

    $curLocal = $curLocal.AddDays(1)
}

# Optionally include the next month's first morning block
if ($IncludeNextMorning) {

    $firstOfNextLocal = $nextMonthLocal

    $nextMorningStartLocal = New-LocalDate $firstOfNextLocal 0
    $nextMorningEndLocal   = New-LocalDate $firstOfNextLocal $WORK_START

    $nextMorningStartUtc = [TimeZoneInfo]::ConvertTimeToUtc($nextMorningStartLocal, $tz)
    $nextMorningEndUtc   = [TimeZoneInfo]::ConvertTimeToUtc($nextMorningEndLocal, $tz)

    $nmStartStr = $nextMorningStartUtc.ToString("yyyy-MM-ddTHH:mm:ss'Z'")
    $nmEndStr   = $nextMorningEndUtc.ToString("yyyy-MM-ddTHH:mm:ss'Z'")

    $blocks.Add("(Sent>=$nmStartStr AND Sent<$nmEndStr)")
}

# Build final KQL
$kql = [string]::Join(" OR ", $blocks)

# Output filename
$filename = "KQL_OutsideWorkingHours_9to5_{0}-{1:D2}.txt" -f $Year, $Month

# Save file
$kql | Out-File -FilePath $filename -Encoding utf8

Write-Host "Saved: $filename"