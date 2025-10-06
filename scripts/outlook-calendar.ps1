# Outlook Calendar COM Automation Script
# This script provides calendar operations through Outlook COM objects

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('list', 'get', 'create', 'update', 'delete', 'search')]
    [string]$Action,

    [string]$StartDate,
    [string]$EndDate,
    [string]$EventId,
    [string]$Subject,
    [string]$Body,
    [string]$Location,
    [string]$Attendees,
    [string]$Query,
    [bool]$IsAllDay = $false
)

# Error handling
$ErrorActionPreference = 'Stop'

function Get-OutlookApplication {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        return $outlook
    }
    catch {
        Write-Error "Failed to connect to Outlook. Make sure Outlook is installed and accessible."
        exit 1
    }
}

function ConvertTo-JsonSafe {
    param($Object)
    return $Object | ConvertTo-Json -Depth 10 -Compress
}

function Get-CalendarFolder {
    param($Outlook)
    $namespace = $Outlook.GetNamespace("MAPI")
    return $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
}

function List-Events {
    param($StartDate, $EndDate)

    $outlook = Get-OutlookApplication
    $calendar = Get-CalendarFolder -Outlook $outlook
    $items = $calendar.Items
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true

    $events = @()

    if ($StartDate -and $EndDate) {
        $filter = "[Start] >= '$StartDate' AND [End] <= '$EndDate'"
        $filteredItems = $items.Restrict($filter)
        foreach ($item in $filteredItems) {
            $events += @{
                id = $item.EntryID
                subject = $item.Subject
                start = $item.Start.ToString("o")
                end = $item.End.ToString("o")
                location = $item.Location
                body = $item.Body
                isAllDay = $item.AllDayEvent
            }
        }
    }
    else {
        $count = 0
        foreach ($item in $items) {
            if ($count -ge 50) { break }
            $events += @{
                id = $item.EntryID
                subject = $item.Subject
                start = $item.Start.ToString("o")
                end = $item.End.ToString("o")
                location = $item.Location
                body = $item.Body
                isAllDay = $item.AllDayEvent
            }
            $count++
        }
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    return ConvertTo-JsonSafe -Object $events
}

function Get-Event {
    param($EventId)

    $outlook = Get-OutlookApplication
    $namespace = $outlook.GetNamespace("MAPI")

    try {
        $item = $namespace.GetItemFromID($EventId)
        $event = @{
            id = $item.EntryID
            subject = $item.Subject
            start = $item.Start.ToString("o")
            end = $item.End.ToString("o")
            location = $item.Location
            body = $item.Body
            isAllDay = $item.AllDayEvent
            organizer = $item.Organizer
            required_attendees = $item.RequiredAttendees
            optional_attendees = $item.OptionalAttendees
        }

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        return ConvertTo-JsonSafe -Object $event
    }
    catch {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        Write-Error "Event not found: $EventId"
        exit 1
    }
}

function Create-Event {
    param($Subject, $StartDate, $EndDate, $Body, $Location, $Attendees, $IsAllDay)

    $outlook = Get-OutlookApplication
    $item = $outlook.CreateItem([Microsoft.Office.Interop.Outlook.OlItemType]::olAppointmentItem)

    $item.Subject = $Subject
    $item.Start = [DateTime]::Parse($StartDate)
    $item.End = [DateTime]::Parse($EndDate)

    if ($Body) { $item.Body = $Body }
    if ($Location) { $item.Location = $Location }
    if ($IsAllDay) { $item.AllDayEvent = $true }

    if ($Attendees) {
        $item.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
        $attendeeList = $Attendees -split ';'
        foreach ($email in $attendeeList) {
            $recipient = $item.Recipients.Add($email.Trim())
            $recipient.Type = [Microsoft.Office.Interop.Outlook.OlMeetingRecipientType]::olRequired
        }
        $item.Recipients.ResolveAll() | Out-Null
    }

    $item.Save()

    $result = @{
        id = $item.EntryID
        subject = $item.Subject
        start = $item.Start.ToString("o")
        end = $item.End.ToString("o")
        message = "Event created successfully"
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    return ConvertTo-JsonSafe -Object $result
}

function Update-Event {
    param($EventId, $Subject, $StartDate, $EndDate, $Body, $Location)

    $outlook = Get-OutlookApplication
    $namespace = $outlook.GetNamespace("MAPI")

    try {
        $item = $namespace.GetItemFromID($EventId)

        if ($Subject) { $item.Subject = $Subject }
        if ($StartDate) { $item.Start = [DateTime]::Parse($StartDate) }
        if ($EndDate) { $item.End = [DateTime]::Parse($EndDate) }
        if ($Body) { $item.Body = $Body }
        if ($Location) { $item.Location = $Location }

        $item.Save()

        $result = @{
            id = $item.EntryID
            subject = $item.Subject
            message = "Event updated successfully"
        }

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        return ConvertTo-JsonSafe -Object $result
    }
    catch {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        Write-Error "Failed to update event: $EventId"
        exit 1
    }
}

function Delete-Event {
    param($EventId)

    $outlook = Get-OutlookApplication
    $namespace = $outlook.GetNamespace("MAPI")

    try {
        $item = $namespace.GetItemFromID($EventId)
        $item.Delete()

        $result = @{
            success = $true
            message = "Event deleted successfully"
        }

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        return ConvertTo-JsonSafe -Object $result
    }
    catch {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        Write-Error "Failed to delete event: $EventId"
        exit 1
    }
}

function Search-Events {
    param($Query)

    $outlook = Get-OutlookApplication
    $calendar = Get-CalendarFolder -Outlook $outlook
    $items = $calendar.Items

    $filter = "[Subject] LIKE '%$Query%' OR [Body] LIKE '%$Query%'"
    $results = $items.Restrict($filter)

    $events = @()
    foreach ($item in $results) {
        $events += @{
            id = $item.EntryID
            subject = $item.Subject
            start = $item.Start.ToString("o")
            end = $item.End.ToString("o")
            location = $item.Location
        }
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    return ConvertTo-JsonSafe -Object $events
}

# Main execution
try {
    switch ($Action) {
        'list' {
            Write-Output (List-Events -StartDate $StartDate -EndDate $EndDate)
        }
        'get' {
            Write-Output (Get-Event -EventId $EventId)
        }
        'create' {
            Write-Output (Create-Event -Subject $Subject -StartDate $StartDate -EndDate $EndDate -Body $Body -Location $Location -Attendees $Attendees -IsAllDay $IsAllDay)
        }
        'update' {
            Write-Output (Update-Event -EventId $EventId -Subject $Subject -StartDate $StartDate -EndDate $EndDate -Body $Body -Location $Location)
        }
        'delete' {
            Write-Output (Delete-Event -EventId $EventId)
        }
        'search' {
            Write-Output (Search-Events -Query $Query)
        }
    }
}
catch {
    $errorObj = @{
        error = $_.Exception.Message
        type = $_.Exception.GetType().FullName
    }
    Write-Output (ConvertTo-JsonSafe -Object $errorObj)
    exit 1
}
