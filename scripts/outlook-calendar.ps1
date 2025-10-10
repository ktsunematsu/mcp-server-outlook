# Outlook Calendar COM Automation Script
# This script provides calendar operations through Outlook COM objects

param(
    [Parameter(Mandatory=$true)]
    [string]$Action,

    [string]$StartDate,
    [string]$EndDate,
    [string]$EventId,
    [string]$Subject,
    [string]$Body,
    [string]$Location,
    [string]$Attendees,
    [string]$Query,
    [switch]$IsAllDay
)

# 出力エンコーディングをUTF-8に統一
$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Error handling
$ErrorActionPreference = 'Stop'

function Get-OutlookApplication {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        return $outlook
    }
    catch {
        Write-Output (ConvertTo-JsonSafe -Object @{ error = "Failed to connect to Outlook. Make sure Outlook is installed and accessible."; type = $_.Exception.GetType().FullName })
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

    $events = @()
    if ($StartDate -and $EndDate) {
        try {
            # OutlookのRestrict用に日付をyyyy/MM/dd HH:mm形式に変換
            $startDateObj = [datetime]::Parse($StartDate)
            $endDateObj = [datetime]::Parse($EndDate)
            $startStr = $startDateObj.ToString('yyyy/MM/dd HH:mm')
            $endStr = $endDateObj.ToString('yyyy/MM/dd HH:mm')
            $filter = "[End] >= '$startStr' AND [Start] <= '$endStr'"
            $filteredItems = $items.Restrict($filter)
            $count = 0
            foreach ($item in $filteredItems) {
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
        } catch {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
            Write-Output (ConvertTo-JsonSafe -Object @{ error = $_.Exception.Message; type = $_.Exception.GetType().FullName })
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
    # 必ず空配列でもJSONを返す
    Write-Output (ConvertTo-JsonSafe -Object $events)
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
        Write-Output (ConvertTo-JsonSafe -Object @{ error = "Event not found: $EventId"; type = $_.Exception.GetType().FullName })
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
        Write-Output (ConvertTo-JsonSafe -Object @{ error = "Failed to update event: $EventId"; type = $_.Exception.GetType().FullName })
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
        Write-Output (ConvertTo-JsonSafe -Object @{ error = "Failed to delete event: $EventId"; type = $_.Exception.GetType().FullName })
    }
}

function Search-Events {
    param($Query)

    $outlook = Get-OutlookApplication
    $calendar = Get-CalendarFolder -Outlook $outlook
    $items = $calendar.Items

    try {
        # Outlookの検索では、部分一致検索を行う場合はRestrictではなく、
        # 全アイテムをループして手動でフィルタリングする必要があります
        $events = @()
        $count = 0
        
        foreach ($item in $items) {
            # 最大100件まで検索
            if ($count -ge 100) { break }
            
            # 件名または本文にクエリが含まれているかチェック
            if ($item.Subject -like "*$Query*" -or $item.Body -like "*$Query*") {
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
    catch {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        Write-Output (ConvertTo-JsonSafe -Object @{ error = $_.Exception.Message; type = $_.Exception.GetType().FullName })
    }
}


# Main execution
$script:__outlook_calendar_ps1_output = $null
try {
    switch ($Action) {
        'list' {
            $script:__outlook_calendar_ps1_output = List-Events -StartDate $StartDate -EndDate $EndDate
        }
        'get' {
            $script:__outlook_calendar_ps1_output = Get-Event -EventId $EventId
        }
        'create' {
            $script:__outlook_calendar_ps1_output = Create-Event -Subject $Subject -StartDate $StartDate -EndDate $EndDate -Body $Body -Location $Location -Attendees $Attendees -IsAllDay $IsAllDay
        }
        'update' {
            $script:__outlook_calendar_ps1_output = Update-Event -EventId $EventId -Subject $Subject -StartDate $StartDate -EndDate $EndDate -Body $Body -Location $Location
        }
        'delete' {
            $script:__outlook_calendar_ps1_output = Delete-Event -EventId $EventId
        }
        'search' {
            $script:__outlook_calendar_ps1_output = Search-Events -Query $Query
        }
    }
    if ($null -ne $script:__outlook_calendar_ps1_output -and $script:__outlook_calendar_ps1_output -ne "") {
        Write-Output $script:__outlook_calendar_ps1_output
    } else {
        Write-Output (ConvertTo-JsonSafe -Object @())
    }
}
catch {
    $errorObj = @{
        error = $_.Exception.Message
        type = $_.Exception.GetType().FullName
    }
    Write-Output (ConvertTo-JsonSafe -Object $errorObj)
}
