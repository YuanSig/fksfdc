$N_weeks_ago = 1
$work_hours_hash = calc_work_hours(extract_calender_items($N_weeks_ago))
$work_hours_hash.GetEnumerator() | Export-Csv -NoTypeInformation -Path ./work_hours.csv 

#get your calender items as a hash 
Function extract_calender_items($N_weeks_ago){
    $olFolderCalendar = 9
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace('MAPI')
    $Monday = get_lastMonday($N_weeks_ago)
    $Friday = $Monday.AddDays(7).ToShortDateString()
    $Monday = $Monday.ToShortDateString()
    $Filter = "[MessageClass]='IPM.Appointment' AND [Start] > '$Monday' AND [End] < '$Friday'"
    $WeeklyCalendarItems = New-Object System.Collections.ArrayList
    $ns.GetDefaultFolder($olFolderCalendar).Items.Restrict($Filter) | foreach{
        If($_.RecurrenceState -eq 0){
        $WeeklyCalendarItems += $_
        }
    }
    Return $WeeklyCalendarItems
}


Function get_lastMonday ($N_weeks_ago){

    switch ((Get-Date).DayOfWeek){
        "Monday"{$LastMonday = (Get-Date).AddDays(0-$N_weeks_ago*7)}
        "Tuesday"{$LastMonday = (Get-Date).AddDays(-1-$N_weeks_ago*7)}
        "Wednesday"{$LastMonday = (Get-Date).AddDays(-2-$N_weeks_ago*7)}
        "Thursday"{$LastMonday = (Get-Date).AddDays(-3-$N_weeks_ago*7)}
        "Friday"{$LastMonday = (Get-Date).AddDays(-4-$N_weeks_ago*7)}
        "Saturday"{$LastMonday = (Get-Date).AddDays(-5-$N_weeks_ago*7)}
        "Sunday"{$LastMonday = (Get-Date).AddDays(-6-$N_weeks_ago*7)}
        default {Write-Host "What day is it $_ ??"}

    }
    Return $LastMonday
}


#arrange and consolidate "WeeklyCalendarItems"
#get a hash {category, sum of durations}
Function calc_work_hours($WeeklyCalendarItems){
    $i = 0
    $hours_per_category_hash = @{}
    $WeeklyCalendarItems.Categories | foreach{
        If($_ -ne ""){
            #add the duration to the exsisting key value
            If($hours_per_category_hash.ContainsKey($_)){
                $hours_per_category_hash[$_] = $hours_per_category_hash[$_] + $WeeklyCalendarItems.Duration[$i]
            }
            #add the key
            else{
                $hours_per_category_hash.Add($_ ,$WeeklyCalendarItems.Duration[$i])
            }
        }
    $i++
    }
    Return $hours_per_category_hash
}

