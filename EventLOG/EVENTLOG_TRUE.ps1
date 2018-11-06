function event-read ($ID = '4740,4726,4720', $MaxEvents = 20)
{
[int64]$time = [int64](New-TimeSpan (Get-Date).AddHours(-3)).TotalMilliseconds
$Addomain = Get-ADDomainController -Filter {IsReadOnly -eq $false} | select HostName

$out = [PSCustomObject]@{

'Server' = ""
'Time' = ""
'ID' = ""
'Info' = ""
'TargetUserName' = ""
'SubjectUserName' = ""
'TargetDomainName' = ""
}


foreach ($AD in $Addomain)
{
    Write-Host -BackgroundColor Yellow -ForegroundColor black Controller = $AD.HostName
    $event = Get-WinEvent -FilterHashtable @{LogName='Security';ID=4740,4726,4720} -MaxEvents $MaxEvents -ComputerName $AD.HostName -ErrorAction SilentlyContinue

    $t = $event.Count
    $t1 = 1
    if (!($event -eq $null) -and !($event -eq -1))
    {
        for ($i=0; $i -ne $event.Count-1 ; $i++)
        {
        $eventcount = $event[$i]
        $eventxml = [xml]$event[$i].ToXML()

            if ($eventxml.Event.EventData.Data[1].'#text' -eq $null)
            {
                $TargetDomainName = "Null"
            }else
            {
                $TargetDomainName = $eventxml.Event.EventData.Data[1].'#text'.ToString()
            }

        [array]$out += [PSCustomObject]@{
        'Server' = $AD.HostName
        'Time' = $eventcount.TimeCreated
        'ID' = $eventcount.Id.ToString()
        'Info' = $eventcount.LevelDisplayName.ToString()
        'TargetUserName' = $eventxml.Event.EventData.Data[0].'#text'.ToString()
        'SubjectUserName' = $eventxml.Event.EventData.Data[4].'#text'.ToString()
        'TargetDomainName' = $TargetDomainName
        }
        
        Write-Progress -Activity "Создание списка из $t пользователей" -Status "Обработка пользователя: $t1" -CurrentOperation $eventxml.Event.EventData.Data[0].'#text'.ToString() -PercentComplete ($i / $event.Count*100)
        $t1++
        }
    }
    Write-Host -BackgroundColor Yellow -ForegroundColor black Controller = $AD.HostName - COMPLETE
}


#$out | where {$_.Time -ge (Get-Date).AddHours(-1)} | Out-GridView
$out | Out-GridView
}