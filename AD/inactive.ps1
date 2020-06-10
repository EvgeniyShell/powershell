#Define
$xMF_inactive_txtbox_ou.Text = "OU=OU_RT,DC=ASO,DC=RT,DC=LOCAL"
$xMF_inactive_months.Text = "6"
$global:sortcheck = $true
$xMF_inactive_txtbox_filter_Descr.Text = "декрет,общий"


###############################################################################
# Действия по сортировке Listview
###############################################################################
#
#[Windows.RoutedEventHandler]$evt2 = 	{
#	$view1 = [System.Windows.Data.CollectionViewSource]::GetDefaultView($xMF_lstv_inactive.Items)
#	$sort1 = $view1.SortDescriptions[0].Direction
#	$direction1 = if($sort1 -and 'Descending' -eq $sort1){'Ascending'}else{'Descending'}
#	$view1.SortDescriptions.Clear()
#	$sortDescription1 = New-Object System.ComponentModel.SortDescription($_.OriginalSource.Column.Header, $direction1)
#	$view1.SortDescriptions.Add($sortDescription1)
#}
#
 #   $xMF_lstv_inactive.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $evt2)



#Функция для сортировки
function sortirovka ($sort ="WhenCreated")
{
    
    $Global:spisoksort = @()

if ($global:sortcheck -eq $true)
{
    $global:sortcheck = $false

    for($i=0;$i -ne $xMF_lstv_inactive.Items.Count;$i++)
    {
        if ($xMF_lstv_inactive.Items[$i].LastLogonAD -eq ""){$LastLogonAD = ""}else{$LastLogonAD = [datetime]::Parse($xMF_lstv_inactive.Items[$i].LastLogonAD)}
        if ($xMF_lstv_inactive.Items[$i].LastLogonEX -eq ""){$LastLogonEX = ""}else{$LastLogonEX = [datetime]::Parse($xMF_lstv_inactive.Items[$i].LastLogonEX)}
        if ($xMF_lstv_inactive.Items[$i].WhenCreated -eq ""){$WhenCreated = ""}else{$WhenCreated = [datetime]::Parse($xMF_lstv_inactive.Items[$i].WhenCreated)}

        $Global:spisoksort += [PSCustomObject]@{
                'Enabled' = $xMF_lstv_inactive.Items[$i].Enabled
                'DisplayName'= $xMF_lstv_inactive.Items[$i].DisplayName
                'SamAccountname' = $xMF_lstv_inactive.Items[$i].SamAccountname
                'LastLogonAD' = $LastLogonAD
                'LastLogonEX' = $LastLogonEX
                'WhenCreated' = $WhenCreated
                'Description' = $xMF_lstv_inactive.Items[$i].Description
                'mail' = $xMF_lstv_inactive.Items[$i].mail
                'DistinguishedName' = $xMF_lstv_inactive.Items[$i].DistinguishedName
                'status' = $xMF_lstv_inactive.Items[$i].status
        }
    }

    $tempsorts = $spisoksort | sort $sort

    $xMF_lstv_inactive.Items.Clear()
    foreach ($tempsort in $tempsorts)
    {
        if($tempsort.LastLogonAD -eq ""){$tempsort.LastLogonAD = ""}else{$tempsort.LastLogonAD = $tempsort.LastLogonAD.ToString('dd/MM/yyyy')}
        if($tempsort.LastLogonEX -eq ""){$tempsort.LastLogonEX = ""}else{$tempsort.LastLogonEX = $tempsort.LastLogonEX.ToString('dd/MM/yyyy')}
        if($tempsort.WhenCreated -eq ""){$tempsort.WhenCreated = ""}else{$tempsort.WhenCreated = $tempsort.WhenCreated.ToString('dd/MM/yyyy')}

        $xMF_lstv_inactive.items.Add($tempsort)
    }

    Write-Host "Сортировка прошла по возрастанию"



}
elseif ($global:sortcheck -eq $false)
{
    $global:sortcheck = $true

    for($i=0;$i -ne $xMF_lstv_inactive.Items.Count;$i++)
    {
        if ($xMF_lstv_inactive.Items[$i].LastLogonAD -eq ""){$LastLogonAD = ""}else{$LastLogonAD = [datetime]::Parse($xMF_lstv_inactive.Items[$i].LastLogonAD)}
        if ($xMF_lstv_inactive.Items[$i].LastLogonEX -eq ""){$LastLogonEX = ""}else{$LastLogonEX = [datetime]::Parse($xMF_lstv_inactive.Items[$i].LastLogonEX)}
        if ($xMF_lstv_inactive.Items[$i].WhenCreated -eq ""){$WhenCreated = ""}else{$WhenCreated = [datetime]::Parse($xMF_lstv_inactive.Items[$i].WhenCreated)}

        $Global:spisoksort += [PSCustomObject]@{
                'Enabled' = $xMF_lstv_inactive.Items[$i].Enabled
                'DisplayName'= $xMF_lstv_inactive.Items[$i].DisplayName
                'SamAccountname' = $xMF_lstv_inactive.Items[$i].SamAccountname
                'LastLogonAD' = $LastLogonAD
                'LastLogonEX' = $LastLogonEX
                'WhenCreated' = $WhenCreated
                'Description' = $xMF_lstv_inactive.Items[$i].Description
                'mail' = $xMF_lstv_inactive.Items[$i].mail
                'DistinguishedName' = $xMF_lstv_inactive.Items[$i].DistinguishedName
                'status' = $xMF_lstv_inactive.Items[$i].status
         }
    }

    $tempsorts = $spisoksort | sort $sort -Descending

    $xMF_lstv_inactive.Items.Clear()
    foreach ($tempsort in $tempsorts)
    {
        if($tempsort.LastLogonAD -eq ""){$tempsort.LastLogonAD = ""}else{$tempsort.LastLogonAD = $tempsort.LastLogonAD.ToString('dd/MM/yyyy')}
        if($tempsort.LastLogonEX -eq ""){$tempsort.LastLogonEX = ""}else{$tempsort.LastLogonEX = $tempsort.LastLogonEX.ToString('dd/MM/yyyy')}
        if($tempsort.WhenCreated -eq ""){$tempsort.WhenCreated = ""}else{$tempsort.WhenCreated = $tempsort.WhenCreated.ToString('dd/MM/yyyy')}

        $xMF_lstv_inactive.items.Add($tempsort)
    }


    Write-Host "Сортировка прошла по убыванию"
}

}




#
#Функция для тестирования, загружено ли удаленное управление Exchange
#
function testconnection
{
    if (!($global:Session -eq ""))
    {
        if ($global:Session.State -eq "Opened")
        {
            $xMF_inactive_exch_btn.Content = "Статус: подключено"
            $xMF_inactive_exch_btn.Background = "green"
            $xMF_menuitem_exchange.Background = "green"
            return "on"
        }elseif ($global:Session.State -eq "Closed")
        {
            $xMF_inactive_exch_btn.Content = "Статус: не подключено"
            $xMF_inactive_exch_btn.Background = "red"
            $xMF_menuitem_exchange.Background = "red"
            return "off"
        }
    }else
    {
        $xMF_inactive_exch_btn.Content = "Статус: не подключено"
        $xMF_inactive_exch_btn.Background = "red"
        return "off"
    }
}


#
# Функция подключения к Exchange
#
function exch_connect
{
    $statusconnect = "False"
    try
    {
        $global:Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://pso-ex-02/PowerShell/ -Authentication Kerberos
        $statusconnect = "OK"
    }catch
    {
        write-host -BackgroundColor Red -ForegroundColor Yellow "Подключение к Exchange не удалось: ->" $error[0].Exception.Message
    }
    if ($statusconnect -eq "OK")
    {
        Import-PSSession $global:Session
    }
}

#
#Функция для проверки на пустое значение
#
function empty ($source)
{

    if ($source -eq $null)
    {
        ""
    }else
    {
        $source
    }

}

#
# Основная функция загрузки данных в ListView
#
function inactive ($enabled = "all")
{
$inactives = ""
$connection = testconnection
if ($connection -eq "on")
{
    $ok = [System.Windows.Forms.MessageBox]::Show("Показать с проверкой из Exchange?","Уведомление","YESNO","Information")
    if ($ok -eq "Yes")
    {
        $connection = "on"
    }
    if ($ok -eq "No")
    {
        $connection = "off"
    }
}

    if (!($xMF_inactive_txtbox_ou.Text -eq ""))
    {
         switch ($enabled)
            {
                "all" { $inactives = Get-ADUser -SearchBase $xMF_inactive_txtbox_ou.Text -filter * -Properties mail,DisplayName,LastLogonDate,Description,enabled,SamAccountname,DistinguishedName,Whencreated | sort LastLogonDate -Descending | select mail,DisplayName,LastLogonDate,Description,Enabled,SamAccountname,DistinguishedName,Whencreated}
                "false" { $inactives = Get-ADUser -SearchBase $xMF_inactive_txtbox_ou.Text -filter {(Enabled -eq $false)} -Properties mail,DisplayName,LastLogonDate,Description,enabled,SamAccountname,DistinguishedName,Whencreated | sort LastLogonDate -Descending | select mail,DisplayName,LastLogonDate,Description,Enabled,SamAccountname,DistinguishedName,Whencreated }
                "time" {$inactives = Get-ADUser -SearchBase $xMF_inactive_txtbox_ou.Text -Filter {(Enabled -eq $True)} -Properties mail,DisplayName,LastLogonDate,Description,enabled,SamAccountname,DistinguishedName,Whencreated | ?{($_.LastLogonDate -lt $time) -and ($_.Whencreated -lt $timeCreated)} | sort LastLogonDate -Descending | select mail,DisplayName,LastLogonDate,Description,enabled,SamAccountname,DistinguishedName,Whencreated}
                "default" { [System.Windows.Forms.MessageBox]::Show("Неправильно значение в функции","Уведомление","OK","ERROR")}
            }

            #Для точной фильтрации
            if (!($xMF_inactive_txtbox_filter.Text.Length -eq 0))
            {
                [array]$list = $xMF_inactive_txtbox_filter.Text -split ","
                for ($j=0 ;$j -ne $list.Count ; $j++)
                {
                [string]$list2 = $list[$j]
                $inactives = $inactives | ?{!($_.DistinguishedName -like "*$list2*")}
                }
            }

            if (!($xMF_inactive_txtbox_filter_Descr.Text.Length -eq 0))
            {
                [array]$list = $xMF_inactive_txtbox_filter_Descr.Text -split ","
                for ($j=0 ;$j -ne $list.Count ; $j++)
                {
                [string]$list2 = $list[$j]
                $inactives = $inactives | ?{!($_.Description -like "*$list2*")}
                }
            }

    $xMF_prBar.Value = 0
    $xMF_prBar.Maximum = $inactives.Count
    $SCount = $inactives.Count
    $Spisok = @()
    $Global:spisoksort= @()
    $xMF_Inactive_button_clear.Content = "Всего записей: "+$SCount

    if (!($inactives.Count -lt 0))
    {
    for ($i=0 ;$i -ne $inactives.Count ; $i++)
    {
        if ($inactives[$i].LastLogonDate -eq $null)
        {
            [string]$dt = ""
        }
        else
        {
            [string]$dt = ($inactives[$i].LastLogonDate).ToString('dd/MM/yyyy')
        }

        
        if ($connection -eq "on")
        {
            if ($inactives[$i].mail -like "*@sakhalin.gov.ru"){
            try{$mailbox = Get-MailboxStatistics $inactives[$i].SamAccountname -ErrorAction Ignore}catch{}
            }

            if (!($mailbox.LastLogonTime -eq $null))
            {
                $mailbox = $mailbox.LastLogonTime.ToString('dd/MM/yyyy')
                $mailbox2 = $mailbox.LastLogonTime
            }else
            {
                $mailbox = ""
            }
        }else
        {
            $mailbox = ""
        }

        
        $Spisok = [PSCustomObject]@{
                'Enabled' = $inactives[$i].Enabled
                'DisplayName'= empty($inactives[$i].DisplayName)
                'SamAccountname' = empty($inactives[$i].SamAccountname)
                'LastLogonAD' = empty($dt)
                'LastLogonEX' = $mailbox
                'WhenCreated' = $inactives[$i].Whencreated.ToString('dd/MM/yyyy')
                'Description' = empty($inactives[$i].Description)
                'mail' = empty($inactives[$i].mail)
                'DistinguishedName' = empty($inactives[$i].DistinguishedName)
                'status' = ""
                }
                $xMF_lstv_inactive.Items.Add($Spisok)
                $xMF_lstv_inactive.Items.Refresh()
        
        if ($connection -eq "on")
        {
            $SCount2 = $SCount2 + 1
            $xMF_prBar.Value = $SCount2
            $xMF_label_prBar.Content = "Обработано: "+$SCount2+" из "+$SCount
            $Form_Main.Dispatcher.Invoke([action]{},"Render")
        }elseif ($connection -eq "off")
        {
            $SCount = $SCount - 1
            Write-Host $SCount
        }
        
     }
     }#if (!($inactives.Count -lt 0))


    }
}


$xMF_inactive_btn_sort.add_click({

if (!($xMF_lstv_inactive.Items.Count -eq 0))
{

if ($xMF_inactive_rb_displayname.IsChecked)
{

    sortirovka -sort "DisplayName"

}

if ($xMF_inactive_rb_lastloggonad.IsChecked)
{

    sortirovka -sort "LastLogonAD"

}

if ($xMF_inactive_rb_lastlogonex.IsChecked)
{

    sortirovka -sort "LastLogonEX"

}

if ($xMF_inactive_rb_whencreated.IsChecked)
{

    sortirovka -sort "WhenCreated"

}


}

})



$xMF_inactive_exch_btn.add_click({
    $testconn = testconnection
    if ($testconn -eq "off")
    {
        $ok = [System.Windows.Forms.MessageBox]::Show("Подключиться к почтовому серверу?","Уведомление","OKCANCEL","Information")
        if ($ok -eq "OK")
        {
            exch_connect
            testconnection
        }
    }
    if ($testconn -eq "on")
    {
        $ok = [System.Windows.Forms.MessageBox]::Show("Отключиться от почтового сервера?","Уведомление","OKCANCEL","Information")
        if ($ok -eq "OK")
        {
            Remove-PSSession $Global:Session
            testconnection
        }
    }

})

$xMF_inactive_months.add_TextChanged({

if ($xMF_inactive_months.Text[0] -notmatch "[0-9]")
{
    $xMF_inactive_months.Text = ""

}elseif (!($xMF_inactive_months.Text[1] -eq $null))
{
    if ($xMF_inactive_months.Text[1] -notmatch "[0-9]")
    {
        $xMF_inactive_months.Text = $xMF_inactive_months.Text[0]
    }
}

})

$xMF_inactive_OU.add_click({
Get-OUDialog-start
$windowSelectOU.Show()
$windowSelectOU.Activate()
})

$xMF_inactive_load_disabled.add_click({

inactive -enabled "false"
$xMF_Inactive_button_clear.Content = "Всего записей: "+$xMF_lstv_inactive.items.count
})


$xMF_inactive_load_6months.add_click({
$time = ([System.DateTime]::Today).AddMonths(-$xMF_inactive_months.Text)
$timeCreated = ([System.DateTime]::Today).AddMonths(-$xMF_inactive_months.Text+3)
inactive -enabled "time"
$xMF_Inactive_button_clear.Content = "Всего записей: "+$xMF_lstv_inactive.items.count
})

#
#Кнопка удаления записей
#
$xMF_lstv_inactive_del.add_click({
    #Проверка количества выделенных записей
    $count = $xMF_lstv_inactive.SelectedItems.Count


    if ($count -gt 0)
    {
        #Пока количество выделеных записей больше нуля, удалять записи
        while ($count -gt 0)
        {
            $xMF_lstv_inactive.Items.Remove($xMF_lstv_inactive.SelectedItem)
            $count = $xMF_lstv_inactive.SelectedItems.Count
        }
    }

    $xMF_Inactive_button_clear.Content = "Всего записей: "+$xMF_lstv_inactive.items.count
})


#
#Кнопка очистки lstv
#
$xMF_Inactive_button_clear.add_click({

    if (!($xMF_lstv_inactive.items.Count -eq 0))
    {
    $info = [System.Windows.Forms.MessageBox]::Show("Очистить список?","Уведомление","OKCANCEL","information")
    if ($info -eq "OK")
    {
        $xMF_lstv_inactive.items.Clear()
        $xMF_Inactive_button_clear.Content = "Всего записей: "+$xMF_lstv_inactive.items.count
        $xmf_label_prBar.Content = ""
        $xMF_prBar.Value = 0
    }
    }

})

#
# Основная кнопка
#
$xMF_inactive_go.add_click({

if (!($xMF_lstv_inactive.Items.Count -eq 0))
{
    if ($xMF_inactive_radio_disable.IsChecked -eq $true)
    {

    $ok = [System.Windows.Forms.MessageBox]::Show("Обработать (Блокировка)?","Уведомление","OKCANCEL","information")
    if ($ok -eq "OK")
    {
    #Переменные для счетчиков выполнения
    $count = 0
    $counterr = 0
    $xMF_label_prBar.Content = ""

        for ($i=0 ;$i -ne $xMF_lstv_inactive.Items.Count;  $i++)
        {
            $complete = $False
            if (($xMF_lstv_inactive.Items[$i].Enabled -eq $True) -and !($xMF_lstv_inactive.Items[$i].status -eq "disabled"))
            {
                Write-Host "----------------------" $xMF_lstv_inactive.Items[$i].SamAccountname "--------------------------------"
                $user = $xMF_lstv_inactive.Items[$i].SamAccountname
                try
                {
                    Disable-ADAccount -Identity $user
                    write-host "Пользователь отключен:" $user
                    $complete = $True
                }catch{write-host -BackgroundColor red -ForegroundColor Yellow "Ошибка отключения ->" $Error[0].Exception.Message}

                if ($complete -eq $True)
                {
                    
                    try
                    {
                        $lastlogon = $xMF_lstv_inactive.Items[$i].LastLogonAd
                        $lastlogonEX = $xMF_lstv_inactive.Items[$i].LastLogonEX
                        $Created = $xMF_lstv_inactive.Items[$i].Whencreated
                        if (!($lastlogon -eq "") -and !($lastlogon -eq $null)){
                        Set-ADUser $user -Description "Посл. вход в УЗ: $lastlogon, УЗ создана: $Created, Почта: $lastlogonEX"
                        write-host "Description: Посл. вход в УЗ: $lastlogon, УЗ создана: $Created, Почта: $lastlogonEX"
                        }else{Set-ADUser $user -Description "Посл. вход в УЗ: Никогда, УЗ создана: $Created, Почта: $lastlogonEX"
                        write-host "Description: Посл. вход в УЗ: Никогда, УЗ создана: $Created, Почта: $lastlogonEX" }
                        $count++

                    }catch{write-host -BackgroundColor red -ForegroundColor Yellow "Ошибка Description ->" $Error[0].Exception.Message
                    $counterr++}

                    if ($xMF_inactive_cbox_dannie.IsChecked -eq $true)
                    {
                        try
                        {
                            Set-ADUser $user -Office $null -OfficePhone $null -Department $null -Title $null -City $null -StreetAddress $null -PostalCode $null -Company $null
                            write-host "Данные у пользователя очищены:" $user
                        }catch{write-host -BackgroundColor red -ForegroundColor Yellow "Ошибка очистки ->" $Error[0].Exception.Message}
                    }

                    if ($xMF_inactive_cbox_delman.IsChecked -eq $true)
                    {
                        managers -user $user -empty $true -short $true
                    }

                    if ($xMF_inactive_cbox_delgr.IsChecked -eq $true)
                    {
                        chkboxdeletegroups -user $user -exception "mdaemon"
                    }

                    $xMF_label_prBar.Content = "Обработано: $count , Пропущено: $counterr"
                    $Form_Main.Dispatcher.Invoke([action]{},"Render")
                    $xMF_lstv_inactive.Items[$i].status = "disabled"

                }
            }

        }# for ($i=0 ;$i -ne $xMF_lstv_inactive.Items.Count;  $i++)
            $ok = [System.Windows.Forms.MessageBox]::Show("Очистить записи со статусом Disabled?","Уведомление","OKCANCEL","Information")
            if ($ok -eq "OK")
            {
            $item = 0
            $icount = $xMF_lstv_inactive.Items.Count
            while ($item -lt $icount)
            {
                if ($xMF_lstv_inactive.Items[$item].status -eq "disabled")
                {
                    $xMF_lstv_inactive.Items.Remove($xMF_lstv_inactive.items[$item])
                }
                else
                {
                    $item++
                }
            }
            $xMF_lstv_inactive.Items.Refresh()
            $xMF_Inactive_button_clear.Content = "Всего записей: "+$xMF_lstv_inactive.items.count
            }
    $xMF_lstv_inactive.Items.Refresh()
    } #if ($ok -eq "OK")
    } # ($xMF_inactive_radio_disable.IsChecked -eq $true)

    
    if ($xMF_inactive_radio_move.IsChecked -eq $true)
    {
    $ok = [System.Windows.Forms.MessageBox]::Show("Обработать (Перенос)?","Уведомление","OKCANCEL","information")
    if ($ok -eq "OK")
    {
    #Переменные для счетчиков выполнения
    $count = 0
    $counterr = 0
    $xMF_label_prBar.Content = ""

        for ($i=0 ;$i -ne $xMF_lstv_inactive.Items.Count;  $i++)
        {
            $complete = $False
            if (($xMF_lstv_inactive.Items[$i].Enabled -eq $False))
            {
                $user = $xMF_lstv_inactive.Items[$i].SamAccountname
                Write-Host "----------------------" $xMF_lstv_inactive.Items[$i].SamAccountname "--------------------------------"

                if ($xMF_inactive_cbox_dannie.IsChecked -eq $true)
                {
                    try
                    {
                        Set-ADUser $user -Office $null -OfficePhone $null -Department $null -Title $null -City $null -StreetAddress $null -PostalCode $null -Company $null
                        write-host "Данные у пользователя очищены:" $user
                    }catch{write-host -BackgroundColor red -ForegroundColor Yellow "Ошибка очистки ->" $Error[0].Exception.Message}
                }

                if ($xMF_inactive_cbox_delman.IsChecked -eq $true)
                {
                    managers -user $user -empty $true -short $true
                }

                if ($xMF_inactive_cbox_delgr.IsChecked -eq $true)
                {
                    chkboxdeletegroups -user $user -exception "mdaemon"
                }

                try
                {
                    $TargetOU = "OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"
                    switch ($xMF_inactive_cmbx_move.items[$xMF_inactive_cmbx_move.SelectedIndex].Content)
                    {
                        'OU_RT' {$TargetOU = "OU=OU_RT,OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"}
                        'GBU' {$TargetOU = "OU=GBU,OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"}
                        'MO' {$TargetOU = "OU=MO,OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"}
                        'ALL' {$TargetOU = "OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"}
                    }
                                        
                    $userdn = $xMF_lstv_inactive.Items[$i].distinguishedName
                    Move-ADObject -Identity $userdn -TargetPath $TargetOU
                    $count++
                    Write-Host "Пользователь перенесен: $user"
                    $xMF_lstv_inactive.Items[$i].status = "Moved"
                }catch
                {
                    write-host "Перенос - " $Error[0].Exception
                    $counterr++
                }

                $xMF_label_prBar.Content = "Перенесено: $count , Ошибок: $counterr"
                $Form_Main.Dispatcher.Invoke([action]{},"Render")
                

            } # if (($xMF_lstv_inactive.Items[$i].Enabled -eq "True"))

        } # for ($i=0 ;$i -ne $xMF_lstv_inactive.Items.Count;  $i++)
            
    } #if ($ok -eq "OK")
    }# -----END if ($xMF_inactive_radio_move.IsChecked -eq $true)

}# -------END (!($xMF_lstv_inactive.Items.Count -eq 0))


})

##################################################################
# Кнопка для очистки данных у пользователей в ОУшке Inactive
##################################################################

$xMF_Inactive_btn_cleardannie.add_click({

        $TargetOU = "OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"
        switch ($xMF_inactive_cmbx_move.items[$xMF_inactive_cmbx_move.SelectedIndex].Content)
        {
            'OU_RT' {$TargetOU = "OU=OU_RT,OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"}
            'GBU' {$TargetOU = "OU=GBU,OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"}
            'MO' {$TargetOU = "OU=MO,OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"}
            'ALL' {$TargetOU = "OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"}
        }

        if (($xMF_inactive_cbox_dannie.IsChecked -eq $true) -or ($xMF_inactive_cbox_delman.IsChecked -eq $true) -or ($xMF_inactive_cbox_delgr.IsChecked -eq $true))
        {


        $ok = [System.Windows.Forms.MessageBox]::Show("Данные у УЗ будут очищены:`n$TargetOU","Уведомление","OKCANCEL","information")
        if ($ok -eq "OK")
        {
            $spisok = Get-ADUser -SearchBase $TargetOU -Filter * -Properties title,Department,OfficePhone,Office,Company,memberof,manager,directreports

            if ($xMF_inactive_cbox_dannie.IsChecked -eq $true)
            {
                $spisok | ?{!($_.title -eq $null) -and !($_.title -eq "") -or !($_.company -eq $null) -and !($_.company -eq "") -or !($_.Department -eq $null) -and !($_.Department -eq "") -or !($_.OfficePhone -eq $null) -and !($_.OfficePhone -eq "") -or !($_.Office -eq $null) -and !($_.Office -eq "")} | Set-ADUser -Office $null -OfficePhone $null -Department $null -Title $null -City $null -StreetAddress $null -PostalCode $null -Company $null
            }

            if ($xMF_inactive_cbox_delman.IsChecked -eq $true)
            {
                $spisok | select directreports,manager,samaccountname |?{!($_.directreports[0] -eq $null) -and ($_.manager -eq $null)} | %{managers -user $_.SamAccountName -empty $true -short $true}
            }

            if ($xMF_inactive_cbox_delgr.IsChecked -eq $true)
            {
                $spisok | select memberof,SamAccountName | ?{!($_.memberof[0] -eq $null)} | %{chkboxdeletegroups -user $_.SamAccountName -exception "mdaemon"}
            }

        }

        }else
        {
            [System.Windows.Forms.MessageBox]::Show("Ни одно свойство не выбрано","Уведомление","OK","information")
        }

})


#
# Кнопка переноса пользователя в Inactive через МЕНЮ в LISTVIEW INACTIVE
#
$xMF_lstv_inactive_move.add_click({
#Проверка  выделена ли вообще хотя бы 1 строка
if (!($xMF_lstv_inactive.selecteditems.count -eq 0))
{
    $ok = ""
    #$xMF_prBar.Value = 0
    #$xMF_prBar.Maximum = $xMF_lstv_inactive.selecteditems.count - 1

#Проверка на количество выделеных строк
for ($i=0 ;$i -ne $xMF_lstv_inactive.selecteditems.count;  $i++)
{
    #проверка на заполненость определенных ячеек
    if (($xMF_lstv_inactive.selecteditems[$i].Enabled -eq $False) -and !($xMF_lstv_inactive.selecteditems[$i].status -eq "Moved"))
    {
        #Если соглашаемся с переносом
        if (!($ok -eq "OK"))
        {
        $ok = [System.Windows.Forms.MessageBox]::Show("Обработать (Перенос) пользователя?","Уведомление","OKCANCEL","information")
        }
            if ($ok -eq "OK")
            {
    
                    $user = $xMF_lstv_inactive.selecteditems[$i].SamAccountname
                    Write-Host "----------------------" $xMF_lstv_inactive.selecteditems[$i].SamAccountname "--------------------------------"

                    #Выполняем очистку данных у пользователей если установлен флажок на очистку данных
                    if ($xMF_inactive_cbox_dannie.IsChecked -eq $true)
                    {
                        try
                        {
                            Set-ADUser $user -Office $null -OfficePhone $null -Department $null -Title $null -City $null -StreetAddress $null -PostalCode $null -Company $null
                            write-host "Данные у пользователя очищены:" $user
                        }catch{write-host -BackgroundColor red -ForegroundColor Yellow "Ошибка очистки ->" $Error[0].Exception.Message}
                    }

                    #Выполняем очистку менеджеров у пользователей если установлен флажок на очистку менеджеров
                    if ($xMF_inactive_cbox_delman.IsChecked -eq $true)
                    {
                        managers -user $user -empty $true -short $true
                    }
                    
                    #Выполняем очистку групп у пользователей если установлен флажок на очистку групп
                    if ($xMF_inactive_cbox_delgr.IsChecked -eq $true)
                    {
                        chkboxdeletegroups -user $user -exception "mdaemon"
                    }
                    
                    #Пробуем перенести учетки
                    try
                    {
                        $TargetOU = "OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"
                        switch ($xMF_inactive_cmbx_move.items[$xMF_inactive_cmbx_move.SelectedIndex].Content)
                        {
                            'OU_RT' {$TargetOU = "OU=OU_RT,OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"}
                            'GBU' {$TargetOU = "OU=GBU,OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"}
                            'MO' {$TargetOU = "OU=MO,OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"}
                            'ALL' {$TargetOU = "OU=Users,OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"}
                        }
                                            
                        $userdn = $xMF_lstv_inactive.selecteditems[$i].distinguishedName
                        Move-ADObject -Identity $userdn -TargetPath $TargetOU
                        Write-Host "Пользователь перенесен: $user - " $xMF_inactive_cmbx_move.items[$xMF_inactive_cmbx_move.SelectedIndex].Content
                        $xMF_lstv_inactive.selecteditems[$i].status = "Moved"
                    }catch
                    {
                        write-host "Перенос - " $Error[0].Exception
                        $counterr++
                    }
                    #$xMF_prBar.Value += 1
                    $xMF_label_prBar.Content = "Пользователей перенесено - " + $xMF_lstv_inactive.selecteditems.count
                    $xMF_lstv_inactive.Items.Refresh()
                    #$Form_Main.Dispatcher.Invoke([action]{},"Render")
    
        } # $ok
        else {
            $i = $xMF_lstv_inactive.selecteditems.count-1
        }
    } # if ($xMF_lstv_inactive.Items[$xMF_inactive_cmbx_move.SelectedIndex].Enabled -eq "False")
    
} # for ($i=0 ;$i -ne $xMF_lstv_inactive.selecteditems.count-1;  $i++)
} # if (!($xMF_lstv_inactive.selecteditems.count -eq 0))
else
{
    [System.Windows.Forms.MessageBox]::Show("Список пустой","Уведомление","OK","Warning")
}

})


#
# Кнопка отключения пользователя через МЕНЮ в LISTVIEW INACTIVE
#
$xMF_lstv_inactive_disable.add_click({

#Проверка  выделена ли вообще хотя бы 1 строка
if (!($xMF_lstv_inactive.selecteditems.count -eq 0))
{
    #$xMF_prBar.Value = 0
    #$xMF_prBar.Maximum = $xMF_lstv_inactive.selecteditems.count - 1

#Проверка на количество выделеных строк
for ($i=0 ;$i -ne $xMF_lstv_inactive.selecteditems.count;  $i++)
{

    $complete = $False
    if (($xMF_lstv_inactive.selecteditems[$i].Enabled -eq "True") -and !($xMF_lstv_inactive.selecteditems[$i].status -eq "disabled"))
    {
        Write-Host "----------------------" $xMF_lstv_inactive.selecteditems[$i].SamAccountname "--------------------------------"
        $user = $xMF_lstv_inactive.selecteditems[$i].SamAccountname
        try
        {
            Disable-ADAccount -Identity $user
            write-host "Пользователь отключен:" $user
            $complete = $True
        }catch{write-host -BackgroundColor red -ForegroundColor Yellow "Ошибка отключения ->" $Error[0].Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Ошибка отключения УЗ - "+$user,"Уведомление","OK","information")
        }

        if ($complete -eq $True)
        {
            try
            {
                $lastlogon = $xMF_lstv_inactive.selecteditems[$i].LastLogonAd
                $lastlogonEX = $xMF_lstv_inactive.selecteditems[$i].LastLogonEX
                $Created = $xMF_lstv_inactive.selecteditems[$i].Whencreated
                if (!($lastlogon -eq "") -and !($lastlogon -eq $null)){
                Set-ADUser $user -Description "Посл. вход в УЗ: $lastlogon, УЗ создана: $Created, Почта: $lastlogonEX"
                write-host "Description: Посл. вход в УЗ: $lastlogon, УЗ создана: $Created, Почта: $lastlogonEX"
                }else{Set-ADUser $user -Description "Посл. вход в УЗ: Никогда, УЗ создана: $Created, Почта: $lastlogonEX"
                write-host "Description: Посл. вход в УЗ: Никогда, УЗ создана: $Created, Почта: $lastlogonEX" }
                $xMF_lstv_inactive.selecteditems[$i].status = "disabled"
                #$xMF_lstv_inactive.Items.Refresh()
                #$xMF_label_prBar.Content = "Обработано: $user"      
                        
            }catch{write-host -BackgroundColor red -ForegroundColor Yellow "Ошибка Description ->" $Error[0].Exception.Message
            $xMF_label_prBar.Content = "НЕ обработано: $user"
            #[System.Windows.Forms.MessageBox]::Show("Ошибка в прописывании Description у "+$user,"Уведомление","OK","information")
            }

            if (($xMF_inactive_cbox_delman.IsChecked -eq $true) -and ($complete -eq $True))
            {
                managers -user $user -empty $true -short $true
            }

            if (($xMF_inactive_cbox_delgr.IsChecked -eq $true) -and ($complete -eq $True))
            {
                chkboxdeletegroups -user $user -exception "mdaemon"
            }


            
        }
    }
    
}
$xMF_lstv_inactive.Items.Refresh()
}

})

$xMF_inactive_btn_delstatus.add_click({

if (!($xMF_lstv_inactive.Items.Count -eq 0))
{
    $ok = [System.Windows.Forms.MessageBox]::Show("Очистить записи со статусом Disabled?","Уведомление","OKCANCEL","Information")
            if ($ok -eq "OK")
            {
                $item = 0
                $icount = $xMF_lstv_inactive.Items.Count
                while ($item -lt $icount)
                {
                    if ($xMF_lstv_inactive.Items[$item].status -eq "disabled")
                    {
                        $xMF_lstv_inactive.Items.Remove($xMF_lstv_inactive.items[$item])
                    }
                    else
                    {
                        $item++
                    }
                }
                $xMF_lstv_inactive.Items.Refresh()
                $xMF_Inactive_button_clear.Content = "Всего записей: "+$xMF_lstv_inactive.items.count
                $xMF_label_prBar.Content = ""
                $xMF_prBar.Value = 0
            }

}
}) 