# Exchange script - DO NOT DELETE THIS

[xml]$xaml3 = @"
<Window x:Name="Exchange"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp1"
       Title="Exchange" Height="450" Width="800" ResizeMode="NoResize" WindowStyle="None">
    <Grid UseLayoutRounding="False">
        <GroupBox HorizontalAlignment="Left" Height="450" VerticalAlignment="Top" Width="800" BorderBrush="Black" Margin="0" Header="Exchange">
            <Grid HorizontalAlignment="Left" Height="431.465" VerticalAlignment="Top" Width="790" Margin="0,0,-2,-4.425">
                <Button x:Name="Exchange_btn_addmail" Content="Создать почту" HorizontalAlignment="Left" Margin="285.24,367.401,0,0" VerticalAlignment="Top" Width="125.458" Height="44.674"/>
                <Button x:Name="Exchange_btn_exec_database" Content="&lt;&lt;&lt; Прописать" HorizontalAlignment="Left" VerticalAlignment="Top" Width="125.458" ToolTip="Присвоение базы данных пользователю" Margin="285.24,142.105,0,0" Height="22.019"/>
                <Button x:Name="Exchange_btn_clear_users" Content="&lt;&lt;&lt; Очистить" HorizontalAlignment="Left" VerticalAlignment="Top" Width="125.458" Margin="285.24,169.124,0,0" ToolTip="Очистить список пользователей"/>
            </Grid>
        </GroupBox>
        <ListView x:Name="Exchange_listview_users" Margin="10,49.585,0,0" HorizontalAlignment="Left" Width="258.344" Height="380.227" VerticalAlignment="Top">
            <ListView.ContextMenu>
                <ContextMenu>

                    <MenuItem x:Name="exchange_lstv_menu_user_datexec" Header="Присвоить базу" ToolTip="Прописать базу данных пользователю"/>
                    <MenuItem x:Name="exchange_lstv_menu_user_add" Header="Создать почту" ToolTip="Создать почту пользователям"/>
                    <MenuItem x:Name="exchange_lstv_menu_user_delete" Header="Удалить" ToolTip="Удалить пользователя из списка"/>

                </ContextMenu>
            </ListView.ContextMenu>
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="samaccountname" DisplayMemberBinding="{Binding samaccountname}" Width="100"/>
                    <GridViewColumn Header="database" DisplayMemberBinding="{Binding database}" Width="100"/>
                    <GridViewColumn Header="check" DisplayMemberBinding="{Binding check}" Width="40"/>
                </GridView>
            </ListView.View>
        </ListView>
        <Button x:Name="Exchange_btn_close" Content="Закрыть" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="715,425.04,0,0" ToolTip="Закрытие окна"/>
        <Button x:Name="Exchange_btn_load_users" Content="Загрузить пользователей" HorizontalAlignment="Left" Margin="10,24.625,0,0" VerticalAlignment="Top" Width="258.344" RenderTransformOrigin="0.497,0.526" ToolTip="Загрузка пользователей в список с главного окна"/>
        <ListView x:Name="Exchange_listview_database" HorizontalAlignment="Left" Height="380.465" Margin="448.732,39.575,0,0" VerticalAlignment="Top" Width="341.268" SelectionMode="Single">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="database" DisplayMemberBinding="{Binding database}" Width="100"/>
                    <GridViewColumn Header="Size" DisplayMemberBinding="{Binding size}" Width="100"/>
                    <GridViewColumn Header="available" DisplayMemberBinding="{Binding available}" Width="99"/>
                </GridView>
            </ListView.View>
        </ListView>
        <Button x:Name="Exchange_btn_database_load" Content="Загрузить информацию о базах данных" HorizontalAlignment="Left" Margin="448.732,14.615,0,0" VerticalAlignment="Top" Width="341.268" ToolTip="Загрузить информацию в список о базах данных Exchange"/>

    </Grid>
</Window>

"@


$XReader3=(New-Object System.Xml.XmlNodeReader $xaml3)
$Form_3=[Windows.Markup.XamlReader]::Load( $XReader3 )
$xaml3.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{Set-Variable -Name "xMF3_$($_.Name)" -Value $Form_3.FindName($_.Name)}


############################################################
#Начальная инициализации окон
############################################################

# Переменная для проверки на администрирование, проверка специально, что бы при каждом нажатии не подвисало
if ((Get-Variable -Name sam -ErrorAction Ignore) -eq $null)
{
   $Global:sam = ""
}

$con = testconnection
if ($con -eq "on")
{
    $xMF_menuitem_exchange.Background = "green"
}
else
{
    $xMF_menuitem_exchange.Background = "red"
}

$xMF3_Exchange_btn_close.add_click({

    $Form_3.Hide()

})

$Form_3.add_MouseLeftButtonDown({
    $this.dragmove()
})

############################################################



#
#Кнопка открытия окна для Exchange
#
$xMF_menuitem_exchange.add_click({
    

        #Пробуем сделать условие
        try
        {
            #Проверка на нахождение пользователя в группе администраторов Exchange
            
            if (($Global:sam -eq "") -or ($Global:sam -eq $null))
            {
                $Global:sam = Get-ADGroupMember "Exchange Organization Administrators" -Server (Get-ADForest).RootDomain -Recursive | ?{$_.samaccountname -like "$env:UserName"} | select SamAccountName
            }

            if (($Global:sam.SamAccountName -eq "$env:UserName") -and !($Global:sam -eq "") -and !($Global:sam -eq $null))
            {

            #Проверка на подключение к Exchange
                $con = testconnection
                if ($con -eq "on")
                {

                        $Form_3.Show()
                        $Form_3.Activate()

                } #(testconnection -eq "on")
                else
                {
                    Write-Host -BackgroundColor red -ForegroundColor white "Нет подключения к серверу Exchange"
                    $ok = [System.Windows.Forms.MessageBox]::Show("Нет подключения к серверу Exchange, подключить?","Подтверждение","OKCANCEL","Warning")
                    if ($ok -eq "OK")
                    {
                        exch_connect
                        testconnection
                    }
                }

            }
            else
            {
                [System.Windows.Forms.MessageBox]::Show("Пользователь не имеет прав на администрирование Exchange","Подтверждение","OK","information")
            }
        }
        catch
        {
            [System.Windows.Forms.MessageBox]::Show($Error[0].Exception.Message,"Ошибка","OK","Warning")
            Write-Host -BackgroundColor red -ForegroundColor white "Exchange -"$Error[0].Exception.Message
        }
})


#
# Кнопка добавления пользователей в список создания пользователя
#
$xMF3_Exchange_btn_load_users.add_click({

    #Проверяем, не пустое ли значение в lstv, а так же, что выделены какие либо элементы.
    if (!($xMF_lstv_SingleUser.Items.Count -eq 0) -or !($xMF_lstv_SingleUser.SelectedIndex -eq -1))
    {
        $count = $xMF_lstv_SingleUser.SelectedItems.Count
        $t = 0
        while ($count -gt 0)
        {
            if (!($xMF_lstv_SingleUser.SelectedItems[$t].samaccountname -eq ""))
            {
                $Spisok = [PSCustomObject]@{
                    'samaccountname' = $xMF_lstv_SingleUser.SelectedItems[$t].samaccountname
                    'database' = ""
                    'check' = ""
                    }
                $xMF3_Exchange_listview_users.Items.Add($Spisok)
            }
            else
            {
                Write-host -BackgroundColor Yellow -ForegroundColor black "Exchange - Пользователю -"$xMF_lstv_SingleUser.SelectedItems[$t].Displayname"не присвоен логин"
            }
            $count--
            $t++
            
        }
    }
    else
    {
        [System.Windows.Forms.MessageBox]::Show("Основной список пустой, либо не выделено ни одного пользователя","Подтверждение","OK","Information")
    }


})


#
#Кнопка загрузки информация по базам данных
#
$xMF3_Exchange_btn_database_load.add_click({

    if ($xMF3_Exchange_listview_database.items.IsEmpty)
    {

        $Status = Get-MailboxDatabase -Status |  select Name,DatabaseSize,AvailableNewMailboxSpace

        $array=@()
        foreach ($Status1 in $status)
        {

            $1=$status1.databasesize -split " " -replace "[(,)]",""
            $2=$status1.AvailableNewMailboxSpace -split " " -replace "[(,)]",""

        if (($1[2] / 1024MB) -lt 1)
        {
            $D_size = "{0:N3}" -f $($1[2] / 1024MB);
        }
        else
        {
            $D_size = "{0:N1}" -f $($1[2] / 1024MB);
        }


    
        if (($2[2] / 1024MB) -lt 1)
        {
            $D_size2 = "{0:N3}" -f $($2[2] / 1024MB);
        }
        else
        {
            $D_size2 = "{0:N1}" -f $($2[2] / 1024MB);
        }

            $array += [PSCustomObject]@{database=$status1.Name; Size=[string]$D_size; available=[string]$D_size2}
        }
        $array2 = $array | sort database


        foreach ($array2_1 in $array2)
        {
            $xMF3_Exchange_listview_database.items.Add($array2_1)
        }

    }
    else
    {
        $ok = [System.Windows.Forms.MessageBox]::Show("Очистить список?","Подтверждение","OKCANCEL","Information")
        if ($ok -eq "OK")
        {
            $xMF3_Exchange_listview_database.items.Clear()
        }
    }

})


#
# Кнопка и функция прописывания базы данных пользователям в списке
#

function loaddatabasetousers 
{

#Условия
    if (!($xMF3_Exchange_listview_database.items.IsEmpty))
    {
        if (!($xMF3_Exchange_listview_database.SelectedIndex -eq -1))
        {
            if (!($xMF3_Exchange_listview_users.items.IsEmpty))
            {
                #Рабочий блок кода

                #Храним количество выделенных ячеек
                $selcount = $xMF3_Exchange_listview_users.SelectedItems.count
                if ($selcount -gt 0)
                {
                    while ($selcount -gt 0)
                    {

                        $xMF3_Exchange_listview_users.SelectedItems[$selcount-1].database = $xMF3_Exchange_listview_database.Items[$xMF3_Exchange_listview_database.SelectedIndex].database
                        $selcount--
                        
                    }
                    $xMF3_Exchange_listview_users.Items.Refresh()
                }
                else
                {
                    for ($i=0 ; $i -ne $xMF3_Exchange_listview_users.Items.Count ; $i++)
                    {
                        $xMF3_Exchange_listview_users.Items[$i].database = $xMF3_Exchange_listview_database.Items[$xMF3_Exchange_listview_database.SelectedIndex].database
                    }
                    $xMF3_Exchange_listview_users.Items.Refresh()
                }


            }
            else
            {
                [System.Windows.Forms.MessageBox]::Show("Данные не кому присваивать","Уведомление","OK","Information")
            }
        }
        else
        {
            [System.Windows.Forms.MessageBox]::Show("База данных не выбрана","Уведомление","OK","Information")
        }
    }
    else
    {
        [System.Windows.Forms.MessageBox]::Show("Список баз данных пустой","Уведомление","OK","Information")
    }

}

$xMF3_Exchange_btn_exec_database.add_click({

loaddatabasetousers

})

$xMF3_exchange_lstv_menu_user_datexec.add_click({

loaddatabasetousers

})



#
# Кнопка удаления пользователей
#

$xmf3_exchange_lstv_menu_user_delete.add_click({

    if ($xMF3_Exchange_listview_users.SelectedItems.Count -gt 0)
    {

        $count = $xMF3_Exchange_listview_users.SelectedItems.Count
        if ($count -gt 0)
        {
        #Пока количество выделеных записей больше нуля, удалять записи
        while ($count -gt 0)
        {
            $xMF3_Exchange_listview_users.items.Remove($xMF3_Exchange_listview_users.SelectedItem)
            $count = $xMF3_Exchange_listview_users.SelectedItems.Count
        }
    }
    }

})

#Кнопка очистки листвью пользователи
$xmf3_Exchange_btn_clear_users.add_click({

    $xMF3_Exchange_listview_users.items.Clear()

})


#########################################################################################################################
# Блок добавления пользователей
##########################################################################################################################


function create_mail($all = $false) {

    #Проверка, выделены ли какие нибудь строки
    if (($xMF3_Exchange_listview_users.SelectedIndex -eq -1) -or ($all -eq $true))
    {
        #Цикл прохода по всем записям в листвью
        for ($i=0 ; $i -ne $xMF3_Exchange_listview_users.Items.Count ; $i++)
        {
            #Проверка значения в столбце CHECK
            if ((!($xMF3_Exchange_listview_users.Items[$i].database -eq "") -and !($xMF3_Exchange_listview_users.Items[$i].database -eq $null)))
            {
            if ((!($xMF3_Exchange_listview_users.Items[$i].check -eq "added") -and !($xMF3_Exchange_listview_users.Items[$i].check -eq "exist")))
            {
                $database = $xMF3_Exchange_listview_users.Items[$i].database
                $user = $xMF3_Exchange_listview_users.Items[$i].samaccountname
                $checkmail = Get-Mailbox $user -ErrorAction SilentlyContinue
                if ($checkmail -eq $null)
                {
                    try
                    {
                        Enable-Mailbox $user -Database $database
                        $xMF3_Exchange_listview_users.Items[$i].check = "added"
                    }
                    catch
                    {
                        Write-Host -BackgroundColor red -ForegroundColor White "Exchange -"$Error[0].Exception.Message
                    }
                }
                else
                {
                    $xMF3_Exchange_listview_users.Items[$i].check = "exist"
                }
            }# END if ((!($xMF3_Exchange_listview_users.Items[$i].check -eq "added") -and !($xMF3_Exchange_listview_users.Items[$i].check -eq "exist")))
            }# END if ((!($xMF3_Exchange_listview_users.Items[$i].database -eq "") -and (!$xMF3_Exchange_listview_users.Items[$i].database -eq $null)))
            else
            {write-host -BackgroundColor Yellow -ForegroundColor Black "Exchange - логину"$xMF3_Exchange_listview_users.Items[$i].samaccountname"не присвоена база данных"}
        }
        $xMF3_Exchange_listview_users.Items.Refresh()
    }
    else #Продолжение, если выделены несколько учетных записей
    {
        for ($i=0 ; $i -ne $xMF3_Exchange_listview_users.SelectedItems.Count ; $i++)
        {
            #Проверка значения в столбце CHECK
            if ((!($xMF3_Exchange_listview_users.SelectedItems[$i].database -eq "") -and !($xMF3_Exchange_listview_users.SelectedItems[$i].database -eq $null)))
            {
            if ((!($xMF3_Exchange_listview_users.SelectedItems[$i].check -eq "added") -and !($xMF3_Exchange_listview_users.SelectedItems[$i].check -eq "exist")))
            {
                $database = $xMF3_Exchange_listview_users.SelectedItems[$i].database
                $user = $xMF3_Exchange_listview_users.SelectedItems[$i].samaccountname
                $checkmail = Get-Mailbox $user -ErrorAction SilentlyContinue
                if ($checkmail -eq $null)
                {
                    try
                    {
                        Enable-Mailbox $user -Database $database
                        $xMF3_Exchange_listview_users.SelectedItems[$i].check = "added"
                        write-host "Exchange - почтовый ящик"$xMF3_Exchange_listview_users.SelectedItems[$i].samaccountname"создан"
                    }
                    catch
                    {
                        Write-Host -BackgroundColor red -ForegroundColor White "Exchange -"$Error[0].Exception.Message
                    }
                }
                else
                {
                    $xMF3_Exchange_listview_users.SelectedItems[$i].check = "exist"
                    write-host "Exchange - почтовый ящик"$xMF3_Exchange_listview_users.SelectedItems[$i].samaccountname"существует"
                }
            }# END if ((!($xMF3_Exchange_listview_users.SelectedItems[$i].check -eq "added") -and !($xMF3_Exchange_listview_users.SelectedItems[$i].check -eq "exist")))
            }# END if ((!($xMF3_Exchange_listview_users.SelectedItems[$i].database -eq "") -and (!$xMF3_Exchange_listview_users.SelectedItems[$i].database -eq $null)))
            else
            {write-host -BackgroundColor Yellow -ForegroundColor Black "Exchange - логину"$xMF3_Exchange_listview_users.SelectedItems[$i].samaccountname"не присвоена база данных"}
        }
        $xMF3_Exchange_listview_users.Items.Refresh()
    }


}


$xMF3_Exchange_btn_addmail.add_click({
    
    create_mail -all $true
    
})


$xMF3_exchange_lstv_menu_user_add.add_click({

    create_mail
})