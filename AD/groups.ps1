Function getgroups
{
    try
    {
        $groups_load_user1 = Get-ADUser $xMF_textbox_Group_newuser.Text -Properties memberof
        $groups_load_user2 = Get-ADUser $xMF_textbox_Group_existuser.Text -Properties memberof
    }
    catch
    {
        [System.Windows.Forms.MessageBox]::Show($error[0].Exception.Message,"Ошибка","OK","Warning")
        Write-Host -ForegroundColor Black -BackgroundColor Yellow "ERROR -> " $error[0].Exception.Message
        return "ERROR"
    }
            if ($groups_load_user2.MemberOf[0] -eq $null)
            {
                Write-Host "У пользователя "$xMF_textbox_Group_existuser.Text" нет групп"
                return "0"
            }
            else
            {
                $groups_load_compare = Compare-Object -ReferenceObject $groups_load_user1.MemberOf -DifferenceObject $groups_load_user2.MemberOf -IncludeEqual
                $global:er = 0
                
                foreach ($strings in $groups_load_compare)
                {
                    try
                    {
                        $strings2 = Get-ADGroup -Identity $strings.InputObject
                        if ($strings.SideIndicator -eq "=>")
                        {
                            write-host -BackgroundColor red -ForegroundColor white "У пользователя" $groups_load_user1.Name "нет групп" $strings2.SamAccountName
                            $xMF_groups_lstv_user1.Items.Add($strings2.SamAccountName)
                        }
                    }
                    catch
                    {
                        $global:er++
                        Write-Host "ERROR -> " $error[0].Exception.Message
                    }
                }
                return "1"
            }
}




function getusersgroups
{

    try
    {
        
    }
    catch
    {
        
    }

}



function chkboxdeletegroups ($user = $null,[string]$exception = $null)
{
    if (!($user -eq $null))
    {
        $usergroups = Get-ADUser $user -Properties memberof | select SamAccountName,memberof
        if (!($usergroups.memberof[0] -eq $null))
        {
            foreach ($usergroup in $usergroups.memberof)
            {
                if ($usergroup -notlike "*$exception*")
                {
                    try
                    {
                        Remove-ADGroupMember -Identity $usergroup -Members $usergroups.SamAccountName -Confirm:$false
                        write-host "Группа удалена -> Пользователь" $usergroups.SamAccountName "группа" $usergroup
                    }
                    catch
                    {
                        Write-Host "ERROR -> " $error[0].Exception.Message
                    }
                 }
                 else
                 {
                    write-host "Группа НЕ удалена согласно правилу -> Пользователь" $usergroups.SamAccountName "группа" $usergroup
                 }
            }
        }
        else
        {
            write-host -BackgroundColor red -ForegroundColor white "У пользователя" $usergroups.SamAccountName "нет групп"
        }
    }
}


function managers ($user = $null,$empty = $false,$short=$false)
{

<#
.SYNOPSIS
	Очистка поля менеджер в учетной записи.

.DESCRIPTION
	Функция для очистки поля менеджер в учетной записи, если является начальником удалит свои данные у подчиненных.

.PARAMETER user
Логин пользователя - SamAccountName

.PARAMETER empty
Выводит информационныые данные, если у пользоватля нет менеджеров.

.PARAMETER short
Выводит информационные данные в укороченном виде, если пользоватль является начальником
#>

    #Если поле User заполнено
    if (!($user -eq $null))
    {
        $usermanager = Get-ADUser $user -Properties directReports,manager,CN | select directReports,manager,SamAccountName,CN
        $t=0
        $t1=0

        #Если у пользователя заполнено поле Менеджер
        if (!($usermanager.manager.count -eq 0))
        {
            try
            {
                Set-ADUser -Identity $user -Manager $null
                write-host "Менеджер удален -> Пользователь" $usermanager.SamAccountName "менеджер" $usermanager.manager
            }
            catch
            {
                Write-Host "ERROR -> " $error[0].Exception.Message
            }
        }
        elseif ($empty -eq $True)
        {
            write-host "Менеджера НЕТ -> Пользователь" $usermanager.SamAccountName
        }

        #Если пользователь является менеджером для других учетных записей
        if (!($usermanager.directReports.count -eq 0))
        {
            write-host "-------------------- Удаление начальника "$usermanager.SamAccountName" у пользователей ------------------"
            #Перебор всех записей
            foreach ($direcreports in $usermanager.directReports)
            {
                $usermanager2 = (Get-ADUser $direcreports).SamAccountName
                
                try
                {
                    Set-ADUser -Identity $usermanager2 -Manager $null
                    $t++
                    if ($short -eq $false)
                    {
                        write-host "Менеджер удален -> Пользователь" $usermanager2 "менеджер" $usermanager.CN "("$usermanager.SamAccountName")"
                    }
                }
                catch
                {
                    if ($short -eq $false)
                    {
                        Write-Host "ERROR -> " $error[0].Exception.Message
                    }
                    $t1++
                }
            }

            if ($short -eq $True)
            {
                write-host Всего удалено менеджеров: $t"," ошибок: $t1
            }
            write-host "-------------------- Завершено Удаление начальника у пользователей ------------------"
        }# КОНЕЦ if (!($usermanager.directReports.count -eq 0))
        elseif ($empty -eq $True)
        {
            write-host "Нет подчиненных -> Пользователь" $usermanager.SamAccountName
        }# КОНЕЦ elseif ($empty -eq $True)

    }

}