Clear-Host

Add-Type -AssemblyName PresentationFramework, PresentationCore, System.Windows.Forms

$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

#Import-Module RemoteAD

# Функция для генерации пароля
. $ScriptPath\New-Password.ps1
# Функция для загрузки списка OU из АД
. $ScriptPath\Get-OUDialog.ps1
# Функция OpenDialog для выбора файла через окно (Не используется)
. $ScriptPath\Get-Filename.ps1
# Транслит
. $ScriptPath\TranslitToLat.ps1
# Создание логина из транслита
. $ScriptPath\login.ps1
# Сравнение логина и ФИО в ад, генерация логина
. $ScriptPath\find_user.ps1
# Добавление пользователей из списка в АД
. $ScriptPath\add_user_ad
# XAML
. $ScriptPath\xml.ps1

#. $ScriptPath\find_user_single.ps1
#. $ScriptPath\MakeCopy.ps1

$xMF_lstv_SingleUser.Items.Clear()
$Global:OU = "OU=Test,OU=OU_OTHER,DC=ASO,DC=RT,DC=LOCAL"
$xMF_txtbox_OU_path.Text = $Global:OU


###############################################################################
#Функции для проверки данных
###############################################################################


Function delete_spaces_FIO($SetFIO)
{
if (!($SetFIO -match "[0-9]")) 
{

    $SplitFIO = $SetFIO -split "\s"
    $FIO = @{}
    $t= 0
    for ($i=0 ; $i -ne $SplitFIO.Count ; $i++ )
    {
        if (!($SplitFIO[$i] -eq ""))
        {
            switch ($t)
            {
                "0" { $FIO += @{'Lastname' = $SplitFIO[$i]}}
                "1" { $FIO += @{'Name' = $SplitFIO[$i]}}
                "2" { $FIO += @{'SecondName' = $SplitFIO[$i]}}
            }
            $t++
        }
    }

    if ($FIO.Count -eq 3)
    {
        return $FIO.Lastname + " " + $FIO.Name + " " + $FIO.SecondName
    }
    elseif ($FIO.Count -eq 2)
    {
        return $FIO.Lastname + " " + $FIO.Name
    }
    else
    {
        return "ERROR - Нет имени и отчества"
    }

}
else
{
   return "ERROR - Цифра в ФИО"
   #return $SetFIO
}

}


function telephone ($number){
$fromwhere = "анив","александровск","долинск","холмск","корсаков","курильск","макаров","невельск","ноглик","охинск","поронайск","шахтерск","северо-курильск","смирных","томари","тымовск","углегорск","южно-курильск","южно-сахалинск"
$telephone = "842441","842434","842442","842433","842435","842454","842443","842436","842444","842437","842431","842432","842453","842452","842446","842447","842432","842455","84242"
$number = $number -replace "\D"

if (($number.Length -cle 10) -and ($number.Length -gt 6))
{

    $aa = "8" + $number -replace "\D"
    return $aa

}
elseif (($number.Length -cle 6) -and ($number.Length -ge 5))
{
    for ($i = 0 ; $i -ne $fromwhere.Count ; $i++)
    {
        if ($dep -like "*"+$fromwhere[$i]+"*")
        {
            $aa = $telephone[$i] + $number -replace "\D"
            return $aa
        }
    }
}
else
{
    return $number
}
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------

###############################################################################
#Загрузка данных в Lstv_view_singleuser из файла
###############################################################################


$xMF_Btn_LoadFiles.add_click({

if ((Test-Path $ScriptPath\users.csv) -eq $true)
{

    $csv = Import-Csv -Delimiter "	" -Path "$ScriptPath\users.csv"
    if ($xMF_lstv_SingleUser.Items.Count -cge 1)
    {
        $message = [System.Windows.Forms.MessageBox]::Show("В списке уже имеются записи, добавить к ним?","Уведомление","OKCANCEL","Warning")
        $combomail = $xMF_txtbox_mail_single.Text
        if ($message -eq "OK")
        {
            foreach ($csvs in $csv)
            {
            $TempFIO = $csvs.displayname -replace "`n|`r",""
            $TempFIO2 = delete_spaces_FIO $TempFIO

                $dep = $csvs.city -replace "`n|`r",""
                $tel = telephone $csvs.OfficePhone

                $Spisok = [PSCustomObject]@{
                'firstname' = $csvs.firstname -replace "`n|`r",""
                'lastname'= $csvs.lastName -replace "`n|`r",""
                'displayname' = $TempFIO2
                'office' = $csvs.office -replace "`n|`r",""
                'OfficePhone' = $tel
                'jobtitle' = $csvs.jobtitle -replace "`n|`r",""
                'department' = $csvs.department -replace "`n|`r",""
                'company' = $csvs.company -replace "`n|`r",""
                'streetaddress' = $csvs.streetaddress -replace "`n|`r",""
                'city' = $csvs.city -replace "`n|`r",""
                'postalcode' = $csvs.postalCode -replace "`n|`r",""
                'adcheck' = "unchecked"
                'obshie' = "no"
                'samaccountname' = ""
                'ini' = ""
                'pass' = ""
                'mail' = $combomail
                }

                $xMF_lstv_SingleUser.Items.Add($Spisok)
                }

            }
     }
     else
     {
        foreach ($csvs in $csv)
            {

            $dep = $csvs.city -replace "`n|`r",""
            $tel = telephone $csvs.OfficePhone

            $TempFIO = $csvs.displayname -replace "`n|`r",""
            $TempFIO2 = delete_spaces_FIO $TempFIO

            $Spisok = [PSCustomObject]@{
            'firstname' = $csvs.firstname -replace "`n|`r",""
            'lastname'= $csvs.lastName -replace "`n|`r",""
            'displayname' = $TempFIO2
            'office' = $csvs.office -replace "`n|`r",""
            'OfficePhone' = $tel
            'jobtitle' = $csvs.jobtitle -replace "`n|`r",""
            'department' = $csvs.department -replace "`n|`r",""
            'company' = $csvs.company -replace "`n|`r",""
            'streetaddress' = $csvs.streetaddress -replace "`n|`r",""
            'city' = $csvs.city -replace "`n|`r",""
            'postalcode' = $csvs.postalCode -replace "`n|`r",""
            'adcheck' = "unchecked"
            'obshie' = "no"
            'samaccountname' = ""
            'ini' = ""
            'pass' = ""
            'mail' = $combomail
            }

            $xMF_lstv_SingleUser.Items.Add($Spisok)
            }
     }
}
else
{
[System.Windows.Forms.MessageBox]::Show("Файла users.csv не существует","Уведомление","OK","Warning")
}

})
#--------------------------------------------------------------------------------------------------------------------------------------------------------------

###############################################################################
#Блок кнопок на второй форме Form2
###############################################################################
$xMF2_Save.add_click({
if ($xMF_lstv_SingleUser.SelectedIndex -eq -1)
{
    [System.Windows.Forms.MessageBox]::Show("Для изменения необходимо выделить строку","Уведомление","OK","Warning")
}
else
{
   
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].firstname = $xMF2_textbox_firstname.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].lastname = $xMF2_textbox_lastname.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].Displayname = $xMF2_textbox_Displayname.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].office = $xMF2_textbox_Office.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].officePhone = $xMF2_textbox_OfficePhone.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].Jobtitle = $xMF2_textbox_Jobtitle.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].department = $xMF2_textbox_Department.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].company = $xMF2_textbox_Company.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].streetAddress = $xMF2_textbox_StreetAddress.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].city = $xMF2_textbox_City.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].postalCode = $xMF2_textbox_PostalCode.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname = $xMF2_textbox_Samaccountname.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].pass = $xMF2_textbox_Pass.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].mail = $xMF2_textbox_mail.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].ini = $xMF2_textbox_ini.Text
    $xMF_lstv_SingleUser.Items.Refresh()
}
})

$xMF2_Cancel.add_click({
    $Form_2.Hide()
})

$xMF2_Saveandclose.add_click({
if ($xMF_New.SelectedIndex -eq -1)
{
    [System.Windows.Forms.MessageBox]::Show("Для изменения необходимо выделить строку","Уведомление","OK","Warning")
}
else
{
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].firstname = $xMF2_textbox_firstname.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].lastname = $xMF2_textbox_lastname.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].Displayname = $xMF2_textbox_Displayname.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].office = $xMF2_textbox_Office.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].officePhone = $xMF2_textbox_OfficePhone.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].Jobtitle = $xMF2_textbox_Jobtitle.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].department = $xMF2_textbox_Department.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].company = $xMF2_textbox_Company.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].streetAddress = $xMF2_textbox_StreetAddress.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].city = $xMF2_textbox_City.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].postalCode = $xMF2_textbox_PostalCode.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname = $xMF2_textbox_Samaccountname.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].pass = $xMF2_textbox_Pass.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].mail = $xMF2_textbox_mail.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].ini = $xMF2_textbox_ini.Text
    $xMF_lstv_SingleUser.Items.Refresh()
    $Form_2.Hide()

}

})

$xMF2_btn_PassGen.add_click({
    $xMF2_textbox_Pass.Text = New-Password
})
#--------------------------------------------------------------------------------------------------------------------------------------------------------------





##############################################################
#Блок кнопок для добавления пользователя из буфера в SingleListView
##############################################################


$xMF_btn_AddSingle.add_click({

if ([Windows.Clipboard]::ContainsText() -eq $true)
{
    $combomail = $xMF_txtbox_mail_single.Text
    $clip1 = [windows.clipboard]::GetText().Split("`n")
    for ($i=0 ;$i -ne $clip1.Count-1 ; $i++)
    {
        $clip2 = $clip1[$i].Split("	")
        $TempFIO = $clip2[2] -replace "`n|`r",""
        if (($clip2[0] -eq "") -or ($clip2[1] -eq ""))
        {
        $TempFIO2 = $TempFIO
        }
        else
        {
        $TempFIO2 = delete_spaces_FIO $TempFIO
        }

            $dep = $clip2[9] -replace "`n|`r",""
            $tel = telephone $clip2[4]
            
            $Spisok = [PSCustomObject]@{
            'firstname' = $clip2[0] -replace "`n|`r",""
            'lastname'= $clip2[1] -replace "`n|`r",""
            'displayname' = $TempFIO2
            'office' = $clip2[3] -replace "`n|`r",""
            'OfficePhone' = $tel
            'jobtitle' = $clip2[5] -replace "`n|`r",""
            'department' = $clip2[6] -replace "`n|`r",""
            'company' = $clip2[7] -replace "`n|`r",""
            'streetaddress' = $clip2[8] -replace "`n|`r",""
            'city' = $clip2[9] -replace "`n|`r",""
            'postalcode' = $clip2[10] -replace "`n|`r",""
            'samaccountname' = $clip2[11] -replace "`n|`r",""
            'pass' = ""
            'mail' = $combomail
            'ini' = ""
            'adcheck' = "unchecked"
            'obshie' = "no"
            }
            $xMF_lstv_SingleUser.Items.Add($Spisok)
    }
}
else
{
    [System.Windows.Forms.MessageBox]::Show("Буфер обмена не содержит текст","Уведомление","OK","Warning")
}

})

$xmf_lstv_Menu_paste.add_click({

if ([Windows.Clipboard]::ContainsText() -eq $true)
{
    $combomail = $xMF_txtbox_mail_single.Text
    $clip1 = [windows.clipboard]::GetText().Split("`n")
    for ($i=0 ;$i -ne $clip1.Count-1 ; $i++)
    {
        $clip2 = $clip1[$i].Split("	")
        $TempFIO = $clip2[2] -replace "`n|`r",""
        if (($clip2[0] -eq "") -or ($clip2[1] -eq ""))
        {
        $TempFIO2 = $TempFIO
        }
        else
        {
        $TempFIO2 = delete_spaces_FIO $TempFIO
        }

            $dep = $clip2[9] -replace "`n|`r",""
            $tel = telephone $clip2[4]

            $Spisok = [PSCustomObject]@{
            'firstname' = $clip2[0] -replace "`n|`r",""
            'lastname'= $clip2[1] -replace "`n|`r",""
            'displayname' = $TempFIO2
            'office' = $clip2[3] -replace "`n|`r",""
            'OfficePhone' = $tel
            'jobtitle' = $clip2[5] -replace "`n|`r",""
            'department' = $clip2[6] -replace "`n|`r",""
            'company' = $clip2[7] -replace "`n|`r",""
            'streetaddress' = $clip2[8] -replace "`n|`r",""
            'city' = $clip2[9] -replace "`n|`r",""
            'postalcode' = $clip2[10] -replace "`n|`r",""
            'samaccountname' = $clip2[11] -replace "`n|`r",""
            'pass' = ""
            'mail' = $combomail
            'ini' = ""
            'adcheck' = "unchecked"
            'obshie' = "no"
            }
            $xMF_lstv_SingleUser.Items.Add($Spisok)

    }
}
else
{
    [System.Windows.Forms.MessageBox]::Show("Буфер обмена не содержит текст","Уведомление","OK","Warning")
}

})


$xmf_btn_clearsingle.add_click({
    $xMF_lstv_SingleUser.Items.Clear()
    $xMF_lstv_SingleUser_Exist.Items.Clear()
    $xMF_prBar.Value = 0
    $xMF_label_prBar.Content = ""
    textcheckad
})

$xMF_lstv_Menu_Clear.add_click({
    $xMF_lstv_SingleUser.Items.Clear()
})

$xMF_lstv_Menu_delete.add_click({
if (!($xMF_lstv_SingleUser.SelectedIndex -eq -1)){
    $xMF_lstv_SingleUser.Items.Remove($xMF_lstv_SingleUser.SelectedItem)}
else{
    [System.Windows.Forms.MessageBox]::Show("Выделите строку с данными","Уведомление","OK","Information")}
})

$xMF_lstv_Menu_Pass.add_click({
if (!($xMF_lstv_SingleUser.SelectedIndex -eq -1)){
    $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].pass = New-Password
    $xmf_lstv_SingleUser.Items.Refresh()}
else{
    [System.Windows.Forms.MessageBox]::Show("Выделите строку с данными","Уведомление","OK","Information")}
})



$xMF_lstv_menu_change.add_click({
#if ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].adcheck -eq "unchecked")
#{
#    [System.Windows.Forms.MessageBox]::Show("Необходимо сначала проверить в АД.","Уведомление","OK","Warning")
#}
if ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].adcheck -eq "Added")
{
    [System.Windows.Forms.MessageBox]::Show("УЗ со статусом Added изменить нельзя, перепроверьте УЗ в АД","Уведомление","OK","Warning")
}
elseif ($xMF_lstv_SingleUser.Items.Count -eq 0)
{
    [System.Windows.Forms.MessageBox]::Show("Список пустой, загрузите файл.","Уведомление","OK","Warning")
}
elseif ($xMF_lstv_SingleUser.SelectedIndex -eq -1)
{
    [System.Windows.Forms.MessageBox]::Show("Выделите строку.","Уведомление","OK","Warning")
}
else
{
    $select = $xMF_lstv_SingleUser.SelectedIndex
    $xMF2_textbox_firstname.Text = $xMF_lstv_SingleUser.Items[$select].firstname
    $xMF2_textbox_lastname.Text = $xMF_lstv_SingleUser.Items[$select].lastname
    $xMF2_textbox_Displayname.Text = $xMF_lstv_SingleUser.Items[$select].Displayname
    $xMF2_textbox_Office.Text = $xMF_lstv_SingleUser.Items[$select].office
    $xMF2_textbox_OfficePhone.Text = $xMF_lstv_SingleUser.Items[$select].officePhone
    $xMF2_textbox_Jobtitle.Text = $xMF_lstv_SingleUser.Items[$select].Jobtitle
    $xMF2_textbox_Department.Text = $xMF_lstv_SingleUser.Items[$select].department
    $xMF2_textbox_Company.Text = $xMF_lstv_SingleUser.Items[$select].company
    $xMF2_textbox_StreetAddress.Text = $xMF_lstv_SingleUser.Items[$select].streetAddress
    $xMF2_textbox_City.Text = $xMF_lstv_SingleUser.Items[$select].city
    $xMF2_textbox_PostalCode.Text = $xMF_lstv_SingleUser.Items[$select].postalCode
    $xMF2_textbox_Samaccountname.Text = $xMF_lstv_SingleUser.Items[$select].samaccountname
    $xMF2_textbox_Pass.Text = $xMF_lstv_SingleUser.Items[$select].pass
    $xMF2_textbox_mail.Text = $xMF_lstv_SingleUser.Items[$select].mail
    $xMF2_textbox_ini.Text = $xMF_lstv_SingleUser.Items[$select].ini

if ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].adcheck -eq "EXIST")
{
    $xMF2_textbox_Samaccountname.IsEnabled = $False
}
else
{
    $xMF2_textbox_Samaccountname.IsEnabled = $True
}

    $Form_2.ShowDialog() | out-null
}
})


$xMF_lstv_menu_changePASS.add_click({

if ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].adcheck -eq "EXIST")
{
    (!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].pass -eq ""))
    {
        $pass = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].pass
            try
            {
                Set-ADAccountPassword $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname -Reset -NewPassword (ConvertTo-SecureString $pass -AsPlainText -force)
                write-host Пароль пользователю "-" $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].DisplayName "("$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname")" сброшен на "("$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].pass")"
            }
            catch
            {
                write-host -BackgroundColor Red -ForegroundColor white [УЗ] Ошибка назначения пароля пользователю $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].DisplayName "("$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname")": $error[0].Exception.Message
            }
    }
    else
    {
        [System.Windows.Forms.MessageBox]::Show("Поле пароль пустое","Уведомление","OK","Information")
    }
}
else
{

    [System.Windows.Forms.MessageBox]::Show("Сначала проверьте существование пользователя","Уведомление","OK","Information")
}


})


$xMF_lstv_menu_changeAD.add_click({

if ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].adcheck -eq "EXIST")
{
    $message = [System.Windows.Forms.MessageBox]::Show("Данные у пользователя`n"+$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].DisplayName+"`nбудут изменены, продолжить?","Уведомление","OKCANCEL","Information")

    if ($message -eq "OK")
    {
        #http://www.computerperformance.co.uk/Logon/LDAP_attributes_active_directory.htm
        $login = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname
        $hash = @{}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].Office -eq "")){$hash.physicalDeliveryOfficeName = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].office;}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].OfficePhone -eq "")){$hash.telephoneNumber = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].officephone;}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].jobtitle -eq "")){$hash.title = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].jobtitle;}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].Department -eq "")){$hash.department = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].department;}else{Set-ADUser $login -Department $null}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].Company -eq "")){$hash.company = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].company;}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].StreetAddress -eq "")){$hash.streetAddress = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].streetaddress;}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].City -eq "")){$hash.l = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].city;}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].PostalCode -eq "")){$hash.postalCode = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].postalcode;}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].mail -eq "")){$hash.mail = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].mail;}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].ini -eq "")){$hash.initials = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].ini;}

        try
        {
            Set-ADUser $login -Replace $hash
            write-host Новые данные пользователю "-" $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].DisplayName "("$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname")" внесены
            $xMF_label_prBar.Content = "Новые данные пользователю - "+$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].DisplayName+"("+$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname+") внесены"
        }
        catch
        {
            write-host -BackgroundColor Red -ForegroundColor white [УЗ] Ошибка внесения новых данных пользователю $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].DisplayName "("$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname")": $error[0].Exception.Message
            $xMF_label_prBar.Content = "Новые данные пользователю не внесены - "+$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].DisplayName+"("+$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname+")"
        }

        if (!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].pass -eq ""))
        {
            $pass = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].pass
            try
            {
                Set-ADAccountPassword $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname -Reset -NewPassword (ConvertTo-SecureString $pass -AsPlainText -force)
                write-host Пароль пользователю "-" $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].DisplayName "("$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname")" сброшен на "("$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].pass")"
            }
            catch
            {
                write-host -BackgroundColor Red -ForegroundColor white [УЗ] Ошибка назначения пароля пользователю $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].DisplayName "("$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname")": $error[0].Exception.Message
            }
        }
    }

}
elseif ($xMF_lstv_SingleUser.Items.Count -eq 0)
{
    [System.Windows.Forms.MessageBox]::Show("Список пустой, загрузите файл.","Уведомление","OK","Warning")
}
elseif ($xMF_lstv_SingleUser.SelectedIndex -eq -1)
{
    [System.Windows.Forms.MessageBox]::Show("Выделите строку.","Уведомление","OK","Warning")
}
elseif (!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].adcheck -eq "EXIST"))
{
    [System.Windows.Forms.MessageBox]::Show("Статус УЗ должен быть EXIST","Уведомление","OK","Warning")
}

})


$xMF_lstvExist_Menu_Clear.add_click({
    $xMF_lstv_SingleUser_Exist.Items.Clear()
})

$xMF_lstvExist_Menu_delete.add_click({
if (!($xMF_lstv_SingleUser_Exist.SelectedIndex -eq -1)){
    $xMF_lstv_SingleUser_Exist.Items.Remove($xMF_lstv_SingleUser_Exist.SelectedItem)}
else{
    [System.Windows.Forms.MessageBox]::Show("Выделите строку с данными","Уведомление","OK","Warning")}
})



#---------------------------------------------------------------------------------------------------------------------------------------------------------

##############################################################
#Двойное нажатие и кнопки меню на lstv_SingleUser и lstv_SingleUser_Exist для копирования в буфер
##############################################################
$xMF_lstv_SingleUser.add_MouseDoubleClick({
if (!($xMF_lstv_SingleUser.SelectedIndex -eq -1))
{
    $item = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].DisplayName
    $item2 = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].SamAccountName
    $item3 = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].Pass
    [Windows.Clipboard]::SetText("$item	$item2	$item3")
    $xMF_label_prBar.Content = "Данные пользователя $item скопированы в буфер обмена"
    Write-Host "Данные пользователя $item скопированы в буфер обмена"
}
})
#--------------------------------#
$xMF_lstv_Menu_copy2bufer.add_click({
if (!($xMF_lstv_SingleUser.SelectedIndex -eq -1))
{
    $item = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].DisplayName
    $item2 = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].SamAccountName
    $item3 = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].Pass
    [Windows.Clipboard]::SetText("$item	$item2	$item3")
    $xMF_label_prBar.Content = "Данные пользователя $item скопированы в буфер обмена"
    Write-Host "Данные пользователя $item скопированы в буфер обмена"
}
})
#--------------------------------#
$xMF_lstv_SingleUser_Exist.add_MouseDoubleClick({
if (!($xMF_lstv_SingleUser_Exist.SelectedIndex -eq -1))
{
    $item = $xMF_lstv_SingleUser_Exist.Items[$xMF_lstv_SingleUser_Exist.SelectedIndex].DisplayName
    $item2 = $xMF_lstv_SingleUser_Exist.Items[$xMF_lstv_SingleUser_Exist.SelectedIndex].SamAccountName
    $item3 = $xMF_lstv_SingleUser_Exist.Items[$xMF_lstv_SingleUser_Exist.SelectedIndex].jobtitle
    $item4 = $xMF_lstv_SingleUser_Exist.Items[$xMF_lstv_SingleUser_Exist.SelectedIndex].Department
    $item5 = $xMF_lstv_SingleUser_Exist.Items[$xMF_lstv_SingleUser_Exist.SelectedIndex].Company
    [Windows.Clipboard]::SetText("$item	$item3	$item4	$item5	$item2")
    $xMF_label_prBar.Content = "Данные существующего пользователя $item скопированы в буфер обмена"
    Write-Host "Данные существующего пользователя $item скопированы в буфер обмена"
}
})
#--------------------------------#
$xMF_lstvExist_Menu_copy2bufer.add_click({
if (!($xMF_lstv_SingleUser_Exist.SelectedIndex -eq -1))
{
    $item = $xMF_lstv_SingleUser_Exist.Items[$xMF_lstv_SingleUser_Exist.SelectedIndex].DisplayName
    $item2 = $xMF_lstv_SingleUser_Exist.Items[$xMF_lstv_SingleUser_Exist.SelectedIndex].SamAccountName
    $item3 = $xMF_lstv_SingleUser_Exist.Items[$xMF_lstv_SingleUser_Exist.SelectedIndex].jobtitle
    $item4 = $xMF_lstv_SingleUser_Exist.Items[$xMF_lstv_SingleUser_Exist.SelectedIndex].Department
    $item5 = $xMF_lstv_SingleUser_Exist.Items[$xMF_lstv_SingleUser_Exist.SelectedIndex].Company
    [Windows.Clipboard]::SetText("$item	$item3	$item4	$item5	$item2")
    $xMF_label_prBar.Content = "Данные существующего пользователя $item скопированы в буфер обмена"
    Write-Host "Данные существующего пользователя $item скопированы в буфер обмена"
}
})
#---------------------------------------------------------------------------------------------------------------------------------------------------------


##############################################################
#Блок кнопок для добавления пользователя из ListView в AD
##############################################################
$xmf_lstv_Menu_add_ad.add_click({

add_user_ad -multi $false

})
#---------------------------------------------------------------------------------------------------------------------------------------------------------


###################################################################
#Кнопка для добавления всех пользователей в списке SingleUser в АД
###################################################################
$xMF_btn_single_add_ad.add_click({

add_user_ad -multi $true

})
#---------------------------------------------------------------------------------------------------------------------------------------------------------

###################################################################
#Сделать всех пользователей ОБЩИЕ УЗ
###################################################################
$xMF_btn_allobshie.add_click({

if (!($xmf_lstv_SingleUser.Items.Count -eq 0))
{

    $message = [System.Windows.Forms.MessageBox]::Show("Сделать всё ОБЩИМИ УЗ?","Уведомление","OKCANCEL","Information")
    if ($message = "OK")
    {
        for ($i=0 ; $i -ne $xmf_lstv_SingleUser.Items.Count ; $i++)
        {

            $xmf_lstv_SingleUser.Items[$i].obshie = "yes"
            
        }
        textcheckad
        $xmf_lstv_SingleUser.Items.Refresh()
    }
}

})
#---------------------------------------------------------------------------------------------------------------------------------------------------------

###################################################################
#Генерировать всем новые пароли
###################################################################

$xMF_btn_GenPasswords.add_click({

if (!($xmf_lstv_SingleUser.Items.Count -eq 0))
{

    $message = [System.Windows.Forms.MessageBox]::Show("Будут сгенерированы всем новые пароли`nПродолжить?","Уведомление","OKCANCEL","Information")
    if ($message = "OK")
    {
        for ($i=0 ; $i -ne $xmf_lstv_SingleUser.Items.Count ; $i++)
        {
            $xmf_lstv_SingleUser.Items[$i].pass = New-Password
        }
        $xmf_lstv_SingleUser.Items.Refresh()
        Write-Host -BackgroundColor Yellow -ForegroundColor Black [PASS] Всем пользователям сгенерированы новые пароли
        $xMF_label_prBar.Content = "[PASS] Всем пользователям сгенерированы новые пароли"
        $xMF_prBar.Value = 0
    }
}

})

#---------------------------------------------------------------------------------------------------------------------------------------------------------

###################################################################
#Экспортировать все строки в коносль для отправки данных пользователю
###################################################################

$xMF_btn_exporttoconsole.add_click({

if (!($xmf_lstv_SingleUser.Items.Count -eq 0))
{

    $message = [System.Windows.Forms.MessageBox]::Show("Выгрузить данные в консоль?","Уведомление","OKCANCEL","Information")
    if ($message = "OK")
    {
        Write-Host -BackgroundColor Yellow -ForegroundColor Black "---------------------------------START---------------------------------"
        for ($i=0 ; $i -ne $xmf_lstv_SingleUser.Items.Count ; $i++)
        {
            $pass = $xmf_lstv_SingleUser.Items[$i].pass
            $displayname = $xmf_lstv_SingleUser.Items[$i].displayname
            $samaccountname = $xmf_lstv_SingleUser.Items[$i].samaccountname

            Write-Host "ФИО: $displayname - Логин: $samaccountname - Пароль: $pass"
        }
        Write-Host -BackgroundColor Yellow -ForegroundColor Black "---------------------------------END-----------------------------------"
        $xMF_label_prBar.Content = "Данные в консоль выгружены"
        $xMF_prBar.Value = 0
    }
}

})

#---------------------------------------------------------------------------------------------------------------------------------------------------------



####################################################################
#Кнопка изменения статуса общей учетной записи в lstview (obshie)
####################################################################
$xMF_lstv_Menu_status.add_click({

if (!($xMF_lstv_SingleUser.Items.Count -eq 0) -or !($xMF_lstv_SingleUser.SelectedIndex -eq -1))
{
    if ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].obshie -eq "yes")
    {
        $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].obshie = "no"
        $xMF_lstv_SingleUser.Items.Refresh()
    }
    elseif ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].obshie -eq "no")
    {
        $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].obshie = "yes"
        $xMF_lstv_SingleUser.Items.Refresh()
    }
    textcheckad
}
else
{
    [System.Windows.Forms.MessageBox]::Show("ОШИБКА","Уведомление","OK","Warning")
}

})

#---------------------------------------------------------------------------------------------------------------------------------------------------------

$xMF_lstvExist_Menu_Disable.add_click({

    if (!($xMF_lstv_SingleUser_Exist.SelectedIndex -eq -1) -or !($xMF_lstv_SingleUser_Exist.Items.Count -eq 0))
    {
        $login = $xMF_lstv_SingleUser_Exist.Items[$xMF_lstv_SingleUser_Exist.SelectedIndex].SamAccountName
        try
        {
            Disable-ADAccount -Identity $login
            Set-ADUser $login -Description $xMF_txtbox_mail_single.Text
            write-host [ОТКЛ] -> Пользователь $login отключен, в поле Description добавлено ($xMF_txtbox_mail_single.Text)
            $xMF_label_prBar.Content = "Пользователь $login отключен, в поле Description добавлено ("+$xMF_txtbox_mail_single.Text+")"
        }
        catch
        {
            write-host -BackgroundColor Red -ForegroundColor white [ОТКЛ] -> Ошибка в отключении пользователя $login
            $xMF_label_prBar.Content = "Ошибка в отключении пользвоателя $login"
        }

    }
})


##############################################################
#Кнопка переноса всех пользователей в другую OU
##############################################################
$xMF_btn_change_OU.add_click({


if (!($xmf_lstv_SingleUser.Items.Count -eq 0))
{
$message = [System.Windows.Forms.MessageBox]::Show("Пользователи со статусом EXIST будут перенесены в`n"+$xMF_txtbox_OU_path.Text+"`nпродолжить?","Подтверждение","OKCANCEL","information")

    if ($message -eq "OK")
    {
    $t=0
    $nt=0
        for ($i=0 ; $i -ne $xmf_lstv_SingleUser.Items.Count ; $i++)
        {
            if ($xMF_lstv_SingleUser.Items[$i].adcheck -eq "EXIST")
            {
                $User = Get-ADUser -Identity $xMF_lstv_SingleUser.Items[$i].samaccountname
                $UserDN = $User.distinguishedName
                $TargetOU = $xMF_txtbox_OU_path.Text

                    if ($user.Enabled -eq $False)
                    {
                        Enable-ADAccount -Identity $UserDN
                    }
                try
                {
                    
                    Move-ADObject -Identity $UserDN -TargetPath $TargetOU
                    write-host Пользователь $xMF_lstv_SingleUser.Items[$i].displayname успешно перенесен в $TargetOU
                    $xMF_label_prBar.Content = "Пользователь "+$xMF_lstv_SingleUser.Items[$i].displayname+" успешно перенесен в $TargetOU"
                    $t+=1
                }
                catch
                {
                    $nt+=1
                    write-host Ошибка, перенести пользователя $xMF_lstv_SingleUser.Items[$i].displayname не удалось.
                    $xMF_label_prBar.Content = "Ошибка, перенести пользователя "+$xMF_lstv_SingleUser.Items[$i].displayname+" не удалось."
                }


            }
            else
            {
                write-host -BackgroundColor Yellow -ForegroundColor White Пользователь $xMF_lstv_SingleUser.Items[$i].displayname не проверен в АД
            }
        }
     [System.Windows.Forms.MessageBox]::Show("Перенесено: $t`nНе перенесено: $nt","Перенос выполнен","OK","Information")
    }
}
else
{
    [System.Windows.Forms.MessageBox]::Show("Список пустой, загрузите файл.","Уведомление","OK","Warning")
}
})


$xMF_lstv_menu_moveOU.add_click({

   if ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].adcheck -eq "EXIST")
   {
        $User = Get-ADUser -Identity $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname
        $UserDN = $User.distinguishedName
        $TargetOU = $xMF_txtbox_OU_path.Text

        $message = [System.Windows.Forms.MessageBox]::Show("Перенести "+$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].displayname+" в:`n"+$xMF_txtbox_OU_path.Text+" ?","Подтверждение","OKCANCEL","information")

        if ($message -eq "OK")
        {
                if ($user.Enabled -eq $False)
                {
                    Enable-ADAccount -Identity $UserDN
                }
            try
            {
                
                Move-ADObject -Identity $UserDN -TargetPath $TargetOU
                write-host Пользователь $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].displayname успешно перенесен в $TargetOU
                $xMF_label_prBar.Content = "Пользователь "+$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].displayname+" успешно перенесен в $TargetOU"
            }
            catch
            {
                [System.Windows.Forms.MessageBox]::Show("Ошибка, перенести пользователя не удалось.","Уведомление","OK","Warning")
                write-host Ошибка, перенести пользователя $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].displayname не удалось.
                $xMF_label_prBar.Content = "Ошибка, перенести пользователя "+$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].displayname+" не удалось."
            }
        }
   }

})
#---------------------------------------------------------------------------------------------------------------------------------------------------------

##############################################################
#Функция проверки пользователей в АД из lstv_Single
##############################################################

$xMF_Btn_single_CheckAD.add_click({

if (!($xmf_lstv_SingleUser.Items.Count -eq 0))
{

        $FileCount = $xMF_lstv_SingleUser.Items.Count
        $SCount = 0
        $xMF_prBar.Value = 0
        $xMF_prBar.Maximum = $FileCount - 1

        for ($i=0 ; $i -ne $xmf_lstv_SingleUser.Items.Count ; $i++)
        {
            if (($xmf_lstv_SingleUser.Items[$i].adcheck -eq "unchecked") -or ($xmf_lstv_SingleUser.Items[$i].adcheck -eq "ERROR"))
            {
                if (!($xmf_lstv_SingleUser.Items[$i].displayname -like "*ERROR*"))
                {
                find_user -SelectMail $xmf_txtbox_mail_single.Text.ToLower() -checkcount $i -multi $true

                $SCount = $SCount + 1
                $xMF_prBar.Value = $SCount
                $xMF_label_prBar.Content = "Обработка - " + $xmf_lstv_SingleUser.Items[$i].DisplayName
                $Form_Main.Dispatcher.Invoke([action]{},"Render")
                }
                else
                {
                    $SCount = $SCount + 1
                    $xMF_prBar.Value = $SCount
                    $xmf_lstv_SingleUser.Items[$i].adcheck = "ERROR"
                    $xmf_lstv_SingleUser.Items.Refresh()
                }

            }
        }

    
}
else
{
    [System.Windows.Forms.MessageBox]::Show("Список пустой, загрузите файл.","Уведомление","OK","Warning")
}


textcheckad

})

$xMF_lstv_Menu_check_ad.add_click({

if (!($xmf_lstv_SingleUser.Items.Count -eq 0) -or !($xmf_lstv_SingleUser.SelectedIndex -eq -1))
{
    if (($xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck -eq "unchecked") -or ($xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck -eq "ERROR"))
    {
        if (!($xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].displayname -like "*ERROR*"))
        {
            $xMF_prBar.Value = 0
            $xMF_label_prBar.Content = ""
            find_user -SelectMail $xmf_txtbox_mail_single.Text.ToLower() -multi $false
        }
        else
        {
            $xMF_prBar.Value = 0
            $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck = "ERROR"
            $xmf_lstv_SingleUser.Items.Refresh()
        }
    }
}
else
{
    [System.Windows.Forms.MessageBox]::Show("Список пустой, загрузите файл.","Уведомление","OK","Warning")
}

textcheckad

})

#---------------------------------------------------------------------------------------------------------------------------------------------------------




##############################################################
#Кнопка открытия OU
##############################################################


$xMF_OU.add_click({

$xMF_txtbox_OU_path.Text = Get-OUDialog-start

})


#---------------------------------------------------------------------------------------------------------------------------------------------------------





##############################################################
#Всякие мелкие функции
##############################################################


function textcheckad {

if ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].obshie -eq "yes")
{
    $xMF_Statustext.Background = "green"
    $xMF_Statustext.Foreground = "white"
    $xMF_Statustext.Text = "Статус: УЗ ОБЩАЯ ("+$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].adcheck+")"
}
elseif ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].adcheck -eq "checked")
{
    $xMF_Statustext.Background = "green"
    $xMF_Statustext.Foreground = "white"
    $xMF_Statustext.Text = "Статус в АД: "+$xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck
}
elseif ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].adcheck -eq "unchecked")
{
    $xMF_Statustext.Background = "red"
    $xMF_Statustext.Foreground = "white"
    $xMF_Statustext.Text = "Статус в АД: "+$xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck
}
elseif ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].adcheck -eq "EXIST")
{
    $xMF_Statustext.Background = "yellow"
    $xMF_Statustext.Foreground = "black"
    $xMF_Statustext.Text = "Статус в АД: "+$xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck
}
elseif ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].adcheck -eq "Added")
{
    $xMF_Statustext.Background = "green"
    $xMF_Statustext.Foreground = "white"
    $xMF_Statustext.Text = "Статус в АД: "+$xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck
}
elseif ($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].adcheck -eq "ERROR")
{
    $xMF_Statustext.Background = "red"
    $xMF_Statustext.Foreground = "black"
    $xMF_Statustext.Text = "Статус в АД: "+$xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck
}
else
{
    $xMF_Statustext.Background = "white"
    $xMF_Statustext.Foreground = "black"
    $xMF_Statustext.Text = "Статус в АД:"
}

}


#---------------------------------------------------------------------------------------------------------------------------------------------------------

##############################################################
#Form2 реакция на нажатие клавиш мыши
##############################################################
$Form_2.add_MouseRightButtonUP({
    $form_2.hide()
})

$Form_2.add_MouseLeftButtonDown({
    $this.dragmove()
})

$Form_2.add_MouseDoubleClick({
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].firstname = $xMF2_textbox_firstname.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].lastname = $xMF2_textbox_lastname.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].Displayname = $xMF2_textbox_Displayname.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].office = $xMF2_textbox_Office.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].officePhone = $xMF2_textbox_OfficePhone.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].Jobtitle = $xMF2_textbox_Jobtitle.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].department = $xMF2_textbox_Department.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].company = $xMF2_textbox_Company.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].streetAddress = $xMF2_textbox_StreetAddress.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].city = $xMF2_textbox_City.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].postalCode = $xMF2_textbox_PostalCode.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname = $xMF2_textbox_Samaccountname.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].pass = $xMF2_textbox_Pass.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].mail = $xMF2_textbox_mail.Text
    $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].ini = $xMF2_textbox_ini.Text
    $xMF_lstv_SingleUser.Items.Refresh()
    $form_2.hide()
})
#---------------------------------------------------------------------------------------------------------------------------------------------------------

##############################################################
#Всякое
##############################################################

$xMF_lstv_SingleUser.add_SelectionChanged({

textcheckad

})


#Сброс статусов, для повторной проверки пользователей в списке
$xMF_status_canc.add_click({

if ($xmf_lstv_SingleUser.Items.Count -cge 1)
{
    $message = [System.Windows.Forms.MessageBox]::Show("Статусы AD будут сброшены, продолжить?","Уведомление","OKCANCEL","Information")

    if ($message -eq "OK")
    {
        for ($i=0 ; $i -ne $xmf_lstv_SingleUser.Items.Count ; $i++)
        {
            $xmf_lstv_SingleUser.Items[$i].adcheck = "unchecked"
            $xmf_lstv_SingleUser.Items.Refresh()
        }
        textcheckad
    }
}


})


$xMF_Btn_Exit.add_click({
    $Form_2.Close()
    $Form_Main.Close()
})


$Form_Main.ShowDialog() | out-null