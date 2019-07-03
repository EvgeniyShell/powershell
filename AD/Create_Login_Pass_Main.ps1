#Clear-Host

Add-Type -AssemblyName PresentationFramework, PresentationCore, System.Windows.Forms
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

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
# Функции для работы с группами
. $ScriptPath\groups.ps1
# Функции для работы с inactive
. $ScriptPath\inactive.ps1

#. $ScriptPath\find_user_single.ps1
#. $ScriptPath\MakeCopy.ps1
$xMF_btn_group_OU.Visibility = "Hidden"
$xMF_lstv_SingleUser.Items.Clear()
$Global:OU = "OU=Test,OU=OU_OTHER,DC=ASO,DC=RT,DC=LOCAL"
$xMF_txtbox_OU_path.Text = $Global:OU
$xMF_textbox_Group_newuser.Text=""
$xMF_textbox_Group_existuser.Text=""
$global:checklog = 0
$global:firstrun = 1
$global:Session = ""

testconnection

###############################################################################
#Функции для проверки данных
###############################################################################

$xMF_chk_hideshow.Add_Checked({$xMF_chk_hideshow.content = "Спрятать консоль"; Show-Console})
$xMF_chk_hideshow.Add_UnChecked({$xMF_chk_hideshow.content = "Показать консоль"; Hide-Console})
function Show-Console {
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 5)
}

function Hide-Console {
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)
}

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
    write-host "Цифра в ФИО -> Удаление ($SetFIO)"
    $SetFIO = $SetFIO -replace "[0-9]",""
    delete_spaces_FIO($SetFIO)
   #return "ERROR - Цифра в ФИО"
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

#Функция проверки количества символов
function symbol ($check)
{
    if ($check.Length -gt 64)
    {
        $global:checklog++
        return "БОЛЬШЕ 64 СИМВОЛОВ"
    }
    else
    {
        return $check -replace "`n|`r",""
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

                $dep = $csvs.city -replace "`n|`r","" #Важная переменная, работает в функции telephone
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

    if ($xMF_pwd_old.IsChecked)
    {
        $xMF2_textbox_Pass.Text = New-Password2
    }
    elseif ($xMF_pwd_new.IsChecked)
    {
        $xMF2_textbox_Pass.Text = New-Password
    }
})
#--------------------------------------------------------------------------------------------------------------------------------------------------------------





##############################################################
#Блок кнопок для добавления пользователя из буфера в SingleListView
##############################################################


$xMF_btn_AddSingle.add_click({

if ([Windows.Clipboard]::ContainsText() -eq $true)
{
    $global:checklog = 0
    $combomail = $xMF_txtbox_mail_single.Text
    $clip1 = [windows.clipboard]::GetText().Split("`n")
    [array]$clip3 = $clip1[0].Split("	") #переменная Для проверки, что скопирована только 1 ячейка с Эксэль
    if ($clip3.Count -eq 1) #проверка на наличие 1 переменной в эксэль
    {
        for ($i=0 ;$i -ne $clip1.Count-1 ; $i++) #Цикл, если скопировано много строк, по 1 ячейке в каждой строке.
        {
        $TempFIO = delete_spaces_FIO $clip1[$i]
        $Spisok = [PSCustomObject]@{
            'firstname' = $TempFIO.Split(" ")[1] -replace "`n|`r",""
            'lastname'= $TempFIO.Split(" ")[0] -replace "`n|`r",""
            'displayname' = $TempFIO
            'office' = ""
            'OfficePhone' = ""
            'jobtitle' = ""
            'department' = ""
            'company' = ""
            'streetaddress' = ""
            'city' = ""
            'postalcode' = ""
            'samaccountname' = ""
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
            'firstname' = symbol($clip2[0])
            'lastname'= symbol($clip2[1])
            'displayname' = symbol($TempFIO2)
            'office' = symbol($clip2[3])
            'OfficePhone' = symbol($tel)
            'jobtitle' = symbol($clip2[5])
            'department' = symbol($clip2[6])
            'company' = symbol($clip2[7])
            'streetaddress' = symbol($clip2[8])
            'city' = symbol($clip2[9])
            'postalcode' = symbol($clip2[10])
            'samaccountname' = symbol($clip2[11])
            'pass' = ""
            'mail' = $combomail
            'ini' = ""
            'adcheck' = "unchecked"
            'obshie' = "no"
            }
            $xMF_lstv_SingleUser.Items.Add($Spisok)
    }
    } #($clip3.Count -eq 1)

    if ($global:checklog -ge 1)
    {
    [System.Windows.Forms.MessageBox]::Show("Ячейки с > 64 символами: $global:checklog","Уведомление","OK","Warning")
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
    $global:checklog = 0
    $combomail = $xMF_txtbox_mail_single.Text
    $clip1 = [windows.clipboard]::GetText().Split("`n")
    [array]$clip3 = $clip1[0].Split("	") #переменная Для проверки, что скопирована только 1 ячейка с Эксэль
    if ($clip3.Count -eq 1) #проверка на наличие 1 переменной в эксэль
    {
        for ($i=0 ;$i -ne $clip1.Count-1 ; $i++) #Цикл, если скопировано много строк, по 1 ячейке в каждой строке.
        {
        $TempFIO = delete_spaces_FIO $clip1[$i]
        $Spisok = [PSCustomObject]@{
            'firstname' = $TempFIO.Split(" ")[1] -replace "`n|`r",""
            'lastname'= $TempFIO.Split(" ")[0] -replace "`n|`r",""
            'displayname' = $TempFIO
            'office' = ""
            'OfficePhone' = ""
            'jobtitle' = ""
            'department' = ""
            'company' = ""
            'streetaddress' = ""
            'city' = ""
            'postalcode' = ""
            'samaccountname' = ""
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
            'firstname' = symbol($clip2[0])
            'lastname'= symbol($clip2[1])
            'displayname' = symbol($TempFIO2)
            'office' = symbol($clip2[3])
            'OfficePhone' = symbol($tel)
            'jobtitle' = symbol($clip2[5])
            'department' = symbol($clip2[6])
            'company' = symbol($clip2[7])
            'streetaddress' = symbol($clip2[8])
            'city' = symbol($clip2[9])
            'postalcode' = symbol($clip2[10])
            'samaccountname' = symbol($clip2[11])
            'pass' = ""
            'mail' = $combomail
            'ini' = ""
            'adcheck' = "unchecked"
            'obshie' = "no"
            }
            $xMF_lstv_SingleUser.Items.Add($Spisok)

    }
    }# ($clip3.Count -eq 1)

    if ($global:checklog -ge 1)
    {
    [System.Windows.Forms.MessageBox]::Show("Ячейки с > 64 символами: $global:checklog","Уведомление","OK","Warning")
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
    
    $count = $xMF_lstv_SingleUser.SelectedItems.Count

    if ($count -gt 0)
    {
        #Пока количество выделеных записей больше нуля, удалять записи
        while ($count -gt 0)
        {
            $xMF_lstv_SingleUser.Items.Remove($xMF_lstv_inactive.SelectedItem)
            $count = $xMF_lstv_SingleUser.SelectedItems.Count
        }
    }
}
else{
    [System.Windows.Forms.MessageBox]::Show("Выделите строку с данными","Уведомление","OK","Information")}
})

$xMF_lstv_Menu_Pass.add_click({
if (!($xMF_lstv_SingleUser.SelectedIndex -eq -1)){
    if ($xMF_pwd_old.IsChecked)
    {
        $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].pass = New-Password2
    }
    elseif ($xMF_pwd_new.IsChecked)
    {
        $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].pass = New-Password
    }
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
    if (!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].pass -eq ""))
    {
        $pass = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].pass
            try
            {
                Set-ADAccountPassword $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname -Reset -NewPassword (ConvertTo-SecureString $pass -AsPlainText -force)
                write-host Пароль пользователю "-" $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].DisplayName "("$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname")" сброшен на "("$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].pass")"
                $xMF_label_prBar.Content = "Пароль пользователю "+ $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].DisplayName + " сброшен на " + "("+$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].pass+")"
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


$xMF_chk_all.add_click({

if ($xMF_chk_all.IsChecked -eq $true)
{
$xMF_chk_all.Content = "Убрать всем"
$xMF_chk_office.IsChecked = $True
$xMF_chk_OfficePhone.IsChecked = $True
$xMF_chk_jobtitle.IsChecked = $True
$xMF_chk_department.IsChecked = $True
$xMF_chk_company.IsChecked = $True
$xMF_chk_streetaddress.IsChecked = $True
$xMF_chk_city.IsChecked = $True
$xMF_chk_postalcode.IsChecked = $True
}
else
{
$xMF_chk_all.Content = "Выделить все"
$xMF_chk_office.IsChecked = $False
$xMF_chk_OfficePhone.IsChecked = $False
$xMF_chk_jobtitle.IsChecked = $False
$xMF_chk_department.IsChecked = $False
$xMF_chk_company.IsChecked = $False
$xMF_chk_streetaddress.IsChecked = $False
$xMF_chk_city.IsChecked = $False
$xMF_chk_postalcode.IsChecked = $False
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
        $UserADinfo = Get-ADUser $login -Properties Description | select Enabled,Description

        $hash = @{}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].Office -eq "")){$hash.physicalDeliveryOfficeName = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].office;}elseif($xmf_chk_office.isChecked -eq $true){Set-ADUser $login -Office $null}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].OfficePhone -eq "")){$hash.telephoneNumber = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].officephone;}elseif($xMF_chk_OfficePhone.isChecked -eq $true){Set-ADUser $login -OfficePhone $null}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].jobtitle -eq "")){$hash.title = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].jobtitle;}elseif($xMF_chk_jobtitle.isChecked -eq $true){Set-ADUser $login -Title $null}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].Department -eq "")){$hash.department = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].department;}elseif($xMF_chk_department.isChecked -eq $true){Set-ADUser $login -Department $null}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].Company -eq "")){$hash.company = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].company;}elseif($xMF_chk_company.isChecked -eq $true){Set-ADUser $login -Company $null}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].StreetAddress -eq "")){$hash.streetAddress = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].streetaddress;}elseif($xMF_chk_streetaddress.isChecked -eq $true){Set-ADUser $login -StreetAddress $null}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].City -eq "")){$hash.l = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].city;}elseif($xMF_chk_city.isChecked -eq $true){Set-ADUser $login -City $null}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].PostalCode -eq "")){$hash.postalCode = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].postalcode;}elseif($xMF_chk_postalcode.isChecked -eq $true){Set-ADUser $login -PostalCode $null}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].mail -eq "")){$hash.mail = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].mail;}
        if(!($xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].ini -eq "")){$hash.initials = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].ini;}

        
        if ($UserADinfo.Enabled -eq $False)
        {
            try
            {
                Enable-ADAccount -Identity $login
                write-host "Учетная запись пользователя" $login "включена"
            }
            catch
            {
                write-host -BackgroundColor Red -ForegroundColor Yellow "Учетную запись пользователя" $login "включить нельзя" -> $error[0].Exception.Message
            }
        }

        if (!($UserADinfo.Description -eq $null))
        {
            try
            {
                Set-ADUser $login -Description $null
                write-host "Учетной записи пользователя" $login "поле Description очищено"
            }
            catch
            {
                write-host -BackgroundColor Red -ForegroundColor Yellow "Учетной записи пользователя" $login "поле Description очистить нельзя" -> $error[0].Exception.Message
            }
        }

        try
        {
            Set-ADUser $login -Replace $hash -Description $null
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
    
    $count = $xMF_lstv_SingleUser_Exist.SelectedItems.Count

    if ($count -gt 0)
    {
        #Пока количество выделеных записей больше нуля, удалять записи
        while ($count -gt 0)
        {
            $xMF_lstv_SingleUser_Exist.Items.Remove($xMF_lstv_SingleUser_Exist.SelectedItem)
            $count = $xMF_lstv_SingleUser_Exist.SelectedItems.Count
        }
    }
     
}
else{
    [System.Windows.Forms.MessageBox]::Show("Выделите строку с данными","Уведомление","OK","Warning")}
})



$xMF_lstvExist_Menu_deletegroup.add_click({

    if (!($xMF_lstv_SingleUser_Exist.SelectedIndex -eq -1))
    {
        $userdelete = $xMF_lstv_SingleUser_Exist.Items[$xMF_lstv_SingleUser_Exist.SelectedIndex].samaccountname
        chkboxdeletegroups -user $userdelete -exception "mdaemon"
    }

})


$xMF_lstvExist_Menu_deletemanager.add_click({

    if (!($xMF_lstv_SingleUser_Exist.SelectedIndex -eq -1))
    {
        $userdelete = $xMF_lstv_SingleUser_Exist.Items[$xMF_lstv_SingleUser_Exist.SelectedIndex].samaccountname
        managers -user $userdelete -empty $true
    }

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
    [Windows.Clipboard]::Clear()
    [Windows.Clipboard]::SetDataObject("$item	$item2	$item3")
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
    [Windows.Clipboard]::Clear()
    [Windows.Clipboard]::SetDataObject("$item	$item2	$item3")
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
    [Windows.Clipboard]::Clear()
    [Windows.Clipboard]::SetDataObject("$item	$item3	$item4	$item5	$item2")
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
    [Windows.Clipboard]::Clear()
    [Windows.Clipboard]::SetDataObject("$item	$item3	$item4	$item5	$item2")
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
            if ($xMF_pwd_old.IsChecked)
            {
                $xmf_lstv_SingleUser.Items[$i].pass = New-Password2
            }
            elseif ($xMF_pwd_new.IsChecked)
            {
                $xmf_lstv_SingleUser.Items[$i].pass = New-Password
            }
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
    if ($message -eq "OK")
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
#Кнопка btn_blockall - Заблокировать всех в нижнем списке
####################################################################

$xMF_btn_blockall.add_click({

if (!($xMF_lstv_SingleUser_Exist.Items.Count -eq 0))
{
    $message = [System.Windows.Forms.MessageBox]::Show("Пользователи будут заблокированы","Уведомление","OKCANCEL","Information")
    if ($message = "OK")
    {
    $xMF_prBar.Value = 0
    $SCount = 0
    $xMF_prBar.Maximum = $xMF_lstv_SingleUser.Items.Count
    for ($i=0 ; $i -ne $xMF_lstv_SingleUser_Exist.Items.Count ; $i++)
    {
        if ($xMF_lstv_SingleUser_Exist.Items[$i].Enabled -eq "True")
        {
            $login = $xMF_lstv_SingleUser_Exist.Items[$i].SamAccountName
            try
            {
                Disable-ADAccount -Identity $login
                Set-ADUser $login -Description $xMF_txtbox_mail_single.Text
                write-host [ОТКЛ] -> Пользователь $login отключен, в поле Description добавлено ($xMF_txtbox_mail_single.Text)
            }
            catch
            {
                write-host -BackgroundColor Red -ForegroundColor white [ОТКЛ] -> Ошибка в отключении пользователя $login ($error[0].Exception.Message)
            }
        } # ---End IF ($xMF_lstv_SingleUser_Exist.Items[$i].Enabled -eq "True")
        elseif ($xMF_lstv_SingleUser_Exist.Items[$i].Enabled -eq "False")
        {
            write-host -BackgroundColor Red -ForegroundColor white [ОТКЛ] -> Пользователь $login не был отключен, он и так заблокирован. Описание - ($xMF_lstv_SingleUser_Exist.Items[$i].Description)
        } # ---End ELSEIF ($xMF_lstv_SingleUser_Exist.Items[$i].Enabled -eq "False")
        $SCount = $SCount + 1
        $xMF_prBar.Value = $SCount

    } # ---end FOR ($i=0 ; $i -ne $xMF_lstv_SingleUser_Exist.Items.Count ; $i++)
    } # ---end if ($message = "OK")
} # ---END (!($xMF_lstv_SingleUser_Exist.Items.Count -eq 0))


})

#---------------------------------------------------------------------------------------------------------------------------------------------------------


####################################################################
#Кнопка btn_changeall - Изменить данные у всех
####################################################################

$xMF_btn_changeall.add_click({

if (!$xMF_lstv_SingleUser.Items.Count -eq 0)
{
$message1 = [System.Windows.Forms.MessageBox]::Show("Изменить у всех пользователей?","Уведомление","OKCANCEL","Information")
if ($message1 -eq "OK")
{
$xMF_prBar.Value = 0
$SCount = 0
$xMF_prBar.Maximum = $xMF_lstv_SingleUser.Items.Count
for ($i=0 ; $i -ne $xMF_lstv_SingleUser.Items.Count ; $i++)
{

if ($xMF_lstv_SingleUser.Items[$i].adcheck -eq "EXIST")
{
    $message = [System.Windows.Forms.MessageBox]::Show("Данные у пользователя`n"+$xMF_lstv_SingleUser.Items[$i].DisplayName+"`nбудут изменены, продолжить?","Уведомление","OKCANCEL","Information")

    if ($message -eq "OK")
    {
        #http://www.computerperformance.co.uk/Logon/LDAP_attributes_active_directory.htm
        $login = $xMF_lstv_SingleUser.Items[$i].samaccountname
        $UserADinfo = Get-ADUser $login -Properties Description | select Enabled,Description

        $hash = @{}
        if(!($xMF_lstv_SingleUser.Items[$i].Office -eq "")){$hash.physicalDeliveryOfficeName = $xMF_lstv_SingleUser.Items[$i].office;}elseif($xmf_chk_office.isChecked -eq $true){Set-ADUser $login -Office $null}
        if(!($xMF_lstv_SingleUser.Items[$i].OfficePhone -eq "")){$hash.telephoneNumber = $xMF_lstv_SingleUser.Items[$i].officephone;}elseif($xMF_chk_OfficePhone.isChecked -eq $true){Set-ADUser $login -OfficePhone $null}
        if(!($xMF_lstv_SingleUser.Items[$i].jobtitle -eq "")){$hash.title = $xMF_lstv_SingleUser.Items[$i].jobtitle;}elseif($xMF_chk_jobtitle.isChecked -eq $true){Set-ADUser $login -Title $null}
        if(!($xMF_lstv_SingleUser.Items[$i].Department -eq "")){$hash.department = $xMF_lstv_SingleUser.Items[$i].department;}elseif($xMF_chk_department.isChecked -eq $true){Set-ADUser $login -Department $null}
        if(!($xMF_lstv_SingleUser.Items[$i].Company -eq "")){$hash.company = $xMF_lstv_SingleUser.Items[$i].company;}elseif($xMF_chk_company.isChecked -eq $true){Set-ADUser $login -Company $null}
        if(!($xMF_lstv_SingleUser.Items[$i].StreetAddress -eq "")){$hash.streetAddress = $xMF_lstv_SingleUser.Items[$i].streetaddress;}elseif($xMF_chk_streetaddress.isChecked -eq $true){Set-ADUser $login -StreetAddress $null}
        if(!($xMF_lstv_SingleUser.Items[$i].City -eq "")){$hash.l = $xMF_lstv_SingleUser.Items[$i].city;}elseif($xMF_chk_city.isChecked -eq $true){Set-ADUser $login -City $null}
        if(!($xMF_lstv_SingleUser.Items[$i].PostalCode -eq "")){$hash.postalCode = $xMF_lstv_SingleUser.Items[$i].postalcode;}elseif($xMF_chk_postalcode.isChecked -eq $true){Set-ADUser $login -PostalCode $null}
        if(!($xMF_lstv_SingleUser.Items[$i].mail -eq "")){$hash.mail = $xMF_lstv_SingleUser.Items[$i].mail;}
        if(!($xMF_lstv_SingleUser.Items[$i].ini -eq "")){$hash.initials = $xMF_lstv_SingleUser.Items[$i].ini;}

        
        if ($UserADinfo.Enabled -eq $False)
        {
            try
            {
                Enable-ADAccount -Identity $login
                write-host "Учетная запись пользователя" $login "включена"
            }
            catch
            {
                write-host -BackgroundColor Red -ForegroundColor Yellow "Учетную запись пользователя" $login "включить нельзя" -> $error[0].Exception.Message
            }
        }

        if (!($UserADinfo.Description -eq $null))
        {
            try
            {
                Set-ADUser $login -Description $null
                write-host "Учетной записи пользователя" $login "поле Description очищено"
            }
            catch
            {
                write-host -BackgroundColor Red -ForegroundColor Yellow "Учетной записи пользователя" $login "поле Description очистить нельзя" -> $error[0].Exception.Message
            }
        }

        try
        {
            Set-ADUser $login -Replace $hash -Description $null
            write-host Новые данные пользователю "-" $xMF_lstv_SingleUser.Items[$i].DisplayName "("$xMF_lstv_SingleUser.Items[$i].samaccountname")" внесены
            $xMF_label_prBar.Content = "Новые данные пользователю - "+$xMF_lstv_SingleUser.Items[$i].DisplayName+"("+$xMF_lstv_SingleUser.Items[$i].samaccountname+") внесены"
        }
        catch
        {
            write-host -BackgroundColor Red -ForegroundColor white [УЗ] Ошибка внесения новых данных пользователю $xMF_lstv_SingleUser.Items[$i].DisplayName "("$xMF_lstv_SingleUser.Items[$i].samaccountname")": $error[0].Exception.Message
            $xMF_label_prBar.Content = "Новые данные пользователю не внесены - "+$xMF_lstv_SingleUser.Items[$i].DisplayName+"("+$xMF_lstv_SingleUser.Items[$i].samaccountname+")"
        }

        if (!($xMF_lstv_SingleUser.Items[$i].pass -eq ""))
        {
            $pass = $xMF_lstv_SingleUser.Items[$i].pass
            try
            {
                Set-ADAccountPassword $xMF_lstv_SingleUser.Items[$i].samaccountname -Reset -NewPassword (ConvertTo-SecureString $pass -AsPlainText -force)
                write-host Пароль пользователю "-" $xMF_lstv_SingleUser.Items[$i].DisplayName "("$xMF_lstv_SingleUser.Items[$i].samaccountname")" сброшен на "("$xMF_lstv_SingleUser.Items[$i].pass")"
            }
            catch
            {
                write-host -BackgroundColor Red -ForegroundColor white [УЗ] Ошибка назначения пароля пользователю $xMF_lstv_SingleUser.Items[$i].DisplayName "("$xMF_lstv_SingleUser.Items[$i].samaccountname")": $error[0].Exception.Message
            }
        }
    } # ---END if ($message -eq "OK")
} # ---END if ($xMF_lstv_SingleUser.Items[$i].adcheck -eq "EXIST")
else
{
write-host -ForegroundColor White -BackgroundColor Red У пользователя $login = $xMF_lstv_SingleUser.Items[$i].samaccountname должен быть статус EXIST
}
    $SCount = $SCount + 1
    $xMF_prBar.Value = $SCount
} # ---END for ($i=0 ; $i -ne $xMF_lstv_SingleUser.Items.Count ; $i++)
} # ---END ($message1 -eq "OK")
} # ---END ($xMF_lstv_SingleUser.Items.Count -eq 0)
else
{
    [System.Windows.Forms.MessageBox]::Show("Список пустой","Уведомление","OK","Information")
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
                $login = $xMF_lstv_SingleUser.Items[$i].samaccountname
                $UserDN = $User.distinguishedName
                $TargetOU = $xMF_txtbox_OU_path.Text
                $complete = $false

                try
                {
                    
                    Move-ADObject -Identity $UserDN -TargetPath $TargetOU
                    write-host Пользователь $xMF_lstv_SingleUser.Items[$i].displayname успешно перенесен в $TargetOU
                    $xMF_label_prBar.Content = "Пользователь "+$xMF_lstv_SingleUser.Items[$i].displayname+" успешно перенесен в $TargetOU"
                    $t+=1
                    $complete = $true
                }
                catch
                {
                    $nt+=1
                    write-host Ошибка, перенести пользователя $xMF_lstv_SingleUser.Items[$i].displayname не удалось.
                    $xMF_label_prBar.Content = "Ошибка, перенести пользователя "+$xMF_lstv_SingleUser.Items[$i].displayname+" не удалось."
                    $complete = $false
                }


                    if (($user.Enabled -eq $False) -and ($complete -eq $true))
                    {
                        Try
                        {
                            Enable-ADAccount -Identity $login
                            write-host УЗ -> Пользователя $xMF_lstv_SingleUser.Items[$i].displayname включена
                        }catch
                        {write-host -BackgroundColor Red -ForegroundColor White УЗ -> Пользователя $xMF_lstv_SingleUser.Items[$i].displayname НЕ включена}
                    }

                    if (($xMF_chk_deletgroupsmove.IsChecked -eq $True) -and ($complete -eq $true))
                    {
                        chkboxdeletegroups -user $User -exception "mdaemon"
                    }

                    if (($xMF_chk_deletmanagermove.IsChecked -eq $True) -and ($complete -eq $true))
                    {
                        managers -user $User -empty $true -short $true
                    }

                    $complete = $false

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
        $login = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].samaccountname
        $UserDN = $User.distinguishedName
        $TargetOU = $xMF_txtbox_OU_path.Text
        $complete = $false

        $message = [System.Windows.Forms.MessageBox]::Show("Перенести "+$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].displayname+" в:`n"+$xMF_txtbox_OU_path.Text+" ?","Подтверждение","OKCANCEL","information")

        if ($message -eq "OK")
        {
              
            try
            {
                
                Move-ADObject -Identity $UserDN -TargetPath $TargetOU
                write-host Пользователь $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].displayname успешно перенесен в $TargetOU
                $xMF_label_prBar.Content = "Пользователь "+$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].displayname+" успешно перенесен в $TargetOU"
                $complete = $True
            }
            catch
            {
                [System.Windows.Forms.MessageBox]::Show("Ошибка, перенести пользователя не удалось.","Уведомление","OK","Warning")
                write-host Ошибка, перенести пользователя $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].displayname не удалось.
                $xMF_label_prBar.Content = "Ошибка, перенести пользователя "+$xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].displayname+" не удалось."
                $complete = $false
            }

            if (($user.Enabled -eq $False) -and ($complete -eq $true))
            {
                Try
                {
                Enable-ADAccount -Identity $login
                write-host УЗ -> Пользователя $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].displayname включена
                }catch
                {write-host -BackgroundColor Red -ForegroundColor White УЗ -> Пользователя $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].displayname НЕ включена}
            }

            if (($xMF_chk_deletgroupsmove.IsChecked -eq $True) -and ($complete -eq $true))
            {
                chkboxdeletegroups -user $User -exception "mdaemon"
            }

            if (($xMF_chk_deletmanagermove.IsChecked -eq $True) -and ($complete -eq $true))
            {
                managers -user $User -empty $true -short $true
            }

            $complete = $false
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
Get-OUDialog-start
$windowSelectOU.Show()
$windowSelectOU.Activate()

#$xMF_txtbox_OU_path.Text = Get-OUDialog-start
})

$windowSelectOU.add_MouseLeftButtonDown({
        try
        {
            $this.dragmove()
        }catch
        {
        }
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
#ГРУППЫ
##############################################################



$xMF_btn_groups_load.add_click({

if ($xMF_group_radio_one.IsChecked)
{

if (!($xMF_textbox_Group_newuser.Text -eq "") -and !($xMF_textbox_Group_existuser.Text -eq ""))
{
    if (($xMF_groups_lstv_user1.Items.IsEmpty -eq $False) -or ($xMF_groups_lstv_user2.Items.IsEmpty -eq $False))
    {
        $message = [System.Windows.Forms.MessageBox]::Show("Список имеет записи, он будет очищен для нового списка","Подтверждение","OKCANCEL","information")
        if ($message -eq "OK")
        {
            $xMF_groups_lstv_user1.Items.Clear()
            $xMF_groups_lstv_user2.Items.Clear()
            $gr = getgroups
            if ($gr -eq "ERROR")
            {
                $xMF_groups_lstv_user1.Items.Clear()
                $xMF_groups_lstv_user2.Items.Clear()
            }
            elseif ($gr -eq 0)
            {
                [System.Windows.Forms.MessageBox]::Show("У пользователя"+$xMF_textbox_Group_existuser.Text+"нет групп","Подтверждение","OK","information")
            }
            
        }

    }
    else
    {
        $gr = getgroups
        if ($gr -eq 0)
        {
            [System.Windows.Forms.MessageBox]::Show("У пользователя"+$xMF_textbox_Group_existuser.Text+"нет групп","Подтверждение","OK","information")
        }
    }

}
else
{
    [System.Windows.Forms.MessageBox]::Show("Не заполнено одно из полей","information","OK","information")
}

} # конец $xMF_group_radio_one.IsChecked


if ($xMF_group_radio_many.IsChecked)
{
    if (!($xMF_textbox_Group_newuser.Text -eq "") -and !($xMF_textbox_Group_existuser.Text -eq ""))
    {
        
    }
}

})

$xMF_group_btn_right.add_click({

    if ($xMF_groups_lstv_user1.Items.IsEmpty)
    {
        [System.Windows.Forms.MessageBox]::Show("Нечего переносить","Уведомление","OK","information")
    }
    else
    {
        $xMF_groups_lstv_user2.Items.Add($xMF_groups_lstv_user1.Items[$xMF_groups_lstv_user1.SelectedIndex])
        $xMF_groups_lstv_user1.Items.Remove($xMF_groups_lstv_user1.SelectedItem)
    }

})

$xMF_group_right.add_click({
    if ($xMF_groups_lstv_user1.Items.IsEmpty)
    {
        [System.Windows.Forms.MessageBox]::Show("Нечего переносить","Уведомление","OK","information")
    }
    else
    {
        $xMF_groups_lstv_user2.Items.Add($xMF_groups_lstv_user1.Items[$xMF_groups_lstv_user1.SelectedIndex])
        $xMF_groups_lstv_user1.Items.Remove($xMF_groups_lstv_user1.SelectedItem)
    }
})

$xMF_group_left.add_click({
    if ($xMF_groups_lstv_user2.Items.IsEmpty)
    {
        [System.Windows.Forms.MessageBox]::Show("Нечего переносить","Уведомление","OK","information")
    }
    else
    {
        $xMF_groups_lstv_user1.Items.Add($xMF_groups_lstv_user2.Items[$xMF_groups_lstv_user2.SelectedIndex])
        $xMF_groups_lstv_user2.Items.Remove($xMF_groups_lstv_user2.SelectedItem)
    }
})


$xMF_group_btn_left.add_click({

    if ($xMF_groups_lstv_user2.Items.IsEmpty)
    {
        [System.Windows.Forms.MessageBox]::Show("Нечего переносить","Уведомление","OK","information")
    }
    else
    {
        $xMF_groups_lstv_user1.Items.Add($xMF_groups_lstv_user2.Items[$xMF_groups_lstv_user2.SelectedIndex])
        $xMF_groups_lstv_user2.Items.Remove($xMF_groups_lstv_user2.SelectedItem)
    }

})

$xMF_group_btn_moveall.add_click({

    if ($xMF_groups_lstv_user1.Items.IsEmpty)
    {
        [System.Windows.Forms.MessageBox]::Show("Нечего переносить","Уведомление","OK","information")
    }
    else
    {
        for ($i=0 ; $i -ne $xMF_groups_lstv_user1.Items.Count ; $i++)
        {
        $xMF_groups_lstv_user2.Items.Add($xMF_groups_lstv_user1.Items[$i])
        }
        $xMF_groups_lstv_user1.Items.Clear()
    }

})

$xMF_group_btn_moveall_left.add_click({

if ($xMF_groups_lstv_user2.Items.IsEmpty)
    {
        [System.Windows.Forms.MessageBox]::Show("Нечего переносить","Уведомление","OK","information")
    }
    else
    {
        for ($i=0 ; $i -ne $xMF_groups_lstv_user2.Items.Count ; $i++)
        {
        $xMF_groups_lstv_user1.Items.Add($xMF_groups_lstv_user2.Items[$i])
        }
        $xMF_groups_lstv_user2.Items.Clear()
    }

})


$xMF_group_btn_deleteall.add_click({

$xMF_groups_lstv_user1.Items.Clear()
$xMF_groups_lstv_user2.Items.Clear()

})



$xMF_groups_lstv_user1.add_SelectionChanged({
#$xMF_label_prBar.Content = $xMF_listbox_groups_newuser.Items[$xMF_listbox_groups_newuser.SelectedIndex]
#$xMF_listbox_groups_newuser.Items.Remove($xMF_listbox_groups_newuser.SelectedIndex)
})


$xMF_btn_groups_save.add_click({

if ($xMF_group_radio_one.IsChecked)
{
    if ($xMF_groups_lstv_user2.Items.IsEmpty)
    {
        [System.Windows.Forms.MessageBox]::Show("Список пустой","Уведомление","OK","information")
    }
    else
    {
        $message = [System.Windows.Forms.MessageBox]::Show("Добавить группы пользователю?","Уведомление","OKCANCEL","information")
        if ($message -eq "OK")
        {
            $countOK = 0
            $CountError = 0
            $usertemp = $xMF_textbox_Group_newuser.Text
            for ($i=0 ; $i -ne $xMF_groups_lstv_user2.Items.Count ; $i++)
            {
                $grouptemp = $xMF_groups_lstv_user2.Items[$i]
                try
                {
                    Add-ADGroupMember -Identity $grouptemp -Members $usertemp
                    write-host "ОК группа -> Пользователь" $usertemp "добавлен в группу" $grouptemp
                    $countOK++
                }
                catch
                {
                    write-host -BackgroundColor red -ForegroundColor Yellow "Ошибка добавления ->" $Error[0].Exception.Message
                    $CountError++
                }
            }
        $xMF_groups_lstv_user1.Items.Clear()
        $xMF_groups_lstv_user2.Items.Clear()
        $usertemp = $null
        [System.Windows.Forms.MessageBox]::Show("Добавлено групп: $countOK`nНе добавлено: $CountError","Уведомление","OK","information")
        }
    }

} # Конец $xMF_group_radio_one.IsChecked

})

$xMF_group_radio_one.add_click({

$xMF_btn_group_OU.Visibility = "Hidden"
$xMF_group_right.IsEnabled = $True
$xMF_group_left.IsEnabled = $True
$xMF_group_label_user1.Content = "Новый пользователь"
$xMF_group_label_user2.Content = "Существующий пользователь"
$xMF_group_label_lstv1.Content = "Группы существующего пользователя"
$xMF_group_label_lstv2.Content = "Группы для нового пользователя"
$xMF_group_btn_moveall_left.IsEnabled = $True
$xMF_group_btn_moveall.IsEnabled = $True
$xMF_group_btn_right.IsEnabled = $True
$xMF_group_btn_left.IsEnabled = $True

})

$xMF_group_radio_many.add_click({

$xMF_btn_group_OU.Visibility = "Visible"
$xMF_group_right.IsEnabled = $false
$xMF_group_left.IsEnabled = $false
$xMF_group_label_user1.Content = "Существующий пользователь"
$xMF_group_label_user2.Content = "Путь к OU"
$xMF_group_label_lstv1.Content = "Пользователи"
$xMF_group_label_lstv2.Content = "Группы для добавления"
$xMF_group_btn_moveall_left.IsEnabled = $false
$xMF_group_btn_moveall.IsEnabled = $false
$xMF_group_btn_right.IsEnabled = $false
$xMF_group_btn_left.IsEnabled = $false

})

$xMF_btn_group_OU.add_click({
Get-OUDialog-start
$windowSelectOU.Show()
$windowSelectOU.Activate()
})


#---------------------------------------------------------------------------------------------------------------------------------------------------------



##############################################################
#Всякое
##############################################################

$xMF_btn_clearOU.add_click({

if (!($global:domstruct -eq $null))
{
    $clearOU = [System.Windows.Forms.MessageBox]::Show("Сбросить OU список?","Уведомление","OKCANCEL","Information")
    if ($clearOU -eq "OK")
    {
        $global:domstruct = $null
        $treeviewOUs.Items.Clear()
        
    }
}

})

$xMF_lstv_SingleUser.add_SelectionChanged({

textcheckad

})

$xMF_btn_genpass.add_click({

if ($xMF_pwd_old.IsChecked)
{
$xMF_textbox_pass.Text = New-Password2
}
elseif ($xMF_pwd_new.IsChecked)
{
$xMF_textbox_pass.Text = New-Password
}

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

$xMF_Tab_Inactive.add_MouseLeftButtonup({

    $xMF_menu_settings.IsEnabled = $false
    $xMF_menu_obrab.IsEnabled = $false
    
})

$xMF_Tab_singleUser.add_MouseLeftButtonup({

    $xMF_menu_settings.IsEnabled = $True
    $xMF_menu_obrab.IsEnabled = $True

})

$xMF_Tab_Groups.add_MouseLeftButtonup({

    $xMF_menu_settings.IsEnabled = $false
    $xMF_menu_obrab.IsEnabled = $false

})

$xMF_Btn_Exit.add_click({

    $Form_2.Close()
    $Form_Main.Close()

})


$Form_Main.ShowDialog() | out-null

#Select-Xml $xaml -xpath "//*[@*[contains(translate(name(.),'n','N'),'Name')]]" | Foreach {$_.Node} | Foreach {Remove-Variable -Name "xMF_$($_.Name)"}
#Select-Xml $xaml2 -xpath "//*[@*[contains(translate(name(.),'n','N'),'Name')]]" | Foreach {$_.Node} | Foreach {Remove-Variable -Name "xMF2_$($_.Name)"}
#Remove-Variable -Name xaml
#Remove-Variable -Name xaml2
#Remove-Variable -Name XReader
#Remove-Variable -Name XReader2
#Remove-Variable -Name Form_Main
#Remove-Variable -Name Form_2
