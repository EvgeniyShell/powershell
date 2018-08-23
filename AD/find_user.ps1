
Function find_user([string]$SelectMail="",[int]$checkcount=0,[bool]$multi=$true)
{


#Забираем пока что только ФИО
if ($multi -eq $true)
{
    $displayname = $xMF_lstv_SingleUser.Items[$checkcount].displayname -replace "`n|`r",""

}
else
{
    $displayname = $xMF_lstv_SingleUser.Items[$xMF_lstv_SingleUser.SelectedIndex].displayname -replace "`n|`r",""
}

#Ищем пользователя в АД
$user = Get-ADUser -filter "Name -like '$displayname*'" -Properties CanonicalName,Name,SamAccountName,Initials,UserPrincipalName,DisplayName,GivenName,surName,OfficePhone,EmailAddress,department,Title,City,StreetAddress,PostalCode,Country,Company,office,description

#Создаем инициалы
$ini = $displayname.Split(" ")
$ini2 = $ini[1].Chars(0)
try
    {
        $ini3 = $ini[2].Chars(0)
    }
catch
    {
        Write-Host -BackgroundColor DarkYellow "Отчество не может быть получено - ($displayname). Добавьте отчество пользователю для правильного отображения инициалов."
        $ini3 = ""
    }
if ($ini3 -eq "")
{
    $ini4 = "$ini2"+"."
}
else
{
    $ini4 = "$ini2"+"."+"$ini3"+"."
}



# Если пользователь по ФИО в АД не найден
if ($user -eq $null)
{
#Генерируем пароль в переменную

if ($xMF_pwd_old.IsChecked)
{
    $pass = New-Password2
}
elseif ($xMF_pwd_new.IsChecked)
{
    $pass = New-Password -PasswordLength 8
}

#Генерируем логин в переменную
$login = login "$displayname"

#Условие для генерации столбика mail в файле
if ($selectmail -eq "")
{$smail = ""}
else
{$smail = $login + "@" + "$selectmail"}

#Ищем, занят ли логин
$user2 = Get-ADUser -filter{samaccountname -eq $login} | select name,samaccountname
    #Если найден совпадающий логин в АД
    if ($user2.samaccountname -eq $login)
    {
    #переменная для добавления символов имени (см. функцию login)
    $i=1
    #Переменная для работы с функцией логин т.к. у функции несколько параметров (см. функцию login)
    $iteration=0
        #Пока логины в АД совпадают с генерированым
        while ($user2.samaccountname -eq $login)
        {
            #Сравнение, первый ли раз прошла итерация, если да то используем функцию Login с одними параметрами
            if ($iteration -eq "0")
            {
            $login = login "$displayname" -Pexist $true
            $user2 = Get-ADUser -filter{samaccountname -eq $login} | select name,samaccountname
            }
            #Иначе используем другие параметры функции Login с переменной i
            else
            {
            $login = login "$displayname" -Pexist $true -PexistNum $i
            $user2 = Get-ADUser -filter{samaccountname -eq $login} | select name,samaccountname
            $i=$i+1
            }
        $iteration=$iteration+1
        }
        #Выводим данные с присвоением уникального логина
        

        if ($multi -eq $true)
        {
            $xmf_lstv_SingleUser.Items[$checkcount].samaccountname = $login
            if ($xmf_lstv_SingleUser.Items[$checkcount].pass -eq "")
            {
                $xmf_lstv_SingleUser.Items[$checkcount].pass = $pass
            }
            $xmf_lstv_SingleUser.Items[$checkcount].mail = $smail
            $xmf_lstv_SingleUser.Items[$checkcount].ini = $ini4
            $xmf_lstv_SingleUser.Items[$checkcount].adcheck = "Checked"
            $xmf_lstv_SingleUser.Items.Refresh()
            Write-Host -BackgroundColor Yellow -ForegroundColor black "Пользователю $displayname сгенерирован логин - $login (пароль"$xmf_lstv_SingleUser.Items[$checkcount].pass")"

            
        }
        else
        {
            $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].samaccountname = $login
            if ($xmf_lstv_SingleUser.Items[$checkcount].pass -eq "")
            {
                $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].pass = $pass
            }
            $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].mail = $smail
            $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].ini = $ini4
            $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck = "Checked"
            $xmf_lstv_SingleUser.Items.Refresh()
        }

        
    }
    #Если логин с самого начала уникальный
    else
    {
        

        if ($multi -eq $true)
        {
            $xmf_lstv_SingleUser.Items[$checkcount].samaccountname = $login
            if ($xmf_lstv_SingleUser.Items[$checkcount].pass -eq "")
            {
                $xmf_lstv_SingleUser.Items[$checkcount].pass = $pass
            }
            $xmf_lstv_SingleUser.Items[$checkcount].mail = $smail
            $xmf_lstv_SingleUser.Items[$checkcount].ini = $ini4
            $xmf_lstv_SingleUser.Items[$checkcount].adcheck = "Checked"
            $xmf_lstv_SingleUser.Items.Refresh()

        }
        else
        {
            $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].samaccountname = $login
            if ($xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].pass -eq "")
            {
                $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].pass = $pass
            }
            $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].mail = $smail
            $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].ini = $ini4
            $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck = "Checked"
            $xmf_lstv_SingleUser.Items.Refresh()
        }
        Write-Host -BackgroundColor Yellow -ForegroundColor black "Пользователю $displayname сгенерирован логин - $login (пароль"$xmf_lstv_SingleUser.Items[$checkcount].pass")"
        
    }
}
#Если ФИО существует, выводим по нему данные в Listview_Exist
else 
    {
     

        $Spisok = [PSCustomObject]@{
            'firstname' = $user.GivenName
            'lastname'= $user.surName
            'displayname' = $user.DisplayName
            'office' = $user.office
            'OfficePhone' = $user.OfficePhone
            'jobtitle' = $user.Title
            'department' = $user.department
            'company' = $user.Company
            'streetaddress' = $user.StreetAddress
            'city' = $user.City
            'postalcode' = $user.PostalCode
           'samaccountname' = $user.SamAccountName
           #'pass' = $pass
           'mail' = $user.EmailAddress
           'ini' = $user.Initials
           'enabled' = $user.enabled
           'OU' = $user.CanonicalName
           'Description' = $user.description
         }
        $xMF_lstv_SingleUser_Exist.Items.Add($Spisok)

        if ($multi -eq $true)
        {
            $xmf_lstv_SingleUser.Items[$checkcount].adcheck = "EXIST"
            $xmf_lstv_SingleUser.Items[$checkcount].samaccountname = $user.SamAccountName
            $xmf_lstv_SingleUser.Items[$checkcount].ini = $user.Initials
            
        }
        else
        {
            $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck = "EXIST"
            $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].samaccountname = $user.SamAccountName
            $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].ini = $user.Initials
        }
          
        $xmf_lstv_SingleUser.Items.Refresh()
       # $xMF_lstv_SingleUser.items.remove($xMF_lstv_SingleUser.Items[$checkcount])

        Write-Host Пользователь $displayname существует в АД"," логин - $user.samaccountname
    }



}