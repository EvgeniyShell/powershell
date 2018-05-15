function login($Pname,[bool]$Pexist=$false,[int]$PexistNum=1,[bool]$PExistOtch=$true)
{
    #Для получения транслита, что бы можно было далее посчитать кол-во символов в имени.
    if ($pname -ne $null)
    {
        $translit = TranslitToLAT($pname)
        $translitSplit = $translit.Split(" ").ToLower()
    }

    #Основное исключение, если поле пустое вывести ошибку, далее основной блок кода
    if ($pname -eq $null)
    {
        Write-Host "Введите ФИО полностью"
    }
    elseif ($pexist -eq $false)
    {
        $translitFIO = $translitSplit[0]
        $translitNAME = $translitSplit[1].Chars(0)
        $translitALL = $translitNAME+"."+$translitFIO
        $translitALL
    }
    elseif ($pexist -eq $true)
    {
        if (($PexistNum -lt "1") -or ($PexistNum -gt $translitSplit[1].Length)){Write-Host "Количество символов не может быть меньше 1 или больше символов в имени"}

        else{
        $translitFIO = $translitSplit[0]
        $i=0
        while ($i -ne $PexistNum)
        {
        $translitNAME += $translitSplit[1].Chars($i)
        $i = $i+1
        }
        $translitOTCH = $translitSplit[2].Chars(0)

        if ($PExistOtch -eq $false){ 
        $translitALL = $translitNAME+"."+$translitFIO
        $translitALL
        }
        elseif ($PExistOtch -eq $true){
        $translitALL = $translitNAME+"."+$translitOTCH+"."+$translitFIO
        $translitALL
        }

            }
    }

}