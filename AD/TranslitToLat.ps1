function global:TranslitToLAT
{
param([string]$inString)
 
$Translit_To_LAT = @{
[char]'а' = "a"
[char]'А' = "A"
[char]'б' = "b"
[char]'Б' = "B"
[char]'в' = "v"
[char]'В' = "V"
[char]'г' = "g"
[char]'Г' = "G"
[char]'д' = "d"
[char]'Д' = "D"
[char]'е' = "e"
[char]'Е' = "E"
[char]'ё' = "e"
[char]'Ё' = "E"
[char]'ж' = "zh"
[char]'Ж' = "Zh"
[char]'з' = "z"
[char]'З' = "Z"
[char]'и' = "i"
[char]'И' = "I"
[char]'й' = "y"
[char]'Й' = "Y"
[char]'к' = "k"
[char]'К' = "K"
[char]'л' = "l"
[char]'Л' = "L"
[char]'м' = "m"
[char]'М' = "M"
[char]'н' = "n"
[char]'Н' = "N"
[char]'о' = "o"
[char]'О' = "O"
[char]'п' = "p"
[char]'П' = "P"
[char]'р' = "r"
[char]'Р' = "R"
[char]'с' = "s"
[char]'С' = "S"
[char]'т' = "t"
[char]'Т' = "T"
[char]'у' = "u"
[char]'У' = "U"
[char]'ф' = "f"
[char]'Ф' = "F"
[char]'х' = "kh"
[char]'Х' = "Kh"
[char]'ц' = "ts"
[char]'Ц' = "Ts"
[char]'ч' = "ch"
[char]'Ч' = "Ch"
[char]'ш' = "sh"
[char]'Ш' = "Sh"
[char]'щ' = "sch"
[char]'Щ' = "Sch"
[char]'ъ' = "787" # "``"
[char]'Ъ' = "787" # "``"
[char]'ы' = "y" # "y`"
[char]'Ы' = "Y" # "Y`"
[char]'ь' = "787" # "`"
[char]'Ь' = "787" # "`"
[char]'э' = "e" # "e`"
[char]'Э' = "E" # "E`"
[char]'ю' = "yu"
[char]'Ю' = "Yu"
[char]'я' = "ya"
[char]'Я' = "Ya"
}
 
$outChars=""
 
foreach ($c in $inChars = $inString.ToCharArray())
{
if ($Translit_To_LAT[$c] -cne $Null )
{
$outChars += $Translit_To_LAT[$c] -replace "787",""
}
else
{
$outChars += $c
}
 
}
 
Write-Output $outChars
 
}
#$inS = "Иванов Иван"
#TranslitToLAT($inS)