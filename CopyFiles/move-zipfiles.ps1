function move-zipfiles ([string]$SourcePTH="C:\Program Files\Microsoft\Exchange Server\V15\Logging\",[string]$DestPTH="\\pso-fs-101\E$\Exchange\Archive_LOGS_2016EX\",$filter="zip",[bool]$show=$true)
{
#IIS logs - C:\inetpub\logs\LogFiles\
#Exchange logs - C:\Program Files\Microsoft\Exchange Server\V15\Logging\

[array]$1 = Get-ChildItem $SourcePTH -Filter ("*."+$filter) -Recurse | select name,Directory,Length

if ($show -eq $False)
{

for ($i=0; $i -ne $1.Count ; $i++)
{
    #To find subfolders. We delete SourcePTH from full Directory and get folders that we will create in DestPTH
    $DIREC = $1[$i].directory.Fullname.Replace($SourcePTH,"")
    $DIREC2 = $DestPTH+$DIREC+"\"
    
    #Path to file that we move
    $PATH = $1[$i].directory.Fullname+"\"+$1[$i].name
    #Deleting extension in filename to add digits in the end of filename if exists.
    $FileName = $1[$i].name.Replace("."+$filter,"")

    #Checking file existence
    if (Test-Path ($PATH))
    {
        #if there is no folder in destPTH. Creating a folders.
        if (!(Test-Path ($DIREC2)))
        {
            New-Item -Path $DIREC2 -ItemType Directory | out-null
            Write-host -BackgroundColor Yellow "Created Directory $DIREC2"
        }
        
        #if folder, in destPTH, exists. Moving files.
        if (Test-Path ($DIREC2))
        {
            #If files exists, we add digits to the end of the filename.
            if (Test-Path ($DIREC2+$FileName+"."+$filter))
            {
                $TMP = $DIREC2+$FileName+"-1."+$filter
                $TMPCOUNT = 2
                while (Test-Path $TMP)
                {
                    $TMP = $DIREC2+$FileName+"-$TMPCOUNT."+$filter
                    $TMPCOUNT++
                }
                Move-Item -Path $PATH -Destination $TMP
                write-host "----------------------"
                write-host "Copying $PATH"
                write-host "to $TMP"
            }
            else
            {
                $TMP2= $DIREC2+$FileName+"."+$filter
                Move-Item -Path $PATH -Destination $TMP2
                write-host "----------------------"
                write-host "Copying $PATH"
                write-host "to $TMP2"
            }
        }
    }
    
}


} # End of (if $show)
else #Just shows info about files 
{
    foreach ($array in $1)
    {
        "Name = "+$array.name+" Length = "+"{0:N2}" -f $($array.Length/1MB) +"MB"
    }
    $sum = $1 | measure -Property length -Sum
    Write-Host "-------------------------"
    "{0:N2}" -f $($sum.Sum/1GB) +"GB";

}

}
