Function New-Password {
<#
.SYNOPSIS
	Generates passwords at required strength

.DESCRIPTION
	This function creates a password. If it fails to get get characters from all character sets, it will rerun itself. For a maximum of 100 loops.

.PARAMETER PasswordLength
Specifies number of characters in returned password.

.PARAMETER SelectSets
Allows you to select which characters sets to use.

.PARAMETER RequiredNumberofSets
Number different sets password has to be formed of.
.PARAMETER CustomSets
Input an array of strings to generate password from.
.PARAMETER Invocation
Count which number invocation this is, to stop runaway loops

.INPUTS
Most common used are SelectSets and PasswordLength

.OUTPUTS
	[String]

.EXAMPLE
New-Password

b3M3mc1<

.EXAMPLE
New-Password -PasswordLength 12 -SelectSets UpperCase,LowerCase,Special

%zXWKF^rtgpp

.EXAMPLE
New-Password -CustomSets 'ABCDE','abcde','12345'

ae2A3b42

.EXAMPLE
New-Password -PasswordLength 3 -SelectSets UpperCase,LowerCase,Special -RequiredNumberofSets 1

q@v

.LINK
Script center: http://gallery.technet.microsoft.com/scriptcenter/New-Password-Yet-Another-fdc44a10
My Blog: http://virot.eu
Blog Entry: http://virot.eu/wordpress/new-password-yet-another-password-function

.NOTES
Author:	Oscar Virot - virot@virot.com
Filename: New-Password.ps1
Version: 2013-09-20

.FUNCTIONALITY
   Generates random passwords
#>
	Param(
		[Parameter(Mandatory=$False)]
		[Int32]
		[Alias('Length')]
		$PasswordLength=8,
		[parameter(Mandatory=$False)]
		[String[]]
		[ValidateSet('All','UpperCase','LowerCase','Numerical','Special')]
		$SelectSets='All',
		[Parameter(Mandatory=$False)]
		[Alias('Required')]
		[Int32]
		$RequiredNumberofSets=-1,
		[parameter(Mandatory=$False)]
		[String[]]
		$CustomSets,
		[Parameter(Mandatory=$False)]
		[Int32]
		$Invocation=0

	)	
Begin
	{
		$CharSets = $Null; $CharSets = @()
		if ($CustomSets -eq $Null)
		{
			if ($SelectSets -contains 'All' -or $SelectSets -contains 'UpperCase')
			{
				$CharSets += @{'Chars'=@('A','B','C','D','E','F','G','H','J','K','L','M','N','P','Q','R','S','T','U','V','W','X','Y','Z');'useCount'=0}
			}
			if ($SelectSets -contains 'All' -or $SelectSets -contains 'LowerCase')
			{
				$CharSets += @{'Chars'=@('a','b','c','d','e','f','g','h','i','j','k','m','n','p','q','r','s','t','u','v','w','x','y','z');'useCount'=0}
			}
			if ($SelectSets -contains 'All' -or $SelectSets -contains 'Numerical')
			{
				$CharSets += @{'Chars'=@('1','2','3','4','5','6','7','8','9');'useCount'=0}
			}
			if ($SelectSets -contains 'All' -or $SelectSets -contains 'Special')
			{
				$CharSets += @{'Chars'=@('!','#','%','&','+','*');'useCount'=0}
			}
		}
		else
		{
			ForEach($Set in $CustomSets)
			{
				$CharSets += @{'Chars'=[char[]]$Set;'useCount'=0}
			}
		}
		if ($RequiredNumberofSets -gt $CharSets.Count)
		{
			Throw('Required number of sets is greater then number of sets')
		}
		elseif ($RequiredNumberofSets -eq -1)
		{
			$RequiredNumberofSets = $CharSets.Count
		}
		if ($RequiredNumberofSets -gt $PasswordLength)
		{
			Throw('Required number of sets is greater then length of requested password.')
		}

	}
	Process
	{
		$Password = ''
#Failsafe for Iteration 50
		if ($Invocation -eq 50)
		{
#Do the following loop until password requirements has been met
			Do
			{
#Build the next available character from unused charactersets only
				$CharsLeft = $Charsets| ? {$_.useCount -eq 0} | % {$_.Chars}
				$passwordchar = $CharsLeft[(Get-Random -Min 0 -Max (([array]$CharsLeft).count -1))]
				$Password += $passwordchar
				ForEach ($tempCS in $CharSets)
				{
					if ($tempCS.Chars -ccontains $passwordchar)
					{
						$tempCS.useCount=1
					}
				}
			}While (($charsets | % {$_.useCount}| Measure-Object -Sum| Select-Object -expand sum) -lt $RequiredNumberofSets)
		}
#Build a complete set of characters.
		$AllChars = $Charsets | % {$_.Chars}
		For ($i = $password.Length; $i -lt $PasswordLength; $i++)
		{
			$passwordchar = $AllChars[(Get-Random -Min 0 -Max ($AllChars.count -1))]
			$Password += $passwordchar
			ForEach ($tempCS in $CharSets)
			{
				if ($tempCS.Chars -ccontains $passwordchar)
				{
					$tempCS.useCount=1
				}
			}
		}
		if (($charsets | % {$_.useCount}| Measure-Object -Sum| Select-Object -expand sum) -ge $RequiredNumberofSets)
		{
			return $Password
		}
		return New-Password -PasswordLength:$PasswordLength -SelectSets:$SelectSets -Invocation:($invocation+1) -CustomSets:$CustomSets -RequiredNumberofSets:$RequiredNumberofSets
	}
}



function New-SWRandomPassword {
    <#
    .Synopsis
       Generates one or more complex passwords designed to fulfill the requirements for Active Directory
    .DESCRIPTION
       Generates one or more complex passwords designed to fulfill the requirements for Active Directory
    .EXAMPLE
       New-SWRandomPassword
       C&3SX6Kn

       Will generate one password with a length between 8  and 12 chars.
    .EXAMPLE
       New-SWRandomPassword -MinPasswordLength 8 -MaxPasswordLength 12 -Count 4
       7d&5cnaB
       !Bh776T"Fw
       9"C"RxKcY
       %mtM7#9LQ9h

       Will generate four passwords, each with a length of between 8 and 12 chars.
    .EXAMPLE
       New-SWRandomPassword -InputStrings abc, ABC, 123 -PasswordLength 4
       3ABa

       Generates a password with a length of 4 containing atleast one char from each InputString
    .EXAMPLE
       New-SWRandomPassword -InputStrings abc, ABC, 123 -PasswordLength 4 -FirstChar abcdefghijkmnpqrstuvwxyzABCEFGHJKLMNPQRSTUVWXYZ
       3ABa

       Generates a password with a length of 4 containing atleast one char from each InputString that will start with a letter from 
       the string specified with the parameter FirstChar
    .OUTPUTS
       [String]
    .NOTES
       Written by Simon Wåhlin, blog.simonw.se
       I take no responsibility for any issues caused by this script.
    .FUNCTIONALITY
       Generates random passwords
    .LINK
       http://blog.simonw.se/powershell-generating-random-password-for-active-directory/
   
    #>
    [CmdletBinding(DefaultParameterSetName='FixedLength',ConfirmImpact='None')]
    [OutputType([String])]
    Param
    (
        # Specifies minimum password length
        [Parameter(Mandatory=$false,
                   ParameterSetName='RandomLength')]
        [ValidateScript({$_ -gt 0})]
        [Alias('Min')] 
        [int]$MinPasswordLength = 8,
        
        # Specifies maximum password length
        [Parameter(Mandatory=$false,
                   ParameterSetName='RandomLength')]
        [ValidateScript({
                if($_ -ge $MinPasswordLength){$true}
                else{Throw 'Max value cannot be lesser than min value.'}})]
        [Alias('Max')]
        [int]$MaxPasswordLength = 12,

        # Specifies a fixed password length
        [Parameter(Mandatory=$false,
                   ParameterSetName='FixedLength')]
        [ValidateRange(1,2147483647)]
        [int]$PasswordLength = 8,
        
        # Specifies an array of strings containing charactergroups from which the password will be generated.
        # At least one char from each group (string) will be used.
        [String[]]$InputStrings = @('abcdefghijkmnpqrstuvwxyz', 'ABCEFGHJKLMNPQRSTUVWXYZ', '23456789', '!#%&'),

        # Specifies a string containing a character group from which the first character in the password will be generated.
        # Useful for systems which requires first char in password to be alphabetic.
        [String] $FirstChar,
        
        # Specifies number of passwords to generate.
        [ValidateRange(1,2147483647)]
        [int]$Count = 1
    )
    Begin {
        Function Get-Seed{
            # Generate a seed for randomization
            $RandomBytes = New-Object -TypeName 'System.Byte[]' 4
            $Random = New-Object -TypeName 'System.Security.Cryptography.RNGCryptoServiceProvider'
            $Random.GetBytes($RandomBytes)
            [BitConverter]::ToUInt32($RandomBytes, 0)
        }
    }
    Process {
        For($iteration = 1;$iteration -le $Count; $iteration++){
            $Password = @{}
            # Create char arrays containing groups of possible chars
            [char[][]]$CharGroups = $InputStrings

            # Create char array containing all chars
            $AllChars = $CharGroups | ForEach-Object {[Char[]]$_}

            # Set password length
            if($PSCmdlet.ParameterSetName -eq 'RandomLength')
            {
                if($MinPasswordLength -eq $MaxPasswordLength) {
                    # If password length is set, use set length
                    $PasswordLength = $MinPasswordLength
                }
                else {
                    # Otherwise randomize password length
                    $PasswordLength = ((Get-Seed) % ($MaxPasswordLength + 1 - $MinPasswordLength)) + $MinPasswordLength
                }
            }

            # If FirstChar is defined, randomize first char in password from that string.
            if($PSBoundParameters.ContainsKey('FirstChar')){
                $Password.Add(0,$FirstChar[((Get-Seed) % $FirstChar.Length)])
            }
            # Randomize one char from each group
            Foreach($Group in $CharGroups) {
                if($Password.Count -lt $PasswordLength) {
                    $Index = Get-Seed
                    While ($Password.ContainsKey($Index)){
                        $Index = Get-Seed                        
                    }
                    $Password.Add($Index,$Group[((Get-Seed) % $Group.Count)])
                }
            }

            # Fill out with chars from $AllChars
            for($i=$Password.Count;$i -lt $PasswordLength;$i++) {
                $Index = Get-Seed
                While ($Password.ContainsKey($Index)){
                    $Index = Get-Seed                        
                }
                $Password.Add($Index,$AllChars[((Get-Seed) % $AllChars.Count)])
            }
            Write-Output -InputObject $(-join ($Password.GetEnumerator() | Sort-Object -Property Name | Select-Object -ExpandProperty Value))
        }
    }
}




function new-password2 
{
$CharSets = 'A','B','C','D','E','F','G','H','J','K','L','M','N','P','Q','R','S','T','U','V','W','X','Y','Z','a','b','c','d','e','f','g','h','i','j','k','m','n','p','q','r','s','t','u','v','w','x','y','z'
$CharSetshigh = 'A','B','C','D','E','F','G','H','J','K','L','M','N','P','Q','R','S','T','U','V','W','X','Y','Z'
$CharSetslow = 'a','b','c','d','e','f','g','h','i','j','k','m','n','p','q','r','s','t','u','v','w','x','y','z'
$CharSets2 = '1','2','3','4','5','6','7','8','9'
$CharSets3 = '!','#','%','&','+','*',"@"

$Letter = ""
$Number = ""
$Symbol = ""
#for ($i=0 ; $i -ne 5 ; $i++)
#{
$Letter += -join ($CharSets | Get-Random -Count 5)
#}

#for ($i=0 ; $i -ne 2 ; $i++)
#{
$Number += -join ($CharSets2 | Get-Random -Count 2)
#}

#for ($i=0 ; $i -ne 1 ; $i++)
#{
$Symbol += -join ($CharSets3 | Get-Random -Count 1)
#}
$ALLtogether = $Number +  $Symbol +  $Letter

return $ALLtogether
}