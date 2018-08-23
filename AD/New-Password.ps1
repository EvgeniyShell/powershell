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




function new-password2 
{
$CharSets = 'A','B','C','D','E','F','G','H','J','K','L','M','N','P','Q','R','S','T','U','V','W','X','Y','Z','a','b','c','d','e','f','g','h','i','j','k','m','n','p','q','r','s','t','u','v','w','x','y','z'
$CharSets2 = '1','2','3','4','5','6','7','8','9'
$CharSets3 = '!','#','%','&','+','*'

$Letter = ""
$Number = ""
$Symbol = ""
for ($i=0 ; $i -ne 5 ; $i++)
{
$Letter += Get-Random $CharSets
}

for ($i=0 ; $i -ne 2 ; $i++)
{
$Number += Get-Random $CharSets2
}

for ($i=0 ; $i -ne 1 ; $i++)
{
$Symbol += Get-Random $CharSets3
}
$ALLtogether = $Number +  $Symbol +  $Letter

return $ALLtogether
}