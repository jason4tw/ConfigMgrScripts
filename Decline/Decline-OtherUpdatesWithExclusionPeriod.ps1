# ===============================================
# Script to decline superseeded updates in WSUS.
# ===============================================
# It's recommended to run the script with the -SkipDecline switch to see how many superseded updates are in WSUS and to TAKE A BACKUP OF THE SUSDB before declining the updates.
# Parameters:

# $UpdateServer             = Specify WSUS Server Name
# $UseSSL                   = Specify whether WSUS Server is configured to use SSL
# $Port                     = Specify WSUS Server Port
# $SkipDecline              = Specify this to do a test run and get a summary of how many superseded updates we have
# $DeclineLastLevelOnly     = Specify whether to decline all superseded updates or only last level superseded updates
# $ExclusionPeriod          = Specify the number of days between today and the release date for which the superseded updates must not be declined. Eg, if you want to keep superseded updates published within the last 2 months, specify a value of 60 (days)
# $DeclineItanium           = Specify whether to decline all itanium updates
# $DeclineBeta              = Specify whether to decline all beta updates
# $DeclineOther             = Specify whether to decline all other updates and the string to match on (caution! verify this first in the output file)

# Supersedence chain could have multiple updates. 
# For example, Update1 supersedes Update2. Update2 supersedes Update3. In this scenario, the Last Level in the supersedence chain is Update3. 
# To decline only the last level updates in the supersedence chain, specify the DeclineLastLevelOnly switch

# Usage:
# =======

# To do a test run against WSUS Server without SSL
# Decline-SupersededUpdates.ps1 -UpdateServer SERVERNAME -Port 8530 -SkipDecline

# To do a test run against WSUS Server using SSL
# Decline-SupersededUpdates.ps1 -UpdateServer SERVERNAME -UseSSL -Port 8531 -SkipDecline

# To decline all superseded updates on the WSUS Server using SSL
# Decline-SupersededUpdates.ps1 -UpdateServer SERVERNAME -UseSSL -Port 8531

# To decline only Last Level superseded updates on the WSUS Server using SSL
# Decline-SupersededUpdates.ps1 -UpdateServer SERVERNAME -UseSSL -Port 8531 -DeclineLastLevelOnly

# To decline all superseded updates on the WSUS Server using SSL but keep superseded updates published within the last 2 months (60 days)
# Decline-SupersededUpdates.ps1 -UpdateServer SERVERNAME -UseSSL -Port 8531 -ExclusionPeriod 60


[CmdletBinding()]
Param(
	[Parameter(Mandatory=$False,Position=1)]
    [string] $UpdateServer = $($env:computername),
	
	[Parameter(Mandatory=$False)]
    [switch] $UseSSL,
	
	[Parameter(Mandatory=$False, Position=2)]
    [int]$Port = 8530,
	
    [switch] $SkipDecline,
    [switch] $DeclineItanium,
    [switch] $DeclineBeta,
    [string] $DeclineOther,	
	
    [Parameter(Mandatory=$False)]
    [int] $ExclusionPeriod = 0
)

Write-Host ""

$outPath = Split-Path $script:MyInvocation.MyCommand.Path
$outUpdateList = Join-Path $outPath "OtherUpdates.csv"
$outUpdateListBackup = Join-Path $outPath "OtherUpdatesBackup.csv"
"UpdateID, RevisionNumber, Title, KBArticle, SecurityBulletin, LastLevel" | Out-File $outUpdateList

try {
    
    if ($UseSSL) {
        Write-Host "Connecting to WSUS server $UpdateServer on Port $Port using SSL... " -NoNewLine
    } Else {
        Write-Host "Connecting to WSUS server $UpdateServer on Port $Port... " -NoNewLine
    }
    
    [reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null
    $wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($UpdateServer, $UseSSL, $Port);
}
catch [System.Exception] 
{
    Write-Host "Failed to connect."
    Write-Host "Error:" $_.Exception.Message
    Write-Host "Please make sure that WSUS Admin Console is installed on this machine"
	Write-Host ""
    $wsus = $null
}

if ($wsus -eq $null) { return } 

Write-Host "Connected."

$countAllUpdates = 0
$countDeclined = 0

$countBetaAll =0 
$countOtherAll = 0
$countItaniumAll = 0
$countBetaDeclined =0 
$countOtherDeclined = 0
$countItaniumDeclined = 0
$countBetaExclusionPeriod =0 
$countOtherExclusionPeriod = 0
$countItaniumExclusionPeriod = 0

Write-Host "Getting a list of all updates... " -NoNewLine

try {
	$allUpdates = $wsus.GetUpdates()
}

catch [System.Exception]
{
	Write-Host "Failed to get updates."
	Write-Host "Error:" $_.Exception.Message
    Write-Host "If this operation timed out, please decline the superseded updates from the WSUS Console manually."
	Write-Host ""
	return
}

Write-Host "Done"

Write-Host "Parsing the list of updates... " -NoNewLine

foreach($update in $allUpdates) {
    
    $countAllUpdates++

    if ($update.IsDeclined) {
        $countDeclined++
    }

    $addToOutputFile = $false

    if ($DeclineItanium -and ($update.Title -like '*Itanium*' -or $update.Title -like '*ia64*')) {

        $countItaniumAll++

        if ($update.IsDeclined) {
            $countItaniumDeclined++
        }
        
        if (!$update.IsDeclined) {
            if ($update.CreationDate -lt (get-date).AddDays(-$ExclusionPeriod))  {
                $countItaniumExclusionPeriod++

                $addToOutputFile = $true
            }
        }
    }
    
    if ($DeclineOther.Length -gt 0 -and $update.Title -like "*$DeclineOther*") {

        $countOtherAll++

        if ($update.IsDeclined) {
            $countOtherDeclined++
        }
        
        if (!$update.IsDeclined) {
            if ($update.CreationDate -lt (get-date).AddDays(-$ExclusionPeriod))  {
                $countOtherExclusionPeriod++

                $addToOutputFile = $true
            }
        }
    }

    if ($DeclineBeta -and $update.Title -like '*Beta*') {

        $countBetaAll++

        if ($update.IsDeclined) {
            $countBetaDeclined++
        }
        
        if (!$update.IsDeclined) {
            if ($update.CreationDate -lt (get-date).AddDays(-$ExclusionPeriod))  {
                $countBetaExclusionPeriod++

                $addToOutputFile = $true
            }
        }
    }

    if($addToOutputFile) {
        "$($update.Id.UpdateId.Guid), $($update.Id.RevisionNumber), $($update.Title), $($update.KnowledgeBaseArticles), $($update.SecurityBulletins), $($update.HasSupersededUpdates)" | Out-File $outUpdateList -Append       
    }

}

Write-Host "Done."
Write-Host "List of identified updates: $outUpdateList"

Write-Host ""
Write-Host "Summary:"
Write-Host "========"

Write-Host "All Updates =" $countAllUpdates
Write-Host "Any except Declined =" ($countAllUpdates - $countDeclined)

if($DeclineItanium) {
    Write-Host "All Itanium Updates =" $countItaniumAll
    Write-Host "    Itanium Updates (not declined) =" ($countItaniumAll - $countItaniumDeclined)
    Write-Host "    Itanium Updates (not declined and older than $ExclusionPeriod days) =" $countItaniumExclusionPeriod
}

if($DeclineBeta) {
    Write-Host "All Beta Updates =" $countBetaAll
    Write-Host "    Beta Updates (not declined) =" ($countBetaAll - $countBetaDeclined)
    Write-Host "    Beta Updates (not declined and older than $ExclusionPeriod days) =" $countBetaExclusionPeriod
}

if($DeclineOther.Length -gt 0) {
    Write-Host "All Other Updates =" $countOtherAll
    Write-Host "    Other Updates (not declined) =" ($countOtherAll - $countOtherDeclined)
    Write-Host "    Other Updates (not declined and older than $ExclusionPeriod days) =" $countOtherExclusionPeriod
}

Write-Host ""

$i = 0
if (!$SkipDecline) {
    
    Write-Host "SkipDecline flag is set to $SkipDecline. Continuing with declining updates"
    $updatesDeclined = 0
    $updatesToDecline = $countItaniumExclusionPeriod + $countBetaExclusionPeriod + $countOtherExclusionPeriod
        
    foreach ($update in $allUpdates) {
        
        if (!$update.IsDeclined -and
            (($DeclineItanium -and ($update.Title -like '*Itanium*' -or $update.Title -like '*ia64*')) -or
            ($DeclineBeta -and $update.Title -like '*Beta*') -or
            ($DeclineOther.Length -gt 0 -and $update.Title -like "*$DeclineOther*"))) {
            if ($update.CreationDate -lt (get-date).AddDays(-$ExclusionPeriod))  {   
            
                $i++
                $percentComplete = "{0:N2}" -f (($updatesDeclined/$updatesToDecline) * 100)
                Write-Progress -Activity "Declining Updates" -Status "Declining update #$i/$updatesToDecline - $($update.Id.UpdateId.Guid)" -PercentComplete $percentComplete -CurrentOperation "$($percentComplete)% complete"
                try 
                {
                    $update.Decline()
                    $updatesDeclined++
                }
                catch [System.Exception]
                {
                    Write-Warning "Failed to decline update $($update.Id.UpdateId.Guid). Error:" $_.Exception.Message
                }
            }              
        }
    }   
    
    Write-Host "  Declined $updatesDeclined updates."
    if ($updatesDeclined -ne 0) {
        Copy-Item -Path $outUpdateList -Destination $outUpdateListBackup -Force
		Write-Host "  Backed up list of previosuly identified updates to $outUpdateListBackup"
    }
    
}
else {
    Write-Host "SkipDecline flag is set to $SkipDecline. Skipped declining updates"
}

Write-Host ""
Write-Host "Done"
Write-Host ""