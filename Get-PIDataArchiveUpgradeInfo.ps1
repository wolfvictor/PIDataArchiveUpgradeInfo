<#
	.SYNOPSIS
		Gets important data from PI Data Archive, to verify server health before upgrade

	.DESCRIPTION
		The intent of this script is to save the following data from a PI Data Archive into both .TXT, .CSV & HTML files:
            - Error Messages in log
            - Critical Messages in log
            - Tuning Parameters
            - Archives List
            - Network Manager Statistics
            - PI Services
            - Windows System Info
            - OSIsoft installed products
            - Bad PI Points
            - Stale PI Points
            - AFLink Health Status
            - Collective Members (if it is a collective)
            - License Info
            - Backup Info

        This script works both remotely and local on the PI Data Archive.
        Performance is much better when running locally on the PI Data Archive Server.
        A folder with CSV, HTML and TXT files is saved to current user Desktop folder.

	.NOTES

		Last Modified: 06-Aug-2019

# ************************************************************************************************************************
# * Example created by Victor Wolf (vwolf@osisoft.com)                                                                   *
# * This is just an example. It is not an official script. OSIsoft does not provide ANY specific support on this script. *
# ************************************************************************************************************************

#>


#######################
###### VARIABLES ######
#######################

$script:HTMLStart = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">'

$script:HTMLHead = "<head>
<style>

        body {
            background: #F3F3F4;
            color: #1E1E1F;
            font-family: ""Segoe UI"", Tahoma, Geneva, Verdana, sans-serif;
            padding: 0;
            margin: 0;
        }

        /* For the report title */
        h1 {
            padding: 5px 0px 5px 5px;
            font-size: 21pt;
            background-color: #E2E2E2;
            border-bottom: 1px #C1C1C2 solid;
            color: #201F20;
            margin: 0;
            font-weight: normal;
        }

        /* For Summary and N/NT sections */
        h2 {
            font-size: 18pt;
            font-weight: normal;
            padding: 15px 0 5px 0;
            margin: 0 0 15px 10px;
            color: #1382CE;
        }

        /* For sub-sections for success and failure */
        h3 {
            font-weight: normal;
            font-size: 15pt;
            margin: 0;
            padding: 0 0 10px 50px;
            background-color: transparent;
        }

        .green {
            color: green;
        }

        .yellow {
            color: gold;
        }

        .red {
            color: red;
        }

        table {
            width: 1%;
            border-collapse: collapse;
            white-space: nowrap;
			margin: 15px 0 0 20px;
        }

        th, td {
            padding: 3px;
            border: 1px solid #ddd;
            word-wrap: break-word;
        }

        #summary th, #summary td {
            border: 0;
        }

        th {
            background-color: #007DC3;
            color: white;
            text-align: left;
        }

        td.headerCell {
            border: none;
        }

        a {
            color: #1382CE
        }

        ol {
            padding: 0;
            margin:0;
            word-wrap: break-word;
        }

        ol li {
            padding: 0;
            margin-left: 35px;
            word-wrap: break-word;
        }

.img-common {
    background-repeat: no-repeat;
    height:16px;
    width:16px;
    background-position:center;
}

.img-minus {
    background-image:url(""data:image/gif;base64,R0lGODlhJwAoAPcAAP8A/3t7e4SEhP///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP///yH5BAEAAAAALAAAAAAnACgAAAhsAAEIHEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHJkwgACTKE96FDCgpcsBAjwGeOkywEqaLWN2nEmgZ88BNjuyxKmT40ycQTkOpVl0Y8qnJKNKnUq1qtWrWLNq3cq1a8GAADs="");
}

.img-plus {
    background-image:url(""data:image/gif;base64,R0lGODlhJwAoAPcAAP8A/3t7e4SEhP///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP///yH5BAEAAAAALAAAAAAnACgAAAhwAAEIHEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHJkwgACTKE96FDCgpcsBAjwGcEnAZYCVNF3G7DiTgE+fA252ZNmyZsudHGcWtYnz5VGZJ6OiJEm1qtWrWLNq3cq1q9evYAsGBAA7"");
}
</style>
<script type=""text/javascript"" language=""javascript"">
function toggleDisplayById(expandId, contentId) {
    var exItem = document.getElementById(expandId);
    var isVisible;
	if(!exItem) {
	    return;
	}

    isVisible = exItem.style.display == '';
    exItem.style.display = isVisible ? 'none' : '';

    var contentItem = document.getElementById(contentId);
    if (!contentItem) {
         return;
    }

    if (isVisible) {
        // if it was visible, it is now collapsed
        contentItem.className = ""img-common img-plus"";
    } else {
        contentItem.className = ""img-common img-minus"";
    }
}
</script>
</head><body>
<H1>PI Data Archive Information</H1>"

$script:HTMLEnd = "</body></html>"

#######################
###### FUNCTIONS ######
#######################


# Function ExportToFiles exports gathered data to csv and txt files

Function ExportToFiles($FileName, $FileINFO, $DisableOutputToCSV) {
    $DateString = Get-Date -Format "yyyy_MM_dd_HH_mm"

    $OutputFilePath = $script:OutputFolderTXT + $DateString + $FileName + ".txt"

    $FileINFO | Format-Table -Property * -AutoSize ` | Out-String -Width 4096 ` | Out-File $OutputFilePath

    if(!$DisableOutputToCSV) {
        $OutputFilePath = $script:OutputFolderCSV + $DateString + $FileName + ".csv"
        $FileINFO | Export-Csv $OutputFilePath -NoTypeInformation -ErrorAction SilentlyContinue
    }
}


# Function ExportTo-HTMLStart exports gathered data to csv and txt files in a new way

Function ExportTo-HTMLStart() {
    $DateString = Get-Date -Format "yyyy_MM_dd_HH_mm"
    $script:OutputFilePathHTML = $script:OutputFolderHTML + $DateString + "PIDataArchiveBasicInfo" + ".html"
    $script:HTMLStart + $script:HTMLHead | Out-File $script:OutputFilePathHTML
}


# Function ExportTo-HTMLAppend exports gathered data to csv and txt files in a new way

Function ExportTo-HTMLAppend($FileINFO) {
    $FileINFO | Out-File $script:OutputFilePathHTML -Append
}


# Function ExportTo-HTMLEnd exports gathered data to csv and txt files in a new way

Function ExportTo-HTMLEnd() {
    $script:HTMLEnd = "</body></html>" | Out-File $script:OutputFilePathHTML -Append
    Invoke-Expression $script:OutputFilePathHTML
}


# Function Create-HTMLH2 exports data to a Header 2

Function Create-HTMLH2($ID, $Name) {
    $H2 = "<H2><a name=""ntfTemplates"" onclick=""toggleDisplayById('$ID', 'notificationTemplatesTableExpander')""><span id=""notificationTemplatesTableExpander"" class=""img-common img-minus"" style=""display:inline-block""></span></a>$Name</H2>"
    return $H2
}


# Function Create-HTMLH3 exports data to a Header 3

Function Create-HTMLH3($ID, $Status) {
    $H3 = "<H3 id=""$ID"">$Status</H3>"
    return $H3
}


# Function NotifyUser updates progress bar

Function NotifyUser($DisplayMessage, $PercentComplete) {
    Write-Progress -Activity "Getting $DisplayMessage..." -PercentComplete $PercentComplete -CurrentOperation "$PercentComplete% complete" -Status "Please wait"
}


# Function VerifyLocalAdminRights verifies if current powershell session has local admin rights

Function VerifyLocalAdminRights {
    if(!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {

        $title = "No local admin rights"
        $message = "You don't have local admin rights. Are you sure you want to proceed? Subject to errors."

        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
            "Proceed."

        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
            "End script."

        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

        $result = $host.ui.PromptForChoice($title, $message, $options, 0)

        switch ($result) {
                0 {"You selected Yes."}
                1 {"You selected No."}
        }

        if($result) {
            break
        }
    }
}


# Function ConnectToPIDataArchive verifies connects to PI Data Archive, if possible

Function ConnectToPIDataArchive {
    $script:Server = Read-Host -Prompt 'Input your server name (FQDN, Hostname or IP Address)'

    if(!$script:Server) {
        "Using default: localhost"
        $script:Server = "localhost"
    }

    $Error=@()

    $script:con = Connect-PIDataArchive $script:Server -ErrorAction SilentlyContinue -ErrorVariable Error

    if(!$script:con) {
        Write-Host "Could not connect to $script:Server PI Data Archive. $Error" -ForegroundColor Yellow -BackgroundColor Black
        break
    }
}


# Function VerifyGet-ServiceRights verifies if the current powershell session has access to the Windows Services in the Server machine

Function VerifyGet-ServiceRights {
    try {
        $script:PIServices = Get-Service -ComputerName $script:Server -DisplayName pi*
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Yellow -BackgroundColor Black
        break
    }
}


# Function DefineParameters defines some parameters that will be used throughout the script

function DefineParameters {
    $NumberOfDays = 1

    if(!$script:StartTime) {
        $script:StartTime = (Get-Date).AddDays(-$NumberOfDays)
    }

    if(!$script:EndTime) {
        $script:EndTime = Get-Date
    }

    $OutputFolder = $env:USERPROFILE + "\Desktop\" + $script:con.CurrentRole.Name + "_Upgrade_INFO\"

    if(!(Test-Path $OutputFolder)) {
        New-Item $OutputFolder -type directory | Out-Null
    }

    $script:OutputFolderTXT = $OutputFolder + "TXT\"

    if(!(Test-Path $script:OutputFolderTXT)) {
        New-Item $script:OutputFolderTXT -type directory | Out-Null
    }

    $script:OutputFolderCSV = $OutputFolder + "CSV\"

    if(!(Test-Path $script:OutputFolderCSV)) {
        New-Item $script:OutputFolderCSV -type directory | Out-Null
    }

    $script:OutputFolderHTML = $OutputFolder + "HTML\"

    if(!(Test-Path $script:OutputFolderHTML)) {
        New-Item $script:OutputFolderHTML -type directory | Out-Null
    }
}


# Function ExportMessageLog exports messages in log with Severity equal or greater than $Severity

Function ExportMessageLog ($Severity) {
    $ExportString = Get-PIMessage -Connection $script:con -Starttime $script:StartTime -Endtime $script:EndTime -SeverityType $Severity | Select-Object Severity,TimeStamp,ProgramName,Message,ID,Category,Source1,Source2,Source3,ProcessPIUser,ProcessOSUser,ProcessHost,Priority,ProcessID,OriginatingPIUser,OriginatingOSUser,OriginatingHost,Parameters
    ExportToFiles "$($Severity)pigetmsg" $ExportString
}


# Function ExportTuningParameters exports Tuning Parameters

Function ExportTuningParameters {
    $ExportString = Get-PITuningParameter -Connection $script:con | Select-Object Name,Value,Min,Max,Default,Units,TakesEffect,Description
    ExportToFiles "TuningParameters" $ExportString
}


# Function ExportArchivesList exports Archives List

Function ExportArchivesList {
    $script:HTMLMiddleArchiveInfo = Get-PIArchiveFileInfo -Connection $script:con
    $ExportString = $script:HTMLMiddleArchiveInfo | Select-Object Index,StartTime,EndTime,RecordSize,PercentFull,IsCorrupt,Path,LastModifiedTime,LastBackupTime,ArchiveSet,TotalEvents,PrimaryIndexCount,OverflowIndexCount,OverlappingPrimaryMax,AverageEventsPerRecordCount,AddRatePerHour,AnnotationMax,AnnotationFileSize,AnnotationsUsed,AnnotationUid,Version,Type,State,IsWriteable,IsShiftable,RecordCount,FreePrimaryRecords,FreeOverflowRecords,MaxPrimaryRecords,MaxOverflowRecords
    ExportToFiles "ArchivesList" $ExportString
}


# Function ExportNetworkManagerStatistic formats and exports Network Manager Statistics

Function ExportNetworkManagerStatistics {
    $NetStats = Get-PIConnectionStatistics -Connection $script:con

    $CSVObject = @()

    # Formatting of Network Manager Statistic for csv export

    $NetworkObjectcounter = 0

    foreach($script:connection in $NetStats) {
        $AuxObject = New-Object -TypeName PSObject

        $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value $script:connection.Name
        $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $script:connection.Value
        $AuxObject | Add-Member -MemberType NoteProperty -Name StatisticType -Value $script:connection.StatisticType

        $CSVObject += $AuxObject

        $NetworkObjectcounter++
    }


    $NetworkObject = @()

    for($i = 0; $i -lt $NetworkObjectcounter; $i++) {
        $AuxObject = New-Object -TypeName PSObject

        for($j = 0; $j -le $CSVObject[$j].Name.Count; $j++) {
            if($CSVObject[$i].Name[$j]) {
                $AuxObject | Add-Member -MemberType NoteProperty -Name $CSVObject[$i].Name[$j] -Value $CSVObject[$i].Value[$j]
            }
        }

        $NetworkObject += $AuxObject
    }

    $ExportString = $NetworkObject | Select-Object ID,PIPath,Name,PID,RegAppName,RegAppType,RegAppID,ProtocolVersion,PeerAddress,PeerPort,ConType,NetType,ConStatus,ConTime,LastCall,ElapsedTime,BytesSent,BytesRecv,MsgSent,MsgRecv,RecvErrors,SendErrors,APICount,SDKCount,ServerID,PIVersion,OSSysName,OSVersion,OSBuild,User,OSUser,Trust,NumConnections,IsTCPListenerOpen,IsStandAlone

    ExportToFiles "NetworkStatistics" $ExportString
}


# Function ExportPIServices formats and exports all PI Services on the machine


Function ExportPIServices {
    $CSVObject = @()

    # Formatting of PI Services for csv export

    for($i = 0; $i -lt $script:PIServices.DependentServices.Count; $i++) {
        $AuxObject = New-Object -TypeName PSObject
        $RequiredServices = ""
        $DependentServices = ""
        $ServicesDependedOn = ""

        for($j = 0; $j -lt $script:PIServices[$i].DependentServices.ServiceName.Count; $j++) {
            if($j -lt $script:PIServices[$i].DependentServices.ServiceName.Count - 1){
                $DependentServices += "$($script:PIServices[$i].DependentServices[$j].ServiceName); "
            }
            else {
                $DependentServices += $script:PIServices[$i].DependentServices[$j].ServiceName
            }
        }

        for($j = 0; $j -lt $script:PIServices[$i].RequiredServices.ServiceName.Count; $j++) {
            if($j -lt $script:PIServices[$i].RequiredServices.ServiceName.Count - 1){
                $RequiredServices += "$($script:PIServices[$i].RequiredServices[$j].ServiceName); "
            }
            else {
                $RequiredServices += $script:PIServices[$i].RequiredServices[$j].ServiceName
            }
        }

        for($j = 0; $j -lt $script:PIServices[$i].ServicesDependedOn.ServiceName.Count; $j++) {
            if($j -lt $script:PIServices[$i].ServicesDependedOn.ServiceName.Count - 1){
                $ServicesDependedOn += "$($script:PIServices[$i].ServicesDependedOn[$j].ServiceName); "
            }
            else {
                $ServicesDependedOn += $script:PIServices[$i].ServicesDependedOn[$j].ServiceName
            }
        }

        $AuxObject | Add-Member -MemberType NoteProperty -Name DisplayName -Value $script:PIServices[$i].DisplayName
        $AuxObject | Add-Member -MemberType NoteProperty -Name ServiceName -Value $script:PIServices[$i].ServiceName
        $AuxObject | Add-Member -MemberType NoteProperty -Name Status -Value $script:PIServices[$i].Status
        $AuxObject | Add-Member -MemberType NoteProperty -Name RequiredServices -Value $RequiredServices
        $AuxObject | Add-Member -MemberType NoteProperty -Name DependentServices -Value $DependentServices
        $AuxObject | Add-Member -MemberType NoteProperty -Name ServicesDependedOn -Value $ServicesDependedOn
        $AuxObject | Add-Member -MemberType NoteProperty -Name CanPauseAndContinue -Value $script:PIServices[$i].CanPauseAndContinue
        $AuxObject | Add-Member -MemberType NoteProperty -Name CanShutdown -Value $script:PIServices[$i].CanShutdown
        $AuxObject | Add-Member -MemberType NoteProperty -Name CanStop -Value $script:PIServices[$i].CanStop
        $AuxObject | Add-Member -MemberType NoteProperty -Name MachineName -Value $script:PIServices[$i].MachineName
        $AuxObject | Add-Member -MemberType NoteProperty -Name ServiceHandle -Value $script:PIServices[$i].ServiceHandle
        $AuxObject | Add-Member -MemberType NoteProperty -Name ServiceType -Value $script:PIServices[$i].ServiceType
        $AuxObject | Add-Member -MemberType NoteProperty -Name Site -Value $script:PIServices[$i].Site
        $AuxObject | Add-Member -MemberType NoteProperty -Name Container -Value $script:PIServices[$i].Container
        $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value $script:PIServices[$i].Name

        $CSVObject += $AuxObject
    }

    ExportToFiles "PIServices" $CSVObject
}


# Function ExportOSIProductsList exports a list of OSIsoft Products installed on the machine

Function ExportOSIProductsList {
    $ExportString = Get-WmiObject -Class Win32_Product -ComputerName $script:Server | Where-Object -FilterScript {$_.Vendor -like 'OSI*'} | Select-Object Name,Vendor,Version | Sort-Object Name
    ExportToFiles "PIPrograms" $ExportString
}


# Function ExportPIPointsInBadStatus exports a list of PI Points in Bad Status

Function ExportPIPointsInBadStatus {
    $ExportString = ""

    $AllTags = (Get-PIPoint -Connection $script:con -Name * -Attributes tag).Point

    $CSVObject = @()

    foreach($tag in $AllTags) {
        $PointID = $false
        $PointID = Get-PIValue -PointName $tag.Name -Connection $script:con -Time $script:EndTime | Where-Object {$_.IsGood -eq $false} | Select-Object StreamId
        if($PointID) {
            $AuxObject = New-Object -TypeName PSObject

            $AuxObject | Add-Member -MemberType NoteProperty -Name BadTagName -Value $tag.Name

            $CSVObject += $AuxObject
        }
    }

    ExportToFiles "BadPIPoints" $CSVObject
}


# Function ExportStalePIPoints exports a list of Stale PI Points

Function ExportStalePIPoints {
    $ExportString = ""

    $AllTags = (Get-PIPoint -Connection $script:con -Name * -Attributes tag).Point

    $CSVObject = @()

    foreach($tag in $AllTags) {
        if(!$tag.IsFuture) {
            $LastTimestamp = (Get-PIValue -PointName $tag.Name -Connection $script:con -Time $script:EndTime -ArchiveMode Previous).Timestamp
            if($LastTimestamp -le (Get-Date).AddHours(-4)){
                $ExportString += $tag.Name + "`r`n"

                $AuxObject = New-Object -TypeName PSObject

                $AuxObject | Add-Member -MemberType NoteProperty -Name StaleTagName -Value $tag.Name

                $CSVObject += $AuxObject
            }
        }
    }

    ExportToFiles "StalePIPoints" $CSVObject
}


# Function ExportSystemInfo exports the result of systeminfo cmd command to a txt file

Function ExportSystemInfo {
    $sys = systeminfo /S $script:Server
    ExportToFiles "SystemInfo" $sys 1
}


# Function GetDataArchiveGeneralInfoHTML exports basic PI System info related but not limited to license, backup and archives.

Function GetDataArchiveGeneralInfoHTML {

    ExportTo-HTMLStart

    $ExportString = Create-HTMLH2 "AFLink" "AFLink Health Status"

    ExportTo-HTMLAppend $ExportString

    $Error = $null

    $AFLinkStatus = Get-PIAFLinkModuleDatabaseStatistics -Connection $script:con -ErrorAction SilentlyContinue -ErrorVariable Error

    if ($Error) {
        $ExportString = Create-HTMLH3 "AFLink" $Error
    }

    else {
        $ExportString = Create-HTMLH3 "AFLink" $AFLinkStatus.AFLink.HealthStatus.Value
    }

    ExportTo-HTMLAppend $ExportString


    $ExportString = Create-HTMLH2 "CollectiveMembers" "Collective Members"

    ExportTo-HTMLAppend $ExportString "CollectiveMembers"

    if($script:con.CurrentRole.Type -eq "UnSpecified") {
        $ExportString = Create-HTMLH3 "CollectiveMembers" "Standalone Server: $($script:con.Configuration.Name)"
    }
    else {        
        $collective = Get-PICollective -Connection $script:con
        $CollectiveMembers = $collective.Members.Name -join "; "
        $ExportString = Create-HTMLH3 "CollectiveMembers" $CollectiveMembers
    }

    ExportTo-HTMLAppend $ExportString


    $ExportString = Create-HTMLH2 "License" "License Info"

    ExportTo-HTMLAppend $ExportString "License"

    $lic = Get-PILicenseReport -Connection $script:con
    $licEntry = Get-PILicenseEntry -Connection $script:con
    $maxpointcount = $LicEntry | Where-Object {$_.Name -eq "pibasess.maxpointcount"}
    $MaxAggregatePointModuleCount = $LicEntry | Where-Object {$_.Name -eq "pibasess.MaxAggregatePointModuleCount"}

    $CSVObject = @()

    $AuxObject = New-Object -TypeName PSObject
    $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value "Installation Id"
    $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $lic.FileInfo.InstallationId
    $CSVObject += $AuxObject

    $AuxObject = New-Object -TypeName PSObject
    $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value "Expiration Time"
    $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $lic.FileInfo.ExpirationTime
    $CSVObject += $AuxObject

    $AuxObject = New-Object -TypeName PSObject
    $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value "Required Percent Match"
    $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $lic.FileInfo.RequiredPercentMatch
    $CSVObject += $AuxObject

    $AuxObject = New-Object -TypeName PSObject
    $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value "Current Percent Match"
    $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $lic.Usage.CurrentPercentMatch
    $CSVObject += $AuxObject

    $AuxObject = New-Object -TypeName PSObject
    $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value "Amount of tags used"
    $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $maxpointcount.AmountUsed
    $CSVObject += $AuxObject

    $AuxObject = New-Object -TypeName PSObject
    $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value "Licensed amount of tags"
    $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $maxpointcount.TotalAmount
    $CSVObject += $AuxObject

    $AuxObject = New-Object -TypeName PSObject
    $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value "Amount of tags used (Aggregate Point Module)"
    $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $MaxAggregatePointModuleCount.AmountUsed
    $CSVObject += $AuxObject

    $AuxObject = New-Object -TypeName PSObject
    $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value "Licensed amount of tags (Aggregate Point Module)"
    $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $MaxAggregatePointModuleCount.TotalAmount
    $CSVObject += $AuxObject

    $ExportString = $CSVObject | ConvertTo-Html -Fragment

    $ExportString = $ExportString -replace '<table>', '<table id="License">'
    $ExportString = $ExportString -replace "<tr><th>\D*</th></tr>", ""

    ExportTo-HTMLAppend $ExportString



    if(!$script:HTMLMiddleArchiveInfo) {
        $script:HTMLMiddleArchiveInfo = Get-PIArchiveFileInfo -Connection $script:con
    }

    $ArchiveGapsFlag = $false

    

    $ExportString = Create-HTMLH2 "Archive" "Archive Info"

    ExportTo-HTMLAppend $ExportString "Archive"

    $CSVObject = @()
    $AuxObject = New-Object -TypeName PSObject

    $ArchiveInfo = $script:HTMLMiddleArchiveInfo | Where {$_.EndTime -ge 1} | Select-Object Index,StartTime,EndTime,RecordSize,PercentFull,IsCorrupt,Path,LastModifiedTime,LastBackupTime,ArchiveSet,TotalEvents,PrimaryIndexCount,OverflowIndexCount,OverlappingPrimaryMax,AverageEventsPerRecordCount,AddRatePerHour,AnnotationMax,AnnotationFileSize,AnnotationsUsed,AnnotationUid,Version,Type,State,IsWriteable,IsShiftable,RecordCount,FreePrimaryRecords,FreeOverflowRecords,MaxPrimaryRecords,MaxOverflowRecords

    for($i = 0; $i -lt $ArchiveInfo.EndTime.Count - 1; $i++) {
        if($ArchiveInfo.EndTime[$i]) {
            if($ArchiveInfo.EndTime[$i+1] -ne $ArchiveInfo.StartTime[$i]) {
                $AuxObject = New-Object -TypeName PSObject
                $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value "There is an archive gap from $($ArchiveInfo.EndTime[$i+1]) to $($ArchiveInfo.StartTime[$i]) UTC"
                $ArchiveGapsFlag = $true
                $CSVObject += $AuxObject
            }
        }
    }

    if(!$ArchiveGapsFlag) {
        $ExportString = Create-HTMLH3 "Archive" "There are no Archive Gaps"
    }

    else {
        $ExportString = $CSVObject | ConvertTo-Html -Fragment
        $ExportString = $ExportString -replace '<table>', '<table id="Archive">'
        $ExportString = $ExportString -replace "<tr><th>\D*</th></tr>", ""
    }    

    ExportTo-HTMLAppend $ExportString



    $ExportString = Create-HTMLH2 "Backup" "Backup Info"

    ExportTo-HTMLAppend $ExportString "Backup"

    $BackupInfo = Get-PIBackupSummary -Connection $script:con | Sort-Object index -Descending | Select-Object -First 1

    $CSVObject = @()
    $AuxObject = New-Object -TypeName PSObject

    if(!$BackupInfo) {
        "It was not possible to retrieve backup info. Please check manually"
    }

    else {
        $AuxObject = New-Object -TypeName PSObject
        $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value "Last Backup Start"
        $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $BackupInfo.BackupStart
        $CSVObject += $AuxObject

        $AuxObject = New-Object -TypeName PSObject
        $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value "Last Backup End"
        $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $BackupInfo.BackupEnd
        $CSVObject += $AuxObject

        $AuxObject = New-Object -TypeName PSObject
        $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value "Last Backup Status Message"
        $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $BackupInfo.StatusMessage
        $CSVObject += $AuxObject

        $AuxObject = New-Object -TypeName PSObject
        $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value "Last Backup Type"
        $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $BackupInfo.Type
        $CSVObject += $AuxObject
    }

    $ExportString = $CSVObject | ConvertTo-Html -Fragment

    $ExportString = $ExportString -replace '<table>', '<table id="Backup">'
    $ExportString = $ExportString -replace "<tr><th>\D*</th></tr>", ""

    ExportTo-HTMLAppend $ExportString

    ExportTo-HTMLEnd
}


#######################
######### MAIN ########
#######################


# Initialization

Write-Progress -Activity "Initializing..." -PercentComplete 0 -CurrentOperation "0% complete" -Status "Please wait"

"`r`nThis script is considerably faster to run locally on the PI Data Archive Server"

VerifyLocalAdminRights

ConnectToPIDataArchive

VerifyGet-ServiceRights

DefineParameters


# Gather and Export Data

NotifyUser "Error messages in log" 8

ExportMessageLog "Error"

NotifyUser "Critical messages in log" 16

ExportMessageLog "Critical"

NotifyUser "Tuning Parameters" 24

ExportTuningParameters

NotifyUser "Archives List" 32

ExportArchivesList

NotifyUser "Network Statistics" 40

ExportNetworkManagerStatistics

NotifyUser "System Info" 48

ExportSystemInfo

NotifyUser "Basic PI Server Info" 56

GetDataArchiveGeneralInfoHTML

NotifyUser "PI Services" 64

ExportPIServices

NotifyUser "PI Programs" 72

ExportOSIProductsList

NotifyUser "Bad PI Points" 80

ExportPIPointsInBadStatus

NotifyUser "Stale PI Points" 88

ExportStalePIPoints


# Complete

Write-Progress -Activity "Complete" -PercentComplete 100 -Status "Very Nice! Great Success"

Start-Sleep 3