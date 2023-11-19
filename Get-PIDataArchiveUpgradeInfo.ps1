Function Get-PIDataArchiveUpgradeInfo {

    [cmdletbinding()]

    param(
        [Int32]$NumberOfDays
    )

    Function ExportToFiles($FileName, $FileINFO) {

        $DateString = Get-Date -Format "yyyy_MM_dd_HH_mm"

        $OutputFilePath = $OutputFolderTXT + $DateString + $FileName + ".txt"

        $FileINFO | Format-Table -Property * -AutoSize ` | Out-String -Width 4096 ` | Out-File $OutputFilePath

        $OutputFilePath = $OutputFolderCSV + $DateString + $FileName + ".csv"

        $FileINFO | Export-Csv $OutputFilePath -NoTypeInformation -ErrorAction SilentlyContinue

    }

    Function NotifyUser($DisplayMessage, $PercentComplete) {

        Write-Progress -Activity "Getting $DisplayMessage..." -PercentComplete $PercentComplete -CurrentOperation "$PercentComplete% complete" -Status "Please wait"
    }

    # PI Data Archive Connection

    NotifyUser "Initializing..." 0

    $con = Connect-PIDataArchive localhost -ErrorAction SilentlyContinue

    if(!$con) {
        "Could not connect to localhost PI Data Archive"
        break
    }

    # Parameter Initialization

    if(!$NumberOfDays) {
        $NumberOfDays = 1
    }

    if(!$StartTime) {
        $StartTime = (Get-Date).AddDays(-$NumberOfDays)
    }

    if(!$EndTime) {
        $EndTime = Get-Date
    }

    $OutputFolder = $env:USERPROFILE + "\Desktop\" + $env:computername + "_Upgrade_INFO\"

    if(!(Test-Path $OutputFolder)) {
        New-Item $OutputFolder -type directory | Out-Null
    }

    $OutputFolderTXT = $OutputFolder + "TXT\"

    if(!(Test-Path $OutputFolderTXT)) {
        New-Item $OutputFolderTXT -type directory | Out-Null
    }

    $OutputFolderCSV = $OutputFolder + "CSV\"

    if(!(Test-Path $OutputFolderCSV)) {
        New-Item $OutputFolderCSV -type directory | Out-Null
    }


    # Get Error PI Messages

    NotifyUser "Error messages in log" 10

    $ExportString = Get-PIMessage -Connection $con -Starttime $StartTime -Endtime $EndTime -SeverityType Error | Select-Object Severity,TimeStamp,ProgramName,Message,ID,Category,Source1,Source2,Source3,ProcessPIUser,ProcessOSUser,ProcessHost,Priority,ProcessID,OriginatingPIUser,OriginatingOSUser,OriginatingHost,Parameters

    ExportToFiles "Errorpigetmsg" $ExportString


    # Get Critical PI Messages

    NotifyUser "Critical messages in log" 20

    $ExportString = Get-PIMessage -Connection $con -Starttime $StartTime -Endtime $EndTime -SeverityType Critical | Select-Object Severity,TimeStamp,ProgramName,Message,ID,Category,Source1,Source2,Source3,ProcessPIUser,ProcessOSUser,ProcessHost,Priority,ProcessID,OriginatingPIUser,OriginatingOSUser,OriginatingHost,Parameters

    ExportToFiles "Critpigetmsg" $ExportString


    # Get Tuning Parameters

    NotifyUser "Tuning Parameters" 30

    $ExportString = Get-PITuningParameter -Connection $con | Select-Object Name,Value,Min,Max,Default,Units,TakesEffect,Description
    
    ExportToFiles "TuningParameters" $ExportString


    # Get Archives List

    NotifyUser "Archives List" 40

    $ExportString = Get-PIArchiveFileInfo -Connection $con | Select-Object Index,StartTime,EndTime,RecordSize,PercentFull,IsCorrupt,Path,LastModifiedTime,LastBackupTime,ArchiveSet,TotalEvents,PrimaryIndexCount,OverflowIndexCount,OverlappingPrimaryMax,AverageEventsPerRecordCount,AddRatePerHour,AnnotationMax,AnnotationFileSize,AnnotationsUsed,AnnotationUid,Version,Type,State,IsWriteable,IsShiftable,RecordCount,FreePrimaryRecords,FreeOverflowRecords,MaxPrimaryRecords,MaxOverflowRecords

    ExportToFiles "ArchivesList" $ExportString


    # Get Network Statistics

    NotifyUser "Network Statistics" 50

    $NetStats = Get-PIConnectionStatistics -Connection $con

    $CSVObject = @()

    foreach($connection in $NetStats) {
        $AuxObject = New-Object -TypeName PSObject

        $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value $connection.Name
        $AuxObject | Add-Member -MemberType NoteProperty -Name Value -Value $connection.Value
        $AuxObject | Add-Member -MemberType NoteProperty -Name StatisticType -Value $connection.StatisticType

        $CSVObject += $AuxObject
    }

    $NetworkObject = @()

    for($i = 0; $i -lt $CSVObject[$i].Name.Count; $i++) {
        
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


    #Get PI Services

    NotifyUser "PI Services" 60

    $CSVObject = @()

    $PIServices = Get-Service -DisplayName pi*

    for($i = 0; $i -lt $PIServices.DependentServices.Count; $i++) {
        $AuxObject = New-Object -TypeName PSObject
        $RequiredServices = ""
        $DependentServices = ""
        $ServicesDependedOn = ""

        for($j = 0; $j -lt $PIServices[$i].DependentServices.ServiceName.Count; $j++) {
            if($j -lt $PIServices[$i].DependentServices.ServiceName.Count - 1){
                $DependentServices += "$($PIServices[$i].DependentServices[$j].ServiceName); "
            }
            else {
                $DependentServices += $PIServices[$i].DependentServices[$j].ServiceName
            }
        }

        for($j = 0; $j -lt $PIServices[$i].RequiredServices.ServiceName.Count; $j++) {
            if($j -lt $PIServices[$i].RequiredServices.ServiceName.Count - 1){
                $RequiredServices += "$($PIServices[$i].RequiredServices[$j].ServiceName); "
            }
            else {
                $RequiredServices += $PIServices[$i].RequiredServices[$j].ServiceName
            }
        }

        for($j = 0; $j -lt $PIServices[$i].ServicesDependedOn.ServiceName.Count; $j++) {
            if($j -lt $PIServices[$i].ServicesDependedOn.ServiceName.Count - 1){
                $ServicesDependedOn += "$($PIServices[$i].ServicesDependedOn[$j].ServiceName); "
            }
            else {
                $ServicesDependedOn += $PIServices[$i].ServicesDependedOn[$j].ServiceName
            }
        }

        $AuxObject | Add-Member -MemberType NoteProperty -Name DisplayName -Value $PIServices[$i].DisplayName
        $AuxObject | Add-Member -MemberType NoteProperty -Name ServiceName -Value $PIServices[$i].ServiceName
        $AuxObject | Add-Member -MemberType NoteProperty -Name Status -Value $PIServices[$i].Status
        $AuxObject | Add-Member -MemberType NoteProperty -Name RequiredServices -Value $RequiredServices
        $AuxObject | Add-Member -MemberType NoteProperty -Name DependentServices -Value $DependentServices
        $AuxObject | Add-Member -MemberType NoteProperty -Name ServicesDependedOn -Value $ServicesDependedOn
        $AuxObject | Add-Member -MemberType NoteProperty -Name CanPauseAndContinue -Value $PIServices[$i].CanPauseAndContinue
        $AuxObject | Add-Member -MemberType NoteProperty -Name CanShutdown -Value $PIServices[$i].CanShutdown
        $AuxObject | Add-Member -MemberType NoteProperty -Name CanStop -Value $PIServices[$i].CanStop
        $AuxObject | Add-Member -MemberType NoteProperty -Name MachineName -Value $PIServices[$i].MachineName
        $AuxObject | Add-Member -MemberType NoteProperty -Name ServiceHandle -Value $PIServices[$i].ServiceHandle
        $AuxObject | Add-Member -MemberType NoteProperty -Name ServiceType -Value $PIServices[$i].ServiceType
        $AuxObject | Add-Member -MemberType NoteProperty -Name Site -Value $PIServices[$i].Site
        $AuxObject | Add-Member -MemberType NoteProperty -Name Container -Value $PIServices[$i].Container
        $AuxObject | Add-Member -MemberType NoteProperty -Name Name -Value $PIServices[$i].Name
        
        $CSVObject += $AuxObject     
    }

    ExportToFiles "PIServices" $CSVObject


    #Get PI Products

    NotifyUser "PI Programs" 70

    $ExportString = Get-WmiObject -Class Win32_Product | Where-Object -FilterScript {$_.Vendor -like 'OSI*'} | Select-Object Name,Vendor,Version | Sort-Object Name

    ExportToFiles "PIPrograms" $ExportString


    #Get Bad Points

    NotifyUser "Bad PI Points" 80

    $ExportString = ""
    
    $AllTags = (Get-PIPoint -Connection $con -Name * -Attributes tag).Point

    $CSVObject = @()

    foreach($tag in $AllTags) {
        $PointID = $false
        $PointID = Get-PIValue -PointName $tag.Name -Connection $con -Time $EndTime | Where-Object {$_.IsGood -eq $false} | Select-Object StreamId
        if($PointID) {
            $AuxObject = New-Object -TypeName PSObject

            $AuxObject | Add-Member -MemberType NoteProperty -Name BadTagName -Value $tag.Name

            $CSVObject += $AuxObject
        }
    }

    ExportToFiles "BadPIPoints" $CSVObject


    #Get Stale Points

    NotifyUser "Stale PI Points" 90

    $ExportString = ""
    
    $AllTags = (Get-PIPoint -Connection $con -Name * -Attributes tag).Point

    $CSVObject = @()

    foreach($tag in $AllTags) {
        $LastTimestamp = (Get-PIValue -PointName $tag.Name -Connection $con -Time $EndTime -ArchiveMode Previous).Timestamp
        if($LastTimestamp -le (Get-Date).AddHours(-4)){
            $ExportString += $tag.Name + "`r`n"
            
            $AuxObject = New-Object -TypeName PSObject

            $AuxObject | Add-Member -MemberType NoteProperty -Name StaleTagName -Value $tag.Name

            $CSVObject += $AuxObject
        }
    }

    ExportToFiles "StalePIPoints" $CSVObject


    #Complete

    Write-Progress -Activity "Complete" -PercentComplete 100 -Status "Very Nice! Great Success"
}