<#
.NOTES
    Author: Ben Owens
    LinkedIn: https://www.linkedin.com/in/owensben/
    Creation Date: 04/07/2017
    Purpose: This script will query all the mailboxes in a migration batch 
    and detail which of the reported corruptions are genuine and which 
    corruptions related to ACL security principals that don't exist in the source or target forests.  
    This is to help manage the  change in behaviour as covered in the article 
    https://blogs.technet.microsoft.com/exchange/2017/05/30/toomanybaditemspermanentexception-error-when-migrating-to-exchange-online/?replytocom=310635#respond 
.INPUTS
    You will be prompted to enter a batch name when you ruin the script
    
.OUTPUTS
    Configure the $LogDirectory variable at the top of the script.
    3 files will be output to directory specified with a subfolder with a timestamp prepended.
    1. A summary CSV output including the mailboxes queried along with the details in the table above.
    2. An output of all corrupt bad items found, including ACL security principals, should they need to be queried.
    3. A migrations XML output for each mailbox migration - this can imported into any PowerShell session at a later date and queried; this allows you to remove the migration job but keep a logged record.

#>

$LogDate = get-date -f yyyyMMddhhmmss
$LogDirectory = Split-Path -parent "C:\BenTemp\$LogDate\*.*"
$Results = @()

$BatchID = Read-Host -Prompt 'Enter the batch name here'
Mkdir $LogDirectory

$MigrationUsers = Get-MigrationUser -BatchID $BatchID

ForEach ($User in $MigrationUsers) {
    $Statistics = Get-MoveRequestStatistics -Identity $User.Identity -IncludeReport
    #$MoveStatistics = Get-MoveRequestStatistics -Identity $User.Identity
    $MoveStatus = $Statistics | Select -ExpandProperty Status | Select -ExpandProperty Value
    $MoveStatusDetail = $Statistics | Select -ExpandProperty StatusDetail | Select -ExpandProperty Value
    $Identity = Get-User $User.Identity | Select -ExpandProperty Name
    $XMLPath = $logdirectory + "\" + $Identity + "_MigrationReport.xml"
    $Statistics | Export-CliXml $XMLPath
    $ReportedFailures = $Statistics.Report.BadItems
    $CSVPath = $logdirectory + "\" + $Identity + "_CorruptItemsOverview.csv"
    If ($ReportedFailures -eq $NULL) {
        Write-Host "No bad items to report for" $Identity
    }
    Else {
        $ReportedFailures | Select Kind,FolderName,WellKnownFolderType,Subject,Failure,Category | Export-CSV $CSVPath -NoTypeInformation
    }
    $ReportedSourceACLFailures = $ReportedFailures | Where {$_.Category -like "SourcePrincipalError"}
    $ReportedTargetACLFailures = $ReportedFailures | Where {$_.Category -like "TargetPrincipalError"}
    $FilteredReportedFailures = $ReportedFailures | Where {$_.Category -notlike "SourcePrincipalError"}
    $FilteredReportedFailures = $FilteredReportedFailures | Where {$_.Category -notlike "TargetPrincipalError"}
    
    If ($ReportedFailures -eq $NULL) {
        $Result = new-object PSObject -Property @{
                        #Identity = $MigrationUsers.Identity;
                        Identity = $Identity;
                        MoveStatus = $MoveStatus;
                        MoveStatusDetail = $MoveStatusDetail;
                        Corruptions = $ReportedFailures.count
                        GenuineCorruptions = $FilteredReportedFailures.count
                        SourcePrincipalErrors = $ReportedSourceACLFailures.count
                        TargetPrincipalErrors = $ReportedTargetACLFailures.count
                        Comment = "No corruptions to investigate"
                    }
                    $Results += $Result
    }

    ElseIf ($FilteredReportedFailures -eq $NULL) {
        $Result = new-object PSObject -Property @{
                        #Identity = $MigrationUsers.Identity;
                        Identity = $Identity;
                        MoveStatus = $MoveStatus;
                        MoveStatusDetail = $MoveStatusDetail
                        Corruptions = $ReportedFailures.count
                        GenuineCorruptions = $FilteredReportedFailures.count
                        SourcePrincipalErrors = $ReportedSourceACLFailures.count
                        TargetPrincipalErrors = $ReportedTargetACLFailures.count
                        Comment = "Only ACL issues - no need to investigate"
                    }
                    $Results += $Result
    }
    Else {
        $Result = new-object PSObject -Property @{
                        #Identity = $MigrationUsers.Identity;
                        Identity = $Identity;
                        MoveStatus = $MoveStatus;
                        MoveStatusDetail = $MoveStatusDetail;
                        Corruptions = $ReportedFailures.count
                        GenuineCorruptions = $FilteredReportedFailures.count
                        SourcePrincipalErrors = $ReportedSourceACLFailures.count
                        TargetPrincipalErrors = $ReportedTargetACLFailures.count
                        Comment = "Investigation required!"
                    }
                    $Results += $Result
    }
}

$Results | ft Identity, MoveStatus, MoveStatusDetail,Corruptions,SourcePrincipalErrors,TargetPrincipalErrors,GenuineCorruptions,Comment
$ResultsPath = $logdirectory + "\" + "_ReportSummary.csv"
$Results | Select Identity, MoveStatus, MoveStatusDetail,Corruptions,SourcePrincipalErrors,TargetPrincipalErrors,GenuineCorruptions,Comment | Export-CSV $ResultsPath -NoTypeInformation

Write-Host "Go to" $LogDirectory "for output log files" -ForegroundColor Yellow
