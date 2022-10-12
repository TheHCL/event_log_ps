﻿$header = @"
<style>
    h1{
        font-family: Calibri, Helvetica,sans-serif;
        color: #e68a00;
        font-size: 28px;
    }

    h2{
        font-family: Calibri, Helvetica,sans-serif;
        color: #000099;
        font-size: 24px;
    }

    table{
        font-size: 14px;
        border: 0px;
        font-family: Calibri, Helvetica,sans-serif;
    }

    td{
        padding: 4px;
        margin: 0px;
        color: #000099;
        border: solid 1px;
    }

</style>
"@


Add-Type -AssemblyName System.Windows.Forms

Function Sel_File
 {
 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.InitialDirectory = “C:\Users\XX”
 $OpenFileDialog.Title = "Please Select File"
 $OpenFileDialog.filter = “All files (*.evtx*)| *.evtx*”
 If ($OpenFileDialog.ShowDialog() -eq "Cancel") 
 {
  [System.Windows.Forms.MessageBox]::Show("No File Selected. Please select a file !", "Error", 0, 
  [System.Windows.Forms.MessageBoxIcon]::Exclamation)
  }   $Global:SelectedFile = $OpenFileDialog.SafeFileName
  return $OpenFileDialog.FileName
    
}

$filename = Sel_File

#$_.Id -eq '1' -or 

if ($filename -ne $null){
    $startTime = (Get-Date)
    

    $basename = ($filename.Split("\"))[-1]
    $report_title = "<h1>Report Source: $basename</h1>"
    $bsod_ev = Get-WinEvent -Path $filename | Where-Object {$_.Id -eq '1001' -and $_.LogName -eq 'System'}
    $whea_ev = Get-WinEvent -Path $filename | Where-Object {$_.Id -eq '17'-and $_.ProviderName -eq 'Microsoft-Windows-WHEA-Logger' -and $_.LogName -eq 'System'}
    $tdr_ev = Get-WinEvent -Path $filename | Where-Object {$_.Id -eq '4101' -and $_.LogName -eq 'System'}
    $unexpected_ev = Get-WinEvent -Path $filename | Where-Object {$_.Id -eq '6008' -and $_.LogName -eq 'System' }
    $app_ev = Get-WinEvent -Path $filename | Where-Object {$_.Id -eq '1002' -and $_.LogName -eq 'Application'}

    #count 
    $bsod_c = $bsod_ev.Count
    $whea_c = $whea_ev.Count
    $tdr_c =  $tdr_ev.Count
    $unexpected_c=$unexpected_ev.Count
    $app_c = $app_ev.Count
    
    $summary = "<h2>Summary</h2>
                <h2>
                <table>
                <tbody>
                <tr>
                <th>Event Type</th>
                <th>Count</th>
                </tr>
                <tr><td>BSOD</td><td>$bsod_c</td></tr>
                <tr><td>WHEA</td><td>$whea_c</td></tr>
                <tr><td>TDR</td><td>$tdr_c</td></tr>
                <tr><td>Unexpected Shutdown</td><td>$unexpected_c</td></tr>
                <tr><td>App Hang</td><td>$app_c</td></tr>
                </tbody>
                </table>
                </h2>
    "

    $bsod = $bsod_ev | ConvertTo-Html -Property TimeCreated,Id,LevelDisplayName,Message -PreContent "<h2>BSOD:$bsod_c times<h2>"
    $whea = $whea_ev | ConvertTo-Html -Property TimeCreated,Id,LevelDisplayName,Message -PreContent "<h2>WHEA:$whea_c times<h2>"
    $tdr = $tdr_ev | ConvertTo-Html -Property TimeCreated,Id,LevelDisplayName,Message -PreContent "<h2>TDR:$tdr_c times<h2>"
    $unexpected = $unexpected_ev | ConvertTo-Html -Property TimeCreated,Id,LevelDisplayName,Message -PreContent "<h2>Unexpected Shutdown:$unexpected_c times<h2>"
    $app = $app_ev | ConvertTo-Html -Property TimeCreated,Id,LevelDisplayName,Message -PreContent "<h2>App Hang:$app_c times<h2>"

    $all_ev = (Get-WinEvent -Path $filename | Where-Object {(($_.Id -eq '4101'-or ($_.Id -eq '17'-and $_.ProviderName -eq 'Microsoft-Windows-WHEA-Logger') -or $_.Id -eq '1001' -or $_.Id -eq '6008') -and $_.LogName -eq 'System') -or $_.Id -eq '1002'} | Sort-Object -Property TimeCreated) | ConvertTo-Html -Property TimeCreated,Id,LevelDisplayName,Message -PreContent "<h2>All Errors</h2>"
    $report = ConvertTo-Html -Body "$report_title $summary $bsod $whea $tdr $unexpected $app $all_ev" -Head $header
    $report | Out-File $PSScriptRoot\report.html
    
}
$endTime = (Get-Date)
$ElapsedTime = $endTime-$startTime
'Duration: {0:mm} min {0:ss} sec' -f $ElapsedTime

Pause



