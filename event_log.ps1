$header = @"
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
    $basename = ($filename.Split("\"))[-1]
    $report_title = "<h1>Report Source: $basename</h1>"
    $bsod = Get-WinEvent -Path $filename | Where-Object {$_.Id -eq '1001'} | ConvertTo-Html -Property TimeCreated,Id,Message -PreContent "<h2>BSOD<h2>"
    $whea = Get-WinEvent -Path $filename | Where-Object {$_.Id -eq '17'-and $_.ProviderName -eq 'Microsoft-Windows-WHEA-Logger'} | ConvertTo-Html -Property TimeCreated,Id,Message -PreContent "<h2>WHEA<h2>"
    $tdr = Get-WinEvent -Path $filename | Where-Object {$_.Id -eq '4101'} | ConvertTo-Html -Property TimeCreated,Id,Message -PreContent "<h2>TDR<h2>"
    $unexpected = Get-WinEvent -Path $filename | Where-Object {$_.Id -eq '6008'} | ConvertTo-Html -Property TimeCreated,Id,Message -PreContent "<h2>Unexpected Shutdown<h2>"

    $tmp = (Get-WinEvent -Path $filename | Where-Object {$_.Id -eq '4101'-or ($_.Id -eq '17'-and $_.ProviderName -eq 'Microsoft-Windows-WHEA-Logger') -or $_.Id -eq '1001' -or $_.Id -eq '6008'} | Sort-Object -Property TimeCreated) | ConvertTo-Html -Property TimeCreated,Id,Message -PreContent "<h2>All Errors</h2>"
    $report = ConvertTo-Html -Body "$report_title $bsod $whea $tdr $unexpected $tmp" -Head $header
    $report | Out-File $PSScriptRoot\report.html
    
}

Pause



