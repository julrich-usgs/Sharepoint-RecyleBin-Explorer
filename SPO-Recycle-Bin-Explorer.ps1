$scriptpath = $MyInvocation.MyCommand.Path

#I'd like to Switch to 32-bit powershell automatically, but I'm not having a good time making that happen
if([IntPtr]::size -eq 8) {
    Write-Host "This Script must be run in the a 32-bit powershell session."
    #Start-Process "C:\windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" " -file $scriptpath"
    
    exit
}


#Check to see if PnP.Powershell module is installed and install it if not.
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "PnP.PowerShell module (and maybe dependencies) needs to be installed...  Installing"
    Install-Module PnP.Powershell -Scope CurrentUser
} 


$dir = Split-Path $scriptpath
Write-Host ""
Write-Host "You need to provide a Sharepoint or OneDrive site URL"
Write-Host "--- Example Teams Site URL: https://<your-org>.sharepoint.com/sites/<your-team-name>"
Write-Host "--- Example OneDrive Site URL: https://<your-org>.sharepoint.com/personal/<youruserid>"
Write-Host ""
$SPOSite = Read-Host "Enter your site URL"
Write-Host ""
Write-Host "In the Grid-View that opens, you can search and add filters at the top.  "
Write-Host "Columns can be sorted by clicking the column header.  "
Write-Host "Select your item(s) and click OK to restore them to their original location"

#Connect to the site entered
Connect-PnPOnline -Url $SPOSite -UseWebLogin

#Get List of things in RecycleBin
$file = "$dir\contents.csv"
Get-PnPRecycleBinItem | export-csv $file -Force

#Strip out the top line in the CSV file returned so column headers are read-in properly
$firstRow = Get-Content $file -First 1
if($firstRow -eq "#TYPE Microsoft.SharePoint.Client.RecycleBinItem")
{
    get-content $file |
    select -Skip 1 |
    set-content "$file-temp"
    move "$file-temp" $file -Force
}

#Query the CSV file and read it into a data table
$ConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$dir;Extended properties='text;HDR=Yes'"
$Conn = new-object System.Data.OleDb.OleDbConnection($connString)
$conn.open()
$name = "contents.csv"
$cmd = new-object System.Data.OleDb.OleDbCommand("Select Title,AuthorEmail,DeletedByEmail,DeletedDate,DirName,Id from [$name]",$Conn)
$da = new-object System.Data.OleDb.OleDbDataAdapter($cmd)
$dt = new-object System.Data.dataTable
[void]$da.fill($dt)

#Display the recycle bin in a GridView and pass the selections back to the script
$items = $dt | Out-GridView -Title "Select item(s) to restore, press OK at bottom when finished" -PassThru


$i=0
foreach($item in $items)
{
    Write-Host "Restoring: " $item.Title
    #For some reason, Restore-PnPRecycleBinItem has to be run twice on the first item 
    if($i -eq 0)
    { 
        Restore-PnPRecycleBinItem -Identity $item.Id -Force -ErrorAction SilentlyContinue

        Restore-PnPRecycleBinItem -Identity $item.Id -Force
    }
    else
    {
        Restore-PnPRecycleBinItem -Identity $item.Id -Force
    }
    $i++
     
}