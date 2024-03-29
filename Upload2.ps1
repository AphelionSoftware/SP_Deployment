Add-PSSnapin Microsoft.SharePoint.PowerShell -EA 0

#Variables
$Identity = Read-Host 'Specify the site to upload to'
$docLib = Read-Host 'Specify the name of the document library in the destination site'
$Location = Read-Host 'Source folder for the file'
$filename = Read-Host 'Name of file to be uploaded'
$upload = $Location + "\" + $filename

#Upload Excel file to SharePoint
$spWeb = Get-SPWeb -Identity $Identity
$spFolder = $spWeb.GetFolder($docLib)
$spFileCollection = $spFolder.Files
$file = Get-ChildItem $upload
$spFileCollection.Add($docLib + "/" + $filename, $file.OpenRead(), $true)
