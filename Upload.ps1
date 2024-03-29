Add-PSSnapin Microsoft.SharePoint.PowerShell -EA 0

#Variables
$Identity = Read-Host 'Specify the site to upload to'
$docLib = Read-Host 'Specify the name of the document library in the destination site'
$uploadLocation = Read-Host 'Source folder for the file'
$upload = $uploadLocation + "\" + $filename
$filename = Read-Host 'Name of file to be uploaded'
$ExcelApp = Read-Host 'Specify the name of the excel services service application in the farm'
$ExcelFile = Read-Host 'Specify the new trusted file location for excel workbooks'
$ExcelDCL = Read-Host 'New data connection library to be used by excel workbooks'

#Upload Excel file to SharePoint
$spWeb = Get-SPWeb -Identity $Identity
$spFolder = $spWeb.GetFolder($docLib)
$spFileCollection = $spFolder.Files
$file = Get-ChildItem $upload
$spFileCollection.Add($docLib+"/"+$filename,$file.OpenRead(),$true)

#Set the trusted the new workbook and data connection locations
$NewLocation = Get-SPExcelServiceApplication -identity $ExcelApp
$NewLocation | New-SPExcelFileLocation -address $ExcelFile -includechildren -locationType SharePoint -workbooksizemax 50 -externaldataallowed DclAndEmbedded -WarnOnDataRefresh:$false
New-SPExcelDataConnectionLibrary -Address $ExcelDCL -ExcelServiceApplication $ExcelApp