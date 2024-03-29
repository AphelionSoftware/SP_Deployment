Add-PSSnapin Microsoft.SharePoint.PowerShell -EA 0

#Define Variables
$destination = Read-Host 'Insert the destination folder for the file'
$webUrl = Read-Host 'Specify the site where document library is located'
$listName = Read-Host 'Insert name of document library'
$filename = Read-Host 'Which file do you want to download'
$listurl = $webUrl + "/" + $listName + "/Forms/AllItems.aspx"
$temp = $destination + "\Temp"
$tempFile = $temp + "\" + $filename
$allfiles = $temp + "\*"

# New Directories
New-Item $temp -type Directory

$web = Get-SPWeb -Identity $webUrl
$list = $web.GetList($listUrl)

function ProcessFolder {
    param($folderUrl)
    $folder = $web.GetFolder($folderUrl)
    foreach ($file in $folder.Files) {
            #Download file        
            $binary = $file.OpenBinary()        
            $stream = New-Object System.IO.FileStream($temp + "/" + $file.Name), Create        
            $writer = New-Object System.IO.BinaryWriter($stream)        
            $writer.write($binary)        
            $writer.Close()        
            }
}

#Download root files
ProcessFolder($list.RootFolder.Url)

#Download files in folders
foreach ($folder in $list.Folders) 
{
    ProcessFolder($folder.Url)
}

#Copy Item
Copy-Item $tempFile $destination

#Delete unnecessary files
remove-item $allfiles -exclude $filename
remove-item $temp -recurse
