Add-PSSnapin Microsoft.SharePoint.PowerShell -EA 0

#Define Variables
$destination = "C:\ExcelUpdateTemp"
$webUrl = Read-Host 'Specify the site where document library is located'
$listUrl = Read-Host 'Insert full url of document library (e.g http://site/library/Forms/AllItems.aspx)'
$allfiles = $destination + "\*"
$filename = Read-Host 'Which file do you want to download'
$xfilename = "connections.xml"
$xmlpath = $destination + "\" + $xfilename
$oldPath = Read-Host 'Old location of the odc file'
$newPath = Read-Host 'New location of the odc file'

# New Directories
New-Item $destination -type Directory

$web = Get-SPWeb -Identity $webUrl
$list = $web.GetList($listUrl)

function ProcessFolder {
    param($folderUrl)
    $folder = $web.GetFolder($folderUrl)
    foreach ($file in $folder.Files) {
            #Download file        
            $binary = $file.OpenBinary()        
            $stream = New-Object System.IO.FileStream($destination + "/" + $file.Name), Create        
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

#Delete unnecessary files
remove-item $allfiles -exclude $filename, $xfilename

#using simple text replacement
$con = Get-Content $xmlpath
$con | % { $_.Replace($oldPath, $newPath) } | Set-Content $xmlpath
