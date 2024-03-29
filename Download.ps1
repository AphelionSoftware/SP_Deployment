Add-PSSnapin Microsoft.SharePoint.PowerShell -EA 0

#Define Variables
$destination = "C:\ExcelUpdateTemp"
$webUrl = Read-Host 'Specify the site where document library is located'
$listName = Read-Host 'Insert name of document library'
$allfiles = $destination + "\*"
$filename = Read-Host 'Which file do you want to download'
$xfilename = "connections.xml"
$oldPath = Read-Host 'Old location of the odc file'
$newPath = Read-Host 'New location of the odc file'
$listurl = $webUrl + "/" + $listName + "/Forms/AllItems.aspx"
$extempPath = $destination + "\Temp"
$efile = $destination + "\" + $filename
$tempfile = $extempPath + "\" + $filename
$xmlpath = $extempPath + "\xl\" + $xfilename


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
remove-item $allfiles -exclude $filename

# Create new directory and copy file
New-Item $extempPath -type Directory
Copy-Item $efile $extempPath

#Change xlsx to zip file
Dir $tempfile | rename-item -newname {  $_.name  -replace ".xlsx",".zip"  }
$fileName = $filename
$newExtension = "zip"
$newzip = [System.IO.Path]::ChangeExtension($fileName,$newExtension)
$newfile = $extempPath + "\" + $newzip

#extract zip file contents
function Extract-Zip 
{ 
    param([string]$zipfilename, [string] $destination) 
    if(test-path($zipfilename)) 
    { 
        $shellApplication = new-object -com shell.application 
        $zipPackage = $shellApplication.NameSpace($zipfilename) 
        $destinationFolder = $shellApplication.NameSpace($destination) 
        $destinationFolder.CopyHere($zipPackage.Items()) 
    }
    else 
    {   
        Write-Host $zipfilename "not found"
    }
}
Extract-Zip $newfile $extempPath

#using simple text replacement
$con = Get-Content $xmlpath
$con | % { $_.Replace($oldPath, $newPath) } | Set-Content $xmlpath
