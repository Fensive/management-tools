Set-StrictMode -Version "2.0"
Clear-Host

#$PathToOneNote = "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE" #OneNote executable
$BasePath = "D:\Fensive Security\Fensive Shared - Documents\Founders\Backups\OneNote" #alternative: $evn:TEMP (for copying and deleting)


echo "Starting OneNote for API access"

Start-Sleep -Seconds 5

[void][reflection.assembly]::LoadWithPartialName("Microsoft.Office.Interop.Onenote")
$OneNote = New-Object Microsoft.Office.Interop.Onenote.ApplicationClass

[Xml]$Xml = $Null
$OneNote.GetHierarchy($Null, [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsNotebooks, [ref] $Xml)

$Date = Get-Date -Format "dd.MM.yyyy HH-mm"
echo ("Starting Backup, date: " + $Date)

ForEach($Notebook in ($Xml.Notebooks.Notebook)) {
    
    if ($Notebook.path -match "fensive")
    {
        echo ("Starting export: " + $Notebook.name)
        $File = $BasePath + "\" + $Date + "\" + $Notebook.name + ".onepkg"
        $OneNote.Publish($Notebook.ID, $File, 1) #1 = .onepkg
        echo "Finished export"
        Start-Sleep -Seconds 3
    }
}

$FolderCount = (Get-ChildItem -Path $BasePath | Measure-Object).count

if ($FolderCount -gt 30)
{
    
    # Delete Old Folder 
    $OldestFolder = $BasePath + "\" + (Get-ChildItem -Path $BasePath | Sort CreationTime | select -First 1)
    $Items = Get-ChildItem -LiteralPath $OldestFolder -Recurse
    foreach ($Item in $Items) {
        Remove-Item -LiteralPath $Item.Fullname
    }
    $Items = Get-Item -LiteralPath $OldestFolder
    $Items.Delete($true)
}